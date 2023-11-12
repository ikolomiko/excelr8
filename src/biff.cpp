#include "excelr8/biff.hpp"
#include <algorithm>
#include <cstdio>
#include <format>
#include <iostream>
#include <map>
#include <string>
#include <tuple>
#include <vector>

using namespace excelr8::util;

namespace excelr8::biff {

int calc_nchars(const data_t& data, int pos, int lenlen)
{
    // nchars:int = unpack('<' + 'BH'[lenlen-1], data[pos:pos+lenlen])[0]
    if (lenlen == 1) {
        // 'B' == unsigned char
        return std::get<0>(unpack<pytype_B>(vslice(data, pos, pos + lenlen)));
    } else { // lenlen = 2
        // 'H' == unsigned short
        return std::get<0>(unpack<pytype_H>(vslice(data, pos, pos + lenlen)));
    }
}

bool is_cell_opcode(int c)
{
    return _cell_opcode_set.count(c) == 1;
}

void upkbits(void* tgt_obj, int src, std::vector<Manifest>& manifests)
{
    /* TODO
    for n, mask, attr in manifest:
        local_setattr(tgt_obj, attr, (src & mask) >> n)
    */
    for (const auto& [n, mask, attr] : manifests) {
        // tgt_obj->attr = (src & mask) >> n;
    }
}

void upkbitsL(void* tgt_obj, int src, std::vector<Manifest>& manifests)
{
    upkbits(tgt_obj, src, manifests); // but make sure the result is integer
}

std::string unpack_string(const data_t& data, int pos, const std::string& encoding, int lenlen = 1)
{
    int nchars = calc_nchars(data, pos, lenlen);
    pos += lenlen;
    auto slice = vslice(data, pos, pos + nchars);
    return unicode(slice, encoding);
}

std::pair<std::string, int> unpack_string_update_pos(const data_t& data, int pos, const std::string& encoding, int lenlen = 1, int known_len = -1)
{
    int nchars;
    if (known_len != -1) {
        nchars = known_len;
    } else {
        nchars = calc_nchars(data, pos, lenlen);
        pos += lenlen;
    }

    int newpos = pos + nchars;
    auto slice = vslice(data, pos, newpos);
    return { unicode(slice, encoding), newpos };
}

std::string unpack_unicode(const data_t& data, int pos, int lenlen = 2)
{
    int nchars = calc_nchars(data, pos, lenlen);

    if (nchars == 0) {
        // Ambiguous whether 0-length string should have an "options" byte.
        // Avoid crash if missing.
        return "";
    }

    pos += lenlen;
    auto options = (unsigned char)data[pos];
    pos += 1;

    std::string strg;
    // phonetic = options & 0x04
    // richtext = options & 0x08
    if (options & 0x08) {
        // rt = unpack('<H', data[pos:pos+2])[0] # unused
        pos += 2;
    }
    if (options & 0x04) {
        // sz = unpack('<i', data[pos:pos+4])[0] # unused
        pos += 4;
    }
    if (options & 0x01) {
        // Uncompressed UTF-16-LE
        auto rawstrg = vslice(data, pos, pos + 2 * nchars);
        // if DEBUG: print "nchars=%d pos=%d rawstrg=%r" % (nchars, pos, rawstrg)
        strg = unicode(rawstrg, "utf_16_le");
        // pos += 2*nchars
    } else {
        // Note: this is COMPRESSED (not ASCII!) encoding!!!
        // Merely returning the raw bytes would work OK 99.99% of the time
        // if the local codepage was cp1252 -- however this would rapidly go pear-shaped
        // for other codepages so we grit our Anglocentric teeth and return Unicode :-)

        auto slice = vslice(data, pos, pos + nchars);
        strg = unicode(slice, "latin_1");
        // pos += nchars
    }
    /*
     if richtext:
         pos += 4 * rt
     if phonetic:
         pos += sz
     return (strg, pos)
    */
    return strg;
}

std::pair<std::string, int> unpack_unicode_update_pos(const data_t& data, int pos, int lenlen = 2, int known_len = -1)
{
    int nchars;
    if (known_len != -1) {
        nchars = known_len;
    } else {
        nchars = calc_nchars(data, pos, lenlen);
        pos += lenlen;
    }

    if (nchars == 0 and data.size() <= pos + 1) {
        return { "", pos };
    }

    auto options = (unsigned char)data[pos];
    pos += 1;
    auto phonetic = options & 0x04;
    auto richtext = options & 0x08;

    int rt, sz;
    std::string strg;
    if (richtext != 0) {
        rt = std::get<0>(unpack<pytype_H>(vslice(data, pos, pos + 2)));
        pos += 2;
    }
    if (phonetic) {
        sz = std::get<0>(unpack<pytype_i>(vslice(data, pos, pos + 4)));
        pos += 4;
    }
    if (options & 0x01) {
        // Uncompressed UTF-16-LE
        strg = unicode(vslice(data, pos, pos + 2 * nchars), "utf_16_le");
        pos += 2 * nchars;
    } else {
        // Note: this is COMPRESSED (not ASCII) encoding!!!
        strg = unicode(vslice(data, pos, pos + nchars), "latin_1");
        pos += nchars;
    }

    if (richtext) {
        pos += 4 * rt;
    }
    if (phonetic) {
        pos += sz;
    }

    return { strg, pos };
}

int unpack_cell_range_address_list_update_pos(std::vector<pytype_H>& output_list, const data_t data, int pos, int addr_size = 6)
{
    // output_list is updated in situ
    if (addr_size != 6 and addr_size != 8) {
        // Used to assert size == 6 if not BIFF8, but pyWLWriter writes
        // BIFF8-only MERGEDCELLS records in a BIFF5 file!
        std::cerr << "assertion failed at biff.cpp:unpack_cell_range_address_list_update_pos" << std::endl;
        return -1;
    }

    uint16_t n = std::get<0>(unpack<pytype_H>(vslice(data, pos, pos + 2)));
    pos += 2;

    for (int i = 0; i < n; i++) {
        auto slice = vslice(data, pos, pos + addr_size);
        int ra, rb, ca, cb;
        if (addr_size == 6) {
            // <HHBB
            std::tie(ra, rb, ca, cb) = unpack<pytype_H, pytype_H, pytype_B, pytype_B>(slice);
        } else { // addr_size = 8
            // <HHHH
            std::tie(ra, rb, ca, cb) = unpack<pytype_H, pytype_H, pytype_H, pytype_H>(slice);
        }
        output_list.push_back(ra);
        output_list.push_back(rb);
        output_list.push_back(ca);
        output_list.push_back(cb);
        pos += addr_size;
    }

    return pos;
}

void hex_char_dump(const std::string& strg, int ofs, int dlen, int base = 0, std::ostream& fout = std::cout, bool unnumbered = false)
{
    int endpos = std::min<int>(ofs + dlen, strg.length());
    int pos = ofs;
    bool numbered = not unnumbered;
    std::string num_prefix = "";
    while (pos < endpos) {
        int endsub = std::min<int>(pos + 16, endpos);
        int lensub = endsub - pos;
        std::string substrg = strg.substr(pos, lensub);
        if (lensub <= 0 or lensub != substrg.length()) {
            fout << std::format(
                "'??? hex_char_dump: ofs=%d dlen=%d base=%d -> endpos=%d pos=%d endsub=%d substrg=%r\n",
                ofs, dlen, base, endpos, pos, endsub, substrg);
            break;
        }

        std::string hexd;
        std::string chard;
        for (char c : substrg) {
            hexd += std::format("%02x ", c);

            if (c == '\0') {
                c = '~';
            } else if (not(' ' <= c) and not(c <= '~')) {
                c = '?';
            }
            chard += c;
        }
        if (numbered) {
            num_prefix = std::format("%5d: ", base + pos - ofs);
        }

        fout << std::format("%s     %-48s %s\n", num_prefix, hexd, chard);
        pos = endsub;
    }
}

void biff_dump(const data_t& mem, int stream_offset, int stream_len, int base = 0, std::ostream& fout = std::cout, bool unnumbered = false)
{
    int pos = stream_offset;
    int stream_end = stream_offset + stream_len;
    int adj = base - stream_offset;
    int dummies = 0;
    bool numbered = not unnumbered;
    int savpos, length, rc;
    std::string num_prefix;

    while (stream_end - pos >= 4) {
        std::tie(rc, length) = unpack<pytype_H, pytype_H>(vslice(mem, pos, pos + 4));
        if (rc == 0 and length == 0) {
            bool allNull = true;
            for (auto it = mem.begin() + pos; it != mem.end(); it++) {
                if ((unsigned char)*it != '\0') {
                    allNull = false;
                    break;
                }
            }

            if (allNull) {
                dummies = stream_end - pos;
                savpos = pos;
                pos = stream_end;
                break;
            }
            if (dummies != 0) {
                dummies += 4;
            } else {
                savpos = pos;
                dummies = 4;
            }
            pos += 4;
        } else {
            if (dummies != 0) {
                if (numbered) {
                    num_prefix = std::format("%5d: ", adj + savpos);
                }
                fout << std::format("%s---- %d zero bytes skipped ----\n", num_prefix, dummies);
                dummies = 0;
            }
            std::string recname = biff_rec_name_dict.at(rc);
            if (recname.empty()) {
                recname = "<UNKNOWN>";
            }

            if (numbered) {
                num_prefix = std::format("%5d: ", adj + pos);
            }
            fout << std::format("%s%04x %s len = %04x (%d)\n", num_prefix, rc, recname, length, length);
            pos += 4;
            std::string strg(reinterpret_cast<const char*>(mem.data()), mem.size());
            hex_char_dump(strg, pos, length, adj + pos, fout, unnumbered);
            pos += length;
        }
    }
    if (dummies != 0) {
        if (numbered) {
            num_prefix = std::format("%5d: ", adj + savpos);
        }
        fout << std::format("%s---- %d zero bytes skipped ----\n", num_prefix, dummies);
    }
    if (pos < stream_end) {
        if (numbered) {
            num_prefix = std::format("%5d: ", adj + pos);
        }
        fout << std::format("%s---- Misc bytes at end ----\n", num_prefix);
        std::string strg(reinterpret_cast<const char*>(mem.data()), mem.size());
        hex_char_dump(strg, pos, stream_end - pos, adj + pos, fout, unnumbered);
    } else if (pos > stream_end) {
        fout << std::format("Last dumped record has length (%d) that is too large\n", length);
    }
}

void biff_count_records(const data_t& mem, int stream_offset, int stream_len, std::ostream& fout)
{
    int pos = stream_offset;
    int stream_end = stream_offset + stream_len;
    std::map<std::string, int> tally;

    while (stream_end - pos >= 4) {
        std::string recname;
        auto [rc, length] = unpack<pytype_H, pytype_H>(vslice(mem, pos, pos + 4));
        if (rc == 0 and length == 0) {
            bool allNull = true;
            for (auto it = mem.begin() + pos; it != mem.end(); it++) {
                if ((unsigned char)*it != '\0') {
                    allNull = false;
                    break;
                }
            }
            if (allNull) {
                break;
            }

            recname = "<Dummy (zero)>";
        } else {
            recname = biff_rec_name_dict.at(rc);
            if (recname.empty()) {
                recname = std::format("Unknown_0x%04X", rc);
            }
        }
        if (tally.contains(recname)) {
            tally[recname] += 1;
        } else {
            tally.insert({ recname, 1 });
        }
        pos += length + 4;
    }
    for (const auto& [recname, count] : tally) {
        fout << std::format("%8d %s\n", count, recname);
    }
}

}
