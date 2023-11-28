#include "excelr8/compdoc.hpp"
#include "excelr8/data.hpp"
#include "excelr8/util.hpp"
#include <cassert>
#include <format>
#include <sstream>
#include <string>
#include <tuple>
#include <utility>
#include <vector>

namespace excelr8::compdoc {

DirNode::DirNode(int did, const data_t dent, int debug, std::ostream& logfile)
    : logfile(logfile)
    , did(did)
{
    pytype_H cbufsize;
    std::tie(cbufsize, etype, color, left_did, right_did, root_did)
        = dent.slice(64, 80).unpack<pytype_H, pytype_B, pytype_B, pytype_i, pytype_i, pytype_i>();
    std::tie(first_sid, tot_size) = dent.slice(116, 124).unpack<pytype_i, pytype_i>();

    if (cbufsize != 0) {
        name = util::unicode(dent.slice(0, cbufsize - 2), "utf_16_le"); // omit the trailing U+0000
    }
    std::tie(tsinfo[0], tsinfo[1], tsinfo[2], tsinfo[3]) = dent.slice(100, 116).unpack<pytype_I, pytype_I, pytype_I, pytype_I>();

    if (debug) {
        dump(debug);
    }
}

void DirNode::dump(int debug)
{
    logfile << std::format("DID=%d name=%s etype=%d DIDs(left=%d right=%d root=%d parent=%d kids size=%d) first_SID=%d tot_size=%d\n",
        did, name, etype, left_did, right_did, root_did, parent, children.size(), first_sid, tot_size);

    if (debug == 2) {
        // cre_lo, cre_hi, mod_lo, mod_hi = tsinfo
        logfile << std::format("timestamp info %u %u %u %u\n", tsinfo[0], tsinfo[1], tsinfo[2], tsinfo[3]);
    }
}

void _build_family_tree(const std::vector<DirNode*>& dirlist, int parent_did, int child_did)
{
    if (child_did < 0)
        return;

    _build_family_tree(dirlist, parent_did, dirlist[child_did]->left_did);
    dirlist[parent_did]->children.push_back(child_did);
    dirlist[child_did]->parent = parent_did;

    _build_family_tree(dirlist, parent_did, dirlist[child_did]->right_did);
    if (dirlist[child_did]->etype == 1) { // storage
        _build_family_tree(dirlist, child_did, dirlist[child_did]->root_did);
    }
}

CompDoc::CompDoc(const data_t& mem, std::ostream& logfile, int debug, bool ignore_workbook_corruption)
    : logfile(logfile)
    , ignore_workbook_corruption(ignore_workbook_corruption)
    , debug(debug)
    , mem(mem)
{
    if (mem.slice(0, 8) != SIGNATURE) {
        throw CompDocError("Not an OLE2 compound document");
    }
    auto le_marker = mem.slice(28, 30);
    if (le_marker != data_t("\xFE\xFF")) {
        throw CompDocError("Expected 'little-endian' marker, found " + le_marker.to_string());
    }
    auto [revision, version] = mem.slice(24, 28).unpack<pytype_H, pytype_H>();

    if (debug) {
        logfile << std::format("\nCompDoc format: version=0x%04x revision=0x%04x\n", version, revision);
    }

    auto [ssz, sssz] = mem.slice(30, 34).unpack<pytype_H, pytype_H>();
    if (ssz > 20) { // allows for 2**20 bytes i.e. 1MB
        logfile << std::format("WARNING: sector size (2**%d) is preposterous; assuming 512 and continuing ...\n", ssz);
        ssz = 9;
    }
    if (sssz > ssz) {
        logfile << std::format("WARNING: short stream sector size (2**%d) is preposterous; assuming 64 and continuing ...\n", sssz);
        sssz = 6;
    }

    sec_size = 1 << ssz;
    short_sec_size = 1 << sssz;
    if (sec_size != 512 or short_sec_size != 64) {
        logfile << std::format("@@@@ sec_size=%d short_sec_size=%d\n", sec_size, short_sec_size);
    }

    auto _info = mem.slice(44, 76).unpack_vec<pytype_i>(8);
    int SAT_tot_secs = _info[0];
    dir_first_sec_sid = _info[1];
    // _info[2] is unused
    min_size_std_stream = _info[3];
    int SSAT_first_sec_sid = _info[4];
    int SSAT_tot_secs = _info[5];
    int MSATX_first_sec_sid = _info[6];
    int MSATX_tot_secs = _info[7];

    size_t mem_data_len = mem.size() - 512;
    size_t mem_data_secs = mem_data_len / sec_size;
    int left_over = mem_data_len % sec_size;
    if (left_over) {
        // throw CompDocError("Not a whole number of sectors");
        mem_data_secs += 1;
        logfile << std::format("WARNING *** file size (%d) not 512 + multiple of sector size (%d)\n", mem.size(), sec_size);
    }

    this->mem_data_secs = mem_data_secs; // use for checking later
    this->mem_data_len = mem_data_len;
    seen = std::vector<unsigned char>(mem_data_secs, 0);

    if (debug) {
        logfile << std::format("sec sizes %d %d %d %d\n", ssz, sssz, sec_size, short_sec_size);
        logfile << std::format("mem data: %d bytes == %d sectors\n", mem_data_len, mem_data_secs);
        logfile << std::format("SAT_tot_secs=%d, dir_first_sec_sid=%d, min_size_std_stream=%d\n", SAT_tot_secs, dir_first_sec_sid, min_size_std_stream);
        logfile << std::format("SSAT_first_sec_sid=%d, SSAT_tot_secs=%d\n", SSAT_first_sec_sid, SSAT_tot_secs);
        logfile << std::format("MSATX_first_sec_sid=%d, MSATX_tot_secs=%d\n", MSATX_first_sec_sid, MSATX_tot_secs);
    }

    int nent = sec_size / 4; // number of SID entries in a sector
    int trunc_warned = 0;

    //
    // === build the MSAT ===
    //
    std::vector<int> MSAT = mem.slice(76, 512).unpack_vec<int>(109);
    int SAT_sectors_reqd = (mem_data_secs + nent - 1) / nent;
    int expected_MSATX_sectors = std::max(0, (SAT_sectors_reqd - 109 + nent - 2) / (nent - 1));
    int actual_MSATX_sectors = 0;

    if (MSATX_tot_secs == 0 and (MSATX_first_sec_sid == 0 or MSATX_first_sec_sid == EOCSID or MSATX_first_sec_sid == FREESID)) {
        // Strictly, if there is no MSAT extension, then MSATX_first_sec_sid
        // should be set to EOCSID ... FREESID and 0 have been met in the wild.
        // pass # Presuming no extension
    } else {
        int sid = MSATX_first_sec_sid;
        while (sid != EOCSID and sid != FREESID and sid != MSATSID) {
            // Above should be only EOCSID according to MS & OOo docs
            // but Excel doesn't complain about FREESID. Zero is a valid
            // sector number, not a sentinel.
            if (debug > 1) {
                logfile << std::format("MSATX: sid=%d (0x%08X)\n", sid, sid);
            }
            if (sid >= mem_data_secs) {
                std::string msg = std::format("MSAT extension: accessing sector %d but only %d in file", sid, mem_data_secs);
                if (debug > 1) {
                    logfile << msg << std::endl;
                }
                throw CompDocError(msg);
            } else if (sid < 0) {
                throw CompDocError("MSAT extension: invalid sector id: " + std::to_string(sid));
            }
            if (seen[sid]) {
                throw CompDocError(std::format("MSAT corruption: seen[%d] == %d", sid, seen[sid]));
            }
            seen[sid] = 1;
            actual_MSATX_sectors += 1;
            if (debug and actual_MSATX_sectors > expected_MSATX_sectors) {
                logfile << "[1]===>>> " << mem_data_secs << " " << nent << " " << SAT_sectors_reqd << " " << expected_MSATX_sectors << " " << actual_MSATX_sectors << std::endl;
            }
            int offset = 512 + sec_size * sid;
            auto extension = mem.slice(offset, offset + sec_size).unpack_vec<int>(nent);
            MSAT.insert(MSAT.end(), extension.begin(), extension.end());
            sid = MSAT.back();
            MSAT.pop_back(); // last sector id is sid of next sector in the chain
        }
    }

    if (debug and actual_MSATX_sectors != expected_MSATX_sectors) {
        logfile << "[2]===>>> " << mem_data_secs << " " << nent << " " << SAT_sectors_reqd << " " << expected_MSATX_sectors << " " << actual_MSATX_sectors << std::endl;
    }
    if (debug) {
        logfile << "MSAT: len = " << MSAT.size() << std::endl;
        dump_list(MSAT, 10, logfile);
    }

    //
    // === build the SAT ===
    //
    int actual_SAT_sectors = 0;
    int dump_again = 0;
    for (size_t msidx = 0; msidx < MSAT.size(); msidx++) {
        int msid = MSAT[msidx];
        if (msid == FREESID or msid == EOCSID) {
            // Specification: the MSAT array may be padded with trailing FREESID entries.
            // Toleration: a FREESID or EOCSID entry anywhere in the MSAT array will be ignored.
            continue;
        }
        if (msid >= mem_data_secs) {
            if (!trunc_warned) {
                logfile << "WARNING *** File is truncated, or OLE2 MSAT is corrupt!!" << std::endl;
                logfile << std::format("INFO: Trying to access sector %d but only %d available\n", msid, mem_data_secs);
                trunc_warned = 1;
            }
            MSAT[msidx] = EVILSID;
            dump_again = 1;
            continue;
        } else if (msid < -2) {
            throw CompDocError("MSAT: invalid sector id: " + std::to_string(msid));
        }
        if (seen[msid]) {
            throw CompDocError(std::format("MSAT extension corruption: seen[%d] == %d", msid, seen[msid]));
        }
        seen[msid] = 2;
        actual_SAT_sectors += 1;
        if (debug and actual_SAT_sectors > SAT_sectors_reqd) {
            logfile << std::format("[3]===>>> %d %d %d %d %d %d %d\n", mem_data_secs, nent, SAT_sectors_reqd, expected_MSATX_sectors, actual_MSATX_sectors, actual_SAT_sectors, msid);
        }
        int offset = 512 + sec_size * msid;
        std::vector<int> extension = mem.slice(offset, offset + sec_size).unpack_vec<int>(nent);
        SAT.insert(SAT.end(), extension.begin(), extension.end());
    }

    if (debug) {
        logfile << "SAT: len = " << SAT.size() << std::endl;
        dump_list(SAT, 10, logfile);
        // print >> logfile, "SAT ",
        // for i, s in enumerate(self.SAT):
        //     print >> logfile, "entry: %4d offset: %6d, next entry: %4d" % (i, 512 + sec_size * i, s)
        //     print >> logfile, "%d:%d " % (i, s),
        logfile << std::endl;
    }
    if (debug and dump_again) {
        logfile << "MSAT: len = " << MSAT.size() << std::endl;
        dump_list(MSAT, 10, logfile);
        for (size_t satx = mem_data_secs; satx < SAT.size(); satx++) {
            SAT[satx] = EVILSID;
        }
        logfile << "SAT: len = " << SAT.size() << std::endl;
        dump_list(SAT, 10, logfile);
    }

    //
    // === build the directory ===
    //
    data_t dbytes = _get_stream(mem, 512, SAT, sec_size, dir_first_sec_sid, -1, "directory", 3);
    std::vector<DirNode*> dirlist;
    int did = -1;
    for (size_t pos = 0; pos < dbytes.size(); pos += 128) {
        did += 1;
        dirlist.push_back(new DirNode(did, dbytes.slice(pos, pos + 128), 0, logfile));
    }
    this->dirlist = dirlist;
    _build_family_tree(dirlist, 0, dirlist[0]->root_did); // and stand well back ...
    if (debug) {
        for (const auto& d : dirlist) {
            d->dump(debug);
        }
    }

    //
    // === get the SSCS ===
    //
    auto sscs_dir = this->dirlist[0];
    assert(sscs_dir->etype == 5); // root entry
    if (sscs_dir->first_sid < 0 or sscs_dir->tot_size == 0) {
        // Problem reported by Frank Hoffsuemmer: some software was
        // writing -1 instead of -2 (EOCSID) for the first_SID
        // when the SCCS was empty. Not having EOCSID caused assertion
        // failure in _get_stream.
        // Solution: avoid calling _get_stream in any case when the
        // SCSS appears to be empty.
        SSCS = new data_t();
    } else {
        SSCS = &_get_stream(mem, 512, SAT, sec_size, sscs_dir->first_sid, sscs_dir->tot_size, "SSCS", 4);
    }
    // if DEBUG: print >> logfile, "SSCS", repr(self.SSCS)

    //
    // === build the SSAT ===
    //
    if (SSAT_tot_secs > 0 and sscs_dir->tot_size == 0) {
        logfile << "WARNING *** OLE2 inconsistency: SSCS size is 0 but SSAT size is non-zero" << std::endl;
    }
    if (sscs_dir->tot_size > 0) {
        int sid = SSAT_first_sec_sid;
        int nsecs = SSAT_tot_secs;
        while (sid >= 0 and nsecs > 0) {
            if (seen[sid]) {
                throw CompDocError(std::format("SSAT corruption: seen[%d] == %d\n", sid, seen[sid]));
            }
            seen[sid] = 5;
            nsecs -= 1;
            int start_pos = 512 + sid * sec_size;
            auto news = mem.slice(start_pos, start_pos + sec_size).unpack_vec<int>(nent);
            SSAT.insert(SSAT.end(), news.begin(), news.end());
            sid = SAT[sid];
        }
        if (debug) {
            logfile << std::format("SSAT last sid %d; remaining sectors %d\n", sid, nsecs);
        }
        assert(nsecs == 0 and sid == EOCSID);
    }
    if (debug) {
        logfile << "SSAT" << std::endl;
        dump_list(SSAT, 10, logfile);

        logfile << "seen" << std::endl;
        dump_list(seen, 20, logfile);
    }
}

data_t& CompDoc::_get_stream(const data_t& mem, int base, std::vector<int>& sat, int sec_size, int start_sid, int size, std::string name, int seen_id)
{
    // print >> self.logfile, "_get_stream", base, sec_size, start_sid, size
    data_t* sectors = new data_t();
    int s = start_sid;
    if (size == -1) {
        // nothing to check agains
        while (s >= 0) {
            if (seen_id != -1) {
                if (seen[s]) {
                    throw CompDocError(std::format("%s corruption: seen[%d] == %d", name, s, seen[s]));
                }
                seen[s] = seen_id;
            }
            int start_pos = base + s * sec_size;
            sectors->append(mem.slice(start_pos, start_pos + sec_size));
            if (s < sat.size()) {
                s = sat[s];
            } else {
                throw CompDocError(std::format("OLE2 stream %s: sector allocation table invalid entry (%d)", name, s));
            }
        }
        assert(s == EOCSID);
    } else {
        int todo = size;
        while (s >= 0) {
            if (seen_id != -1) {
                if (seen[s]) {
                    throw CompDocError(std::format("%s corruption: seen[%d] == %d", name, s, seen[s]));
                }
                seen[s] = seen_id;
            }
            int start_pos = base + s * sec_size;
            int grab = sec_size;
            if (grab > todo) {
                grab = todo;
            }
            todo -= grab;
            sectors->append(mem.slice(start_pos, start_pos + grab));
            if (s < sat.size()) {
                s = sat[s];
            } else {
                throw CompDocError(std::format("OLE2 stream %s: sector allocation table invalid entry (%d)", name, s));
            }
        }
        assert(s == EOCSID);
        if (todo != 0) {
            logfile << std::format("WARNING *** OLE2 stream %s: expected size %d, actual size %d\n",
                name, size, size - todo);
        }
    }

    return *sectors;
}

std::tuple<const data_t*, int, int> CompDoc::_locate_stream(const data_t& mem, int base, std::vector<int>& sat,
    int sec_size, int start_sid, int expected_stream_size, const std::string& qname, int seen_id)
{
    // print >> self.logfile, "_locate_stream", base, sec_size, start_sid, expected_stream_size
    int s = start_sid;
    if (s < 0) {
        throw CompDocError(std::format("_locate_stream: start_sid (%d) is -ve", start_sid));
    }
    int p = -99; // dummy previous SID
    int start_pos = -9999;
    int end_pos = -8888;
    std::vector<std::pair<int, int>> slices;
    size_t tot_found = 0;
    size_t found_limit = (expected_stream_size + sec_size - 1) / sec_size;

    while (s >= 0) {
        if (seen[s]) {
            if (!ignore_workbook_corruption) {
                logfile << std::format("_locate_stream(%s): seen\n", qname);
                dump_list(seen, 20, logfile);
                throw CompDocError(std::format("%s corruption: seen[%d] == %d", qname, s, seen[s]));
            }
        }
        seen[s] = seen_id;
        tot_found += 1;
        if (tot_found > found_limit) {
            // Note: expected size rounded up higher sector
            throw CompDocError(std::format("%s: size exceeds expected %d bytes; corrupt?",
                qname, found_limit * sec_size));
        }
        if (s == p + 1) {
            // contiguous sectors
            end_pos += sec_size;
        } else {
            // start new slice
            if (p >= 0) {
                // not first time
                slices.push_back({ start_pos, end_pos });
            }
            start_pos = base + s * sec_size;
            end_pos = start_pos + sec_size;
        }
        p = s;
        s = sat[s];
    }
    assert(s == EOCSID);
    assert(tot_found == found_limit);
    // print >> self.logfile, "_locate_stream(%s): seen" % qname; dump_list(self.seen, 20, self.logfile)
    if (slices.empty()) {
        // The stream is contiguous ... just what we like!
        return { &mem, start_pos, expected_stream_size };
    }
    slices.push_back({ start_pos, end_pos });
    // print >> self.logfile, "+++>>> %d fragments" % len(slices)

    data_t* data = new data_t();
    for (const auto& [start_pos, end_pos] : slices) {
        data->append(mem.slice(start_pos, end_pos));
    }
    return { data, 0, expected_stream_size };
}

bool ichar_equals(char a, char b)
{
    return std::tolower(static_cast<unsigned char>(a)) == std::tolower(static_cast<unsigned char>(b));
}

// Case-insensitive string equality check
bool iequals(std::string_view lhs, std::string_view rhs)
{
    return std::ranges::equal(lhs, rhs, ichar_equals);
}

std::vector<std::string> split(const std::string& input, char delimiter)
{
    std::vector<std::string> result;
    std::istringstream ss(input);
    std::string token;

    while (std::getline(ss, token, delimiter)) {
        result.push_back(token);
    }

    return result;
}

DirNode* CompDoc::_dir_search(const std::vector<std::string>& path, int storage_did)
{
    // Return matching DirNode instance, or nullptr
    std::string head = path[0];
    std::vector<std::string> tail = { path.begin() + 1, path.end() };
    for (const auto& child : dirlist[storage_did]->children) {
        if (iequals(dirlist[child]->name, head)) {
            const auto& et = dirlist[child]->etype;
            if (et == 2) {
                return dirlist[child];
            }
            if (et == 1) {
                if (tail.empty()) {
                    throw CompDocError("Requested component is a 'storage'");
                }
                return _dir_search(tail, child);
            }
            dirlist[child]->dump(1);
            throw CompDocError("Requested stream is not a 'user stream'");
        }
    }

    return nullptr;
}

/**
    Interrogate the compound document's directory; return the stream as a
    string if found, otherwise return ``None``.

    :param qname:
        Name of the desired stream e.g. ``'Workbook'``.
        Should be in Unicode or convertible thereto.
*/
data_t* CompDoc::get_named_stream(const std::string& qname)
{
    const auto d = _dir_search(split(qname, '/'));
    if (d == nullptr) {
        return nullptr;
    }
    if (d->tot_size >= min_size_std_stream) {
        return &_get_stream(mem, 512, SAT, sec_size, d->first_sid, d->tot_size, qname, d->did + 6);
    } else {
        return &_get_stream(*SSCS, 0, SSAT, short_sec_size, d->first_sid, d->tot_size, qname + " (from SSCS)");
    }
}

/**
    Interrogate the compound document's directory.

    If the named stream is not found, ``(None, 0, 0)`` will be returned.

    If the named stream is found and is contiguous within the original
    byte sequence (``mem``) used when the document was opened,
    then ``(mem, offset_to_start_of_stream, length_of_stream)`` is returned.

    Otherwise a new string is built from the fragments and
    ``(new_string, 0, length_of_stream)`` is returned.

    :param qname:
        Name of the desired stream e.g. ``'Workbook'``.
        Should be in Unicode or convertible thereto.
*/
std::tuple<const data_t*, int, int> CompDoc::locate_named_stream(const std::string& qname)
{
    const auto d = _dir_search(split(qname, '/'));
    if (d == nullptr) {
        return { nullptr, 0, 0 };
    }
    if (d->tot_size > mem_data_len) {
        throw CompDocError(std::format("%s stream length (%d bytes) > file data size (%d bytes)", qname, d->tot_size, mem_data_len));
    }
    if (d->tot_size >= min_size_std_stream) {
        auto result = _locate_stream(mem, 512, SAT, sec_size, d->first_sid, d->tot_size, qname, d->did + 6);
        if (debug) {
            logfile << "\nseen\n";
            dump_list(seen, 20, logfile);
        }
        return result;
    } else {
        return {
            &_get_stream(
                *SSCS, 0, SSAT, short_sec_size, d->first_sid,
                d->tot_size, qname + " (from SSCS)"),
            0,
            d->tot_size
        };
    }
}

template <typename T>
void dump_list(std::vector<T>& list, int stride, std::ostream& f)
{
    auto _dump_line = [&](int dpos, int equal = 0) {
        f << std::format("%5d%s ", dpos, " ="[equal]);
        for (int i = dpos; i < dpos + stride; i++) {
            auto& value = list[i];
            f << std::to_string(value) << " ";
        }
        f << std::endl;
    };

    int pos = 0, oldpos = -1;
    for (pos = 0; pos < list.size(); pos += stride) {
        if (oldpos == -1) {
            _dump_line(pos);
            oldpos = pos;
        } else {
            std::vector<T> v1 = { list.begin() + pos, list.begin() + pos + stride + 1 };
            std::vector<T> v2 = { list.begin() + oldpos, list.begin() + oldpos + stride + 1 };
            if (v1 != v2) {
                if (pos - oldpos > stride) {
                    _dump_line(pos - stride, 1);
                }
                _dump_line(pos);
                oldpos = pos;
            }
        }
    }
    if (oldpos != -1 and pos != 0 and pos != oldpos) {
        _dump_line(pos, 1);
    }
}
}