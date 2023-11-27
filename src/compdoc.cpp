#include "excelr8/compdoc.hpp"
#include "excelr8/data.hpp"
#include "excelr8/util.hpp"
#include <format>
#include <tuple>
#include <utility>
#include <vector>

namespace excelr8::compdoc {

DirNode::DirNode(int did, const data_t dent, int debug, std::ostream& logfile)
    : did(did)
    , logfile(logfile)
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

    int SAT_tot_secs, SSAT_first_sec_sid, SSAT_tot_secs, MSATX_first_sec_sid, MSATX_tot_secs, _unused;
    std::tie(SAT_tot_secs, dir_first_sec_sid, _unused, min_size_std_stream,
        SSAT_first_sec_sid, SSAT_tot_secs,
        MSATX_first_sec_sid, MSATX_tot_secs)
        = mem.slice(44, 76).unpack<pytype_i, pytype_i, pytype_i, pytype_i, pytype_i, pytype_i, pytype_i, pytype_i>();

    size_t mem_data_len = mem.size() - 512;
    int mem_data_secs = mem_data_len / sec_size;
    int left_over = mem_data_len % sec_size;
    if (left_over) {
        // throw CompDocError("Not a whole number of sectors");
        mem_data_secs += 1;
        logfile << std::format("WARNING *** file size (%d) not 512 + multiple of sector size (%d)\n", mem.size(), sec_size);
    }

    this->mem_data_secs = mem_data_secs; // use for checking later
    this->mem_data_len = mem_data_len;
    seen = std::vector<std::byte>(mem_data_secs, std::byte(0));

    if (debug) {
        logfile << std::format("sec sizes %d %d %d %d\n", ssz, sssz, sec_size, short_sec_size);
        logfile << std::format("mem data: %d bytes == %d sectors\n", mem_data_len, mem_data_secs);
        logfile << std::format("SAT_tot_secs=%d, dir_first_sec_sid=%d, min_size_std_stream=%d\n", SAT_tot_secs, dir_first_sec_sid, min_size_std_stream);
        logfile << std::format("SSAT_first_sec_sid=%d, SSAT_tot_secs=%d\n", SSAT_first_sec_sid, SSAT_tot_secs);
        logfile << std::format("MSATX_first_sec_sid=%d, MSATX_tot_secs=%d\n", MSATX_first_sec_sid, MSATX_tot_secs);
    }

    int nent = sec_size / 4; // number of SID entries in a sector
}
}