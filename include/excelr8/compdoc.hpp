#pragma once

/*
Implements the minimal functionality required
to extract a "Workbook" or "Book" stream (as one big string)
from an OLE2 Compound Document file.
*/

#include "excelr8/data.hpp"
#include <format>
#include <iostream>
#include <memory>
#include <ostream>
#include <stdexcept>
#include <vector>

namespace excelr8::compdoc {

// Magic cookie that should appear in the first 8 bytes of the file.
const data_t SIGNATURE = { "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" };

const int EOCSID = -2;
const int FREESID = -1;
const int SATSID = -3;
const int MSATSID = -4;
const int EVILSID = -5;

class dllexport CompDocError : public std::runtime_error {
    using std::runtime_error::runtime_error;
};

class dllexport DirNode {
private:
    std::ostream& logfile;
    unsigned char color;
    uint32_t tsinfo[4];

public:
    std::string name;
    int32_t first_sid, tot_size;
    int32_t did, left_did, right_did, root_did;
    int32_t parent = -1; // -1 indicates orphan, fixed up later
    std::vector<int32_t> children;
    unsigned char etype;

    // dent is the 128-byte directory entry
    DirNode(int did, const data_t dent, int debug = 0, std::ostream& logfile = std::cout);
    void dump(int debug = 1);
};

class dllexport CompDoc {
    // Compound document handler

private:
    std::ostream& logfile;
    bool ignore_workbook_corruption;
    int debug;
    int sec_size, short_sec_size;
    int32_t dir_first_sec_sid, min_size_std_stream;
    const data_t& mem;
    int mem_data_secs, mem_data_len;
    std::vector<unsigned char> seen;
    std::vector<int> SAT, SSAT;
    std::vector<DirNode*> dirlist;
    data_t* SSCS = nullptr;

    data_t& _get_stream(const data_t& mem, int base, std::vector<int>& sat, int sec_size, int start_sid, int size = -1, std::string name = "", int seen_id = -1);
    DirNode* _dir_search(const std::vector<std::string>& path, int storage_did = 0);
    std::tuple<const data_t*, int, int> _locate_stream(const data_t& mem, int base, std::vector<int>& sat, int sec_size, int start_sid, int expected_stream_size, const std::string& qname, int seen_id);

public:
    CompDoc(const data_t& mem, std::ostream& logfile = std::cout, int debug = 0, bool ignore_workbook_corruption = false);
    data_t* get_named_stream(const std::string& qname);
    std::tuple<const data_t*, int, int> locate_named_stream(const std::string& qname);
};

void _build_family_tree(const std::vector<DirNode*>& dirlist, int parent_did, int child_did);

template <typename T>
void dump_list(std::vector<T>& list, int stride, std::ostream& f = std::cout);

}