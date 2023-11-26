#pragma once

#include "excelr8/data.hpp"
#include <fstream>
#include <ostream>
#include <stdexcept>
#include <string>
#include <tuple>
#include <unordered_map>
#include <unordered_set>
#include <vector>

namespace excelr8::biff {

class dllexport Excelr8Error : public std::runtime_error {
    // An exception indicating problems reading data from an Excel file.
};

class BaseObject {
    // Not implemented yet
};

class dllexport Manifest {
public:
    int n;
    int mask;
    std::string attr;
};

const int FUN = 0; // unknown
const int FDT = 1; // date
const int FNU = 2; // number
const int FGE = 3; // general
const int FTX = 4; // text
const int DATEFORMAT = FDT;
const int NUMBERFORMAT = FNU;

const int XL_CELL_EMPTY = 0;
const int XL_CELL_TEXT = 1;
const int XL_CELL_NUMBER = 2;
const int XL_CELL_DATE = 3;
const int XL_CELL_BOOLEAN = 4;
const int XL_CELL_ERROR = 5;
const int XL_CELL_BLANK = 6;

const std::unordered_map<int, std::string> biff_text_from_num = {
    { 0, "(not BIFF)" },
    { 20, "2.0" },
    { 21, "2.1" },
    { 30, "3" },
    { 40, "4S" },
    { 45, "4W" },
    { 50, "5" },
    { 70, "7" },
    { 80, "8" },
    { 85, "8X" },
};

const std::unordered_map<int, std::string> error_text_from_code = {
    { 0x00, "#NULL!" }, // Intersection of two cell ranges is empty
    { 0x07, "#DIV/0!" }, // Division by zero
    { 0x0F, "#VALUE!" }, // Wrong type of operand
    { 0x17, "#REF!" }, // Illegal or deleted cell reference
    { 0x1D, "#NAME?" }, // Wrong function or range name
    { 0x24, "#NUM!" }, // Value range overflow
    { 0x2A, "#N/A" }, // Argument or function not available
};

const int BIFF_FIRST_UNICODE = 80;

const int XL_WORKBOOK_GLOBALS = 0x5;
const int WBKBLOBAL = 0x5;
const int XL_WORKBOOK_GLOBALS_4W = 0x100;
const int XL_WORKSHEET = 0x10;
const int WRKSHEET = 0x10;

const int XL_BOUNDSHEET_WORKSHEET = 0x00;
const int XL_BOUNDSHEET_CHART = 0x02;
const int XL_BOUNDSHEET_VB_MODULE = 0x06;

// XL_RK2 = 0x7e
const int XL_ARRAY = 0x0221;
const int XL_ARRAY2 = 0x0021;
const int XL_BLANK = 0x0201;
const int XL_BLANK_B2 = 0x01;
const int XL_BOF = 0x809;
const int XL_BOOLERR = 0x205;
const int XL_BOOLERR_B2 = 0x5;
const int XL_BOUNDSHEET = 0x85;
const int XL_BUILTINFMTCOUNT = 0x56;
const int XL_CF = 0x01B1;
const int XL_CODEPAGE = 0x42;
const int XL_COLINFO = 0x7D;
const int XL_COLUMNDEFAULT = 0x20; // BIFF2 only
const int XL_COLWIDTH = 0x24; // BIFF2 only
const int XL_CONDFMT = 0x01B0;
const int XL_CONTINUE = 0x3c;
const int XL_COUNTRY = 0x8C;
const int XL_DATEMODE = 0x22;
const int XL_DEFAULTROWHEIGHT = 0x0225;
const int XL_DEFCOLWIDTH = 0x55;
const int XL_DIMENSION = 0x200;
const int XL_DIMENSION2 = 0x0;
const int XL_EFONT = 0x45;
const int XL_EOF = 0x0a;
const int XL_EXTERNNAME = 0x23;
const int XL_EXTERNSHEET = 0x17;
const int XL_EXTSST = 0xff;
const int XL_FEAT11 = 0x872;
const int XL_FILEPASS = 0x2f;
const int XL_FONT = 0x31;
const int XL_FONT_B3B4 = 0x231;
const int XL_FORMAT = 0x41e;
const int XL_FORMAT2 = 0x1E; // BIFF2, BIFF3
const int XL_FORMULA = 0x6;
const int XL_FORMULA3 = 0x206;
const int XL_FORMULA4 = 0x406;
const int XL_GCW = 0xab;
const int XL_HLINK = 0x01B8;
const int XL_QUICKTIP = 0x0800;
const int XL_HORIZONTALPAGEBREAKS = 0x1b;
const int XL_INDEX = 0x20b;
const int XL_INTEGER = 0x2; // BIFF2 only
const int XL_IXFE = 0x44; // BIFF2 only
const int XL_LABEL = 0x204;
const int XL_LABEL_B2 = 0x04;
const int XL_LABELRANGES = 0x15f;
const int XL_LABELSST = 0xfd;
const int XL_LEFTMARGIN = 0x26;
const int XL_TOPMARGIN = 0x28;
const int XL_RIGHTMARGIN = 0x27;
const int XL_BOTTOMMARGIN = 0x29;
const int XL_HEADER = 0x14;
const int XL_FOOTER = 0x15;
const int XL_HCENTER = 0x83;
const int XL_VCENTER = 0x84;
const int XL_MERGEDCELLS = 0xE5;
const int XL_MSO_DRAWING = 0x00EC;
const int XL_MSO_DRAWING_GROUP = 0x00EB;
const int XL_MSO_DRAWING_SELECTION = 0x00ED;
const int XL_MULRK = 0xbd;
const int XL_MULBLANK = 0xbe;
const int XL_NAME = 0x18;
const int XL_NOTE = 0x1c;
const int XL_NUMBER = 0x203;
const int XL_NUMBER_B2 = 0x3;
const int XL_OBJ = 0x5D;
const int XL_PAGESETUP = 0xA1;
const int XL_PALETTE = 0x92;
const int XL_PANE = 0x41;
const int XL_PRINTGRIDLINES = 0x2B;
const int XL_PRINTHEADERS = 0x2A;
const int XL_RK = 0x27e;
const int XL_ROW = 0x208;
const int XL_ROW_B2 = 0x08;
const int XL_RSTRING = 0xd6;
const int XL_SCL = 0x00A0;
const int XL_SHEETHDR = 0x8F; // BIFF4W only
const int XL_SHEETPR = 0x81;
const int XL_SHEETSOFFSET = 0x8E; // BIFF4W only
const int XL_SHRFMLA = 0x04bc;
const int XL_SST = 0xfc;
const int XL_STANDARDWIDTH = 0x99;
const int XL_STRING = 0x207;
const int XL_STRING_B2 = 0x7;
const int XL_STYLE = 0x293;
const int XL_SUPBOOK = 0x1AE; // aka EXTERNALBOOK in OOo docs
const int XL_TABLEOP = 0x236;
const int XL_TABLEOP2 = 0x37;
const int XL_TABLEOP_B2 = 0x36;
const int XL_TXO = 0x1b6;
const int XL_UNCALCED = 0x5e;
const int XL_UNKNOWN = 0xffff;
const int XL_VERTICALPAGEBREAKS = 0x1a;
const int XL_WINDOW2 = 0x023E;
const int XL_WINDOW2_B2 = 0x003E;
const int XL_WRITEACCESS = 0x5C;
const int XL_WSBOOL = XL_SHEETPR;
const int XL_XF = 0xe0;
const int XL_XF2 = 0x0043; // BIFF2 version of XF record
const int XL_XF3 = 0x0243; // BIFF3 version of XF record
const int XL_XF4 = 0x0443; // BIFF4 version of XF record

const std::unordered_map<int, int> boflen = {
    { 0x0809, 8 },
    { 0x0409, 6 },
    { 0x0209, 6 },
    { 0x0009, 4 }
};
const std::vector<int> bofcodes = { 0x0809, 0x0409, 0x0209, 0x0009 };

const std::vector<int> XL_FORMULA_OPCODES = { 0x0006, 0x0406, 0x0206 };

const std::unordered_set<int> _cell_opcode_set = {
    XL_BOOLERR,
    XL_FORMULA,
    XL_FORMULA3,
    XL_FORMULA4,
    XL_LABEL,
    XL_LABELSST,
    XL_MULRK,
    XL_NUMBER,
    XL_RK,
    XL_RSTRING,
};

const std::unordered_map<int, std::string> biff_rec_name_dict = {
    { 0x0000, "DIMENSIONS_B2" },
    { 0x0001, "BLANK_B2" },
    { 0x0002, "INTEGER_B2_ONLY" },
    { 0x0003, "NUMBER_B2" },
    { 0x0004, "LABEL_B2" },
    { 0x0005, "BOOLERR_B2" },
    { 0x0006, "FORMULA" },
    { 0x0007, "STRING_B2" },
    { 0x0008, "ROW_B2" },
    { 0x0009, "BOF_B2" },
    { 0x000A, "EOF" },
    { 0x000B, "INDEX_B2_ONLY" },
    { 0x000C, "CALCCOUNT" },
    { 0x000D, "CALCMODE" },
    { 0x000E, "PRECISION" },
    { 0x000F, "REFMODE" },
    { 0x0010, "DELTA" },
    { 0x0011, "ITERATION" },
    { 0x0012, "PROTECT" },
    { 0x0013, "PASSWORD" },
    { 0x0014, "HEADER" },
    { 0x0015, "FOOTER" },
    { 0x0016, "EXTERNCOUNT" },
    { 0x0017, "EXTERNSHEET" },
    { 0x0018, "NAME_B2,5+" },
    { 0x0019, "WINDOWPROTECT" },
    { 0x001A, "VERTICALPAGEBREAKS" },
    { 0x001B, "HORIZONTALPAGEBREAKS" },
    { 0x001C, "NOTE" },
    { 0x001D, "SELECTION" },
    { 0x001E, "FORMAT_B2-3" },
    { 0x001F, "BUILTINFMTCOUNT_B2" },
    { 0x0020, "COLUMNDEFAULT_B2_ONLY" },
    { 0x0021, "ARRAY_B2_ONLY" },
    { 0x0022, "DATEMODE" },
    { 0x0023, "EXTERNNAME" },
    { 0x0024, "COLWIDTH_B2_ONLY" },
    { 0x0025, "DEFAULTROWHEIGHT_B2_ONLY" },
    { 0x0026, "LEFTMARGIN" },
    { 0x0027, "RIGHTMARGIN" },
    { 0x0028, "TOPMARGIN" },
    { 0x0029, "BOTTOMMARGIN" },
    { 0x002A, "PRINTHEADERS" },
    { 0x002B, "PRINTGRIDLINES" },
    { 0x002F, "FILEPASS" },
    { 0x0031, "FONT" },
    { 0x0032, "FONT2_B2_ONLY" },
    { 0x0036, "TABLEOP_B2" },
    { 0x0037, "TABLEOP2_B2" },
    { 0x003C, "CONTINUE" },
    { 0x003D, "WINDOW1" },
    { 0x003E, "WINDOW2_B2" },
    { 0x0040, "BACKUP" },
    { 0x0041, "PANE" },
    { 0x0042, "CODEPAGE" },
    { 0x0043, "XF_B2" },
    { 0x0044, "IXFE_B2_ONLY" },
    { 0x0045, "EFONT_B2_ONLY" },
    { 0x004D, "PLS" },
    { 0x0051, "DCONREF" },
    { 0x0055, "DEFCOLWIDTH" },
    { 0x0056, "BUILTINFMTCOUNT_B3-4" },
    { 0x0059, "XCT" },
    { 0x005A, "CRN" },
    { 0x005B, "FILESHARING" },
    { 0x005C, "WRITEACCESS" },
    { 0x005D, "OBJECT" },
    { 0x005E, "UNCALCED" },
    { 0x005F, "SAVERECALC" },
    { 0x0063, "OBJECTPROTECT" },
    { 0x007D, "COLINFO" },
    { 0x007E, "RK2_mythical_?" },
    { 0x0080, "GUTS" },
    { 0x0081, "WSBOOL" },
    { 0x0082, "GRIDSET" },
    { 0x0083, "HCENTER" },
    { 0x0084, "VCENTER" },
    { 0x0085, "BOUNDSHEET" },
    { 0x0086, "WRITEPROT" },
    { 0x008C, "COUNTRY" },
    { 0x008D, "HIDEOBJ" },
    { 0x008E, "SHEETSOFFSET" },
    { 0x008F, "SHEETHDR" },
    { 0x0090, "SORT" },
    { 0x0092, "PALETTE" },
    { 0x0099, "STANDARDWIDTH" },
    { 0x009B, "FILTERMODE" },
    { 0x009C, "FNGROUPCOUNT" },
    { 0x009D, "AUTOFILTERINFO" },
    { 0x009E, "AUTOFILTER" },
    { 0x00A0, "SCL" },
    { 0x00A1, "SETUP" },
    { 0x00AB, "GCW" },
    { 0x00BD, "MULRK" },
    { 0x00BE, "MULBLANK" },
    { 0x00C1, "MMS" },
    { 0x00D6, "RSTRING" },
    { 0x00D7, "DBCELL" },
    { 0x00DA, "BOOKBOOL" },
    { 0x00DD, "SCENPROTECT" },
    { 0x00E0, "XF" },
    { 0x00E1, "INTERFACEHDR" },
    { 0x00E2, "INTERFACEEND" },
    { 0x00E5, "MERGEDCELLS" },
    { 0x00E9, "BITMAP" },
    { 0x00EB, "MSO_DRAWING_GROUP" },
    { 0x00EC, "MSO_DRAWING" },
    { 0x00ED, "MSO_DRAWING_SELECTION" },
    { 0x00EF, "PHONETIC" },
    { 0x00FC, "SST" },
    { 0x00FD, "LABELSST" },
    { 0x00FF, "EXTSST" },
    { 0x013D, "TABID" },
    { 0x015F, "LABELRANGES" },
    { 0x0160, "USESELFS" },
    { 0x0161, "DSF" },
    { 0x01AE, "SUPBOOK" },
    { 0x01AF, "PROTECTIONREV4" },
    { 0x01B0, "CONDFMT" },
    { 0x01B1, "CF" },
    { 0x01B2, "DVAL" },
    { 0x01B6, "TXO" },
    { 0x01B7, "REFRESHALL" },
    { 0x01B8, "HLINK" },
    { 0x01BC, "PASSWORDREV4" },
    { 0x01BE, "DV" },
    { 0x01C0, "XL9FILE" },
    { 0x01C1, "RECALCID" },
    { 0x0200, "DIMENSIONS" },
    { 0x0201, "BLANK" },
    { 0x0203, "NUMBER" },
    { 0x0204, "LABEL" },
    { 0x0205, "BOOLERR" },
    { 0x0206, "FORMULA_B3" },
    { 0x0207, "STRING" },
    { 0x0208, "ROW" },
    { 0x0209, "BOF" },
    { 0x020B, "INDEX_B3+" },
    { 0x0218, "NAME" },
    { 0x0221, "ARRAY" },
    { 0x0223, "EXTERNNAME_B3-4" },
    { 0x0225, "DEFAULTROWHEIGHT" },
    { 0x0231, "FONT_B3B4" },
    { 0x0236, "TABLEOP" },
    { 0x023E, "WINDOW2" },
    { 0x0243, "XF_B3" },
    { 0x027E, "RK" },
    { 0x0293, "STYLE" },
    { 0x0406, "FORMULA_B4" },
    { 0x0409, "BOF" },
    { 0x041E, "FORMAT" },
    { 0x0443, "XF_B4" },
    { 0x04BC, "SHRFMLA" },
    { 0x0800, "QUICKTIP" },
    { 0x0809, "BOF" },
    { 0x0862, "SHEETLAYOUT" },
    { 0x0867, "SHEETPROTECTION" },
    { 0x0868, "RANGEPROTECTION" },
};

const std::unordered_map<int, std::string> encoding_from_codepage {
    {1200 , "utf_16_le"},
    {10000, "mac_roman"},
    {10006, "mac_greek"}, // guess
    {10007, "mac_cyrillic"}, // guess
    {10029, "mac_latin2"}, // guess
    {10079, "mac_iceland"}, // guess
    {10081, "mac_turkish"}, // guess
    {32768, "mac_roman"},
    {32769, "cp1252"},
};
/*
    some more guessing, for Indic scripts
    codepage 57000 range:
    2 Devanagari [0]
    3 Bengali [1]
    4 Tamil [5]
    5 Telegu [6]
    6 Assamese [1] c.f. Bengali
    7 Oriya [4]
    8 Kannada [7]
    9 Malayalam [8]
    10 Gujarati [3]
    11 Gurmukhi [2]
*/

dllexport bool is_cell_opcode(int c);
dllexport void upkbits(void* tgt_obj, int src, std::vector<Manifest>& manifests);
dllexport void upkbitsL(void* tgt_obj, int src, std::vector<Manifest>& manifests);
dllexport std::string unpack_string(const data_t& data, int pos, const std::string& encoding, int lenlen);
dllexport std::pair<std::string, int> unpack_string_update_pos(const data_t& data, int pos, const std::string& encoding, int lenlen, int known_len);
dllexport std::string unpack_unicode(const data_t& data, int pos, int lenlen);
dllexport std::pair<std::string, int> unpack_unicode_update_pos(const data_t& data, int pos, int lenlen, int known_len);
dllexport int unpack_cell_range_address_list_update_pos(std::vector<pytype_H>& output_list, const data_t data, int pos, int addr_size);
dllexport void hex_char_dump(const std::string& strg, int ofs, int dlen, int base, std::ostream& fout, bool unnumbered);
dllexport void biff_dump(const data_t& mem, int stream_offset, int stream_len, int base, std::ostream& fout, bool unnumbered);
dllexport void biff_count_records(const data_t& mem, int stream_offset, int stream_len, std::ostream& fout);

}
