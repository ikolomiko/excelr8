#pragma once

/*
    Module for formatting information.
*/

#include "excelr8/biff.hpp"
#include <cstdint>
#include <iostream>
#include <tuple>
#include <unordered_map>
#include <vector>

using namespace excelr8::biff;
using color_t = std::tuple<unsigned char, unsigned char, unsigned char>;
using palette_t = std::vector<color_t>;

// Forward declaration
namespace excelr8::book {
class Book;
}

namespace excelr8::formatting {

const std::unordered_map<int, int> _cellty_from_fmtty = {
    { FNU, XL_CELL_NUMBER },
    { FUN, XL_CELL_NUMBER },
    { FGE, XL_CELL_NUMBER },
    { FDT, XL_CELL_DATE },
    { FTX, XL_CELL_NUMBER }, // Yes, a number can be formatted as text.
};

// clang-format off
const palette_t excel_default_palette_b5 = {
    {  0,   0,   0}, {255, 255, 255}, {255,   0,   0}, {  0, 255,   0},
    {  0,   0, 255}, {255, 255,   0}, {255,   0, 255}, {  0, 255, 255},
    {128,   0,   0}, {  0, 128,   0}, {  0,   0, 128}, {128, 128,   0},
    {128,   0, 128}, {  0, 128, 128}, {192, 192, 192}, {128, 128, 128},
    {153, 153, 255}, {153,  51, 102}, {255, 255, 204}, {204, 255, 255},
    {102,   0, 102}, {255, 128, 128}, {  0, 102, 204}, {204, 204, 255},
    {  0,   0, 128}, {255,   0, 255}, {255, 255,   0}, {  0, 255, 255},
    {128,   0, 128}, {128,   0,   0}, {  0, 128, 128}, {  0,   0, 255},
    {  0, 204, 255}, {204, 255, 255}, {204, 255, 204}, {255, 255, 153},
    {153, 204, 255}, {255, 153, 204}, {204, 153, 255}, {227, 227, 227},
    { 51, 102, 255}, { 51, 204, 204}, {153, 204,   0}, {255, 204,   0},
    {255, 153,   0}, {255, 102,   0}, {102, 102, 153}, {150, 150, 150},
    {  0,  51, 102}, { 51, 153, 102}, {  0,  51,   0}, { 51,  51,   0},
    {153,  51,   0}, {153,  51, 102}, { 51,  51, 153}, { 51,  51,  51},
};

const palette_t excel_default_palette_b2 = {
    {  0,   0,   0}, {255, 255, 255}, {255,   0,   0}, {  0, 255,   0},
    {  0,   0, 255}, {255, 255,   0}, {255,   0, 255}, {  0, 255, 255},
    {128,   0,   0}, {  0, 128,   0}, {  0,   0, 128}, {128, 128,   0},
    {128,   0, 128}, {  0, 128, 128}, {192, 192, 192}, {128, 128, 128},
};

// Following table borrowed from Gnumeric 1.4 source.
// Checked against OOo docs and MS docs.
const palette_t excel_default_palette_b8 = { // {red, green, blue}
    {  0,  0,  0}, {255,255,255}, {255,  0,  0}, {  0,255,  0}, // 0
    {  0,  0,255}, {255,255,  0}, {255,  0,255}, {  0,255,255}, // 4
    {128,  0,  0}, {  0,128,  0}, {  0,  0,128}, {128,128,  0}, // 8
    {128,  0,128}, {  0,128,128}, {192,192,192}, {128,128,128}, // 12
    {153,153,255}, {153, 51,102}, {255,255,204}, {204,255,255}, // 16
    {102,  0,102}, {255,128,128}, {  0,102,204}, {204,204,255}, // 20
    {  0,  0,128}, {255,  0,255}, {255,255,  0}, {  0,255,255}, // 24
    {128,  0,128}, {128,  0,  0}, {  0,128,128}, {  0,  0,255}, // 28
    {  0,204,255}, {204,255,255}, {204,255,204}, {255,255,153}, // 32
    {153,204,255}, {255,153,204}, {204,153,255}, {255,204,153}, // 36
    { 51,102,255}, { 51,204,204}, {153,204,  0}, {255,204,  0}, // 40
    {255,153,  0}, {255,102,  0}, {102,102,153}, {150,150,150}, // 44
    {  0, 51,102}, { 51,153,102}, {  0, 51,  0}, { 51, 51,  0}, // 48
    {153, 51,  0}, {153, 51,102}, { 51, 51,153}, { 51, 51, 51}, // 52
};
// clang-format on

const std::unordered_map<int, palette_t> default_palette = {
    { 80, excel_default_palette_b8 },
    { 70, excel_default_palette_b5 },
    { 50, excel_default_palette_b5 },
    { 45, excel_default_palette_b2 },
    { 40, excel_default_palette_b2 },
    { 30, excel_default_palette_b2 },
    { 21, excel_default_palette_b2 },
    { 20, excel_default_palette_b2 },
};

// 00H = Normal
// 01H = RowLevel_lv (see next field)
// 02H = ColLevel_lv (see next field)
// 03H = Comma
// 04H = Currency
// 05H = Percent
// 06H = Comma [0] (BIFF4-BIFF8)
// 07H = Currency [0] (BIFF4-BIFF8)
// 08H = Hyperlink (BIFF8)
// 09H = Followed Hyperlink (BIFF8)
const std::vector<std::string> built_in_style_names = {
    "Normal",
    "RowLevel_",
    "ColLevel_",
    "Comma",
    "Currency",
    "Percent",
    "Comma [0]",
    "Currency [0]",
    "Hyperlink",
    "Followed Hyperlink",
};

void initialize_color_map(excelr8::book::Book& book);
uint16_t nearest_color_index(std::unordered_map<uint16_t, color_t*>& color_map, color_t rgb, int debug = 0);

/**
    An Excel "font" contains the details of not only what is normally
    considered as a font, but also several other display attributes.
    Items correspond to those in the Excel UI's Format -> Cells -> Font tab
*/
class Font {
public:
    /// true = Characters are bold. Redundant; see "weight" attribute.
    bool bold = false;

    /**
        Values:

        0 = ANSI Latin
        1 = System default
        2 = Symbol,
        77 = Apple Roman,
        128 = ANSI Japanese Shift-JIS,
        129 = ANSI Korean (Hangul),
        130 = ANSI Korean (Johab),
        134 = ANSI Chinese Simplified GBK,
        136 = ANSI Chinese Traditional BIG5,
        161 = ANSI Greek,
        162 = ANSI Turkish,
        163 = ANSI Vietnamese,
        177 = ANSI Hebrew,
        178 = ANSI Arabic,
        186 = ANSI Baltic,
        204 = ANSI Cyrillic,
        222 = ANSI Thai,
        238 = ANSI Latin II (Central European),
        255 = OEM Latin I
    */
    uint8_t character_set = 0;

    /// An explanation of "color index" is given in 'palette'
    uint16_t color_index = 0;

    /// 1 = Superscript, 2 = Subscript.
    uint16_t escapement = 0;

    /**
        Values:
        0 = None (unknown or don't care)
        1 = Roman (variable width, serifed)
        2 = Swiss (variable width, sans-serifed)
        3 = Modern (fixed width, serifed or sans-serifed)
        4 = Script (cursive)
        5 = Decorative (specialised, for example Old English, Fraktur)
    */
    uint8_t family = 0;

    /// The 0-based index used to refer to this Font() instance.
    /// Note that index 4 is never used; excelr8 supplies a dummy place-holder.
    int font_index = 0;

    /// Height of the font (in twips). A twip = 1/20 of a point.
    uint16_t height = 0;

    /// true = Characters are italic.
    bool italic = false;

    /// The name of the font. Example: "Arial".
    std::string name;

    /// true = Characters are struck out.
    bool struck_out = false;

    /**
        Values:
        0 = None
        1 = Single; 0x21 (33) = Single accounting
        2 = Double; 0x22 (34) = Double accounting
    */
    uint8_t underline_type = 0;

    /// true = Characters are underlined. Redundant;
    /// see underline_type attribute.
    bool underlined = false;

    /// Font weight (100-1000). Standard values are 400 for normal text
    /// and 700 for bold text.
    uint16_t weight = 400;

    /// true = Font is outline style (Macintosh only)
    bool outline = false;

    /// true = Font is shadow style (Macintosh only)
    bool shadow = false;
};

void handle_efont(book::Book& book, const data_t& data); // BIFF2 only

void handle_font(book::Book& book, const data_t& data);

/**
    "Number format" information from a FORMAT record.
*/
class Format {
public:
    /// The key into excelr8::book::Book::format_map
    int format_key = 0;

    /**
        A classification that has been inferred from the format string.
        Currently, this is used only to distinguish between numbers and dates.
        Values:
        FUN = 0; unknown
        FDT = 1; date
        FNU = 2; number
        FGE = 3; general
        FTX = 4; text
    */
    int type = FUN;

    /// The format string
    std::string format_str;

    Format(int format_key, int ty, const std::string& format_str);
    bool operator==(const Format& other) const;
    bool operator!=(const Format& other) const;
};

/// "std" == "standard for US English locale"
/// #### todo ... a lot of work to tailor these to the user's locale.
/// See e.g. gnumeric-1.x.y/src/formats.c
const std::unordered_map<int, std::string> std_format_strings = {
    { 0x00, "General" },
    { 0x01, "0" },
    { 0x02, "0.00" },
    { 0x03, "#,##0" },
    { 0x04, "#,##0.00" },
    { 0x05, "$#,##0_);($#,##0)" },
    { 0x06, "$#,##0_);[Red]($#,##0)" },
    { 0x07, "$#,##0.00_);($#,##0.00)" },
    { 0x08, "$#,##0.00_);[Red]($#,##0.00)" },
    { 0x09, "0%" },
    { 0x0a, "0.00%" },
    { 0x0b, "0.00E+00" },
    { 0x0c, "# ?/?" },
    { 0x0d, "# ??/??" },
    { 0x0e, "m/d/yy" },
    { 0x0f, "d-mmm-yy" },
    { 0x10, "d-mmm" },
    { 0x11, "mmm-yy" },
    { 0x12, "h:mm AM/PM" },
    { 0x13, "h:mm:ss AM/PM" },
    { 0x14, "h:mm" },
    { 0x15, "h:mm:ss" },
    { 0x16, "m/d/yy h:mm" },
    { 0x25, "#,##0_);(#,##0)" },
    { 0x26, "#,##0_);[Red](#,##0)" },
    { 0x27, "#,##0.00_);(#,##0.00)" },
    { 0x28, "#,##0.00_);[Red](#,##0.00)" },
    { 0x29, "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)" },
    { 0x2a, "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)" },
    { 0x2b, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)" },
    { 0x2c, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)" },
    { 0x2d, "mm:ss" },
    { 0x2e, "[h]:mm:ss" },
    { 0x2f, "mm:ss.0" },
    { 0x30, "##0.0E+0" },
    { 0x31, "@" },
};

/// both-inclusive ranges of "standard" format codes
/// Source: the openoffice.org doc't
/// and the OOXML spec Part 4, section 3.8.30
const std::vector<std::tuple<int, int, int>> fmt_code_ranges = {
    { 0, 0, FGE },
    { 1, 13, FNU },
    { 14, 22, FDT },
    { 27, 36, FDT }, // CJK date formats
    { 37, 44, FNU },
    { 45, 47, FDT },
    { 48, 48, FNU },
    { 49, 49, FTX },
    // Gnumeric assumes (or assumed) that built-in formats finish at 49, not at 163
    { 50, 58, FDT }, // CJK date formats
    { 59, 62, FNU }, // Thai number (currency?) formats
    { 67, 70, FNU }, // Thai number (currency?) formats
    { 71, 81, FDT }, // Thai date formats
};

extern std::unordered_map<int, int> std_format_code_types;

class XF;
}