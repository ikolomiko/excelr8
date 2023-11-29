#pragma once

#include "excelr8/name.hpp"
#include "excelr8/formatting.hpp"
#include <string>
#include <unordered_map>
#include <vector>

// Forward declaration
namespace excelr8::formatting {
    class Font;
    class XF;
    class Format;
}

namespace excelr8::book {

/**
    Contents of a "workbook"

    Warning:
        You should never instantiate this class yourself. You use the `Book`
        object that was returned when you called `excelr8::open_workbook()`.
*/
class Book {

public:
    /// The number of worksheets present in the workbook file.
    /// This information is available even when no sheets have been loaded.
    size_t nsheets = 0;

    /// Which date system was in force when this file was last saved.
    /// 0: 1900 system (the Excel for Windows default).
    /// 1: 1904 system (the Excel for Macintosh default).
    /// Defaults to 0 in case it's not specified in the file.
    int8_t datemode = 0;

    /// Version of BIFF (Binary Interchange File Format) used to create the file.
    /// Latest is 8.0 (represented here as 80), introduced with Excel 97.
    /// Earliest supported by this module: 2.0 (represented as 20).
    int8_t biff_version = 0;

    /// Vector containing a Name object for each NAME record in the workbook
    std::vector<Name> name_obj_list;

    /// An integer denoting the character set used for strings in this file.
    /// For BIFF 8 and later, this will be 1200, meaning Unicode;
    /// more precisely, UTF_16_LE.
    /// For earlier versions, this is used to derive the appropriate Python
    /// encoding to be used to convert to Unicode.
    /// Examples: 1252 -> "cp1252", 10000 -> "mac_roman"
    int codepage = -1;

    /// The encoding that was derived from the codepage.
    std::string encoding;

    /// A pair containing the telephone country code for:
    /// [0]: the user-interface setting when the file was created.
    /// [1]: the regional settings.
    /// Example: (1, 61) meaning (USA, Australia).
    /// This information may give a clue to the correct encoding for an
    /// unknown codepage. For a long list of observed values, refer to the
    /// OpenOffice.org documentation for the COUNTRY record.
    std::pair<int, int> countries = {0, 0};

    /// What (if anything) is recorded as the name of the last user to
    /// save to file.
    std::string user_name;

    /// A vector of excelr8::formatting::Font class instances,
    /// each corresponding to a FONT record.
    std::vector<formatting::Font> font_list;

    /// A vector of excelr8::formatting::XF class instances,
    /// each corresponding to an XF record
    std::vector<formatting::XF> xf_list;

    /// A vector of excelr8::formatting::Format objects, each corresponding
    /// to a FORMAT record, in the order that they appear in the input file.
    /// It does *not* contain builtin formats.
    /// If you are creating an output file using (for example) xlwt,
    /// use this list.
    /// The collection to be used for all visual rendering purposes is
    /// 'format_map'.
    std::vector<formatting::Format> format_list;

    /// The mapping from excelr8::formatting::XF::format_key
    /// to excelr8::formatting::Format object.
    std::unordered_map<int, formatting::Format&> format_map;

    /// This provides access via name to the extended format information for
    /// both built-in styles and user-defined styles.
    ///
    /// It maps 'name' to '{built_in, xf_index}', where 'name' is either
    /// the name of a user-defined style, or the name of the built-in styles.
    /// Known built-in names are "Normal", "RowLevel_1" to "RowLevel_7",
    /// "ColLevel_1" to "ColLevel_7", "Comma", "Currency", "Percent", "Comma [0]",
    /// "Currency [0]", "Hyperlink" and "Followed Hyperlink".
    ///
    /// 'built_in' has the following meanings:
    /// true: built-in style; false: user-defined style
    ///
    /// 'xf_index' is an index to Book::xf_list
    /// References: OOo docs s6.99 (STYLE record); Excel UI Format/Style
    /// Extracted only if open_workbook(..., formatting_info=true)
    std::unordered_map<std::string, std::pair<bool, int>> style_name_map;

    /// This provides definitions for color indexes. Please refer to
    /// 'palette' for an explanation of how colors are represented in Excel.
    ///
    /// Color indexes into the palette map into {red, green, blue} tuples.
    /// "Magic" indexes e.g. 0x7FFF map to nullptr.
    ///
    /// color_map is what you need if you want to render cells on screen or
    /// in a PDF file. If you are writing an output XLS file, use palette_record.
    ///
    /// Note: Extracted only if open_workbook(..., formatting_info=true)
    std::unordered_map<uint16_t, color_t*> color_map;

    /// If the user has changed any of the colors in the standard palette, the
    /// XLS file will contain a PALETTE record with 56 (16 for Excel 4.0 and
    /// earlier) RGB values in it, and this list will be e.g.
    /// [(r0, b0, g0), ..., (r55, b55, g55)].
    /// Otherwise this list will be empty. This is what you need if you are
    /// writing an output XLS file. If you want to render cells on screen or in a
    /// PDF file, use color_map.
    ///
    /// Note: Extracted only if open_workbook(..., formatting_info=true)
    std::vector<color_t> palette_record;

    /// Time in seconds to extract the XLS image as a contiguous string
    /// (or mmap equivalent).
    float load_time_stage_1 = -1.0;

    /// Time in seconds to parse the data from the contiguous string
    /// (or mmap equivalent).
    float load_time_stage_2 = -1.0;

    bool formatting_info = false;

    void derive_encoding();

    int verbosity = 0;
    
};
}