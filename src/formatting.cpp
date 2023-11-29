#include "excelr8/formatting.hpp"
#include "excelr8/biff.hpp"
#include "excelr8/book.hpp"
#include "excelr8/data.hpp"
#include <format>
#include <iostream>
#include <unordered_map>

namespace excelr8::book {
class Book;
}

namespace excelr8::formatting {

int DEBUG = 0;

void initialize_color_map(book::Book& book)
{
    book.color_map.clear();
    // book.color_indexes_used = {}; WONT IMPLEMENT
    if (!book.formatting_info) {
        return;
    }

    // Add the 8 invariant colors
    for (int i = 0; i < 8; i++) {
        book.color_map[i] = new color_t(excel_default_palette_b8[i]);
    }

    // Add the default palette depending on the version
    auto& dpal = default_palette.at(book.biff_version);
    size_t ndpal = dpal.size();
    for (int i = 0; i < ndpal; i++) {
        book.color_map[i + 8] = new color_t(dpal[i]);
    }

    // Add the specials -- nullptr means the RGB value is not known
    // System window text color for border lines
    book.color_map[ndpal + 8] = nullptr;

    // System window background color for pattern background
    book.color_map[ndpal + 8 + 1] = nullptr;

    // System ToolTip text color (used in note objects)
    book.color_map[0x51] = nullptr;

    // 32767, system window text color for fonts
    book.color_map[0x7FFF] = nullptr;
}

/**
    General purpose function. Uses Euclidean distance.
    So far only used for pre-BIFF8 WINDOW2 record.
    Doesn't have to be fast.
    Doesn't have to be fancy.
*/
uint16_t nearest_color_index(std::unordered_map<uint16_t, color_t*>& color_map, color_t& rgb, int debug)
{
    int best_metric = 3 * 256 * 256;
    uint16_t best_colorx = 0;
    for (auto& [colorx, cand_rgb] : color_map) {
        if (cand_rgb == nullptr) {
            continue;
        }

        int metric = 0;
        const auto& [r1, g1, b1] = rgb;
        const auto& [r2, g2, b2] = *cand_rgb;
        metric += (r1 - r2) * (r1 - r2);
        metric += (g1 - g2) * (g1 - g2);
        metric += (b1 - b2) * (b1 - b2);

        if (metric < best_metric) {
            best_metric = metric;
            best_colorx = colorx;
            if (metric == 0) {
                break;
            }
        }
    }
    if (debug) {
        const auto& rrgb = *color_map[best_colorx];
        std::cout << std::format("nearest_color_index for RGB(%d,%d,%d) is %d -> RGB(%d,%d%d); best_metric is %d\n",
            std::get<0>(rgb), std::get<1>(rgb), std::get<2>(rgb), best_colorx,
            std::get<0>(rrgb), std::get<1>(rrgb), std::get<2>(rrgb), best_metric);
    }
    return best_colorx;
}

void handle_efont(book::Book& book, const data_t& data)
{
    if (!book.formatting_info) {
        return;
    }
    auto& last_font = book.font_list.back();
    last_font.color_index = std::get<0>(data.unpack<pytype_H>());
}

void handle_font(book::Book& book, const data_t& data)
{
    if (!book.formatting_info) {
        return;
    }

    if (book.encoding.empty()) {
        book.derive_encoding();
    }

    bool blah = DEBUG or book.verbosity >= 2;
    auto& bv = book.biff_version;
    size_t k = book.font_list.size();

    if (k == 4) {
        Font f;
        f.name = "Dummy Font";
        f.font_index = k;
        book.font_list.push_back(f);
        k += 1;
    }

    Font f;
    f.font_index = k;
    book.font_list.push_back(f);
    if (bv >= 50) {
        pytype_H option_flags;
        std::tie(f.height, option_flags, f.color_index, f.weight, f.escapement, f.underline_type, f.family, f.character_set)
            = data.slice(0, 13).unpack<pytype_H, pytype_H, pytype_H, pytype_H, pytype_H, pytype_B, pytype_B, pytype_B>();

        f.bold = option_flags & 1;
        f.italic = (option_flags & 2) >> 1;
        f.underlined = (option_flags & 4) >> 2;
        f.struck_out = (option_flags & 8) >> 3;
        f.outline = (option_flags & 16) >> 4;
        f.shadow = (option_flags & 32) >> 5;
        if (bv >= 80) {
            f.name = unpack_unicode(data, 14, 1);
        } else {
            f.name = unpack_string(data, 14, book.encoding, 1);
        }
    } else if (bv >= 30) {
        uint16_t option_flags;
        std::tie(f.height, option_flags, f.color_index) = data.slice(0, 6).unpack<pytype_H, pytype_H, pytype_H>();
        f.bold = option_flags & 1;
        f.italic = (option_flags & 2) >> 1;
        f.underlined = (option_flags & 4) >> 2;
        f.struck_out = (option_flags & 8) >> 3;
        f.outline = (option_flags & 16) >> 4;
        f.shadow = (option_flags & 32) >> 5;
        f.name = unpack_string(data, 6, book.encoding, 1);

        // Now cook up the remaining attributes ...
        f.weight = f.bold ? 700 : 400;
        f.escapement = 0; // None
        f.underline_type = f.underlined; // None or Single
        f.family = 0; // Unknown / don't care
        f.character_set = 1; // System default (0 means "ANSI Latin")
    } else { // BIFF2
        uint16_t option_flags;
        std::tie(f.height, option_flags) = data.slice(0, 4).unpack<pytype_H, pytype_H>();
        f.color_index = 0x7FFF; // "system window text color"
        f.bold = option_flags & 1;
        f.italic = (option_flags & 2) >> 1;
        f.underlined = (option_flags & 4) >> 2;
        f.struck_out = (option_flags & 8) >> 3;
        f.outline = 0;
        f.shadow = 0;
        f.name = unpack_string(data, 4, book.encoding, 1);

        // Now cook up the remaining attributes ...
        f.weight = f.bold ? 700 : 400;
        f.escapement = 0; // None
        f.underline_type = f.underlined; // None or Single
        f.family = 0; // Unknown / don't care
        f.character_set = 1; // System default (0 means "ANSI Latin")
    }

    if (blah) {
        // f.dump() WONT IMPLEMENT
    }
}

Format::Format(int format_key, int ty, const std::string& format_str)
    : format_key(format_key)
    , type(ty)
    , format_str(format_str)
{
}

bool Format::operator==(const Format& other) const
{
    return (format_key == other.format_key) and (type == other.type) and (format_str == other.format_str);
}

bool Format::operator!=(const Format& other) const
{
    return (format_key != other.format_key) or (type != other.type) or (format_str != other.format_str);
}

std::unordered_map<int, int> std_format_code_types; // TODO fill this

}