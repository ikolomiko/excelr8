#pragma once

#include "excelr8/data.hpp"
#include <cstddef>
#include <cstdint>
#include <string>

namespace excelr8::book {
    
class Book;


/**
    Information relating to a named reference, formula, macro, etc.

    Note:
        Name information is **not** extracted from files older than
        Excel 5.0 (Book::biff_version < 50)
*/
class Name {
private:
    Book* book = nullptr; // parent

    /// false = visible; true = hidden
    bool hidden = false;

    /// false = Command macro; true = Function macro. Relevant only if macro == true
    bool func = false;

    /// false = Sheet macro; true = VisualBasic macro. Relevant only if macro == true
    bool vbasic = false;

    /// false = Standard name; true = Macro name
    bool macro = false;

    /// false = Simple formula; true = Complex formula (array formula or user defined).
    /// Note: No examples have been sighted.
    bool complex = false;

    /// false = User-defined name; true = Built-in name
    /// Common examples: "Print_Area", "Print_Titles"; see OOo docs for full list
    bool builtin = false;

    /// Function group. Relevant only if macro == true; see OOo docs for values.
    int8_t funcgroup = false;

    /// false = Formula definition; true = Binary data
    /// Note: No examples have been sighted.
    bool binary = false;

    /// The index of this object in Book::name_obj_list
    size_t name_index = 0;

    /// A Unicode string. If builtin, decoded as per OOo docs.
    std::string name;

    /// An 8-bit string.
    data_t raw_formula;

    /// -1: The name is global (visible in all calculation sheets).
    /// -2: The name belongs to a macro or a VBA sheet.
    /// -3: The name is invalid.
    /// 0 <= scope < Book::nsheets: The name is local to the sheet
    /// whose index is scope.
    int64_t scope = -1;

    /// The result of evaulating the formula, if any.
    /// If no formula, or evaluation of the formula encountered problems,
    /// the result is nullptr. Otherwise, the result is a single instance
    /// of the excelr8::formula::Operand class.
    void* result = nullptr; // TODO change the variable type
};
}