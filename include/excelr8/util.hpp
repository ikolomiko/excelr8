#pragma once

#include "excelr8/data.hpp"
#include <string>
#include <vector>

namespace excelr8::util {

    /*
    template <typename T>
    T _unpack(const char* buffer, size_t offset);

    template <typename... Ts>
    dllexport std::tuple<Ts...> unpack(const char* buffer);

    template <typename... Ts>
    dllexport std::tuple<Ts...> unpack(const data_t& buffer);*/

    dllexport std::string unicode(const data_t&, const std::string& encoding);

}