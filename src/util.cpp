#include "excelr8/util.hpp"
#include <cstring>
#include <iostream>
#include <unicode/ucnv.h>
#include <unicode/utypes.h>
#include <vector>

namespace excelr8::util {
/*
    template <typename T>
    T _unpack(const char* buffer, size_t& offset)
    {
        T result;
        std::memcpy(&result, buffer + offset, sizeof(T));
        offset += sizeof(T);
        return result;
    }

    template <typename... Ts>
    std::tuple<Ts...> unpack(const char* buffer)
    {
        size_t offset = 0;
        return std::make_tuple(_unpack<Ts>(buffer, offset)...);
    }
    */

dllexport std::string unicode(const data_t& data, const std::string& encoding)
{
    UErrorCode status = U_ZERO_ERROR;
    UConverter* conv = ucnv_open(encoding.c_str(), &status);

    if (U_SUCCESS(status)) {
        int32_t utf8Size = ucnv_toUChars(conv, nullptr, 0, (const char*)data.data(), data.size(), &status);

        if (U_SUCCESS(status)) {
            std::string utf8Str(utf8Size, '\0');
            char* utf8Data = utf8Str.data();
            ucnv_toUChars(conv, reinterpret_cast<UChar*>(utf8Data), utf8Size, reinterpret_cast<const char*>(data.data()), data.size(), &status);

            ucnv_close(conv);
            return utf8Str;
        } else {
            std::cerr << "ucnv_toUChars failed: " << u_errorName(status) << std::endl;
        }

        ucnv_close(conv);
    } else {
        std::cerr << "ucnv_open failed: " << u_errorName(status) << std::endl;
    }

    return ""; // Return an empty string in case of failure
}

}