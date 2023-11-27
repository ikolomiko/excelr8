#pragma once

#if defined _WIN32 || defined __CYGWIN__
#ifdef BUILDING_EXCELR8
#define dllexport __declspec(dllexport)
#else
#define dllexport __declspec(dllimport)
#endif
#else
#ifdef BUILDING_EXCELR8
#define dllexport __attribute__((visibility("default")))
#else
#define dllexport
#endif
#endif

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

// python struct type aliases
// taken from https://docs.python.org/3/library/struct.html
using pytype_c = char;
using pytype_b = signed char;
using pytype_B = unsigned char;
using pytype_questionmark = bool;
using pytype_h = int16_t; // aka short
using pytype_H = uint16_t; // aka ushort
using pytype_i = int32_t; // aka int
using pytype_I = uint32_t; // aka uint
using pytype_l = int32_t; // aka long (but 4 bytes)
using pytype_L = uint32_t; // aka ulong (but 4 bytes)
using pytype_q = int64_t; // aka long long (8 bytes)
using pytype_Q = uint64_t; // aka unsigned long long (8 bytes)
// using pytype_n = ssize_t;
using pytype_N = size_t;
// skipping pytype_e
using pytype_f = float;
using pytype_d = double;
// skipping the rest

namespace excelr8 {
class dllexport data_t {
private:
    std::vector<std::byte> _data;

    template <typename T>
    T _unpack(size_t& offset) const;

public:
    data_t(std::vector<std::byte>& buffer);
    data_t(const std::string& buffer);
    data_t slice(int start, int end) const;

    template <typename... Ts>
    std::tuple<Ts...> unpack(size_t offset = 0) const;

    template <typename T>
    std::vector<T> unpack_vec(size_t count) const;

    const std::byte* data() const;
    size_t size() const;
    std::vector<std::byte>::const_iterator begin() const;
    std::vector<std::byte>::const_iterator end() const;
    bool operator==(const data_t& other) const;
    bool operator!=(const data_t& other) const;
    std::string to_string() const;
};
}