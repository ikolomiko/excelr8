#pragma once

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

using data_t = std::vector<std::byte>;

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
using pytype_n = ssize_t;
using pytype_N = size_t;
// skipping pytype_e
using pytype_f = float;
using pytype_d = double;
// skipping the rest

namespace excelr8::util {

template <typename T>
T unpack(const char* buffer, size_t& offset);

template <typename... Ts>
std::tuple<Ts...> unpack(const char* buffer);

template <typename... Ts>
std::tuple<Ts...> unpack(const data_t& buffer);

template <typename T>
std::vector<T> vslice(const std::vector<T>& v, int start, int end);

std::string unicode(const data_t&, const std::string& encoding);

}