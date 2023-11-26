#include "excelr8/data.hpp"
#include <cstring>
#include <vector>

namespace excelr8 {
data_t::data_t(std::vector<std::byte>& buffer)
    : _data(buffer)
{
}

data_t data_t::slice(int start, int end) const
{
    std::vector<std::byte> vec = { _data.begin() + start, _data.begin() + end + 1 };
    return { vec };
}

template <typename T>
T data_t::_unpack(size_t offset) const
{
    T result;
    const char* buffer = (const char*)(_data.data());
    std::memcpy(&result, buffer + offset, sizeof(T));
    offset += sizeof(T);
    return result;
}

template <typename... Ts>
std::tuple<Ts...> data_t::unpack(size_t offset) const
{
    return std::make_tuple(_unpack<Ts>(offset)...);
}

const std::byte* data_t::data() const
{
    return _data.data();
}

size_t data_t::size() const
{
    return _data.size();
}

std::vector<std::byte>::const_iterator data_t::begin() const
{
    return _data.begin();
}

std::vector<std::byte>::const_iterator data_t::end() const
{
    return _data.end();
}

template pytype_B data_t::_unpack<pytype_B>(size_t offset) const;
template pytype_H data_t::_unpack<pytype_H>(size_t offset) const;
template pytype_i data_t::_unpack<pytype_i>(size_t offset) const;

template std::tuple<pytype_B> data_t::unpack<pytype_B>(size_t) const;
template std::tuple<pytype_H> data_t::unpack<pytype_H>(size_t) const;
template std::tuple<pytype_i> data_t::unpack<pytype_i>(size_t) const;
template std::tuple<pytype_H, pytype_H, pytype_B, pytype_B> data_t::unpack<pytype_H, pytype_H, pytype_B, pytype_B>(size_t) const;
template std::tuple<pytype_H, pytype_H, pytype_H, pytype_H> data_t::unpack<pytype_H, pytype_H, pytype_H, pytype_H>(size_t) const;
template std::tuple<pytype_H, pytype_H> data_t::unpack<pytype_H, pytype_H>(size_t) const;

}