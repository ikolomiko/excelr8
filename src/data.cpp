#include "excelr8/data.hpp"
#include <cstddef>
#include <cstring>
#include <vector>

namespace excelr8 {

data_t::data_t() { }

data_t::data_t(std::vector<std::byte>& buffer)
    : _data(buffer)
{
}

data_t::data_t(const std::string& buffer)
{
    _data = std::vector<std::byte>(buffer.size());
    std::memcpy(_data.data(), buffer.data(), buffer.size());
}

data_t data_t::slice(int start, int end) const
{
    std::vector<std::byte> vec = { _data.begin() + start, _data.begin() + end + 1 };
    return { vec };
}

template <typename T>
T data_t::_unpack(size_t& offset) const
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

template <typename T>
std::vector<T> data_t::unpack_vec(size_t count) const
{
    size_t offset = 0;
    std::vector<T> result(count);
    for (size_t i = 0; i < count; i++) {
        T item = _unpack<T>(offset);
        result.push_back(item);
    }
    return result;
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

bool data_t::operator==(const data_t& other) const
{
    return _data == other._data;
}

bool data_t::operator!=(const data_t& other) const
{
    return _data != other._data;
}

inline data_t& data_t::operator+=(const data_t& other)
{
    _data.insert(_data.end(), other._data.begin(), other._data.end());
    return *this;
}

data_t& data_t::append(const data_t& other)
{
    *this += other;
    return *this;
}

std::string data_t::to_string() const
{
    return std::string(reinterpret_cast<const char*>(_data.data()), _data.size());
}

template pytype_B data_t::_unpack<pytype_B>(size_t&) const;
template pytype_H data_t::_unpack<pytype_H>(size_t&) const;
template pytype_i data_t::_unpack<pytype_i>(size_t&) const;

template std::tuple<pytype_B> data_t::unpack<pytype_B>(size_t) const;
template std::tuple<pytype_H> data_t::unpack<pytype_H>(size_t) const;
template std::tuple<pytype_i> data_t::unpack<pytype_i>(size_t) const;
template std::tuple<pytype_H, pytype_H, pytype_B, pytype_B> data_t::unpack<pytype_H, pytype_H, pytype_B, pytype_B>(size_t) const;
template std::tuple<pytype_H, pytype_H, pytype_H, pytype_H> data_t::unpack<pytype_H, pytype_H, pytype_H, pytype_H>(size_t) const;
template std::tuple<pytype_H, pytype_H> data_t::unpack<pytype_H, pytype_H>(size_t) const;
template std::tuple<pytype_i, pytype_i> data_t::unpack<pytype_i, pytype_i>(size_t) const;
template std::tuple<pytype_H, pytype_B, pytype_B, pytype_i, pytype_i, pytype_i> data_t::unpack<pytype_H, pytype_B, pytype_B, pytype_i, pytype_i, pytype_i>(size_t) const;
template std::tuple<pytype_I, pytype_I, pytype_I, pytype_I> data_t::unpack<pytype_I, pytype_I, pytype_I, pytype_I>(size_t) const;

template std::vector<pytype_i> data_t::unpack_vec<pytype_i>(size_t) const;

}