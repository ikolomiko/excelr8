project('excelr8', 'cpp',
    version : '0.1',
    default_options : ['warning_level=3', 'cpp_std=c++20']
)

# These arguments are only used to build the shared library
# not the executables that use the library.
lib_args = ['-DBUILDING_EXCELR8', '-Wno-sign-compare']

cpp_files = files(
    'src/biff.cpp',
    'src/excelr8.cpp',
    'src/util.cpp',
    'src/data.cpp',
    'src/compdoc.cpp',
)

incl_dir = include_directories('include')

icu_uc_dep = dependency('icu-uc')

shlib = shared_library('excelr8', 
    cpp_files,
    install : true,
    cpp_args : lib_args,
    gnu_symbol_visibility : 'default',
    include_directories: incl_dir,
    dependencies: icu_uc_dep,
)

# Make this library usable as a Meson subproject.
excelr8_dep = declare_dependency(
    include_directories: incl_dir,
    link_with : shlib
)

# Make this library usable from the system's
# package manager.
install_headers('include/excelr8/excelr8.hpp', subdir : 'excelr8')

pkg_mod = import('pkgconfig')
pkg_mod.generate(
    name : 'excelr8',
    filebase : 'excelr8',
    description : '🚀 Blazing fast C++ library for reading old Excel (.xls) files.',
    subdirs : 'excelr8',
    libraries : shlib,
    version : '0.1',
)
