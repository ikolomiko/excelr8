cmake_minimum_required(VERSION 3.15)

project(excelr8 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(ICU 66.1 REQUIRED COMPONENTS uc dt)
include_directories("include")
include_directories(${ICU_INCLUDE_DIRS})

add_compile_definitions(BUILDING_EXCELR8)

file(GLOB SRC_FILES src/*.cpp)
file(GLOB HPP_FILES src/*.hpp)
add_library(excelr8
    SHARED
    ${SRC_FILES}
    ${HPP_FILES}
)

target_compile_options("excelr8" PRIVATE
    -Wall
    -Wextra
    -Wpedantic
    -Wno-sign-compare
)

set_target_properties(excelr8 PROPERTIES
  POSITION_INDEPENDENT_CODE ON
  RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}
  EXPORTED_TARGETS excelr8
)

target_include_directories("excelr8" INTERFACE "include")
target_link_libraries("excelr8" PRIVATE ${ICU_LIBRARIES})
