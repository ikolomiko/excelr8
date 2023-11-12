cmake_minimum_required(VERSION 3.15)

project(excelr8 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

include_directories("include")

find_package(ICU 66.1 REQUIRED COMPONENTS uc dt)

file(GLOB SRC_FILES src/*.cpp)
add_library(excelr8 SHARED
    ${SRC_FILES}
)

set_target_properties(excelr8 PROPERTIES
  POSITION_INDEPENDENT_CODE ON
  RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib
  EXPORTED_TARGETS excelr8
)

target_include_directories("excelr8" INTERFACE "include")
target_include_directories("excelr8" PRIVATE ${ICU_INCLUDE_DIRS})
target_link_libraries("excelr8" PRIVATE ${ICU_LIBRARIES})
