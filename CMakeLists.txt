cmake_minimum_required(VERSION 3.31)
project(AutoHotkey-ProcessRunner LANGUAGES CXX)
include(FetchContent)
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

find_package(Boost REQUIRED COMPONENTS utility locale system process)

file(GLOB_RECURSE src_files CONFIGURE_DEPENDS
  "src/AutoHotkey-ProcessRunner/*.h"
  "src/AutoHotkey-ProcessRunner/*.cpp"
  "src/AutoHotkey-ProcessRunner/*.rc")
add_library(AutoHotkey-ProcessRunner MODULE ${src_files})
target_include_directories(AutoHotkey-ProcessRunner PRIVATE src)
target_compile_features(AutoHotkey-ProcessRunner PUBLIC cxx_std_23)
target_compile_options(AutoHotkey-ProcessRunner PUBLIC
  /utf-8
  /external:anglebrackets
  /external:W0
  /WX
  /W4
  /analyze
  /analyze:external-)
target_compile_definitions(AutoHotkey-ProcessRunner PUBLIC
  UNICODE
  WIN32_LEAN_AND_MEAN
  VC_EXTRALEAN
  NOMINMAX
  _WINDLL
  $<$<CONFIG:Debug>:DEBUG>)
target_link_libraries(AutoHotkey-ProcessRunner PRIVATE Boost::utility Boost::locale Boost::system Boost::process)
target_precompile_headers(AutoHotkey-ProcessRunner PRIVATE src/AutoHotkey-ProcessRunner/PrecompiledHeader.h)

# Compile IDL file
find_program(MIDL midl.exe REQUIRED)
set(MIDL_OUT_DIR ${CMAKE_CURRENT_BINARY_DIR}/midl/AutoHotkey-ProcessRunner/Gen)
file(MAKE_DIRECTORY ${MIDL_OUT_DIR})
set(IDL_FILE ${CMAKE_CURRENT_SOURCE_DIR}/src/AutoHotkey-ProcessRunner/ProcessRunner.idl)
set(MIDL_SOURCES ${MIDL_OUT_DIR}/ProcessRunner_i.cpp ${MIDL_OUT_DIR}/ProcessRunner_i.h)
add_custom_command(
  OUTPUT ${MIDL_SOURCES}
  COMMAND ${MIDL}
          /nologo
          /char signed
          /env x64
          /target NT61
          /h ProcessRunner_i.h
          /iid ProcessRunner_i.cpp
          /out ${MIDL_OUT_DIR}
          ${IDL_FILE}
  DEPENDS ${IDL_FILE}
  WORKING_DIRECTORY ${MIDL_OUT_DIR}
  VERBATIM)
target_sources(AutoHotkey-ProcessRunner PRIVATE ${MIDL_SOURCES})
target_include_directories(AutoHotkey-ProcessRunner PRIVATE ${MIDL_OUT_DIR}/../..)
