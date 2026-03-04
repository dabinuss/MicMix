cmake_minimum_required(VERSION 3.21)
project(MicMix LANGUAGES C CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

if(NOT WIN32)
  message(FATAL_ERROR "MicMix currently supports Windows only.")
endif()

if(MSVC)
  add_compile_options(/utf-8)
endif()

set(TS3_SDK_DIR "${CMAKE_SOURCE_DIR}/third_party/ts3_sdk")
set(SPEEXDSP_DIR "${CMAKE_SOURCE_DIR}/third_party/speexdsp")

if(NOT EXISTS "${TS3_SDK_DIR}/include/ts3_functions.h")
  message(FATAL_ERROR "Missing TeamSpeak 3 Plugin SDK at third_party/ts3_sdk")
endif()

if(NOT EXISTS "${SPEEXDSP_DIR}/libspeexdsp/resample.c")
  message(FATAL_ERROR "Missing speexdsp at third_party/speexdsp")
endif()

add_library(speexdsp_resampler STATIC
  ${SPEEXDSP_DIR}/libspeexdsp/resample.c
)

target_include_directories(speexdsp_resampler PUBLIC
  ${SPEEXDSP_DIR}/include
  ${SPEEXDSP_DIR}/libspeexdsp
  ${SPEEXDSP_DIR}/win32
)
target_compile_definitions(speexdsp_resampler PRIVATE
  HAVE_CONFIG_H
)

add_library(micmix SHARED
  src/plugin_main.cpp
  src/micmix_core.cpp
  src/settings_window.cpp
)

target_include_directories(micmix PRIVATE
  src
  ${TS3_SDK_DIR}/include
  ${TS3_SDK_DIR}/src
  ${SPEEXDSP_DIR}/include
)

target_compile_definitions(micmix PRIVATE
  WIN32_LEAN_AND_MEAN
  NOMINMAX
  _CRT_SECURE_NO_WARNINGS
  UNICODE
  _UNICODE
  BUILDING_DLL
)

target_link_libraries(micmix PRIVATE
  speexdsp_resampler
  ole32
  uuid
  avrt
  mmdevapi
  propsys
  comctl32
  shlwapi
)

set_target_properties(micmix PROPERTIES
  OUTPUT_NAME "micmix"
)
