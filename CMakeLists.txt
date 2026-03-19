cmake_minimum_required(VERSION 3.21)

set(MICMIX_VERSION_OVERRIDE "" CACHE STRING "Explicit MicMix version override (e.g. 1.2.9)")

set(MICMIX_VERSION "1.2.10")
if(MICMIX_VERSION_OVERRIDE)
  set(MICMIX_VERSION "${MICMIX_VERSION_OVERRIDE}")
elseif(EXISTS "${CMAKE_SOURCE_DIR}/VERSION")
  file(READ "${CMAKE_SOURCE_DIR}/VERSION" MICMIX_VERSION_RAW)
  string(STRIP "${MICMIX_VERSION_RAW}" MICMIX_VERSION)
endif()
# Re-run configure when VERSION changes so generated version header stays in sync.
set_property(DIRECTORY APPEND PROPERTY CMAKE_CONFIGURE_DEPENDS "${CMAKE_SOURCE_DIR}/VERSION")

project(MicMix VERSION ${MICMIX_VERSION} LANGUAGES C CXX)

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
  src/ui_shared.cpp
  src/settings_window.cpp
  src/effects_window.cpp
  resources/micmix.rc
)

configure_file(
  ${CMAKE_SOURCE_DIR}/cmake/version.h.in
  ${CMAKE_BINARY_DIR}/generated/micmix_version.h
  @ONLY
)

target_include_directories(micmix PRIVATE
  src
  ${CMAKE_BINARY_DIR}/generated
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

if(MSVC)
  target_compile_options(micmix PRIVATE
    /W4
    /permissive-
    /guard:cf
  )
  target_link_options(micmix PRIVATE
    /GUARD:CF
  )
endif()

target_link_libraries(micmix PRIVATE
  speexdsp_resampler
  ole32
  uuid
  avrt
  winmm
  mmdevapi
  propsys
  comctl32
  shlwapi
)

set_target_properties(micmix PROPERTIES
  OUTPUT_NAME "micmix"
)

add_executable(micmix_vst_host
  src/vst_host_main.cpp
)

target_compile_definitions(micmix_vst_host PRIVATE
  WIN32_LEAN_AND_MEAN
  NOMINMAX
  _CRT_SECURE_NO_WARNINGS
  UNICODE
  _UNICODE
)

if(MSVC)
  target_compile_options(micmix_vst_host PRIVATE
    /W4
    /permissive-
    /guard:cf
  )
  target_link_options(micmix_vst_host PRIVATE
    /GUARD:CF
  )
endif()

set_target_properties(micmix_vst_host PROPERTIES
  OUTPUT_NAME "micmix_vst_host"
)

option(MICMIX_BUILD_TESTS "Build MicMix automated tests" ON)
if(MICMIX_BUILD_TESTS)
  enable_testing()

  add_executable(micmix_config_store_tests
    tests/config_store_tests.cpp
    src/micmix_core.cpp
    src/ui_shared.cpp
    src/settings_window.cpp
    src/effects_window.cpp
  )

  target_include_directories(micmix_config_store_tests PRIVATE
    src
    ${CMAKE_BINARY_DIR}/generated
    ${TS3_SDK_DIR}/include
    ${TS3_SDK_DIR}/src
    ${SPEEXDSP_DIR}/include
  )

  target_compile_definitions(micmix_config_store_tests PRIVATE
    WIN32_LEAN_AND_MEAN
    NOMINMAX
    _CRT_SECURE_NO_WARNINGS
    UNICODE
    _UNICODE
  )

  if(MSVC)
    target_compile_options(micmix_config_store_tests PRIVATE
      /W4
      /permissive-
    )
  endif()

  target_link_libraries(micmix_config_store_tests PRIVATE
    speexdsp_resampler
    ole32
    uuid
    avrt
    winmm
    mmdevapi
    propsys
    comctl32
    shlwapi
    shell32
    uxtheme
  )

  add_test(NAME micmix_config_store_tests COMMAND micmix_config_store_tests)
endif()
