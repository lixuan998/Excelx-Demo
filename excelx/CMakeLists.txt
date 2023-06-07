cmake_minimum_required(VERSION 3.14)
project(Excelx LANGUAGES CXX)

include_directories(zipper)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

file(
    GLOB
    HEADER_DIR_SET
    LIST_DIRECTORIES true
    ${CMAKE_CURRENT_SOURCE_DIR}/include/*
)

file(
    GLOB_RECURSE
    HEADER_FILE_SET
    ${CMAKE_CURRENT_SOURCE_DIR}/include/*
)

file(
    GLOB_RECURSE
    SOURCE_FILE_SET
    ${CMAKE_CURRENT_SOURCE_DIR}/src/*
)

set(ZIPPER_HEADER_DIR ${CMAKE_CURRENT_SOURCE_DIR}/zipper/zipper)

find_package(Qt5 COMPONENTS Core Gui Widgets REQUIRED)
find_package(OpenCV REQUIRED)

include_directories(${HEADER_DIR_SET})
include_directories(${ZIPPER_HEADER_DIR})
include_directories(${CMAKE_CURR})

include_directories(${OpenCV_INCLUDE_DIRS})

add_subdirectory(zipper)

add_library(Excelx STATIC

${HEADER_FILE_SET}
${SOURCE_FILE_SET}
)

target_link_libraries(Excelx Qt5::Core Qt5::Gui Qt5::Widgets ${OpenCV_LIBS} zipper)
