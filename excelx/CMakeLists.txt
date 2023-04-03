cmake_minimum_required(VERSION 3.14)
project(Excelx LANGUAGES CXX)

include_directories(zipper)

set(CMAKE_INCLUDE_CURRENT_DIR ON)
set(XML_DIR_PATH ${CMAKE_CURRENT_SOURCE_DIR}/XML)
set(EXCEL_DIR_PATH ${CMAKE_CURRENT_SOURCE_DIR}/EXCEL)
set(PLUGINS_DIR_PATH ${CMAKE_CURRENT_SOURCE_DIR}/PLUGINS)

aux_source_directory(${XML_DIR_PATH}/src XML_SOURCES)
aux_source_directory(${EXCEL_DIR_PATH}/src EXCEL_SOURCES)
aux_source_directory(${PLUGINS_DIR_PATH}/src PLUGINS_SOURCES)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(Qt5 COMPONENTS Core Gui Widgets REQUIRED)
find_package(OpenCV REQUIRED)

include_directories(${OpenCV_INCLUDE_DIRS})

add_subdirectory(zipper)

add_library(Excelx STATIC

${EXCEL_SOURCES}
${XML_SOURCES}
${PLUGINS_SOURCES}

)

target_link_libraries(Excelx Qt5::Core Qt5::Gui Qt5::Widgets ${OpenCV_LIBS} zipper)
