cmake_minimum_required(VERSION 3.14)

project(Excelx_Demo LANGUAGES CXX)

IF (CMAKE_SYSTEM_NAME MATCHES "Darwin") 
  set(CMAKE_OSX_ARCHITECTURES arm64)
ENDIF (CMAKE_SYSTEM_NAME MATCHES "Darwin")

set(CMAKE_INCLUDE_CURRENT_DIR ON)

set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(XML_MODEL_SOURCE_DIR ${CMAKE_CURRENT_SOURCE_DIR}/excelx/xml_models)
set(XML_MODEL_BINARY_DIR ${CMAKE_CURRENT_BINARY_DIR}/xml_models)

file(
    GLOB
    HEADER_DIR_SET
    LIST_DIRECTORIES true
    ${CMAKE_CURRENT_SOURCE_DIR}/excelx/include/*
)

execute_process(COMMAND ${CMAKE_COMMAND} -E copy_directory  "${XML_MODEL_SOURCE_DIR}" "${XML_MODEL_BINARY_DIR}")

find_package(QT NAMES Qt6 Qt5 COMPONENTS Core Gui Widgets REQUIRED)
find_package(Qt${QT_VERSION_MAJOR} COMPONENTS Core Gui Widgets REQUIRED)
find_package(OpenCV REQUIRED)

include_directories(${HEADER_DIR_SET})
include_directories(${OpenCV_INCLUDE_DIRS})
add_subdirectory(excelx)

add_executable(Excelx_Demo
  excelx_demo.cpp
)

# target_include_directories(Excelx_Demo PRIVATE ${XML_DIR_PATH}/include)
# target_include_directories(Excelx_Demo PRIVATE ${EXCEL_DIR_PATH}/include)
# target_include_directories(Excelx_Demo PRIVATE ${PLUGINS_DIR_PATH}/include)

link_directories(${OpenCV_LIBRARY_DIRS})

target_link_libraries(Excelx_Demo Excelx)



