
# find_package(Python3 COMPONENTS Development REQUIRED)
# target_link_libraries(main PRIVATE Python3::Python)
#
# find_package(pybind11 CONFIG REQUIRED)
# target_link_libraries(main PRIVATE mylib pybind11::lto pybind11::embed pybind11::module)
#
# 

cmake_minimum_required(VERSION 3.14)
project(http LANGUAGES CXX)

set(CMAKE_INCLUDE_CURRENT_DIR ON)

set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)

if(NOT DEFINED INSTALL_EXAMPLESDIR)
  set(INSTALL_EXAMPLESDIR "examples")
endif()

set(INSTALL_EXAMPLEDIR "${INSTALL_EXAMPLESDIR}/network/http")

find_package(Python3 COMPONENTS Development REQUIRED)
find_package(pybind11 CONFIG REQUIRED)
find_package(Qt5 COMPONENTS Core Gui Network Widgets CONFIG REQUIRED)

set(SOURCES
    authenticationdialog.ui
    httpwindow.cpp httpwindow.h
    main.cpp
    http_py11.cpp
)

Python3_add_library(http_py MODULE ${SOURCES})

target_link_libraries(http_py PUBLIC
    Qt::Core Qt::Gui Qt::Network Qt::Widgets
    pybind11::lto pybind11::embed pybind11::module
)

