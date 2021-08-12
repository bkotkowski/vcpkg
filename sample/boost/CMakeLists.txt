cmake_minimum_required(VERSION 3.14)
project(boost_sample LANGUAGES CXX)

set(CMAKE_INCLUDE_CURRENT_DIR ON)

find_package(Boost REQUIRED)

set(SOURCES
    main.cpp
)

add_executable(boost_sample ${SOURCES})
target_link_libraries(boost_sample PRIVATE Boost::headers)

