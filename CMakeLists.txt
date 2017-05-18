cmake_minimum_required(VERSION 3.7)

project(DDE_SAMPLE LANGUAGES CXX)

set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}/bin)

option(USE_UNICODE "Use UNICODE character set." ON)

if(USE_UNICODE)
    add_definitions(-D_UNICODE)
endif()

add_subdirectory(DDEServ)
add_subdirectory(DDECli)
