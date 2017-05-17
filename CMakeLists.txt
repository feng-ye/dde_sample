cmake_minimum_required(VERSION 3.7)

project(DDE_SAMPLE LANGUAGES CXX)

set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}/bin)

add_subdirectory(DDEServ)
add_subdirectory(DDECli)