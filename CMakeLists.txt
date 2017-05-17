cmake_minimum_required(VERSION 3.7)

project(DDE_SAMPLE LANGUAGES C CXX)

add_definitions(
    -D_WIN32_WINNT=0x0601
    -D_CRT_SECURE_NO_WARNINGS
)

set(CMAKE_RUNTIME_OUTPUT_DIRECTORY bin)

add_subdirectory(DDEServ)
add_subdirectory(DDECli)