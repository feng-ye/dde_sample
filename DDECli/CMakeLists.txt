cmake_minimum_required(VERSION 3.7)

set(SRC
    DDECldlg.cpp
    DDECldlg.h
    DDECli.cpp
    DDECli.h
    myclient.cpp
    myclient.h
    myconv.cpp
    myconv.h
    resource.h
    stdafx.cpp
    stdafx.h

    ../StdDDE/stddde.h
    ../StdDDE/stddde.cpp

    DDECli.rc
    res/DDECli.ico
    res/DDECli.rc2
)

set(CMAKE_MFC_FLAG 2)

add_executable(DDECli WIN32 ${SRC})
target_include_directories(DDECli PRIVATE .)
target_compile_definitions(DDECli PRIVATE _AFXDLL)

add_executable(DDECliU WIN32 ${SRC})
target_include_directories(DDECliU PRIVATE .)
target_compile_definitions(DDECliU PRIVATE _AFXDLL _UNICODE)