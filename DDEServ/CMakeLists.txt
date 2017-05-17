cmake_minimum_required(VERSION 3.7)

set(SRC
    datadlg.cpp
    DDEServ.cpp
    DDESvdoc.cpp
    DDESvvw.cpp
    mainfrm.cpp
    myserv.cpp
    stdafx.cpp

    datadlg.h
    DDEServ.h
    DDESvdoc.h
    DDESvvw.h
    mainfrm.h
    myserv.h
    resource.h
    stdafx.h

    ../StdDDE/stddde.h
    ../StdDDE/stddde.cpp

    DDEServ.rc
    res/DDEServ.ico
    res/DDEServ.rc2
    res/toolbar.bmp
)

set(CMAKE_MFC_FLAG 2)

add_executable(DDEServ WIN32 ${SRC})
target_include_directories(DDEServ PRIVATE .)
target_compile_definitions(DDEServ PRIVATE _AFXDLL)

add_executable(DDEServU WIN32 ${SRC})
target_include_directories(DDEServU PRIVATE .)
target_compile_definitions(DDEServU PRIVATE _AFXDLL _UNICODE)