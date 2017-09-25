#-------------------------------------------------
#
# Project created by QtCreator 2017-05-03T13:56:31
#
#-------------------------------------------------

QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = DoQueries
TEMPLATE = app


SOURCES += main.cpp\
        mainwindow.cpp \
    mail.cpp

HEADERS  += mainwindow.h \
    mail.h \
    MAPIAux.h \
    MAPICode.h \
    MAPIDbg.h \
    MAPIDefS.h \
    MAPIForm.h \
    MAPIGuid.h \
    MAPIHook.h \
    MAPINls.h \
    MAPIOID.h \
    MAPISPI.h \
    MAPITags.h \
    MAPIUtil.h \
    MAPIVal.h \
    MAPIWin.h \
    MAPIWz.h \
    MAPIX.h \
    MSPST.h

FORMS    += mainwindow.ui
