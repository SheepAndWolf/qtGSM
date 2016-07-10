#-------------------------------------------------
#
# Project created by QtCreator 2016-07-06T17:10:11
#
#-------------------------------------------------

QT       += core gui sql axcontainer

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = GSM
TEMPLATE = app


SOURCES += main.cpp\
        mainwindow.cpp \
    opendb.cpp \
    datain.cpp \
    dataout.cpp \
    excel_rw.cpp

HEADERS  += mainwindow.h \
    opendb.h \
    datain.h \
    dataout.h \
    excel_rw.h

FORMS    += mainwindow.ui \
    datain.ui \
    dataout.ui
