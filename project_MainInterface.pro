QT       += core gui
QT       +=  sql
QT += axcontainer
QT += quick qml
CONFIG += wasm
greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++17

# You can make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

SOURCES += \
    databasemanager.cpp \
    declarationmanagement.cpp \
    employeeeditdialog.cpp \
    employeemanagementwidget.cpp \
    homepage.cpp \
    intervieweedialog.cpp \
    jobrequirement.cpp \
    loginandregister.cpp \
    main.cpp \
    mainwindow.cpp \
    personalapplication.cpp \
    personalinterface.cpp \
    salarymanagement.cpp \
    statusdelegate.cpp \
    taskcardwidget.cpp \
    taskeditdialog.cpp \
    usersession.cpp

HEADERS += \
    databasemanager.h \
    declarationmanagement.h \
    employeeeditdialog.h \
    employeemanagementwidget.h \
    homepage.h \
    intervieweedialog.h \
    jobrequirement.h \
    loginandregister.h \
    mainwindow.h \
    personalapplication.h \
    personalinterface.h \
    salarymanagement.h \
    statusdelegate.h \
    taskcardwidget.h \
    taskeditdialog.h \
    usersession.h

FORMS += \
    declarationmanagement.ui \
    employeeeditdialog.ui \
    employeemanagementwidget.ui \
    homepage.ui \
    intervieweedialog.ui \
    jobrequirement.ui \
    loginandregister.ui \
    mainwindow.ui \
    personalapplication.ui \
    personalinterface.ui \
    salarymanagement.ui \
    taskcardwidget.ui \
    taskeditdialog.ui

# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target

RESOURCES += \
    res.qrc
