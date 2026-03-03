#include <QCoreApplication>
#include <QDebug>
#include <QPluginLoader>
#include <QSql>
#include <QSqlDatabase>
//#include "mainwindow.h"
#include "loginandregister.h"
#include <QFormLayout>
#include <QApplication>
#include <QMessageBox>
#include "databasemanager.h"

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    // 使用DatabaseManager来打开数据库
    if (!DatabaseManager::instance()->openDatabase("localhost", 3306, "fengmeimei", "root", "zgc.3264746822")) {
        QMessageBox::critical(nullptr, "数据库错误", "无法连接到数据库，应用程序即将退出。");
        return -1;
    }

    LoginANDRegister *logandregister = new LoginANDRegister;
    logandregister->show();

    int result = a.exec();

    // 应用程序退出前，可以考虑关闭数据库（虽然析构函数会自动处理）
    DatabaseManager::instance()->closeDatabase();

    return result;
}

