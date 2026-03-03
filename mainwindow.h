#ifndef MAINWINDOW_H
#define MAINWINDOW_H

/***************************************************************************************************
** 头文件依赖 (Header Dependencies)                                  **
***************************************************************************************************/
#include <QMainWindow>
#include <QtSql>

// --- 各功能模块的头文件 ---
#include "homepage.h"
#include "jobrequirement.h"
#include "declarationmanagement.h"
#include "dialogfinancial.h"
#include "employeemanagementwidget.h"
#include "salarymanagement.h"

// --- 全局工具类 ---
#include "usersession.h"
#include "loginandregister.h"


// UI类的标准前向声明
QT_BEGIN_NAMESPACE
namespace Ui {
class MainWindow;
}
QT_END_NAMESPACE

/***************************************************************************************************
** MainWindow 类定义                                           **
***************************************************************************************************/
/**
 * @class MainWindow
 * @brief 应用程序主窗口 (管理员视图)
 *
 * `MainWindow` 是管理员登录后看到的主操作界面。它作为一个容器，通过左侧的导航按钮
 * (QToolButton) 和右侧的 `QStackedWidget` 来切换和展示不同的功能模块。
 *
 * 主要职责:
 * 1.  **框架搭建:** 提供一个包含导航栏和内容显示区的基础布局。
 * 2.  **用户会话加载:** 在构造时，从全局的 `UserSession` 单例中加载当前登录管理员的信息，
 * 并可以利用这些信息更新UI（如显示欢迎语、姓名、头像等）。
 * 3.  **模块导航:** 响应导航按钮的点击事件，在 `QStackedWidget` 中创建并显示对应的
 * 功能模块窗口（如职位需求、员工管理、薪资管理等）。
 */
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    /**
     * @brief 构造函数
     * @param parent 父窗口指针, 默认为 nullptr。
     */
    MainWindow(QWidget *parent = nullptr);

    /**
     * @brief 析构函数
     */
    ~MainWindow();

    // --- 已废弃或由 UserSession 替代的旧方法 (注释保留作为参考) ---
    // void getMessage(QString username,int employeeID, ...);
    // void setLoggedInUser(const QString &username);
    // QString getLoggedInUser() const;

    // ===============================================================================================
    //                                       槽函数 (SLOTS)
    // ===============================================================================================
private slots:
    void on_tbtnJobRequirement_clicked();        // “职位需求” 导航按钮点击事件
    void on_tbtnPersonCenter_2_clicked();        // “个人主页” 导航按钮点击事件
    void on_tbtnDeclarationManagement_clicked(); // “申报审批” 导航按钮点击事件
    void on_tbtn_clicked();                      // “员工管理” 导航按钮点击事件 (按钮命名建议修改为 on_tbtnEmployeeManagement_clicked)
    void on_tbtnSalaryManagement_clicked();      // “薪资管理” 导航按钮点击事件


    // ===============================================================================================
    //                                   私有成员 (PRIVATE MEMBERS)
    // ===============================================================================================
private:
    Ui::MainWindow *ui; // 指向UI界面的指针

    // --- 当前登录用户的信息成员 (由 loadCurrentUserData 函数从 UserSession 初始化) ---
    QString username;
    int employeeID;
    QString department;
    QString employeeStatus;
    bool isPartyMember;
    QString workNature;
    QString position;
    QDate joinDate;
    QPixmap image;

    // --- 各功能模块的窗口实例指针 (按需创建) ---
    // 注意: 当前实现是在槽函数中创建局部实例，这些成员变量并未被使用，可以移除以简化代码
    HomePage *homepage;
    JobRequirement *jobRequirement;
    DeclarationManagement *declarationmanagement;
    EmployeeManagementWidget *employeemanagementwidget;
    // SalaryManagement *salarymanagement; // 注意，原始代码中为栈对象，非指针

    // ===============================================================================================
    //                                   私有函数 (PRIVATE FUNCTIONS)
    // ===============================================================================================
private:
    /**
     * @brief 加载当前登录用户的数据
     *
     * 从 UserSession 单例获取用户信息，并填充到 MainWindow 的成员变量中。
     * 未来可扩展此函数来更新UI上的用户信息显示。
     */
    void loadCurrentUserData();
};

#endif // MAINWINDOW_H
