#include "loginandregister.h"
#include "ui_loginandregister.h"
#include "mainwindow.h"         // 管理员主窗口
#include "personalinterface.h"  // 员工个人界面
#include "usersession.h"        // 用于存储全局用户信息的单例类
#include "databasemanager.h"    // <-- 引入DatabaseManager

#include <QMovie>               // 如果UI中有GIF动图，会需要这个
#include <QMessageBox>          // 用于显示提示、警告和错误信息框
#include <QSqlRecord>           // 用于接收返回的用户信息

/***************************************************************************************************
** 构造函数与析构函数                                            **
***************************************************************************************************/

/**
 * @brief LoginANDRegister 构造函数
 * @param parent 父窗口指针
 */
LoginANDRegister::LoginANDRegister(QWidget *parent)
    : QWidget(parent)
    , ui(new Ui::LoginANDRegister)
{
    ui->setupUi(this);
    // 可以在这里进行额外的UI初始化，例如设置窗口标题、加载GIF等
}

/**
 * @brief LoginANDRegister 析构函数
 */
LoginANDRegister::~LoginANDRegister()
{
    delete ui;
}

/***************************************************************************************************
** 槽函数实现 (Slot Implementations)                                **
***************************************************************************************************/

/**
 * @brief "注册"按钮的槽函数
 *
 * 当前的实现非常简单，只是弹出一个消息框，提示用户联系管理员进行账号注册。
 */
void LoginANDRegister::on_btnRegister_clicked()
{
    QMessageBox::information(nullptr, "注册？", "如有账号问题，请联系管理员！");
}

/**
 * @brief "登录"按钮的槽函数
 *
 * 这是登录窗口的核心，执行一整套验证流程：
 * 1.  获取用户输入的账号、密码和选择的身份。
 * 2.  检查输入是否为空，以及是否选择了身份。
 * 3.  **重构核心:** 调用 DatabaseManager 进行统一的验证。
 * 4.  **登录成功处理:** 如果所有验证都通过：
 * a. 使用 DatabaseManager 从数据库中提取该员工的全部信息，并填充到 UserSession 中。
 * b. 根据员工类型，创建并显示对应的窗口（`MainWindow` 或 `PersonalInterface`）。
 * c. 关闭当前的登录窗口。
 * 5.  在验证过程的任何一步失败，都会弹出相应的错误提示。
 */
void LoginANDRegister::on_btnLogin_clicked()
{
    // 1. 获取用户输入 (逻辑不变)
    QString username = ui->usernameEdit->text().trimmed();
    QString password = ui->passwordEdit->text().trimmed();

    if (username.isEmpty() || password.isEmpty()) {
        QMessageBox::warning(this, "登录提示", "用户名和密码不能为空！");
        return;
    }

    QString selectedRole;
    if (ui->rbtnEmployee && ui->rbtnEmployee->isChecked()) {
        selectedRole = "员工";
    } else if (ui->rbtnManager && ui->rbtnManager->isChecked()) {
        selectedRole = "管理员";
    }
   /* else if (ui->rbtnBoss && ui->rbtnBoss->isChecked()) {
        selectedRole = "高管";
    } */
    else {
        QMessageBox::warning(this, "登录提示", "请选择您的登录身份！");
        return;
    }

    // 2. *** 重构核心：调用DatabaseManager进行验证 ***
    QSqlRecord userInfo; // 用于接收返回的用户信息
    bool loginSuccess = DatabaseManager::instance()->validateUser(username, password, selectedRole, userInfo);

    if (loginSuccess) {
        // 3. 验证成功，填充全局会话并打开主窗口
        QMessageBox::information(this, "登录成功", "欢迎回来，" + userInfo.value("Name").toString() + "！");
        this->close();

        // 使用 DatabaseManager 填充 UserSession
        DatabaseManager::instance()->fetchEmployeeDetails(username, UserSession::instance());

        // 4. 根据身份打开不同窗口
        if (selectedRole == "管理员" || selectedRole == "高管") {
            MainWindow *w = new MainWindow;
            w->show();
        } else { // "员工"
            PersonalInterface *pi = new PersonalInterface;
            pi->show();
        }
    } else {
        // 5. 登录失败
        QMessageBox::warning(this, "登录失败", "用户名、密码或选择的身份不正确！");
    }
}
