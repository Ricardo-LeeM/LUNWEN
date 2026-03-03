#ifndef LOGINANDREGISTER_H
#define LOGINANDREGISTER_H

/***************************************************************************************************
** 头文件依赖 (Header Dependencies)                                  **
***************************************************************************************************/
#include <QWidget>
#include <QtSql> // 包含Qt SQL模块的基本功能，如 QSqlQuery

// 前向声明，避免在头文件中包含完整的 personalinterface.h
class PersonalInterface;

// UI类的标准前向声明
namespace Ui {
class LoginANDRegister;
}

/***************************************************************************************************
** LoginANDRegister 类定义                                     **
***************************************************************************************************/
/**
 * @class LoginANDRegister
 * @brief 登录与注册窗口
 *
 * 该类提供了应用程序的用户入口界面，负责处理用户的登录和注册请求。
 * 主要功能包括：
 * 1.  提供用户名、密码输入框和身份选择（员工、管理员、高管）的UI。
 * 2.  验证用户输入的凭据是否正确。
 * 3.  根据用户的身份，在登录成功后，打开对应的应用程序主窗口（管理员界面或员工个人界面）。
 * 4.  将登录成功的用户信息存储在全局的 UserSession 单例中，供其他模块使用。
 * 5.  提供一个（当前被禁用的）注册入口。
 */
class LoginANDRegister : public QWidget
{
    Q_OBJECT

public:
    /**
     * @brief 构造函数
     * @param parent 父窗口指针，默认为 nullptr。
     */
    explicit LoginANDRegister(QWidget *parent = nullptr);

    /**
     * @brief 析构函数
     */
    ~LoginANDRegister();

    // ===============================================================================================
    //                                       槽函数 (SLOTS)
    // ===============================================================================================
private slots:
    /**
     * @brief "注册"按钮点击事件的槽函数。
     * 当前实现仅为提示用户联系管理员。
     */
    void on_btnRegister_clicked();

    /**
     * @brief "登录"按钮点击事件的槽函数。
     * 这是本类的核心功能，负责处理完整的登录验证逻辑。
     */
    void on_btnLogin_clicked();

    // --- 注释掉的旧信号槽，可用于未来重构或参考 ---
    // void openWindow(QString username);
    // void onLoginSuccess(const QString &username);

    // ===============================================================================================
    //                                   私有成员 (PRIVATE MEMBERS)
    // ===============================================================================================
private:
    Ui::LoginANDRegister *ui; // 指向UI界面的指针

    // 指向员工个人界面的指针，根据当前设计，此成员可能不是必需的，
    // 因为窗口是在登录成功时局部创建的。
    PersonalInterface *personalinterface;
};

#endif // LOGINANDREGISTER_H
