#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "homepage.h" // 确保 HomePage 头文件被包含，尽管它已在 mainwindow.h 中

#include <QMenu>
#include <QMessageBox>
#include <QSqlQuery>


/***************************************************************************************************
** 构造函数与析构函数                                            **
***************************************************************************************************/

/**
 * @brief MainWindow 构造函数
 *
 * 负责初始化UI，并立即调用 `loadCurrentUserData()` 从全局会话中加载登录用户信息。
 * 这是确保窗口一显示就拥有正确用户上下文的关键步骤。
 * @param parent 父窗口指针
 */
MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    // 从 UserSession 加载当前登录用户的数据
    loadCurrentUserData();

    // 默认显示主页
    on_tbtnPersonCenter_2_clicked();

    // --- 已注释的旧代码，用于创建下拉菜单，可作为未来功能参考 ---
    // QMenu *menuSel = new QMenu(this);
    // menuSel->addAction(ui->actionApplication);
    // menuSel->addAction(ui->actionBasicInfo);
    // ui->tbtnPersonCenter->setMenu(menuSel);
}

/**
 * @brief MainWindow 析构函数
 */
MainWindow::~MainWindow()
{
    delete ui;
}


/***************************************************************************************************
** 私有函数实现 (Private Function Implementations)                       **
***************************************************************************************************/

/**
 * @brief 从 UserSession 加载当前登录用户的数据
 *
 * 该函数是 MainWindow 初始化逻辑的一部分。它访问 `UserSession` 单例，
 * 检查是否存在有效的用户会话（通过 employeeID 是否有效来判断）。
 * 如果用户已登录，则将其所有信息（姓名、部门、职位、头像等）
 * 复制到 MainWindow 的同名私有成员变量中。
 * 这使得在 MainWindow 的生命周期内，可以方便地访问这些信息，例如传递给子窗口。
 *
 * @note 如果用户会话无效（理论上不应发生，因为 MainWindow 应在成功登录后才打开），
 * 则会弹出一个严重错误提示。
 */
void MainWindow::loadCurrentUserData()
{
    UserSession *session = UserSession::instance();

    // 检查用户是否已有效登录
    if (session->employeeID() != -1) {
        // 从会话中获取信息并填充到本类的成员变量中
        this->username = session->name();
        this->employeeID = session->employeeID();
        this->department = session->department();
        this->employeeStatus = session->employeeStatus();
        this->isPartyMember = session->isPartyMember();
        this->workNature = session->workNature();
        this->position = session->position();
        this->joinDate = session->joinDate();
        this->image = session->userPixmap(); // 获取原始QPixmap

        // --- UI更新示例 (已注释) ---
        // 在这里，你可以使用加载的数据来更新主窗口界面上的元素
        // 例如，显示欢迎信息、用户名、部门和头像等
        // ui->welcomeLabel->setText("欢迎, " + this->username);
        // ui->avatarLabel->setPixmap(this->image.scaled(100, 100, Qt::KeepAspectRatio));

    } else {
        // 理论上不应该出现此情况，因为MainWindow是在成功登录后才被创建的
        QMessageBox::critical(this, "严重错误", "无法加载用户信息，用户会话无效！\n应用程序可能无法正常工作。");
        // 在这里可以考虑关闭窗口或强制返回登录界面
        // close();
    }
}


/***************************************************************************************************
** 槽函数实现 (Slot Implementations) - 导航栏按钮                       **
***************************************************************************************************/

/**
 * @brief "个人主页" 导航按钮的槽函数
 *
 * 清空 QStackedWidget，创建一个新的 HomePage 实例，并将其设置为当前显示的页面。
 * 由于 HomePage 需要显示用户信息，而这些信息已通过 `loadCurrentUserData` 加载到
 * MainWindow 的成员变量中，所以理论上可以直接访问，或通过 UserSession 获取。
 *
 * @note 旧的 `getMessage` 调用已被注释，因为推荐的做法是让 HomePage
 * 自己从 UserSession 中获取数据，以实现更好的解耦。
 */
void MainWindow::on_tbtnPersonCenter_2_clicked()
{
    // 每次点击都创建一个新的HomePage实例并显示
    HomePage *homepage = new HomePage(this);
    // homepage->getMessage(...); // 旧的传参方式，现在 HomePage 应该直接从 UserSession 读取数据
    ui->stackedWidget->addWidget(homepage);
    ui->stackedWidget->setCurrentWidget(homepage);
}

/**
 * @brief "职位需求" 导航按钮的槽函数
 *
 * 创建一个新的 JobRequirement 实例，并将其添加到 QStackedWidget 中作为当前页面。
 */
void MainWindow::on_tbtnJobRequirement_clicked()
{
    JobRequirement *jobRequirement = new JobRequirement;
    ui->stackedWidget->addWidget(jobRequirement);
    ui->stackedWidget->setCurrentWidget(jobRequirement);
}

/**
 * @brief "申报审批" 导航按钮的槽函数
 *
 * 创建一个新的 DeclarationManagement 实例，并将其添加到 QStackedWidget 中作为当前页面。
 */
void MainWindow::on_tbtnDeclarationManagement_clicked()
{
    DeclarationManagement *declarationManagement = new DeclarationManagement;
    ui->stackedWidget->addWidget(declarationManagement);
    ui->stackedWidget->setCurrentWidget(declarationManagement);
}

/**
 * @brief "员工管理" 导航按钮的槽函数
 *
 * 创建一个新的 EmployeeManagementWidget 实例，并将其添加到 QStackedWidget 中作为当前页面。
 * @note 建议在UI设计器中将此按钮的对象名(objectName)修改为 `tbtnEmployeeManagement` 以提高代码可读性。
 */
void MainWindow::on_tbtn_clicked()
{
    EmployeeManagementWidget *employeemanagementwidget = new EmployeeManagementWidget;
    ui->stackedWidget->addWidget(employeemanagementwidget);
    ui->stackedWidget->setCurrentWidget(employeemanagementwidget);
}

/**
 * @brief "薪资管理" 导航按钮的槽函数
 *
 * 创建一个新的 SalaryManagement 实例，并将其添加到 QStackedWidget 中作为当前页面。
 */
void MainWindow::on_tbtnSalaryManagement_clicked()
{
    SalaryManagement *salarymanagement = new SalaryManagement;
    ui->stackedWidget->addWidget(salarymanagement);
    ui->stackedWidget->setCurrentWidget(salarymanagement);
}
