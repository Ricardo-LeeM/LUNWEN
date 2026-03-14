#ifndef HOMEPAGE_H
#define HOMEPAGE_H

#include <QWidget>
#include <QListWidgetItem>
#include "databasemanager.h"
#include <QListWidgetItem>
#include <QTimer>
// 前置声明，避免在头文件中引入完整的头文件，减少编译依赖
QT_BEGIN_NAMESPACE
namespace Ui { class HomePage; }
QT_END_NAMESPACE
class QTimer;
class QSqlTableModel;

/**
 * @brief The HomePage class
 *
 * 员工登录后看到的主界面（首页）部件。
 * 该界面主要功能包括：
 * 1. 显示当前登录员工的基本信息（姓名、工号、部门等）。
 * 2. 实时显示问候语（如“早上好”）。
 * 3. 提供财务、休假等快捷申报入口。
 * 4. 显示员工个人的历史薪资记录表格。
 * 5. 提供将薪资记录导出到Excel的功能。
 */
class HomePage : public QWidget
{
    Q_OBJECT

public:
    /**
     * @brief HomePage 构造函数
     * @param parent 父部件指针，默认为nullptr
     */
    explicit HomePage(QWidget *parent = nullptr);

    /**
     * @brief ~HomePage 析构函数
     */
    ~HomePage();

private slots:
    /**
     * @brief 当定时器触发时，更新界面上的问候语。
     */
    void updateTime();

    /**
     * @brief 更新财务申报次数的显示。
     */
    void updateFinancialNumSlot();

    /**
     * @brief 更新休假申报次数的显示。
     */
    void updateVacationNumslot();

    /**
     * @brief 更新立项申报次数的显示。
     */
    void updateProjectionNumslot();

    /**
     * @brief "财务申报" 按钮的点击槽函数。
     *        处理财务申报的提交逻辑。
     */
    void on_tbtnFinancial_clicked();

    /**
     * @brief "休假申报" 按钮的点击槽函数。
     *        处理休假申报的提交逻辑。
     */
    void on_tbtnVacation_clicked();

    /**
     * @brief "立项申报" 按钮的点击槽函数。
     *        (注意：此功能在您的代码中未实现)
     */
    void on_tbtnProjection_clicked();

    /**
     * @brief "导出到Excel" 按钮的点击槽函数。
     */
    void on_btnToExcel_clicked();

    /**
     * @brief "从Excel导入" 按钮的点击槽函数。
     */
    void on_btnFromExcel_clicked();
    void on_btnSendMessage_clicked();
    void on_listWidgetContacts_itemClicked(QListWidgetItem *item);
    void fetchNewMessages(); // 定时器拉取新消息的槽函数

private:
    /**
     * @brief 初始化薪资历史记录表格。
     *        设置数据模型、过滤器、表头等。
     */
    void initializeSalaryTable();

    /**
     * @brief 加载当前登录用户的数据并显示在界面上。
     *        信息来源为 UserSession 单例。
     */
    void loadCurrentUserData();

    /**
     * @brief 根据当前时间获取相应的问候语。
     * @return 返回如 "早上好"、"下午好" 等字符串。
     */
    QString getGreeting();

    /**
     * @brief 将薪资表格中的数据导出为 Excel 文件。
     */
    void exportToExcel();

    /**
     * @brief 从 Excel 文件导入数据。
     *        (注意：当前实现为提示用户此功能不可用)
     */
    void importFromExcel();

    /**
     * @brief 将一条新的申报记录添加到数据库的 'declarations' 表中。
     * @param applicant 申报人姓名。
     * @param type 申报类型（如 "财务申报"）。
     * @param content 申报的具体内容。
     * @return 如果成功插入数据库，返回 true；否则返回 false。
     */
    bool addDeclarationToDb(const QString &applicant, const QString &type, const QString &content);

    /**
     * @brief 更新 'personal_application' 表中指定员工的申报次数。
     *        如果员工记录不存在，则会尝试创建一条新记录。
     * @param applicantName 申报人姓名，用于查询其 EmployeeID。
     * @param declarationType 申报类型，用于确定要更新哪个计数字段。
     * @return 如果更新或插入成功，返回 true；否则返回 false。
     */
    bool updatePersonalApplicationCount(const QString &applicantName, const QString &declarationType);


private:
    Ui::HomePage *ui;               ///< UI界面指针，由Qt Designer生成
    QTimer *timer;                  ///< 用于定时更新问候语的定时器
    QSqlTableModel *salaryHistoryModel; ///< 用于显示薪资历史记录的表格数据模型


    QTimer *chatTimer;
    int currentChatUserId = -1; // 记录当前正在聊天的对象工号

    void loadContacts(); ///< 加载联系人列表
    void loadChatHistory(int targetUserId); ///< 加载历史聊天记录
    void appendChatMessage(const QString &text, bool isMine); ///< 渲染气泡到界面
};

#endif // HOMEPAGE_H
// HOMEPAGE_H
// #ifndef HOMEPAGE_H
// #define HOMEPAGE_H
// #include <QWidget>
// #include <QLabel>
// #include "dialogvacation.h"
// #include "dialogprojection.h"
// //#include "declarationmanagement.h"
// #include "dialogfinancial.h"
// #include <QtSql>
// #include <QDate>
// #include "usersession.h"
// namespace Ui {
// class HomePage;
// }

// class HomePage : public QWidget
// {
//     Q_OBJECT

// public:
//     explicit HomePage(QWidget *parent = nullptr);
//     ~HomePage();
//     void getMessage(QString username, int employeeID,
//                     QString department, QString employeeStatus,
//                     bool isPartyMember, QString workNature,
//                     QString position, QDate joinDate, QPixmap image);
//     QString getGreeting();
// private slots:
//     void updateTime();  // 更新时间打招呼

//     void on_tbtnVacation_clicked();  // 确定休假申请
//     void on_tbtnProjection_clicked(); // 确定项目申请
//     void on_tbtnFinancial_clicked();  // 确定财务申请



//     void updateFinancialNumSlot(); // 更新报销次数
//     void updateVacationNumslot(); // 更新请假申请次数
//     void updateProjectionNumslot(); // 更新立项申请次数

//     // 表格的导入导出
//     void on_btnToExcel_clicked(); // 导出表格按钮
//     void exportToExcel(); // 导出表格实现
//     void importFromExcel();// 导入表格实现
//     void on_btnFromExcel_clicked(); // 导入表格


// private:
//     Ui::HomePage *ui;
//     QLabel *labelTitle;
//     QTimer *timer; // 定时器

//     // DialogVacation *dialogvacation; // 休假申请界面
//     // DialogProjection *dialogprojection; // 立项申请界面
//     // DialogFinancial *dialogfinancial;  // 财务报销界面


//     void initializeDatabase(); // 用于在表不存在时创建表
//     bool addDeclarationToDb(const QString &applicant, const QString &type, const QString &content); // 添加申报到数据库
//     bool updatePersonalApplicationCount(const QString &applicantName, const QString &declarationType); // 更新数据库内的个人申报次数
//     void loadCurrentUserData();  // 加载信息
//     UserSession *session; // 个人信息类
// };

// #endif // HOMEPAGE_H
