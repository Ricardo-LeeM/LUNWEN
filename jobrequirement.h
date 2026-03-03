#ifndef JOBREQUIREMENT_H
#define JOBREQUIREMENT_H

/***************************************************************************************************
** 头文件依赖 (Header Dependencies)                                  **
***************************************************************************************************/
#include <QWidget>
#include <QSqlDatabase>
#include <QStandardItemModel>
#include <QSqlTableModel> // 用于面试者信息管理的数据模型

// Forward declaration for the UI class generated from the .ui file.
// This is a standard practice in Qt to avoid including the full UI header here.
namespace Ui {
class JobRequirement;
}

/***************************************************************************************************
** 全局常量 (Global Constants)                                  **
***************************************************************************************************/
/**
 * @brief 自定义角色 ID (Custom Role ID)
 *
 * 用于在 QStandardItemModel 的单元格中存储与显示文本 (Qt::DisplayRole) 无关的附加数据。
 * 在这里，我们用它来安全地存储数据库记录的唯一主键 ID。
 * 这样，即使用户修改了界面上可见的数据，我们依然能通过这个隐藏的 ID 准确地在数据库中找到对应的记录。
 * Qt::UserRole 是 Qt 预留的第一个可供用户自定义的角色起始值。
 */
const int DbIdRole = Qt::UserRole + 1;

/***************************************************************************************************
** JobRequirement 类定义                                       **
***************************************************************************************************/
/**
 * @class JobRequirement
 * @brief 职位需求与面试管理窗口
 *
 * 该类负责管理两个核心功能，通过一个 QTabWidget 展示：
 * 1.  **职位需求管理 (Tab 1):**
 * - 从数据库加载、显示和编辑职位需求。
 * - 支持添加、删除职位。
 * - 提供将数据导出到 Excel 和从 Excel 导入的功能。
 * - 能够将职位需求发布为格式化的 HTML 页面，以便于分享和查看。
 *
 * 2.  **面试者管理 (Tab 2):**
 * - 使用 QSqlTableModel 直接与数据库的 'interviewees' 表交互，实现对面试者信息的实时管理。
 * - 支持添加、删除、修改面试者信息。
 * - 提供基于姓名或电话的模糊搜索功能。
 * - 使用自定义委托 (StatusDelegate) 来美化“面试状态”列的显示。
 */
class JobRequirement : public QWidget
{
    Q_OBJECT

public:
    /**
     * @brief 构造函数
     * @param parent 父窗口指针，默认为 nullptr。
     */
    explicit JobRequirement(QWidget *parent = nullptr);

    /**
     * @brief 析构函数
     */
    ~JobRequirement();

    // ===============================================================================================
    //                                       槽函数 (SLOTS)
    // ===============================================================================================
private slots:
    // -------------------------------------------------------------------------------------------
    // -                            Tab 1: 职位需求管理 (Job Requirements)                         -
    // -------------------------------------------------------------------------------------------
    void on_btnAddJob_clicked();                     // "新增职位" 按钮点击事件
    void on_btnDeleteRow_clicked();                  // "删除选中行" 按钮点击事件
    void on_checkBox_clicked(bool checked);          // "开启/关闭编辑" 复选框状态改变事件
    void on_btnAutoAdjustRowHeigth_clicked();        // "自动调整行高" 按钮点击事件
    void on_btnAutoAdjustColumnWidth_clicked();      // "自动调整列宽" 按钮点击事件
    void on_exportButton_clicked();                  // "导出到Excel" 按钮点击事件
    void on_btnFromExcel_clicked();                  // "从Excel导入" 按钮点击事件
    //void on_btnRelease_clicked();                    // "发布为HTML" 按钮点击事件
    //void on_btnCheckInHtml_clicked();                // "在浏览器中查看" 按钮点击事件
    void handleItemChanged(QStandardItem *item);     // TableView 中的单元格数据被修改后触发

    // -------------------------------------------------------------------------------------------
    // -                              Tab 2: 面试者管理 (Interviewees)                            -
    // -------------------------------------------------------------------------------------------
    void on_addIntervieweeButton_clicked();          // "添加面试者" 按钮点击事件
    void on_deleteIntervieweeButton_clicked();       // "删除面试者" 按钮点击事件
    void on_editIntervieweeButton_clicked();         // "修改面试者" 按钮点击事件
    //void on_searchLineEdit_textChanged(const QString &text); // 搜索框文本内容改变事件
    void updateIntervieweeFilter();  // 搜索条件

    // ===============================================================================================
    //                                   私有成员 (PRIVATE MEMBERS)
    // ===============================================================================================
private:
    Ui::JobRequirement *ui; // 指向UI界面的指针

    // -------------------------------------------------------------------------------------------
    // -                             Tab 1: 职位需求相关成员                                     -
    // -------------------------------------------------------------------------------------------
    QStandardItemModel *m_jobRequirementModel;      // 存储和管理职位需求数据的模型
    QString m_persistentHtmlFilePath;               // 持久化存储的HTML文件的完整路径

    // -------------------------------------------------------------------------------------------
    // -                             Tab 2: 面试者管理相关成员                                   -
    // -------------------------------------------------------------------------------------------
    QSqlTableModel *m_intervieweeModel;             // 直接映射数据库 'interviewees' 表的模型

    // ===============================================================================================
    //                                   私有函数 (PRIVATE FUNCTIONS)
    // ===============================================================================================
private:
    /** @name 初始化函数 (Initialization Functions) */
    ///@{
    void setupJobRequirementsTab(); // 初始化 "职位需求" 选项卡的用户界面和数据模型
    void setupIntervieweesTab();    // 初始化 "面试者管理" 选项卡的用户界面和数据模型
    ///@}

    /** @name 辅助函数 (Helper Functions) */
    ///@{
    void loadJobRequirementsData();       // 从数据库加载职位需求数据到 m_jobRequirementModel
    void exportToExcel();                 // 将职位需求数据导出为 Excel 文件
    void importFromExcel();               // 从 Excel 文件导入数据到职位需求表格
    QString determinePersistentHtmlPath() const; // 确定用于保存招聘信息的HTML文件的存储位置
    ///@}
};

#endif // JOBREQUIREMENT_H
// #ifndef JOBREQUIREMENT_H
// #define JOBREQUIREMENT_H

// #include <QWidget>
// #include <QSqlDatabase>
// #include <QStandardItemModel>
// #include <QSqlTableModel> // 新增：用于面试者管理

// namespace Ui {
// class JobRequirement;
// }

// // 将这个常量放在头文件中，以便 .cpp 文件中的所有函数都能访问
// const int DbIdRole = Qt::UserRole + 1;

// class JobRequirement : public QWidget
// {
//     Q_OBJECT

// public:
//     explicit JobRequirement(QWidget *parent = nullptr);
//     ~JobRequirement();

// private slots:
//     // --- Tab 1: 职位需求功能槽函数 ---
//     void on_btnAddJob_clicked();
//     void on_btnDeleteRow_clicked();
//     void on_checkBox_clicked(bool checked);
//     void on_btnAutoAdjustRowHeigth_clicked();
//     void on_btnAutoAdjustColumnWidth_clicked();
//     void on_exportButton_clicked();
//     void on_btnFromExcel_clicked();
//     void on_btnRelease_clicked();
//     void on_btnCheckInHtml_clicked();
//     void handleItemChanged(QStandardItem *item);

//     // === 新增：Tab 2: 面试者管理功能槽函数 ===
//     void on_addIntervieweeButton_clicked();    // "添加面试者" 按钮
//     void on_deleteIntervieweeButton_clicked(); // "删除面试者" 按钮
//     void on_editIntervieweeButton_clicked();   // "修改面试者" 按钮
//     void on_searchLineEdit_textChanged(const QString &text); // 搜索框文本变化

// private:
//     Ui::JobRequirement *ui;

//     // --- Tab 1: 职位需求相关成员 ---
//     QStandardItemModel *m_jobRequirementModel; // 建议重命名以区分
//     QString m_persistentHtmlFilePath;

//     // --- Tab 2: 面试者管理相关成员 ---
//     QSqlTableModel *m_intervieweeModel;

//     // --- 优化：私有初始化函数 ---
//     void setupJobRequirementsTab(); // 初始化职位需求Tab
//     void setupIntervieweesTab();    // 初始化面试者Tab

//     // --- 辅助函数 ---
//     void loadJobRequirementsData(); // 重命名 loadDataFromDatabase
//     void exportToExcel();
//     void importFromExcel();
//     QString determinePersistentHtmlPath() const;
// };

/*#endif*/ // JOBREQUIREMENT_H
// #ifndef JOBREQUIREMENT_H
// #define JOBREQUIREMENT_H

// #include <QStandardItemModel>
// #include <QWidget>

// #include <QStandardItemModel> // 需要包含，因为 model 是 QStandardItemModel* 类型
// #include <QSqlDatabase>      // 需要包含，因为 getDatabaseConnection 返回 QSqlDatabase
// #include <QMap> // 新增
// #include <QSqlTableModel>

// #include <QTemporaryFile>
// namespace Ui {
// class JobRequirement;
// }

// // 定义一个结构体来存储面试者信息
// struct InterviewerInfo {
//     QString name;
//     QString email;
//     QString phoneNumber;
//     QString position;
//     QString status;
// };

// const int DbIdRole = Qt::UserRole + 1; // 保持这个常量
// class JobRequirement : public QWidget
// {
//     Q_OBJECT

// public:
//     explicit JobRequirement(QWidget *parent = nullptr);
//     ~JobRequirement();

// private slots:
//     void on_btnAddJob_clicked(); // 添加工作
//     void on_btnDeleteRow_clicked();  // 删除行
//     void on_checkBox_clicked(bool checked); // 可编辑
//     void on_btnAutoAdjustRowHeigth_clicked();
//     void on_btnAutoAdjustColumnWidth_clicked();
//     void on_exportButton_clicked(); // 点击导出
//     void exportToExcel(); // 导出表格
//     void on_btnFromExcel_clicked(); // 点击导入
//     void importFromExcel(); // 导入表格


//     void handleItemChanged(QStandardItem *item); // 当模型中的项发生改变时

//     void on_btnRelease_clicked();

//     void on_btnCheckInHtml_clicked();

// private:
//     Ui::JobRequirement *ui;
//     QStandardItemModel *model;
//     QSqlDatabase m_db; // 将数据库连接作为成员变量
//     QSqlTableModel *m_intervieweeModel; // 声明一个成员变量来持有模型

//     void loadDataFromDatabase();

//     // QString m_generatedHtmlContent; // 移除
//     // std::unique_ptr<QTemporaryFile> m_currentTempHtmlFile; // 移除

//     QString m_persistentHtmlFilePath; // 用于存储持久化HTML文件的路径
//     QString determinePersistentHtmlPath() const; // 辅助函数获取路径



//     void setupDatabaseConnection();// 建议将数据库连接逻辑封装

//     // 确保 DbIdRole 在整个类中可用
//     // 如果在 .h 顶部定义了 const int DbIdRole，则这里不需要再定义
//     // 如果仅在 .cpp 中定义，则需要确保在 .h 中可以访问或者在每个需要它的地方都能正确获取
//     // 最佳实践是将其定义在 .h 文件中，或者作为类的一个 static const int 成员。
//     // 为了简单，我们假设 DbIdRole 已经在您的 .cpp 文件的顶部全局定义，或者您会把它移到 .h
//     // 例如： static const int DbIdRole = Qt::UserRole + 1; (在类声明内部)
//     // 或者如您所做，在 .h 的命名空间之外定义。

// };

// #endif // JOBREQUIREMENT_H
