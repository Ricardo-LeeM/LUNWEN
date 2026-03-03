#ifndef SALARYMANAGEMENT_H
#define SALARYMANAGEMENT_H

/***************************************************************************************************
** 头文件依赖 (Header Dependencies)                                  **
***************************************************************************************************/
#include <QMainWindow>
#include <QSqlTableModel>

// UI类的标准前向声明
namespace Ui {
class SalaryManagement;
}

/***************************************************************************************************
** SalaryManagement 类定义                                     **
***************************************************************************************************/
/**
 * @class SalaryManagement
 * @brief 薪资管理窗口
 *
 * 该类负责薪资记录的创建、计算、和管理。它使用 `QSqlTableModel` 直接与数据库的
 * `PayrollRecords` 表进行交互，并实现了一系列智能化功能来简化薪资专员的操作。
 *
 * 核心功能:
 * 1.  **数据模型:** 使用 `QSqlTableModel` 以表格形式展示和编辑特定薪资周期的数据。
 * 2.  **智能填充:** 当用户在“员工ID”列输入一个ID并确认后，系统会自动从 `EmployeesSalary`
 * 表中查询该员工的姓名、部门和基本工资，并填充到当前行，极大提高了录入效率。
 * 3.  **自动计算:** 提供一个“一键计算”按钮，可以遍历表格中的所有记录，根据预设的公式
 * (应发 = 各项应发之和, 实发 = 应发 - 各项应扣之和) 自动计算并填充“应发工资”和“实发工资”。
 * 4.  **手动提交:** 编辑策略为 `OnManualSubmit`，意味着所有在表格中的修改（增、删、改）
 * 都需要点击“保存”按钮后才会一次性提交到数据库，提供了撤销更改的机会。
 * 5.  **自定义编辑器:** 使用自定义委托（Delegate）替换了数字列默认的微调框（SpinBox），
 * 提供了更流畅的数字输入体验。
 */
class SalaryManagement : public QMainWindow
{
    Q_OBJECT

public:
    /**
     * @brief 构造函数
     * @param parent 父窗口指针，默认为 nullptr。
     */
    explicit SalaryManagement(QWidget *parent = nullptr);

    /**
     * @brief 析构函数
     */
    ~SalaryManagement();

    // ===============================================================================================
    //                                       槽函数 (SLOTS)
    // ===============================================================================================
private slots:
    void on_loadDataButton_clicked();   // “加载数据”按钮点击事件
    void on_calculateButton_clicked();  // “计算工资”按钮点击事件
    void on_saveButton_clicked();       // “保存更改”按钮点击事件
    void on_addButton_clicked();        // “新增记录”按钮点击事件
    void on_deleteButton_clicked();     // “删除记录”按钮点击事件

    /**
     * @brief 当表格数据被用户修改时触发的槽函数。
     * 主要用于实现：当用户输入员工ID后，自动填充员工姓名、部门和基本工资。
     * @param topLeft 被修改区域的左上角模型索引。
     * @param bottomRight 被修改区域的右下角模型索引。
     * @param roles 发生变化的角色。
     */
    void on_dataChanged(const QModelIndex &topLeft, const QModelIndex &bottomRight, const QVector<int> &roles = QVector<int>());


    // ===============================================================================================
    //                                   私有成员 (PRIVATE MEMBERS)
    // ===============================================================================================
private:
    Ui::SalaryManagement *ui; // 指向UI界面的指针
    QSqlTableModel *model;    // 薪资数据模型，直接与数据库表关联

    /**
     * @brief 自动填充标志位
     *
     * 这是一个非常重要的标志。当程序内部代码（如自动填充姓名、计算工资）修改模型数据时，
     * `dataChanged`信号也会被触发。为防止此时再次执行`on_dataChanged`槽内的逻辑导致死循环或意外行为，
     * 我们在自动填充前后将此标志设为true，并在`on_dataChanged`的开头检查此标志。
     */
    bool isAutoFilling;

    // ===============================================================================================
    //                                   私有函数 (PRIVATE FUNCTIONS)
    // ===============================================================================================
private:
    /**
     * @brief 初始化数据模型和表格视图
     *
     * 负责设置`QSqlTableModel`，关联数据库表，设置编辑策略，设置列标题，
     * 以及为特定列应用自定义委托（Delegate）。
     */
    void initializeModel();
};

#endif // SALARYMANAGEMENT_H
// #ifndef SALARYMANAGEMENT_H
// #define SALARYMANAGEMENT_H

// #include <QMainWindow>
// #include <QSqlTableModel>

// namespace Ui {
// class SalaryManagement;
// }

// class SalaryManagement : public QMainWindow
// {
//     Q_OBJECT

// public:
//     explicit SalaryManagement(QWidget *parent = nullptr);
//     ~SalaryManagement();

// private slots:
//     void on_loadDataButton_clicked();
//     void on_calculateButton_clicked();
//     void on_saveButton_clicked();
//     void on_addButton_clicked();
//     void on_deleteButton_clicked();

//     // 新增的槽函数，用于在数据改变时触发
//     void on_dataChanged(const QModelIndex &topLeft, const QModelIndex &bottomRight, const QVector<int> &roles);

// private:
//     Ui::SalaryManagement *ui;
//     QSqlTableModel *model;
//     void initializeModel();

//     // 一个标志位，防止在自动填充时再次触发 on_dataChanged
//     bool isAutoFilling;
// };

// #endif // SALARYMANAGEMENT_H
