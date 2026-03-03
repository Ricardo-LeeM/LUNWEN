#ifndef TASKEDITDIALOG_H
#define TASKEDITDIALOG_H

/***************************************************************************************************
** 头文件依赖 (Header Dependencies)                                  **
***************************************************************************************************/
#include <QDialog>
#include <QSqlRecord> // 用于封装和传递单条数据库记录

// UI类的标准前向声明
namespace Ui {
class TaskEditDialog;
}

/***************************************************************************************************
** TaskEditDialog 类定义                                       **
***************************************************************************************************/
/**
 * @class TaskEditDialog
 * @brief 任务创建与编辑对话框
 *
 * 这个类是任务管理功能的核心交互界面。它以一个对话框的形式出现，允许用户创建新任务或查看、编辑现有任务。
 * 它被设计为具有两种模式（通过 TaskMode 枚举控制），并内置了复杂的权限系统。
 *
 * 在项目中的作用:
 * - **被调用方:** 通常由 `PersonalInterface` 在用户双击一个 `TaskCardWidget` 时创建并显示。
 * - **数据交互:**
 * - 接收一个 `QSqlRecord` 对象来初始化界面（对于编辑模式）或作为模板（对于添加模式）。
 * - 通过 `getRecordData()` 方法返回一个填充了用户输入的新 `QSqlRecord`，供调用者（如 `PersonalInterface`）
 * 用来更新数据库。
 * - **权限控制:**
 * - 利用传入的当前用户ID (`currentUserId`) 和任务本身的“分派人ID”、“执行人ID”，动态地设置界面控件的只读状态。
 * - 例如，只有任务的“执行人”才能修改任务状态；只有“分派人”才能在任务完成后填写评价，这确保了业务流程的正确性。
 * - **动态UI:** 界面会根据“任务类型”（常规任务 vs. 出差任务）动态显示或隐藏相关字段（如目的地、出差开始日期）。
 */
class TaskEditDialog : public QDialog
{
    Q_OBJECT

public:
    /**
     * @enum TaskMode
     * @brief 定义对话框的两种工作模式
     */
    enum TaskMode {
        AddTask,  ///< 添加新任务模式
        EditTask  ///< 编辑现有任务模式
    };

    /**
     * @brief 构造函数
     * @param record 包含任务数据的SQL记录。在AddTask模式下，可传入一个空记录作为模板。
     * @param currentUserId 当前登录用户的ID，用于权限判断。
     * @param mode 对话框的工作模式 (AddTask 或 EditTask)。
     * @param parent 父窗口指针。
     */
    explicit TaskEditDialog(const QSqlRecord &record, int currentUserId, TaskMode mode, QWidget *parent = nullptr);

    /**
     * @brief 析构函数
     */
    ~TaskEditDialog();

    /**
     * @brief 获取对话框中的数据
     *
     * 收集所有UI控件中的用户输入，并将其填充到一个 `QSqlRecord` 对象中返回。
     * 在返回前会进行基本的数据有效性校验。
     * @return 返回一个包含界面数据的 QSqlRecord。如果校验失败，可能返回一个空记录。
     */
    QSqlRecord getRecordData();

private slots:
    /**
     * @brief 当任务类型下拉框选项改变时触发的槽函数。
     * @param taskType 新选中的任务类型文本。
     */
    void onTaskTypeChanged(const QString &taskType);



private:
    /**
     * @brief 初始化界面中的所有下拉框（ComboBox）。
     */
    void initComboBoxes();

    /**
     * @brief 根据任务类型更新UI的显示/隐藏状态。
     * @param taskType 当前选中的任务类型。
     */
    void updateUiForTaskType(const QString &taskType);

    /**
     * @brief 在编辑模式下，用传入的记录数据填充UI控件。
     */
    void fillData();

private:
    Ui::TaskEditDialog *ui;   // 指向UI界面的指针
    QSqlRecord m_record;      // 存储当前任务数据的记录对象
    int m_currentUserId;      // 当前登录用户的ID
    TaskMode m_mode;          // 对话框的当前工作模式
};

#endif // TASKEDITDIALOG_H
// #ifndef TASKEDITDIALOG_H
// #define TASKEDITDIALOG_H

// #include <QDialog>
// #include <QSqlRecord>



// namespace Ui {
// class TaskEditDialog;
// }

// class TaskEditDialog : public QDialog
// {
//     Q_OBJECT

// public:
//     enum TaskMode { AddTask, EditTask };

//     explicit TaskEditDialog(const QSqlRecord &record, int currentUserId, TaskMode mode, QWidget *parent = nullptr);
//     ~TaskEditDialog();
//     QSqlRecord getRecordData();

// private slots:
//     void onTaskTypeChanged(const QString &taskType);

// private:
//     void initComboBoxes();
//     void updateUiForTaskType(const QString &taskType);
//     void fillData();

// private:
//     Ui::TaskEditDialog *ui;
//     QSqlRecord m_record;
//     int m_currentUserId;
//     TaskMode m_mode;
// };

// #endif // TASKEDITDIALOG_H
