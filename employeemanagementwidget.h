#ifndef EMPLOYEEMANAGEMENTWIDGET_H
#define EMPLOYEEMANAGEMENTWIDGET_H

#include <QWidget>
#include <QStandardItemModel> // <-- 重要修改
#include <QSqlRecord>

namespace Ui {
class EmployeeManagementWidget;
}
class EmployeeEditDialog;
class TaskEditDialog;

class EmployeeManagementWidget : public QWidget
{
    Q_OBJECT

public:
    explicit EmployeeManagementWidget(QWidget *parent = nullptr);
    ~EmployeeManagementWidget();

private slots:
    // 员工管理槽函数
    void on_queryEmployeeButton_clicked();
    void on_resetEmployeeButton_clicked();
    void on_addEmployeeButton_clicked();
    void on_editEmployeeButton_clicked();
    void on_deleteEmployeeButton_clicked();

    // 任务管理槽函数
    void on_assignTaskButton_clicked();
    void on_editTaskButton_clicked();
    void on_resetTaskButton_clicked();
    void on_queryTaskButton_clicked();
    void on_deleteTaskButton_clicked();

private:
    void initUI(); // 替换 initModels 和 initViews
    void setupModels();
    void setupViews();

    void loadAllEmployeeData();
    void populateEmployeeTable(const QList<QSqlRecord>& records);

    void loadAllTaskData();
    void populateTaskTable(const QList<QSqlRecord>& records);


private:
    Ui::EmployeeManagementWidget *ui;
    // *** 重要修改：使用QStandardItemModel ***
    QStandardItemModel *employeeModel;
    QStandardItemModel *taskModel;

    // Dialog指针保持不变
    EmployeeEditDialog *employeeeditdialog;
    TaskEditDialog *taskeditdialog;
};

#endif // EMPLOYEEMANAGEMENTWIDGET_H
