#include "employeemanagementwidget.h"
#include "ui_employeemanagementwidget.h"
#include "employeeeditdialog.h"
#include "taskeditdialog.h"
#include "usersession.h"
#include <QSqlDatabase>
#include <QSqlField>
#include "databasemanager.h" // 引入

#include <QMessageBox>
#include <QDebug>
#include <QSqlRecord>

EmployeeManagementWidget::EmployeeManagementWidget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::EmployeeManagementWidget),
    employeeModel(nullptr),
    taskModel(nullptr),
    employeeeditdialog(nullptr),
    taskeditdialog(nullptr)
{
    ui->setupUi(this);
    initUI();
}

EmployeeManagementWidget::~EmployeeManagementWidget()
{
    delete ui;
}

void EmployeeManagementWidget::initUI()
{
    setupModels();
    setupViews();

    // 填充下拉框
    ui->deptFilterComboBox->addItem("所有部门");
    ui->deptFilterComboBox->addItems(DatabaseManager::instance()->getAllDepartments());

    ui->taskTypeFilterComboBox->addItem("所有类型", "ALL");
    ui->taskTypeFilterComboBox->addItems({"常规任务", "出差任务"});
    ui->taskStatusFilterComboBox->addItem("所有状态", "ALL");
    ui->taskStatusFilterComboBox->addItems({"未开始", "进行中", "已完成", "已作废"});
    ui->assigneeFilterComboBox->addItem("所有执行人", 0);
    QList<QSqlRecord> employees = DatabaseManager::instance()->getAllActiveEmployees(0); // 传0表示获取所有在职员工
    for(const QSqlRecord& emp : employees){
        ui->assigneeFilterComboBox->addItem(emp.value("Name").toString(), emp.value("EmployeeID"));
    }

    // 初始加载数据
    loadAllEmployeeData();
    loadAllTaskData();
}

void EmployeeManagementWidget::setupModels()
{
    employeeModel = new QStandardItemModel(this);
    taskModel = new QStandardItemModel(this);
}

void EmployeeManagementWidget::setupViews()
{
    ui->employeeTableView->setModel(employeeModel);
    ui->employeeTableView->setSelectionBehavior(QAbstractItemView::SelectRows);
    ui->employeeTableView->setSelectionMode(QAbstractItemView::SingleSelection);
    ui->employeeTableView->setEditTriggers(QAbstractItemView::NoEditTriggers);
    ui->employeeTableView->setAlternatingRowColors(true);
    ui->employeeTableView->setSortingEnabled(true);

    ui->taskTableView->setModel(taskModel);
    ui->taskTableView->setSelectionBehavior(QAbstractItemView::SelectRows);
    ui->taskTableView->setEditTriggers(QAbstractItemView::NoEditTriggers);
    ui->taskTableView->setSortingEnabled(true);
}

void EmployeeManagementWidget::loadAllEmployeeData()
{
    QList<QSqlRecord> records = DatabaseManager::instance()->getFilteredEmployees("", "");
    populateEmployeeTable(records);
}

void EmployeeManagementWidget::populateEmployeeTable(const QList<QSqlRecord>& records)
{
    employeeModel->clear();
    QStringList headers = {"工号", "姓名", "部门", "职位", "入职日期", "在职状态", "工作性质", "员工类型"};
    employeeModel->setHorizontalHeaderLabels(headers);

    for (const QSqlRecord& record : records) {
        QList<QStandardItem*> rowItems;
        rowItems << new QStandardItem(record.value("EmployeeID").toString());
        rowItems << new QStandardItem(record.value("Name").toString());
        rowItems << new QStandardItem(record.value("Department").toString());
        rowItems << new QStandardItem(record.value("Position").toString());
        rowItems << new QStandardItem(record.value("JoinDate").toDate().toString("yyyy-MM-dd"));
        rowItems << new QStandardItem(record.value("EmployeeStatus").toString());
        rowItems << new QStandardItem(record.value("WorkNature").toString());
        rowItems << new QStandardItem(record.value("EmployeeType").toString());
        employeeModel->appendRow(rowItems);
    }
    ui->employeeTableView->resizeColumnsToContents();
}

void EmployeeManagementWidget::loadAllTaskData()
{
    QList<QSqlRecord> records = DatabaseManager::instance()->getFilteredTasks(0, "", "");
    populateTaskTable(records);
}

void EmployeeManagementWidget::populateTaskTable(const QList<QSqlRecord>& records)
{
    taskModel->clear();
    QStringList headers = {"任务ID", "任务名称", "任务类型", "分派人ID", "执行人ID", "状态", "截止日期"};
    taskModel->setHorizontalHeaderLabels(headers);

    for (const QSqlRecord& record : records) {
        QList<QStandardItem*> rowItems;
        rowItems << new QStandardItem(record.value("TaskID").toString());
        rowItems << new QStandardItem(record.value("TaskName").toString());
        rowItems << new QStandardItem(record.value("TaskType").toString());
        rowItems << new QStandardItem(record.value("AssignerID").toString());
        rowItems << new QStandardItem(record.value("AssigneeID").toString());
        rowItems << new QStandardItem(record.value("Status").toString());
        rowItems << new QStandardItem(record.value("DueDate").toDate().toString("yyyy-MM-dd"));
        taskModel->appendRow(rowItems);
    }
    ui->taskTableView->resizeColumnsToContents();
}

void EmployeeManagementWidget::on_queryEmployeeButton_clicked()
{
    QString name = ui->nameFilterLineEdit->text().trimmed();
    QString department = ui->deptFilterComboBox->currentText();
    QList<QSqlRecord> records = DatabaseManager::instance()->getFilteredEmployees(name, department);
    populateEmployeeTable(records);
}

void EmployeeManagementWidget::on_resetEmployeeButton_clicked()
{
    ui->nameFilterLineEdit->clear();
    ui->deptFilterComboBox->setCurrentIndex(0);
    loadAllEmployeeData();
}

void EmployeeManagementWidget::on_addEmployeeButton_clicked()
{
    // 使用局部变量创建对话框，这是最安全的方式
    EmployeeEditDialog dialog(QSqlRecord(), Add, this);

    // 当用户点击"OK"并且对话框关闭时
    if (dialog.exec() == QDialog::Accepted) {
        // 从对话框获取用户填写的所有数据
        QSqlRecord newRecord = dialog.getRecordData();

        // **关键校验**：确保记录不为空，并且ID是有效的
        if (newRecord.isEmpty() || newRecord.value("EmployeeID").toInt() <= 0) {
            // 如果记录为空，或ID无效，说明对话框内部的校验逻辑已经阻止了它
            // 或者用户输入了无效数据。这里直接返回，无需额外提示。
            return;
        }

        int employeeId = newRecord.value("EmployeeID").toInt();

        // 检查工号是否已存在
        if (DatabaseManager::instance()->checkEmployeeIdExists(employeeId)) {
            QMessageBox::critical(this, "操作失败", QString("员工工号 %1 已存在！").arg(employeeId));
        } else {
            // 调用 DatabaseManager 执行数据库操作
            if (DatabaseManager::instance()->addNewEmployee(newRecord)) {
                QMessageBox::information(this, "成功", "新员工及登录账户已成功创建！");
                loadAllEmployeeData(); // 成功后刷新界面
            } else {
                QMessageBox::critical(this, "数据库错误", "添加新员工失败，请检查终端输出获取详细错误信息。");
            }
        }
    }
}

void EmployeeManagementWidget::on_editEmployeeButton_clicked()
{
    // 获取当前选中的行
        QModelIndex currentIndex = ui->employeeTableView->currentIndex();
    if (!currentIndex.isValid()) {
        QMessageBox::warning(this, "提示", "请先选择要修改的员工！");
        return;
    }

    // 从模型的第一列（工号列）获取员工ID
    int employeeId = employeeModel->item(currentIndex.row(), 0)->text().toInt();

    // 使用ID从数据库获取该员工的完整、最新的记录
    QSqlRecord recordToEdit = DatabaseManager::instance()->getEmployeeRecordById(employeeId);

    if (recordToEdit.isEmpty()) {
        QMessageBox::critical(this, "数据错误", QString("在数据库中找不到工号为 %1 的员工记录。").arg(employeeId));
        loadAllEmployeeData(); // 刷新表格以反映最新状态
        return;
    }

    // 创建编辑对话框并填充数据
    EmployeeEditDialog dialog(recordToEdit, Edit, this);

    if (dialog.exec() == QDialog::Accepted) {
        QSqlRecord updatedRecord = dialog.getRecordData();
        // 执行更新操作
        if (DatabaseManager::instance()->updateEmployee(updatedRecord)) {
            QMessageBox::information(this, "成功", "员工信息已成功更新。");
            loadAllEmployeeData(); // 刷新表格
        } else {
            QMessageBox::critical(this, "失败", "更新员工信息失败。");
        }
    }
}

void EmployeeManagementWidget::on_deleteEmployeeButton_clicked()
{
    int currentRow = ui->employeeTableView->currentIndex().row();
    if (currentRow < 0) {
        QMessageBox::warning(this, "提示", "请先选择要设为离职的员工！");
        return;
    }
    int employeeId = employeeModel->item(currentRow, 0)->text().toInt();
    QString name = employeeModel->item(currentRow, 1)->text();

    if (QMessageBox::question(this, "确认操作", QString("您确定要将员工 [%1] 的状态设置为“离职”吗？").arg(name), QMessageBox::Yes | QMessageBox::No) == QMessageBox::Yes) {
        if (DatabaseManager::instance()->setEmployeeStatusToResigned(employeeId)) {
            QMessageBox::information(this, "操作成功", "员工 [" + name + "] 的状态已更新为“离职”。");
            loadAllEmployeeData();
        } else {
            QMessageBox::critical(this, "操作失败", "更新数据库失败。");
        }
    }
}

// --- 任务管理槽函数 ---

void EmployeeManagementWidget::on_assignTaskButton_clicked()
{

    // 【修改 1】获取数据库连接并读取 Tasks 表的结构，而不是创建一个空的记录
    QSqlDatabase db = QSqlDatabase::database(); // 获取默认数据库连接
    QSqlRecord emptyRecord = db.record("tasks"); // 获取 Tasks 表的字段结构（TaskName, Status 等）
    int currentUserId = UserSession::instance()->employeeID();

    TaskEditDialog dialog(emptyRecord, currentUserId, TaskEditDialog::AddTask, this);

    if (dialog.exec() == QDialog::Accepted) {
        QSqlRecord newTaskRecord = dialog.getRecordData();

        // 【修改 2】增加调试打印，确保我们拿到了数据
        qDebug() << "Debug: 新任务字段数:" << newTaskRecord.count();
        qDebug() << "Debug: 任务名称:" << newTaskRecord.value("TaskName").toString();

        if (newTaskRecord.isEmpty()) {
            QMessageBox::warning(this, "错误", "任务记录为空（可能是字段匹配失败）！");
            return;
        }

        // 调用 DatabaseManager 执行数据库操作
        if (DatabaseManager::instance()->addTask(newTaskRecord)) {
            QMessageBox::information(this, "成功", "新任务已成功分派！");
            loadAllTaskData(); // 成功后刷新任务列表
        } else {
            // 这里会显示 addTask 内部打印的详细错误
            QMessageBox::critical(this, "数据库错误", "提交任务失败，请检查终端输出获取详细错误信息。");
        }
    }
}

void EmployeeManagementWidget::on_queryTaskButton_clicked()
{
    int assigneeId = ui->assigneeFilterComboBox->currentData().toInt();
    QString taskType = ui->taskTypeFilterComboBox->currentText();
    QString status = ui->taskStatusFilterComboBox->currentText();
    QList<QSqlRecord> records = DatabaseManager::instance()->getFilteredTasks(assigneeId, taskType, status);
    populateTaskTable(records);
}

void EmployeeManagementWidget::on_resetTaskButton_clicked()
{
    ui->assigneeFilterComboBox->setCurrentIndex(0);
    ui->taskTypeFilterComboBox->setCurrentIndex(0);
    ui->taskStatusFilterComboBox->setCurrentIndex(0);
    loadAllTaskData();
}

void EmployeeManagementWidget::on_deleteTaskButton_clicked()
{
    int currentRow = ui->taskTableView->currentIndex().row();
    if (currentRow < 0) {
        QMessageBox::warning(this, "操作提示", "请先选择一个要作废的任务！");
        return;
    }
    int taskId = taskModel->item(currentRow, 0)->text().toInt();
    QString taskName = taskModel->item(currentRow, 1)->text();

    if (QMessageBox::question(this, "确认操作", QString("您确定要将任务 [%1] 的状态设置为“已作废”吗？").arg(taskName), QMessageBox::Yes | QMessageBox::No) == QMessageBox::Yes) {
        if (DatabaseManager::instance()->setTaskStatusToCancelled(taskId)) {
            QMessageBox::information(this, "操作成功", "任务 [" + taskName + "] 已成功作废。");
            loadAllTaskData();
        } else {
            QMessageBox::critical(this, "数据库错误", "操作失败，无法更新任务状态。");
        }
    }
}

void EmployeeManagementWidget::on_editTaskButton_clicked()
{
    QModelIndex currentIndex = ui->taskTableView->currentIndex();
    if (!currentIndex.isValid()) {
        QMessageBox::warning(this, "操作提示", "请先选择一个要编辑或查看的任务！");
        return;
    }

    int taskId = taskModel->item(currentIndex.row(), 0)->text().toInt();

    QSqlRecord taskRecord = DatabaseManager::instance()->getTaskRecordById(taskId);
    if (taskRecord.isEmpty()) {
        QMessageBox::critical(this, "数据错误", QString("无法找到ID为 %1 的任务，该任务可能已被删除。").arg(taskId));
        loadAllTaskData();
        return;
    }

    int currentUserId = UserSession::instance()->employeeID();
    TaskEditDialog dialog(taskRecord, currentUserId, TaskEditDialog::EditTask, this);

    if (dialog.exec() == QDialog::Accepted) {
        QSqlRecord updatedRecord = dialog.getRecordData();
        if (!updatedRecord.isEmpty()) {
            if (DatabaseManager::instance()->updateTask(updatedRecord)) {
                QMessageBox::information(this, "成功", "任务信息已成功更新！");
                loadAllTaskData(); // 刷新任务列表
            } else {
                QMessageBox::critical(this, "数据库错误", "更新任务失败。");
            }
        }
    }
}

