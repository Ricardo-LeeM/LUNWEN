#include "taskeditdialog.h"
#include "ui_taskeditdialog.h"
#include "usersession.h"        // 需要用它来获取分派人的姓名
#include <QSqlQuery>
#include <QVariant>
#include <QMessageBox>
#include <QDate>
#include <QDebug>
#include <QSqlError>
#include "databasemanager.h"

/***************************************************************************************************
** 构造函数与析构函数                                            **
***************************************************************************************************/
TaskEditDialog::TaskEditDialog(const QSqlRecord &record, int currentUserId, TaskMode mode, QWidget *parent) :
    QDialog(parent),
    ui(new Ui::TaskEditDialog),
    m_record(record),
    m_currentUserId(currentUserId),
    m_mode(mode)
{
    ui->setupUi(this);
    initComboBoxes(); // 初始化所有下拉框
    // 连接信号槽，当任务类型改变时，动态更新UI
    connect(ui->taskTypeComboBox, &QComboBox::currentTextChanged, this, &TaskEditDialog::onTaskTypeChanged);

    if (m_mode == AddTask) {
        // --- 添加模式下的初始化 ---
        setWindowTitle("分派新任务");
        // 分派人默认为当前登录用户，且不可更改
        ui->assignerLineEdit->setText(UserSession::instance()->name());
        ui->assignerLineEdit->setReadOnly(true);
        // 默认截止日期为一周后
        ui->dueDateEdit->setDate(QDate::currentDate().addDays(7));
        // 新任务还没有完成，评价框不可用
        ui->evaluationTextEdit->setReadOnly(true);
        // 新任务状态由系统默认设置，用户不可选
        ui->statusComboBox->setEnabled(false);
    } else {
        // --- 编辑模式下的初始化 ---
        setWindowTitle("任务详情");
        fillData(); // 使用传入的记录填充界面

        // --- 核心权限控制逻辑 ---
        // 判断当前用户是分派人还是执行人
        bool isAssigner = (m_currentUserId == m_record.value("AssignerID").toInt());
        bool isAssignee = (m_currentUserId == m_record.value("AssigneeID").toInt());
        QString status = m_record.value("Status").toString();

        // 只有分派人能在任务“已完成”前，修改任务的核心信息
        bool canEditCoreInfo = isAssigner && (status != "已完成");
        ui->taskNameLineEdit->setReadOnly(!canEditCoreInfo);
        ui->taskTypeComboBox->setEnabled(canEditCoreInfo);
        ui->assigneeComboBox->setEnabled(canEditCoreInfo);
        ui->dueDateEdit->setReadOnly(!canEditCoreInfo);
        ui->descriptionTextEdit->setReadOnly(!canEditCoreInfo);
        ui->destinationLineEdit->setReadOnly(!canEditCoreInfo); // 出差相关
        ui->startDateEdit->setReadOnly(!canEditCoreInfo);      // 出差相关

        // 只有执行人能在任务“已完成”前，修改任务状态
        ui->statusComboBox->setEnabled(isAssignee && (status != "已完成"));

        // 只有分派人能在任务“已完成”后，填写评价
        ui->evaluationTextEdit->setReadOnly(!isAssigner || (status != "已完成"));
    }
    // 根据当前的（或默认的）任务类型，更新UI显示
    updateUiForTaskType(ui->taskTypeComboBox->currentText());
}

TaskEditDialog::~TaskEditDialog()
{
    delete ui;
}

/***************************************************************************************************
** 公有方法实现 (Public Method Implementations)                      **
***************************************************************************************************/

QSqlRecord TaskEditDialog::getRecordData()
{
    // --- 数据校验 ---
    if (ui->taskNameLineEdit->text().trimmed().isEmpty()) {
        QMessageBox::warning(this, "输入错误", "任务/出差名称不能为空！");
        return QSqlRecord(); // 返回空记录表示失败
    }
    if (ui->assigneeComboBox->currentIndex() == -1) {
        QMessageBox::warning(this, "输入错误", "必须为任务选择一个执行人！");
        return QSqlRecord();
    }

    // --- 数据收集 ---
    // 1. 设置通用字段
    m_record.setValue("AssigneeID", ui->assigneeComboBox->currentData().toInt()); // 从 ComboBox 的 UserData 角色获取ID
    m_record.setValue("TaskName", ui->taskNameLineEdit->text().trimmed());
    m_record.setValue("TaskType", ui->taskTypeComboBox->currentText());
    m_record.setValue("DueDate", ui->dueDateEdit->date());
    m_record.setValue("TaskDescription", ui->descriptionTextEdit->toPlainText());
    m_record.setValue("Evaluation", ui->evaluationTextEdit->toPlainText());

    // 2. 根据任务类型处理特殊字段
    QString taskType = ui->taskTypeComboBox->currentText();
    if (taskType == "出差任务") {
        m_record.setValue("Destination", ui->destinationLineEdit->text());
        m_record.setValue("StartDate", ui->startDateEdit->date());
        if (m_mode == AddTask) m_record.setValue("Status", "待审批"); // 出差任务初始状态为待审批
    } else { // 常规任务
        if (m_mode == AddTask) m_record.setValue("Status", "未开始"); // 常规任务初始状态为未开始
    }

    // 3. 根据模式设置分派人和状态
    if (m_mode == AddTask) {
        m_record.setValue("AssignerID", m_currentUserId); // 分派人是当前用户
    }
    if (m_mode == EditTask) {
        // 编辑模式下，状态从下拉框获取
        m_record.setValue("Status", ui->statusComboBox->currentText());
    }

    return m_record; // 返回填充好数据的记录
}

/***************************************************************************************************
** 私有槽函数与辅助函数 (Private Slots and Helper Functions)         **
***************************************************************************************************/

/**
 * @brief 在编辑模式下，用 m_record 的数据填充UI控件
 */
void TaskEditDialog::fillData()
{
    ui->taskNameLineEdit->setText(m_record.value("TaskName").toString());
    ui->taskTypeComboBox->setCurrentText(m_record.value("TaskType").toString());
    ui->descriptionTextEdit->setPlainText(m_record.value("TaskDescription").toString());
    ui->dueDateEdit->setDate(m_record.value("DueDate").toDate());
    ui->statusComboBox->setCurrentText(m_record.value("Status").toString());
    ui->evaluationTextEdit->setPlainText(m_record.value("Evaluation").toString());
    ui->destinationLineEdit->setText(m_record.value("Destination").toString());
    ui->startDateEdit->setDate(m_record.value("StartDate").toDate());

    // 分派人姓名需要单独查询
    QString assignerName;
    DatabaseManager::instance()->getEmployeeNameById(m_record.value("AssignerID").toInt(), assignerName);
    ui->assignerLineEdit->setText(assignerName);
    ui->assignerLineEdit->setReadOnly(true);

    // 根据执行人ID，在下拉框中找到并选中对应的项
    int assigneeId = m_record.value("AssigneeID").toInt();
    int index = ui->assigneeComboBox->findData(assigneeId);
    ui->assigneeComboBox->setCurrentIndex(index);
}

/**
 * @brief 当任务类型下拉框变化时，更新UI
 */
void TaskEditDialog::onTaskTypeChanged(const QString &taskType)
{
    updateUiForTaskType(taskType);
}

/**
 * @brief 初始化界面上的所有下拉框
 */
void TaskEditDialog::initComboBoxes()
{
    // 静态列表
    ui->taskTypeComboBox->clear();
    ui->taskTypeComboBox->addItems({"常规任务", "出差任务"});
    ui->statusComboBox->clear();
    ui->statusComboBox->addItems({"未开始", "进行中", "已完成", "待审批"}); // 添加待审批状态

    // 动态列表：从数据库加载所有在职员工（除了自己）作为执行人候选
    ui->assigneeComboBox->clear();
    QList<QSqlRecord> employees = DatabaseManager::instance()->getAllActiveEmployees(m_currentUserId);
    for(const QSqlRecord& employee : employees) {
        ui->assigneeComboBox->addItem(employee.value("Name").toString(), employee.value("EmployeeID"));
    }
}

/**
 * @brief 根据任务类型，显示或隐藏特定于出差任务的UI控件
 */
void TaskEditDialog::updateUiForTaskType(const QString &taskType)
{
    bool isTrip = (taskType == "出差任务");

    // 使用 setVisible 来控制控件的可见性
    if (ui->destinationLineEdit) ui->destinationLineEdit->setVisible(isTrip);
    if (ui->labelForDestination) ui->labelForDestination->setVisible(isTrip);
    if (ui->startDateEdit) ui->startDateEdit->setVisible(isTrip);
    if (ui->labelForStartDate) ui->labelForStartDate->setVisible(isTrip);
}
// #include "taskeditdialog.h"
// #include "ui_taskeditdialog.h"
// #include "usersession.h"
// #include <QSqlQuery>
// #include <QVariant>
// #include <QMessageBox>
// #include <QDate>
// #include <QDebug>
// #include <QSqlError>

// TaskEditDialog::TaskEditDialog(const QSqlRecord &record, int currentUserId, TaskMode mode, QWidget *parent) :
//     QDialog(parent),
//     ui(new Ui::TaskEditDialog),
//     m_record(record),
//     m_currentUserId(currentUserId),
//     m_mode(mode)
// {
//     ui->setupUi(this);
//     initComboBoxes();
//     connect(ui->taskTypeComboBox, &QComboBox::currentTextChanged, this, &TaskEditDialog::onTaskTypeChanged);

//     if (m_mode == AddTask) {
//         setWindowTitle("分派新任务");
//         ui->assignerLineEdit->setText(UserSession::instance()->name());
//         ui->assignerLineEdit->setReadOnly(true);
//         ui->dueDateEdit->setDate(QDate::currentDate().addDays(7));
//         ui->evaluationTextEdit->setReadOnly(true);
//         ui->statusComboBox->setEnabled(false);
//     } else {
//         setWindowTitle("任务详情");
//         fillData();
//         bool isAssigner = (m_currentUserId == m_record.value("AssignerID").toInt());
//         bool isAssignee = (m_currentUserId == m_record.value("AssigneeID").toInt());
//         QString status = m_record.value("Status").toString();
//         bool canEditCoreInfo = isAssigner && (status != "已完成");
//         ui->taskNameLineEdit->setReadOnly(!canEditCoreInfo);
//         ui->taskTypeComboBox->setEnabled(canEditCoreInfo);
//         ui->assigneeComboBox->setEnabled(canEditCoreInfo);
//         ui->dueDateEdit->setReadOnly(!canEditCoreInfo);
//         ui->descriptionTextEdit->setReadOnly(!canEditCoreInfo);
//         ui->destinationLineEdit->setReadOnly(!canEditCoreInfo);
//         ui->startDateEdit->setReadOnly(!canEditCoreInfo);
//         ui->statusComboBox->setEnabled(isAssignee && (status != "已完成"));
//         ui->evaluationTextEdit->setReadOnly(!isAssigner || (status != "已完成"));
//     }
//     updateUiForTaskType(ui->taskTypeComboBox->currentText());
// }

// TaskEditDialog::~TaskEditDialog()
// {
//     delete ui;
// }

// void TaskEditDialog::fillData()
// {
//     ui->taskNameLineEdit->setText(m_record.value("TaskName").toString());
//     ui->taskTypeComboBox->setCurrentText(m_record.value("TaskType").toString());
//     ui->descriptionTextEdit->setPlainText(m_record.value("TaskDescription").toString());
//     ui->dueDateEdit->setDate(m_record.value("DueDate").toDate());
//     ui->statusComboBox->setCurrentText(m_record.value("Status").toString());
//     ui->evaluationTextEdit->setPlainText(m_record.value("Evaluation").toString());
//     ui->destinationLineEdit->setText(m_record.value("Destination").toString());
//     ui->startDateEdit->setDate(m_record.value("StartDate").toDate());
//     QSqlQuery assignerQuery;
//     assignerQuery.prepare("SELECT Name FROM basicinfo WHERE EmployeeID = ?");
//     assignerQuery.addBindValue(m_record.value("AssignerID"));
//     assignerQuery.exec();
//     if(assignerQuery.next()) ui->assignerLineEdit->setText(assignerQuery.value(0).toString());
//     ui->assignerLineEdit->setReadOnly(true);
//     int assigneeId = m_record.value("AssigneeID").toInt();
//     int index = ui->assigneeComboBox->findData(assigneeId);
//     ui->assigneeComboBox->setCurrentIndex(index);
// }

// QSqlRecord TaskEditDialog::getRecordData()
// {
//     if (ui->taskNameLineEdit->text().trimmed().isEmpty()) {
//         QMessageBox::warning(this, "输入错误", "任务/出差名称不能为空！");
//         return QSqlRecord();
//     }
//     // 强制要求必须选择一个执行人
//     if (ui->assigneeComboBox->currentIndex() == -1) {
//         QMessageBox::warning(this, "输入错误", "必须为任务选择一个执行人！");
//         return QSqlRecord();
//     }

//     // 1. 统一从下拉框获取执行人ID，这行代码对所有任务类型都生效
//     m_record.setValue("AssigneeID", ui->assigneeComboBox->currentData().toInt());

//     // 2. 设置通用字段
//     m_record.setValue("TaskName", ui->taskNameLineEdit->text().trimmed());
//     m_record.setValue("TaskType", ui->taskTypeComboBox->currentText());
//     m_record.setValue("DueDate", ui->dueDateEdit->date());
//     m_record.setValue("TaskDescription", ui->descriptionTextEdit->toPlainText());
//     m_record.setValue("Evaluation", ui->evaluationTextEdit->toPlainText());

//     // 3. 根据任务类型处理各自的特殊字段和默认状态
//     QString taskType = ui->taskTypeComboBox->currentText();
//     if (taskType == "出差任务") {
//         m_record.setValue("Destination", ui->destinationLineEdit->text());
//         m_record.setValue("StartDate", ui->startDateEdit->date());
//         if (m_mode == AddTask) m_record.setValue("Status", "待审批");
//     } else { // 常规任务
//         if (m_mode == AddTask) m_record.setValue("Status", "未开始");
//     }

//     // 4. 设置分派人和状态（如果是编辑模式）
//     if (m_mode == AddTask) {
//         m_record.setValue("AssignerID", m_currentUserId);
//     }
//     if (m_mode == EditTask) {
//         m_record.setValue("Status", ui->statusComboBox->currentText());
//     }

//     return m_record;
// }

// void TaskEditDialog::onTaskTypeChanged(const QString &taskType)
// {
//     updateUiForTaskType(taskType);
// }

// void TaskEditDialog::initComboBoxes()
// {
//     ui->taskTypeComboBox->clear();
//     ui->statusComboBox->clear();
//     ui->assigneeComboBox->clear();
//     ui->taskTypeComboBox->addItems({"常规任务", "出差任务"});
//     ui->statusComboBox->addItems({"未开始", "进行中", "已完成"});
//     QSqlQuery query;
//     query.prepare("SELECT EmployeeID, Name FROM basicinfo WHERE EmployeeStatus = '在职' AND EmployeeID != ?");
//     query.addBindValue(m_currentUserId);
//     if (!query.exec()) {
//         qDebug() << "加载执行人列表失败:" << query.lastError();
//         return;
//     }
//     while(query.next())
//     {
//         ui->assigneeComboBox->addItem(query.value(1).toString(), query.value(0));
//     }
// }

// void TaskEditDialog::updateUiForTaskType(const QString &taskType)
// {
//     bool isTrip = (taskType == "出差任务");
//     // 执行人选择框不再受任务类型影响，所以相关代码已移除

//     // 只根据是否是出差任务，来显示或隐藏出差相关的控件
//     if (ui->destinationLineEdit) ui->destinationLineEdit->setVisible(isTrip);
//     if (ui->labelForDestination) ui->labelForDestination->setVisible(isTrip);
//     if (ui->startDateEdit) ui->startDateEdit->setVisible(isTrip);
//     if (ui->labelForStartDate) ui->labelForStartDate->setVisible(isTrip);
// }






