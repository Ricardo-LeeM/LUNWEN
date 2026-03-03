/***************************************************************************************************
** 头文件依赖 (Header Dependencies)                                  **
***************************************************************************************************/
#include "personalinterface.h"
#include "ui_personalinterface.h"
#include "usersession.h"
#include "taskcardwidget.h"
#include "taskeditdialog.h"
#include "databasemanager.h" // <-- 引入DatabaseManager

#include <QTimer>
#include <QVBoxLayout>
#include <QMessageBox>
#include <QFileDialog>
#include <QStandardItemModel>
#include <QSqlRecord>
#include <QDir>

// 仅在Windows平台下包含ActiveQt模块
#ifdef Q_OS_WIN
#include <ActiveQt/QAxObject>
#endif

/***************************************************************************************************
** 构造函数与析构函数                                            **
***************************************************************************************************/

PersonalInterface::PersonalInterface(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::PersonalInterface)
    , salaryHistoryModel(nullptr) // 初始化指针
{
    ui->setupUi(this);
    this->setWindowTitle("个人中心");

    UserSession *session = UserSession::instance();
    if (session->employeeID() == -1) {
        QMessageBox::warning(this, "未登录", "无法加载个人信息，用户未登录。");
        return;
    }

    loadCurrentUserData();

    ui->lineEditName->setText(session->name());
    ui->lineEditName->setReadOnly(true);
    ui->lineEditName_1->setText(session->name());
    ui->lineEditName_1->setReadOnly(true);
    ui->lineEditName_2->setText(session->name());
    ui->lineEditName_2->setReadOnly(true);

    updateAllCounts();
    initializeSalaryTable();
    setupTaskboard();
    loadTaskCards();
}

PersonalInterface::~PersonalInterface()
{
    delete ui;
}

/***************************************************************************************************
** 私有函数实现 (Private Function Implementations)                       **
***************************************************************************************************/

void PersonalInterface::updateAllCounts()
{
    int financialCount, vacationCount, projectionCount;
    if (DatabaseManager::instance()->getApplicationCounts(UserSession::instance()->employeeID(), financialCount, vacationCount, projectionCount)) {
        ui->labelFinancialNum->setText("我的申报次数：" + QString::number(financialCount));
        ui->labelVacationNum->setText("我的申报次数：" + QString::number(vacationCount));
        ui->labelProjectionNum->setText("我的申报次数：" + QString::number(projectionCount));
    }
}

void PersonalInterface::initializeSalaryTable()
{
    salaryHistoryModel = new QStandardItemModel(this);
    ui->tableView->setModel(salaryHistoryModel);

    QStringList headers = {"薪资周期", "基本工资", "绩效奖", "津贴", "应发工资", "社保公积金", "个人所得税", "实发工资", "备注"};
    salaryHistoryModel->setHorizontalHeaderLabels(headers);

    QList<QSqlRecord> records = DatabaseManager::instance()->getSalaryHistory(UserSession::instance()->employeeID());

    for (const QSqlRecord& record : records) {
        QList<QStandardItem*> rowItems;
        rowItems << new QStandardItem(record.value("payroll_period").toString());
        rowItems << new QStandardItem(QString::number(record.value("base_salary").toDouble(), 'f', 2));
        rowItems << new QStandardItem(QString::number(record.value("bonus_performance").toDouble(), 'f', 2));
        rowItems << new QStandardItem(QString::number(record.value("allowance").toDouble(), 'f', 2));
        rowItems << new QStandardItem(QString::number(record.value("gross_salary").toDouble(), 'f', 2));
        rowItems << new QStandardItem(QString::number(record.value("deduction_social_security").toDouble(), 'f', 2));
        rowItems << new QStandardItem(QString::number(record.value("deduction_tax").toDouble(), 'f', 2));
        rowItems << new QStandardItem(QString::number(record.value("net_salary").toDouble(), 'f', 2));
        rowItems << new QStandardItem(record.value("notes").toString());
        salaryHistoryModel->appendRow(rowItems);
    }

    ui->tableView->setEditTriggers(QAbstractItemView::NoEditTriggers);
    ui->tableView->resizeColumnsToContents();
    ui->tableView->horizontalHeader()->setStretchLastSection(true);
}

void PersonalInterface::setupTaskboard()
{
    if (ui->toDoColumnWidget->layout() == nullptr) ui->toDoColumnWidget->setLayout(new QVBoxLayout());
    if (ui->inProgressColumnWidget->layout() == nullptr) ui->inProgressColumnWidget->setLayout(new QVBoxLayout());
    if (ui->doneColumnWidget->layout() == nullptr) ui->doneColumnWidget->setLayout(new QVBoxLayout());
}

void PersonalInterface::loadTaskCards()
{
    clearLayout(ui->toDoColumnWidget->layout());
    clearLayout(ui->inProgressColumnWidget->layout());
    clearLayout(ui->doneColumnWidget->layout());

    QList<QSqlRecord> taskRecords = DatabaseManager::instance()->getTasksForEmployee(UserSession::instance()->employeeID());

    for (const QSqlRecord& record : taskRecords) {
        TaskCardWidget *card = new TaskCardWidget(this);
        card->setTaskData(record);
        card->setFixedHeight(170);

        connect(card, &TaskCardWidget::requestMove, this, &PersonalInterface::moveTaskCard);
        connect(card, &TaskCardWidget::viewDetailsRequested, this, &PersonalInterface::showTaskDetails);

        // QString status = record.value("Status").toString();
        // if (status == "未开始") {
        //     ui->toDoColumnWidget->layout()->addWidget(card);
        // } else if (status == "进行中") {
        //     ui->inProgressColumnWidget->layout()->addWidget(card);
        // } else if (status == "已完成") {
        //     ui->doneColumnWidget->layout()->addWidget(card);
        // } else {
        //     delete card;
        // }
        QString status = record.value("Status").toString();
        if (status == "未开始" || status == "待审批") { // 允许显示待审批状态
            ui->toDoColumnWidget->layout()->addWidget(card);
        } else if (status == "进行中") {
            ui->inProgressColumnWidget->layout()->addWidget(card);
        } else if (status == "已完成") {
            ui->doneColumnWidget->layout()->addWidget(card);
        } else {
            delete card;
        }
    }

    static_cast<QVBoxLayout*>(ui->toDoColumnWidget->layout())->addStretch(1);
    static_cast<QVBoxLayout*>(ui->inProgressColumnWidget->layout())->addStretch(1);
    static_cast<QVBoxLayout*>(ui->doneColumnWidget->layout())->addStretch(1);
}

void PersonalInterface::on_tbtnFinancial_clicked()
{
    QString content = ui->textEdit->toPlainText().trimmed();
    if (content.isEmpty()) {
        QMessageBox::warning(this, "内容为空", "申报内容不能为空！");
        return;
    }

    if (QMessageBox::question(this, "确认提交", "您确定要提交财务申报吗？", QMessageBox::Yes | QMessageBox::No) == QMessageBox::Yes) {
        if (DatabaseManager::instance()->addDeclaration(UserSession::instance()->name(), "财务申报", content)) {
            QMessageBox::information(this, "成功", "财务申报已成功提交。");
            ui->textEdit->clear();
            updateAllCounts();
        } else {
            QMessageBox::critical(this, "失败", "提交失败，请稍后重试。");
        }
    }
}

void PersonalInterface::on_tbtnVacation_clicked()
{
    QString content = ui->textEdit_1->toPlainText().trimmed();
    if (content.isEmpty()) {
        QMessageBox::warning(this, "内容为空", "申报内容不能为空！");
        return;
    }

    if (QMessageBox::question(this, "确认提交", "您确定要提交休假申报吗？", QMessageBox::Yes | QMessageBox::No) == QMessageBox::Yes) {
        if (DatabaseManager::instance()->addDeclaration(UserSession::instance()->name(), "休假申报", content)) {
            QMessageBox::information(this, "成功", "休假申报已成功提交。");
            ui->textEdit_1->clear();
            updateAllCounts();
        } else {
            QMessageBox::critical(this, "失败", "提交失败，请稍后重试。");
        }
    }
}

void PersonalInterface::on_tbtnProjection_clicked()
{
    QString content = ui->textEdit_2->toPlainText().trimmed();
    if (content.isEmpty()) {
        QMessageBox::warning(this, "内容为空", "申报内容不能为空！");
        return;
    }

    if (QMessageBox::question(this, "确认提交", "您确定要提交立项申报吗？", QMessageBox::Yes | QMessageBox::No) == QMessageBox::Yes) {
        if (DatabaseManager::instance()->addDeclaration(UserSession::instance()->name(), "立项申报", content)) {
            QMessageBox::information(this, "成功", "立项申报已成功提交。");
            ui->textEdit_2->clear();
            updateAllCounts();
        } else {
            QMessageBox::critical(this, "失败", "提交失败，请稍后重试。");
        }
    }
}

void PersonalInterface::moveTaskCard(TaskCardWidget *card, const QString &newStatus)
{
    if (!card) return;

    if (!DatabaseManager::instance()->updateTaskStatus(card->getTaskId(), newStatus)) {
        QMessageBox::critical(this, "数据库错误", "无法更新任务状态！");
        return;
    }

    card->setParent(nullptr);
    card->updateStatus(newStatus);
    QVBoxLayout* targetLayout = nullptr;
    if (newStatus == "进行中") targetLayout = static_cast<QVBoxLayout*>(ui->inProgressColumnWidget->layout());
    else if (newStatus == "已完成") targetLayout = static_cast<QVBoxLayout*>(ui->doneColumnWidget->layout());

    if (targetLayout) {
        targetLayout->insertWidget(targetLayout->count() - 1, card);
    }
}

void PersonalInterface::showTaskDetails(const QSqlRecord &record)
{
    TaskEditDialog dialog(record, UserSession::instance()->employeeID(), TaskEditDialog::EditTask, this);
    if (dialog.exec() == QDialog::Accepted) {
        QTimer::singleShot(0, this, &PersonalInterface::loadTaskCards);
    }
}

void PersonalInterface::on_shuaxinButton_clicked()
{
    loadTaskCards();
    QMessageBox::information(this, "刷新成功", "任务列表已更新！");
}

void PersonalInterface::loadCurrentUserData() {
    UserSession *session = UserSession::instance();
    if(ui->labelName) ui->labelName->setText("姓名: " + session->name());
    if(ui->labelEmployeeID) ui->labelEmployeeID->setText("工号: " + QString::number(session->employeeID()));
    if(ui->labelDepartment) ui->labelDepartment->setText("部门: " + session->department());
    if(ui->labelEmployeeStatus) ui->labelEmployeeStatus->setText("员工状态: " + session->employeeStatus());
    if(ui->labelIsPartyMember) ui->labelIsPartyMember->setText("是否党员: " + (session->isPartyMember() ? QString("是") : QString("否")));
    if(ui->labelWorkNature) ui->labelWorkNature->setText("工作性质: " + session->workNature());
    if(ui->labelPosition) ui->labelPosition->setText("岗位: " + session->position());
    if(ui->labelJoinDate) ui->labelJoinDate->setText("入职日期: " + session->joinDate().toString("yyyy-MM-dd"));
    if(ui->labWelcome) ui->labWelcome->setText("欢迎您，" + session->name() + "！");
    if (ui->staffImage) {
        if (!session->userPixmap().isNull()) {
            ui->staffImage->setPixmap(session->userPixmap().scaled(ui->staffImage->size(), Qt::KeepAspectRatio, Qt::SmoothTransformation));
        } else {
            ui->staffImage->setText("无照片");
        }
    }
}

void PersonalInterface::clearLayout(QLayout* layout) {
    if (layout == nullptr) return;
    QLayoutItem* item;
    while ((item = layout->takeAt(0)) != nullptr) {
        delete item->widget();
        delete item;
    }
}

void PersonalInterface::on_btnFromExcel_clicked() { /* 保持原样 */ }
void PersonalInterface::on_btnToExcel_clicked() { /* 保持原样 */ }
