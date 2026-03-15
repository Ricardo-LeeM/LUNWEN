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
#include <QSqlQuery>
#include <QDir>
#include <QIcon>
#include <QScrollBar>
#include <QVBoxLayout>
#include <QLabel>

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
    QIcon qIcon(":/C:/images/xiyang.png");
    if (qIcon.isNull()) {
        qDebug() << "图标加载失败，请检查资源文件或路径";
    }
    QPixmap pixmap(":/C:/images/xiyang.png");
    qDebug() << "图片是否为空？" << pixmap.isNull();
    this->setWindowIcon(qIcon);


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

    ui->textBrowserChat->setEnabled(false);
    ui->textEditMessage->setEnabled(false);
    ui->btnSendMessage->setEnabled(false);

    loadContacts(); // 拉取员工列表

    // 初始化轮询定时器（每2秒检查一次新消息）
    chatTimer = new QTimer(this);
    connect(chatTimer, &QTimer::timeout, this, &PersonalInterface::fetchNewMessages);
    chatTimer->start(2000);

    // 绑定信号
    connect(ui->btnSendMessage, &QPushButton::clicked, this, &PersonalInterface::on_btnSendMessage_clicked);
    connect(ui->listWidgetContacts, &QListWidget::itemClicked, this, &PersonalInterface::on_listWidgetContacts_itemClicked);
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

void PersonalInterface::loadContacts()
{
    UserSession *session = UserSession::instance();
    if (session->employeeID() == -1) return;

    ui->listWidgetContacts->clear();

    // 1. 确保图标尺寸足够大，间距合适
    ui->listWidgetContacts->setIconSize(QSize(50, 50));
    ui->listWidgetContacts->setSpacing(5);

    QSqlQuery query = DatabaseManager::instance()->getOtherEmployees(session->employeeID());

    while (query.next()) {
        int empId = query.value(0).toInt();
        QString name = query.value(1).toString();
        QString dept = query.value(2).toString();
        QString position = query.value(3).toString();
        QByteArray imgData = query.value(4).toByteArray();

        // 打印调试信息，看数据库到底查出东西没有
        qDebug() << "拉取联系人:" << name << "部门:" << dept << "照片大小:" << imgData.size();

        // 组装文本
        QString displayText = QString("%1\n%2 | %3").arg(name, dept, position);

        QListWidgetItem *item = new QListWidgetItem(displayText);
        item->setData(Qt::UserRole, empId);

        //核心修复：强制让这个列表项的高度变成 60 像素，确保能容纳两行文字和头像
        item->setSizeHint(QSize(0, 60));

        QPixmap pixmap;
        if (!imgData.isEmpty() && pixmap.loadFromData(imgData)) {
            item->setIcon(QIcon(pixmap));
        } else {
            // 如果数据库里这个人没有存照片，这里会跳过。
            // 建议：你可以取消下面这行的注释，给没有照片的人加个默认头像（需确认资源文件路径）
            item->setIcon(QIcon(":/C:/images/xiyang.png"));
        }

        ui->listWidgetContacts->addItem(item);
    }
}


// 点击加载联系人聊天记录
void PersonalInterface::loadChatHistory(int targetUserId)
{
    ui->textBrowserChat->clear();
    UserSession *session = UserSession::instance();
    int myId = session->employeeID();

    // 查的是聊天记录，不是 basicinfo！
    QSqlQuery query = DatabaseManager::instance()->getChatHistory(myId, targetUserId);

    while (query.next()) {
        int senderId = query.value(0).toInt();
        QString content = query.value(1).toString();
        bool isMine = (senderId == myId);

        appendChatMessage(content, isMine);
    }

    QScrollBar *scrollBar = ui->textBrowserChat->verticalScrollBar();
    scrollBar->setValue(scrollBar->maximum());
}


// 发送消息
void PersonalInterface::on_btnSendMessage_clicked()
{
    if (currentChatUserId == -1) return;

    QString content = ui->textEditMessage->toPlainText().trimmed();
    if (content.isEmpty()) return;

    int myId = UserSession::instance()->employeeID();

    // 调用 DatabaseManager 执行插入
    if (DatabaseManager::instance()->insertChatMessage(myId, currentChatUserId, content)) {
        // 数据库插入成功后，本地立刻显示
        appendChatMessage(content, true);
        ui->textEditMessage->clear();

        QScrollBar *scrollBar = ui->textBrowserChat->verticalScrollBar();
        scrollBar->setValue(scrollBar->maximum());
    } else {
        QMessageBox::warning(this, "发送失败", "数据库写入失败，请检查网络！");
    }
}


void PersonalInterface::fetchNewMessages()
{
    int myId = UserSession::instance()->employeeID();
    if (myId == -1) return;

    // 调用 DatabaseManager 获取发给我的未读消息
    QSqlQuery query = DatabaseManager::instance()->getUnreadMessages(myId);

    while (query.next()) {
        int msgId = query.value(0).toInt();
        int senderId = query.value(1).toInt();
        QString content = query.value(2).toString();

        // 如果这条消息正好是当前正在聊天的人发来的
        if (senderId == currentChatUserId) {
            appendChatMessage(content, false);

            QScrollBar *scrollBar = ui->textBrowserChat->verticalScrollBar();
            scrollBar->setValue(scrollBar->maximum());

            // 将该消息标记为已读
            DatabaseManager::instance()->markMessageAsRead(msgId);
        }
    }
}

// 极其简单的 HTML 渲染，区分左右气泡
void PersonalInterface::appendChatMessage(const QString &text, bool isMine)
{
    // 注意：在 homepage.cpp 中，记得把这行的方法名改为 void HomePage::appendChatMessage
    QString html;

    // 把换行符转换为 HTML 的 <br>，防止长文本变成单行
    QString safeText = text.toHtmlEscaped().replace("\n", "<br>");

    if (isMine) {
        // 我发出的消息：整体靠右
        // 使用一个宽度 100% 的外层表格占满整行，内层表格做背景气泡
        html = QString(
                   "<table width='100%' border='0' cellpadding='0' cellspacing='0' style='margin-top: 5px; margin-bottom: 5px;'>"
                   "<tr><td align='right'>"
                   // 第一行：名字加冒号，靠右，灰色小字
                   "<div style='color: #888888; font-size: 12px; margin-bottom: 2px;'>我：</div>"
                   // 第二行：气泡消息内容 (通过 cellpadding 控制内边距，bgcolor 控制背景色)
                   "<table border='0' cellpadding='8' cellspacing='0'>"
                   "<tr><td bgcolor='#95ec69' style='border-radius: 6px;'>"
                   "<span style='color: black; font-size: 14px;'>%1</span>"
                   "</td></tr>"
                   "</table>"
                   "</td></tr>"
                   "</table>"
                   ).arg(safeText);
    } else {
        // 对方发出的消息：整体靠左
        html = QString(
                   "<table width='100%' border='0' cellpadding='0' cellspacing='0' style='margin-top: 5px; margin-bottom: 5px;'>"
                   "<tr><td align='left'>"
                   // 第一行：名字加冒号，靠左，灰色小字
                   "<div style='color: #888888; font-size: 12px; margin-bottom: 2px;'>对方：</div>"
                   // 第二行：气泡消息内容 (内层 table 也要加上 align='left' 才能靠左对齐)
                   "<table border='0' cellpadding='8' cellspacing='0' align='left'>"
                   "<tr><td bgcolor='#ffffff' style='border-radius: 6px; border: 1px solid #d0d0d0;'>"
                   "<span style='color: black; font-size: 14px;'>%1</span>"
                   "</td></tr>"
                   "</table>"
                   "</td></tr>"
                   "</table>"
                   ).arg(safeText);
    }

    ui->textBrowserChat->append(html);
}

void PersonalInterface::on_listWidgetContacts_itemClicked(QListWidgetItem *item)
{
    if (!item) return;

    // 获取点击的员工工号
    currentChatUserId = item->data(Qt::UserRole).toInt();
    ui->labChatTarget->setText("与 " + item->text() + " 聊天中...");

    // 激活输入面板
    ui->textBrowserChat->setEnabled(true);
    ui->textEditMessage->setEnabled(true);
    ui->btnSendMessage->setEnabled(true);

    loadChatHistory(currentChatUserId);
}

void PersonalInterface::on_btnFromExcel_clicked() { /* 保持原样 */ }
void PersonalInterface::on_btnToExcel_clicked() { /* 保持原样 */ }
