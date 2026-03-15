#include "homepage.h"
#include "ui_homepage.h"
#include "usersession.h" // 应当优先包含自定义的单例类头文件

// Qt 核心模块
#include <QTimer>
#include <QMenu>
#include <QDate>
#include <QTime>
#include <QDir>
#include <QFileDialog>
#include <QDebug>

// Qt 数据库模块
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QSqlTableModel>

// Qt UI 相关模块
#include <QMessageBox>
#include <QTableView>
#include <QHeaderView>
#include <QVBoxLayout>
#include <QScrollBar>

// Qt Windows 平台特定模块 (用于操作Excel)
#ifdef Q_OS_WIN
#include <ActiveQt/QAxObject>
#include <ActiveQt/QAxWidget>
#endif

HomePage::HomePage(QWidget *parent)
    : QWidget(parent)
    , ui(new Ui::HomePage)
{
    ui->setupUi(this);

    // 1. 加载当前登录用户的信息到界面
    loadCurrentUserData();

    // 2. 初始化财务申报的下拉菜单
    QMenu *menuSel = new QMenu(this);
    menuSel->addAction(ui->actReimbursementApplication);
    menuSel->addAction(ui->actLoanApplication);
    menuSel->addAction(ui->actSalaryAdjustmentApplication);
    ui->tbtnFinancial->setMenu(menuSel);

    // 3. 设置时间问候语标签样式，并启动定时器
    ui->labTime->setAlignment(Qt::AlignCenter);
    ui->labTime->setStyleSheet("font-size: 24px; color: black;");
    ui->labTime->setText(getGreeting()); // 立即显示一次
    timer = new QTimer(this);
    connect(timer, &QTimer::timeout, this, &HomePage::updateTime);
    timer->start(60000); // 每分钟更新一次

    // 4. 初始化薪资历史记录表格
    initializeSalaryTable();

    // 5. 设置个人申请页面的申报人信息
    UserSession *session = UserSession::instance();
    if (session->employeeID() != -1) { // 确保用户已登录
        // 自动填写当前用户名并设为只读
        ui->lineEditName->setText(session->name());
        ui->lineEditName->setReadOnly(true);
        ui->lineEditName_2->setText(session->name());
        ui->lineEditName_2->setReadOnly(true);
        ui->lineEditName_3->setText(session->name());
        ui->lineEditName_3->setReadOnly(true);


        // 初始化时更新一次申报次数显示
        updateFinancialNumSlot();
        updateVacationNumslot();
        updateProjectionNumslot();
    }

    // 6.隐藏聊天界面的某些部件直到选择了聊天对象
    ui->textBrowserChat->setEnabled(false);
    ui->textEditMessage->setEnabled(false);
    ui->btnSendMessage->setEnabled(false);

    loadContacts(); // 拉取员工列表

    // 初始化轮询定时器（每2秒检查一次新消息）
    chatTimer = new QTimer(this);
    connect(chatTimer, &QTimer::timeout, this, &HomePage::fetchNewMessages);
    chatTimer->start(2000);
    // 手动绑定发送按钮的点击信号
    connect(ui->btnSendMessage, &QPushButton::clicked, this, &HomePage::on_btnSendMessage_clicked);
    connect(ui->listWidgetContacts, &QListWidget::itemClicked, this, &HomePage::on_listWidgetContacts_itemClicked);
}

HomePage::~HomePage()
{
    delete ui;
}

void HomePage::initializeSalaryTable()
{
    UserSession *session = UserSession::instance();
    if (session->employeeID() == -1) {
        // 如果用户未登录，则不加载表格数据
        return;
    }

    // 创建 QSqlTableModel 实例
    salaryHistoryModel = new QSqlTableModel(this);
    salaryHistoryModel->setTable("PayrollRecords"); // 设置要操作的数据库表

    // 设置过滤器，只显示当前登录用户的记录
    QString filter = QString("employee_id = %1").arg(session->employeeID());
    salaryHistoryModel->setFilter(filter);

    // 按薪资周期（payroll_period）降序排序，最新的记录在最上面
    salaryHistoryModel->setSort(salaryHistoryModel->fieldIndex("payroll_period"), Qt::DescendingOrder);

    // 执行查询，加载数据
    salaryHistoryModel->select();

    // 设置表头显示的中文文本
    salaryHistoryModel->setHeaderData(salaryHistoryModel->fieldIndex("payroll_period"), Qt::Horizontal, "薪资周期");
    salaryHistoryModel->setHeaderData(salaryHistoryModel->fieldIndex("base_salary"), Qt::Horizontal, "基本工资");
    salaryHistoryModel->setHeaderData(salaryHistoryModel->fieldIndex("bonus_attendance"), Qt::Horizontal, "全勤奖");
    salaryHistoryModel->setHeaderData(salaryHistoryModel->fieldIndex("bonus_performance"), Qt::Horizontal, "绩效奖");
    salaryHistoryModel->setHeaderData(salaryHistoryModel->fieldIndex("allowance"), Qt::Horizontal, "津贴");
    salaryHistoryModel->setHeaderData(salaryHistoryModel->fieldIndex("deduction_late"), Qt::Horizontal, "迟到扣款");
    salaryHistoryModel->setHeaderData(salaryHistoryModel->fieldIndex("deduction_absence"), Qt::Horizontal, "缺勤扣款");
    salaryHistoryModel->setHeaderData(salaryHistoryModel->fieldIndex("deduction_social_security"), Qt::Horizontal, "社保公积金");
    salaryHistoryModel->setHeaderData(salaryHistoryModel->fieldIndex("deduction_tax"), Qt::Horizontal, "个税");
    salaryHistoryModel->setHeaderData(salaryHistoryModel->fieldIndex("deduction_other"), Qt::Horizontal, "其他扣款");
    salaryHistoryModel->setHeaderData(salaryHistoryModel->fieldIndex("gross_salary"), Qt::Horizontal, "应发工资");
    salaryHistoryModel->setHeaderData(salaryHistoryModel->fieldIndex("net_salary"), Qt::Horizontal, "实发工资");
    salaryHistoryModel->setHeaderData(salaryHistoryModel->fieldIndex("notes"), Qt::Horizontal, "备注");

    // 将模型设置到 QTableView 控件上
    ui->tableView->setModel(salaryHistoryModel);

    // 隐藏对员工无意义或敏感的列（如数据库记录ID，员工ID等）
    ui->tableView->hideColumn(salaryHistoryModel->fieldIndex("record_id"));
    ui->tableView->hideColumn(salaryHistoryModel->fieldIndex("employee_id"));
    ui->tableView->hideColumn(salaryHistoryModel->fieldIndex("employee_name"));
    ui->tableView->hideColumn(salaryHistoryModel->fieldIndex("department"));

    // 设置表格为只读，防止员工误修改自己的工资记录
    ui->tableView->setEditTriggers(QAbstractItemView::NoEditTriggers);

    // 自动调整列宽以适应内容
    ui->tableView->resizeColumnsToContents();
    // 让最后一列自动拉伸以填满剩余空间
    ui->tableView->horizontalHeader()->setStretchLastSection(true);
}

void HomePage::exportToExcel()
{
#ifdef Q_OS_WIN
    // 使用QAxObject与Excel进行COM交互，此方法仅在Windows平台有效
    QAxObject excel("Excel.Application");
    if (excel.isNull())
    {
        QMessageBox::critical(this, "错误", "无法启动 Excel 应用程序。\n请确保您的电脑已安装Microsoft Excel。");
        return;
    }
    excel.setProperty("Visible", true); // 使Excel窗口可见

    QAxObject *workbooks = excel.querySubObject("Workbooks");
    QAxObject *workbook = workbooks->querySubObject("Add");
    QAxObject *worksheet = workbook->querySubObject("Worksheets(int)", 1);

    // 从 QTableView 获取数据模型
    QAbstractItemModel *model = ui->tableView->model();
    if (!model) {
        QMessageBox::warning(this, "警告", "表格中没有数据模型可供导出。");
        return;
    }

    int colCount = model->columnCount();
    int rowCount = model->rowCount();
    int excelCol = 1; // Excel中的列计数器，用于跳过隐藏列

    // 写入表头
    for (int col = 0; col < colCount; ++col) {
        if (ui->tableView->isColumnHidden(col)) continue; // 跳过隐藏的列
        QString headerText = model->headerData(col, Qt::Horizontal).toString();
        worksheet->querySubObject("Cells(int, int)", 1, excelCol)->setProperty("Value", headerText);
        excelCol++;
    }

    // 写入数据
    for (int row = 0; row < rowCount; ++row) {
        excelCol = 1; // 重置Excel列计数器
        for (int col = 0; col < colCount; ++col) {
            if (ui->tableView->isColumnHidden(col)) continue; // 跳过隐藏的列
            QString cellData = model->data(model->index(row, col)).toString();
            worksheet->querySubObject("Cells(int, int)", row + 2, excelCol)->setProperty("Value", cellData);
            excelCol++;
        }
    }

    // 弹出文件保存对话框
    QString fileName = QFileDialog::getSaveFileName(this, "保存 Excel 文件", "", "Excel 文件 (*.xlsx)");
    if (!fileName.isEmpty()) {
        if (!fileName.endsWith(".xlsx", Qt::CaseInsensitive)) {
            fileName += ".xlsx";
        }
        // 保存工作簿并关闭
        workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(fileName));
        QMessageBox::information(this, "成功", "文件已成功导出到 Excel！");
    }
    workbook->dynamicCall("Close()");
    excel.dynamicCall("Quit()");
#else
    QMessageBox::information(this, "功能限制", "Excel导出功能目前仅支持Windows平台。");
#endif
}

void HomePage::importFromExcel()
{
    // 为保证数据安全和一致性，个人工资详情页面通常不提供导入功能。
    // 如果需要此功能，应由管理员在特定界面操作。
    QMessageBox::information(this, "功能说明", "为保证数据安全，个人工资详情页面不支持从Excel导入。");
}

void HomePage::loadCurrentUserData()
{
    UserSession *session = UserSession::instance();
    if (session->employeeID() == -1) {
        // 处理未登录或会话无效的情况
        QMessageBox::warning(this, "用户信息错误", "无法加载当前用户信息，请重新登录。");
        ui->labelName->setText("姓名: N/A");
        ui->labelEmployeeID->setText("工号: N/A");
        ui->labelDepartment->setText("部门: N/A");
        ui->labelEmployeeStatus->setText("员工状态: N/A");
        ui->labelIsPartyMember->setText("是否党员: N/A");
        ui->labelWorkNature->setText("工作性质: N/A");
        ui->labelPosition->setText("岗位: N/A");
        ui->labelJoinDate->setText("入职日期: N/A");
        ui->labWelcome->setText("欢迎！");
        ui->staffImage->clear();
        return;
    }

    // 从 UserSession 获取信息并填充到UI控件
    ui->labelName->setText("姓名: " + session->name());
    ui->labelEmployeeID->setText("工号: " + QString::number(session->employeeID()));
    ui->labelDepartment->setText("部门: " + session->department());
    ui->labelEmployeeStatus->setText("员工状态: " + session->employeeStatus());
    ui->labelIsPartyMember->setText("是否党员: " +  (session->isPartyMember() ? QString("是") : QString("否")));
    ui->labelWorkNature->setText("工作性质: " + session->workNature());
    ui->labelPosition->setText("岗位: " + session->position());
    ui->labelJoinDate->setText("入职日期: " + session->joinDate().toString("yyyy-MM-dd"));
    ui->labWelcome->setText("欢迎来到工资有限公司：" + session->name());

    // 加载并显示用户头像
    if (!session->userPixmap().isNull()) {
        ui->staffImage->setPixmap(session->userPixmap().scaled(ui->staffImage->size(), Qt::KeepAspectRatio, Qt::SmoothTransformation));
    } else {
        ui->staffImage->setText("无照片");
    }
}

void HomePage::updateTime()
{
    // 定时器触发时，更新问候语
    ui->labTime->setText(getGreeting());

}

QString HomePage::getGreeting()
{
    int hour = QTime::currentTime().hour();
    if (hour >= 6 && hour < 11) {
        return "早上好";
    } else if (hour >= 11 && hour < 14) {
        return "中午好";
    } else if (hour >= 14 && hour < 18) {
        return "下午好";
    } else {
        return "晚上好";
    }
}

void HomePage::updateFinancialNumSlot()
{
    UserSession *session = UserSession::instance();
    if (session->employeeID() == -1) {
        ui->labelFinancialNum->setText("我的申报次数：N/A");
        return;
    }

    // 从数据库查询当前用户的财务申报次数
    QSqlQuery queryGetCount;
    queryGetCount.prepare("SELECT FinancialNum FROM personal_application WHERE EmployeeID = :employeeID");
    queryGetCount.bindValue(":employeeID", session->employeeID());

    if (queryGetCount.exec() && queryGetCount.next()) {
        int count = queryGetCount.value(0).toInt();
        ui->labelFinancialNum->setText("我的申报次数：" + QString::number(count));
    } else {
        // 如果查询不到记录，默认为0次
        ui->labelFinancialNum->setText("我的申报次数：0");
    }
}

void HomePage::updateVacationNumslot()
{
    UserSession *session = UserSession::instance();
    if (session->employeeID() == -1) {
        ui->labelVacationNum->setText("我的申报次数：N/A");
        return;
    }

    // 从数据库查询当前用户的休假申报次数
    QSqlQuery queryGetCount;
    queryGetCount.prepare("SELECT VacationNum FROM personal_application WHERE EmployeeID = :employeeID");
    queryGetCount.bindValue(":employeeID", session->employeeID());

    if (queryGetCount.exec() && queryGetCount.next()) {
        int count = queryGetCount.value(0).toInt();
        ui->labelVacationNum->setText("我的申报次数：" + QString::number(count));
    } else {
        // 如果查询不到记录，默认为0次
        ui->labelVacationNum->setText("我的申报次数：0");
    }
}

void HomePage::updateProjectionNumslot()
{
    UserSession *session = UserSession::instance();
    if (session->employeeID() == -1) {
        // ui->labelProjectionNum->setText("我的申报次数：N/A"); // 假设UI中存在此标签
        return;
    }

    // 从数据库查询当前用户的立项申报次数
    QSqlQuery queryGetCount;
    queryGetCount.prepare("SELECT ProjectionNum FROM personal_application WHERE EmployeeID = :employeeID");
    queryGetCount.bindValue(":employeeID", session->employeeID());

    if (queryGetCount.exec() && queryGetCount.next()) {
        int count = queryGetCount.value(0).toInt();
        ui->labelProjectionNum->setText("我的申报次数：" + QString::number(count));
    } else {
        // 如果查询不到记录，默认为0次
        ui->labelProjectionNum->setText("我的申报次数：0");
    }
}

void HomePage::on_tbtnFinancial_clicked()
{
    UserSession *session = UserSession::instance();
    if (session->employeeID() == -1) {
        QMessageBox::warning(this, "错误!", "用户未登录，无法提交申报。");
        return;
    }

    // 获取申报信息
    QString applicantName = session->name();
    QString content = ui->textEdit->toPlainText().trimmed(); // 使用trimmed()去除首尾空白
    if (content.isEmpty()) {
        QMessageBox::warning(this, "错误!", "申报内容不能为空！");
        return;
    }

    QString type = "财务申报";

    // 弹出确认对话框
    QMessageBox::StandardButton reply;
    reply = QMessageBox::question(this, "确认提交", "您确定要提交财务申报吗？", QMessageBox::Yes | QMessageBox::No, QMessageBox::Yes);

    if (reply == QMessageBox::Yes) {
        // 提交到数据库
        if (addDeclarationToDb(applicantName, type, content)) {
            // 更新个人申报次数
            if (updatePersonalApplicationCount(applicantName, type)) {
                qDebug() << "个人财务申报总次数已更新。";
            } else {
                qDebug() << "个人财务申报总次数更新失败或未找到对应员工记录。";
            }
            QMessageBox::information(this, "提交成功", type + "已成功提交。");
            updateFinancialNumSlot(); // 实时更新界面显示
            ui->textEdit->clear();    // 清空输入框
        }
    } else {
        QMessageBox::information(this, "操作取消", "提交已取消。");
    }
}

void HomePage::on_tbtnVacation_clicked()
{
    UserSession *session = UserSession::instance();
    if (session->employeeID() == -1) {
        QMessageBox::warning(this, "错误!", "用户未登录，无法提交申报。");
        return;
    }

    // 获取申报信息
    QString applicantName = session->name();
    QString content = ui->textEdit_2->toPlainText().trimmed(); // 使用trimmed()去除首尾空白
    if (content.isEmpty()) {
        QMessageBox::warning(this, "错误!", "申报内容不能为空！");
        return;
    }

    QString type = "休假申报";

    // 弹出确认对话框
    QMessageBox::StandardButton reply;
    reply = QMessageBox::question(this, "确认提交", "您确定要提交休假申报吗？", QMessageBox::Yes | QMessageBox::No, QMessageBox::Yes);

    if (reply == QMessageBox::Yes) {
        // 提交到数据库
        if (addDeclarationToDb(applicantName, type, content)) {
            // 更新个人申报次数
            if (updatePersonalApplicationCount(applicantName, type)) {
                qDebug() << "个人休假申报总次数已更新。";
            } else {
                qDebug() << "个人休假申报总次数更新失败或未找到对应员工记录。";
            }
            QMessageBox::information(this, "提交成功", type + "已成功提交。");
            updateVacationNumslot(); // 实时更新界面显示
            ui->textEdit_2->clear();   // 清空输入框
        }
    } else {
        QMessageBox::information(this, "操作取消", "提交已取消。");
    }
}

void HomePage::on_tbtnProjection_clicked()
{
    UserSession *session = UserSession::instance();
    if (session->employeeID() == -1) {
        QMessageBox::warning(this, "错误!", "用户未登录，无法提交申报。");
        return;
    }

    // 获取申报信息
    QString applicantName = session->name();
    QString content = ui->textEdit_3->toPlainText().trimmed(); // 使用trimmed()去除首尾空白
    if (content.isEmpty()) {
        QMessageBox::warning(this, "错误!", "申报内容不能为空！");
        return;
    }

    QString type = "立项申报";

    // 弹出确认对话框
    QMessageBox::StandardButton reply;
    reply = QMessageBox::question(this, "确认提交", "您确定要提交立项申报吗？", QMessageBox::Yes | QMessageBox::No, QMessageBox::Yes);

    if (reply == QMessageBox::Yes) {
        // 提交到数据库
        if (addDeclarationToDb(applicantName, type, content)) {
            // 更新个人申报次数
            if (updatePersonalApplicationCount(applicantName, type)) {
                qDebug() << "个人立项申报总次数已更新。";
            } else {
                qDebug() << "个人立项申报总次数更新失败或未找到对应员工记录。";
            }
            QMessageBox::information(this, "提交成功", type + "已成功提交。");
            updateVacationNumslot(); // 实时更新界面显示
            ui->textEdit_2->clear();   // 清空输入框
        }
    } else {
        QMessageBox::information(this, "操作取消", "提交已取消。");
    }
}

void HomePage::on_btnToExcel_clicked()
{
    exportToExcel();
}

void HomePage::on_btnFromExcel_clicked()
{
    importFromExcel();
}

bool HomePage::addDeclarationToDb(const QString &applicant, const QString &type, const QString &content)
{
    // 获取全局数据库连接
    QSqlDatabase db = QSqlDatabase::database();
    if (!db.isOpen()) {
        QMessageBox::critical(this, "数据库错误", "默认数据库连接未打开或已关闭。");
        return false;
    }

    // 使用预处理语句防止SQL注入
    QSqlQuery query(db);
    query.prepare("INSERT INTO declarations (applicant_name, declaration_type, declaration_content, status, submission_timestamp) "
                  "VALUES (:applicant, :type, :content, '待审批', CURRENT_TIMESTAMP)");
    query.bindValue(":applicant", applicant);
    query.bindValue(":type", type);
    query.bindValue(":content", content);

    if (!query.exec()) {
        QMessageBox::critical(this, "提交失败", "无法保存申报: " + query.lastError().text());
        qDebug() << "数据库插入错误 (declarations):" << query.lastError().text() << "Query:" << query.lastQuery();
        return false;
    }

    return true;
}

bool HomePage::updatePersonalApplicationCount(const QString &applicantName, const QString &declarationType)
{
    QSqlDatabase db = QSqlDatabase::database();
    if (!db.isOpen()) {
        QMessageBox::critical(this, "数据库错误", "更新个人申请次数时，默认数据库连接未打开。");
        return false;
    }

    // 1. 根据员工姓名查询 EmployeeID
    QSqlQuery queryGetId(db);
    queryGetId.prepare("SELECT EmployeeID FROM basicinfo WHERE Name = :name");
    queryGetId.bindValue(":name", applicantName);
    int employeeID = -1;
    if (queryGetId.exec()) {
        if (queryGetId.next()) {
            employeeID = queryGetId.value(0).toInt();
        } else {
            // 员工姓名在basicinfo表中不存在，这是异常情况
            QMessageBox::warning(this, "更新失败", QString("在员工基础信息表中未找到姓名为 '%1' 的员工。").arg(applicantName));
            qDebug() << "未找到员工 basicinfo: " << applicantName;
            return false;
        }
    } else {
        QMessageBox::critical(this, "数据库错误", "查询 EmployeeID 失败: " + queryGetId.lastError().text());
        qDebug() << "查询 EmployeeID 错误 (basicinfo):" << queryGetId.lastError().text() << "Query:" << queryGetId.lastQuery();
        return false;
    }

    if (employeeID == -1) return false;

    // 2. 根据申报类型确定要更新的数据库列名
    QString countColumnName;
    if (declarationType == "财务申报") {
        countColumnName = "FinancialNum";
    } else if (declarationType == "休假申报") {
        countColumnName = "VacationNum";
    } else if (declarationType == "立项申报") {
        countColumnName = "ProjectionNum";
    } else {
        QMessageBox::warning(this, "更新失败", "未知的申报类型，无法更新个人申请次数。");
        qDebug() << "未知的申报类型: " << declarationType;
        return false;
    }

    // 3. 尝试更新员工的申报次数
    QSqlQuery queryUpdateCount(db);
    // 使用 IFNULL(column, 0) + 1 来安全地处理可能为NULL的计数字段
    QString sqlUpdate = QString("UPDATE personal_application SET %1 = IFNULL(%1, 0) + 1 WHERE EmployeeID = :employeeID")
                            .arg(countColumnName);
    queryUpdateCount.prepare(sqlUpdate);
    queryUpdateCount.bindValue(":employeeID", employeeID);

    if (queryUpdateCount.exec()) {
        if (queryUpdateCount.numRowsAffected() > 0) {
            // 更新成功
            qDebug() << QString("成功更新 EmployeeID %1 的 %2 计数。").arg(employeeID).arg(countColumnName);
            return true;
        } else {
            // 如果影响行数为0，说明 personal_application 表中没有该员工的记录，需要插入一条新记录
            qDebug() << QString("EmployeeID %1 在 personal_application 中不存在，尝试插入新记录...").arg(employeeID);
            QSqlQuery queryInsert(db);
            QString sqlInsert = QString("INSERT INTO personal_application (EmployeeID, Name, %1) VALUES (:employeeID, :name, 1)")
                                    .arg(countColumnName);
            queryInsert.prepare(sqlInsert);
            queryInsert.bindValue(":employeeID", employeeID);
            queryInsert.bindValue(":name", applicantName);
            if (queryInsert.exec()) {
                qDebug() << QString("成功为 EmployeeID %1 插入新的 %2 计数。").arg(employeeID).arg(countColumnName);
                return true;
            } else {
                QMessageBox::critical(this, "数据库错误", QString("插入新的个人申请计数失败 (EmployeeID: %1): %2")
                                                                   .arg(employeeID).arg(queryInsert.lastError().text()));
                qDebug() << "插入个人申请计数错误:" << queryInsert.lastError().text() << "Query:" << queryInsert.lastQuery();
                return false;
            }
        }
    } else {
        QMessageBox::critical(this, "数据库错误", QString("更新个人申请计数失败 (EmployeeID: %1): %2")
                                                           .arg(employeeID).arg(queryUpdateCount.lastError().text()));
        qDebug() << "更新个人申请计数错误:" << queryUpdateCount.lastError().text() << "Query:" << queryUpdateCount.lastQuery();
        return false;
    }
}

// 拉取员工列表
void HomePage::loadContacts()
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


void HomePage::on_listWidgetContacts_itemClicked(QListWidgetItem *item)
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


// 点击加载联系人聊天记录
void HomePage::loadChatHistory(int targetUserId)
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
void HomePage::on_btnSendMessage_clicked()
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



void HomePage::fetchNewMessages()
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
void HomePage::appendChatMessage(const QString &text, bool isMine)
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
