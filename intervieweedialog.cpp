#include "intervieweedialog.h"
#include "ui_intervieweedialog.h"
#include <QDateTime>
#include <QMessageBox>

IntervieweeDialog::IntervieweeDialog(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::IntervieweeDialog)
{
    ui->setupUi(this);
    setupUI();

    // 连接按钮信号到槽（也可以在UI设计器中完成）
    // Qt::QueuedConnection 可以避免一些潜在的即时关闭问题
    connect(ui->okButton, &QPushButton::clicked, this, &IntervieweeDialog::on_okButton_clicked);
    connect(ui->cancelButton, &QPushButton::clicked, this, &IntervieweeDialog::on_cancelButton_clicked);
    connect(ui->enableTimeCheckBox, &QCheckBox::toggled, this, &IntervieweeDialog::on_enableTimeCheckBox_toggled);

    // 初始状态下，让时间选择器不可用
    ui->interviewTimeEdit->setEnabled(false);
}

IntervieweeDialog::~IntervieweeDialog()
{
    delete ui;
}

// 初始化UI控件，比如下拉框的选项
void IntervieweeDialog::setupUI()
{
    // 为 statusComboBox 添加选项
    // 我们将文字和对应的数据库数字（0, 1, 2...）一起存储
    ui->statusComboBox->addItem("待筛选", 0);
    ui->statusComboBox->addItem("一面", 1);
    ui->statusComboBox->addItem("二面", 2);
    ui->statusComboBox->addItem("HR面", 3);
    ui->statusComboBox->addItem("待定", 4);
    ui->statusComboBox->addItem("已录用", 5);
    ui->statusComboBox->addItem("未录用", 6);

    // 设置日期时间编辑器
    ui->interviewTimeEdit->setDateTime(QDateTime::currentDateTime());
    ui->interviewTimeEdit->setCalendarPopup(true); // 显示一个方便的日历弹出窗口
    ui->interviewTimeEdit->setDisplayFormat("yyyy-MM-dd hh:mm");
}

// 接收主窗口传来的记录，并填充到UI上
void IntervieweeDialog::setRecord(const QSqlRecord &record)
{
    m_record = record; // 将传入的记录保存到成员变量中

    // 从记录中取出数据，填充到对应的控件
    ui->nameLineEdit->setText(record.value("name").toString());
    ui->phoneLineEdit->setText(record.value("phone").toString());
    ui->emailLineEdit->setText(record.value("email").toString());
    ui->positionLineEdit->setText(record.value("applied_position").toString());
    ui->interviewerLineEdit->setText(record.value("interviewer").toString());
    ui->notesTextEdit->setPlainText(record.value("notes").toString());

    // 特殊处理 ComboBox
    int status = record.value("status").toInt();
    int index = ui->statusComboBox->findData(status); // 根据数字找到对应的索引
    if (index != -1) {
        ui->statusComboBox->setCurrentIndex(index);
    }

    // === 修改时间处理逻辑 ===
    QVariant timeVariant = record.value("interview_time");

    // 如果数据库中的时间是 NULL 或无效的
    if (timeVariant.isNull() || !timeVariant.toDateTime().isValid()) {
        ui->enableTimeCheckBox->setChecked(false); // 不勾选复选框
        ui->interviewTimeEdit->setEnabled(false);   // 禁用时间选择器
        // 可以给一个未来的默认时间，以防用户之后勾选
        ui->interviewTimeEdit->setDateTime(QDateTime::currentDateTime().addDays(1));
    } else {
        // 如果时间有效
        ui->enableTimeCheckBox->setChecked(true);  // 勾选复选框
        ui->interviewTimeEdit->setEnabled(true);    // 启用时间选择器
        ui->interviewTimeEdit->setDateTime(timeVariant.toDateTime()); // 设置为数据库中的时间
    }
}

// 从UI控件收集用户填写的数据，并更新到成员变量的记录中
QSqlRecord IntervieweeDialog::getRecord()
{
    // 从UI控件获取数据，并设置到 m_record 的相应字段
    m_record.setValue("name", ui->nameLineEdit->text());
    m_record.setValue("phone", ui->phoneLineEdit->text());
    m_record.setValue("email", ui->emailLineEdit->text());
    m_record.setValue("applied_position", ui->positionLineEdit->text());
    m_record.setValue("interviewer", ui->interviewerLineEdit->text());
    m_record.setValue("notes", ui->notesTextEdit->toPlainText());

    // 从 ComboBox 获取状态对应的数字
    m_record.setValue("status", ui->statusComboBox->currentData().toInt());

    // 【优化】从 DateTimeEdit 获取时间
    if (ui->enableTimeCheckBox->isChecked()) {
        // 如果复选框被勾选，存入 QDateTimeEdit 中的时间
        m_record.setValue("interview_time", ui->interviewTimeEdit->dateTime());
    } else {
        // 如果未勾选，向数据库存入 NULL
        m_record.setValue("interview_time", QVariant(QVariant::DateTime));
    }

    return m_record;
}

// "确定"按钮的槽函数
void IntervieweeDialog::on_okButton_clicked()
{
    // 数据校验
    if (ui->nameLineEdit->text().isEmpty()) {
        QMessageBox::warning(this, "信息不完整", "姓名不能为空！");
        return;
    }

    // 发送 accept 信号，这会让 QDialog::exec() 返回 QDialog::Accepted
    this->accept();
}

// "取消"按钮的槽函数
void IntervieweeDialog::on_cancelButton_clicked()
{
    // 发送 reject 信号，这会让 QDialog::exec() 返回 QDialog::Rejected
    this->reject();
}

void IntervieweeDialog::on_enableTimeCheckBox_toggled(bool checked)
{
    // 当复选框状态改变时，同步启用或禁用时间选择器
    ui->interviewTimeEdit->setEnabled(checked);
}

