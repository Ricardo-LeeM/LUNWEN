#include "employeeeditdialog.h"
#include "ui_employeeeditdialog.h"
#include <QSqlQuery>
#include <QMessageBox>
#include <QDate>

EmployeeEditDialog::EmployeeEditDialog(const QSqlRecord &record, EditMode mode, QWidget *parent) :
    QDialog(parent),
    ui(new Ui::EmployeeEditDialog),
    m_record(record)
{
    ui->setupUi(this);
    initComboBoxes();
    if (mode == Add) {
        setWindowTitle("添加新员工");
        ui->idLineEdit->setReadOnly(false);
    } else {
        setWindowTitle("修改员工信息");
        ui->idLineEdit->setReadOnly(true);
        fillData();
    }
}

EmployeeEditDialog::~EmployeeEditDialog()
{
    delete ui;
}

void EmployeeEditDialog::fillData()
{
    ui->idLineEdit->setText(m_record.value("EmployeeID").toString());
    ui->nameLineEdit->setText(m_record.value("Name").toString());
    ui->departmentComboBox->setCurrentText(m_record.value("Department").toString());
    ui->positionLineEdit->setText(m_record.value("Position").toString());
    ui->statusComboBox->setCurrentText(m_record.value("EmployeeStatus").toString());
    ui->workNatureComboBox->setCurrentText(m_record.value("WorkNature").toString());
    ui->employeeTypeComboBox->setCurrentText(m_record.value("EmployeeType").toString());
    ui->joinDateEdit->setDate(m_record.value("JoinDate").toDate());
    ui->partyMemberCheckBox->setChecked(m_record.value("IsPartyMember").toBool());
}

void EmployeeEditDialog::accept()
{
    // 获取并去除首尾空格
    QString idStr = ui->idLineEdit->text().trimmed();
    QString nameStr = ui->nameLineEdit->text().trimmed();

    if (idStr.isEmpty() || nameStr.isEmpty()) {
        QMessageBox::warning(this, "输入错误", "员工工号和姓名不能为空！");
        return;
    }

    bool isNumber;
    int id = idStr.toInt(&isNumber);

    if (!isNumber || id <= 0) {
        QMessageBox::warning(this, "输入错误", "员工工号必须是一个有效的正整数！");
        return;
    }

    // （此部分代码不变）
    m_record.setValue("EmployeeID", id);
    m_record.setValue("Name", nameStr);
    m_record.setValue("Department", ui->departmentComboBox->currentText());
    m_record.setValue("Position", ui->positionLineEdit->text());
    m_record.setValue("EmployeeStatus", ui->statusComboBox->currentText());
    m_record.setValue("WorkNature", ui->workNatureComboBox->currentText());
    m_record.setValue("EmployeeType", ui->employeeTypeComboBox->currentText());
    m_record.setValue("JoinDate", ui->joinDateEdit->date());
    m_record.setValue("IsPartyMember", ui->partyMemberCheckBox->isChecked());

    // 调用基类的 accept() 以关闭对话框并返回 QDialog::Accepted
    QDialog::accept();
}

QSqlRecord EmployeeEditDialog::getRecordData() const
{
    return m_record;
}

void EmployeeEditDialog::initComboBoxes()
{
    ui->statusComboBox->clear();
    ui->workNatureComboBox->clear();
    ui->employeeTypeComboBox->clear();
    ui->departmentComboBox->clear();
    ui->statusComboBox->addItems({"在职", "离职"});
    ui->workNatureComboBox->addItems({"全职", "兼职"});
    ui->employeeTypeComboBox->addItems({"员工", "管理员"});
    QSqlQuery query;
    query.prepare("SELECT DISTINCT Department FROM basicinfo WHERE Department IS NOT NULL AND Department != ?");
    query.addBindValue("高管");
    query.exec();
    while(query.next()) {
        ui->departmentComboBox->addItem(query.value(0).toString());
    }
}
