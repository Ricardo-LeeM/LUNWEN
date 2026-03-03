#include "salarymanagement.h"
#include "ui_salarymanagement.h"

#include <QSqlDatabase>
#include <QSqlQuery>
#include <QMessageBox>
#include <QSqlError>
#include <QDate>
#include <QSqlRelationalDelegate> // 用于创建只读列
#include <QHeaderView>
#include <QStyledItemDelegate>    // 自定义委托的基类
#include <QLineEdit>              // 用于创建自定义编辑器
#include <QDoubleValidator>       // 用于验证浮点数输入
#include "databasemanager.h"


/***************************************************************************************************
** 自定义委托类定义 (Custom Delegate Class Definition)                 **
***************************************************************************************************/
/**
 * @class NoSpinBoxDelegate
 * @brief 一个自定义的项委托（Item Delegate）
 *
 * 默认情况下，当编辑一个数字类型的单元格时，QTableView会提供一个带上下箭头的微调框(QSpinBox)。
 * 对于薪资这种需要输入大量小数的场景，微调框并不方便。
 *
 * 这个委托类通过重写 `createEditor` 方法，用一个标准的行编辑器 `QLineEdit` 来替换默认的微调框。
 * 同时，它为这个 `QLineEdit` 设置了一个 `QDoubleValidator`，以确保用户只能输入合法的浮点数，
 * 从而提供了更佳的用户体验。
 */
class NoSpinBoxDelegate : public QStyledItemDelegate
{
public:
    explicit NoSpinBoxDelegate(QObject *parent = nullptr) : QStyledItemDelegate(parent) {}

    // 重写 createEditor 函数，以提供自定义的编辑器
    QWidget *createEditor(QWidget *parent, const QStyleOptionViewItem & /*option*/, const QModelIndex & /*index*/) const override
    {
        // 创建一个标准的行编辑器作为我们的编辑器
        QLineEdit *editor = new QLineEdit(parent);
        // 为编辑器设置一个验证器，只允许输入浮点数（包括负数和科学计数法，但通常用于正数）
        editor->setValidator(new QDoubleValidator(parent));
        return editor;
    }
};


/***************************************************************************************************
** 构造函数与析构函数                                            **
***************************************************************************************************/

SalaryManagement::SalaryManagement(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::SalaryManagement),
    isAutoFilling(false)
{
    ui->setupUi(this);
    this->setWindowTitle("薪资管理");

    for (int i = 0; i < 6; ++i) {
        ui->periodComboBox->addItem(QDate::currentDate().addMonths(-i).toString("yyyy-MM"));
    }

    initializeModel();
    on_loadDataButton_clicked();
}

SalaryManagement::~SalaryManagement()
{
    delete ui;
}

/***************************************************************************************************
** 私有函数实现 (Private Function Implementations)                       **
***************************************************************************************************/

void SalaryManagement::initializeModel()
{
    QSqlDatabase db = QSqlDatabase::database();
    if (!db.isOpen()) {
        QMessageBox::critical(this, "数据库错误", "数据库未连接！");
        return;
    }

    model = new QSqlTableModel(this, db);
    model->setTable("PayrollRecords");
    // 设置编辑策略为手动提交。所有更改都缓存在本机，直到调用 submitAll()
    model->setEditStrategy(QSqlTableModel::OnManualSubmit);

    // 连接模型的 dataChanged 信号到我们的槽函数，以实现智能填充
    connect(model, &QSqlTableModel::dataChanged, this, &SalaryManagement::on_dataChanged);

    // 设置所有表头的中文，提高可读性
    model->setHeaderData(model->fieldIndex("employee_id"), Qt::Horizontal, "员工ID");
    model->setHeaderData(model->fieldIndex("employee_name"), Qt::Horizontal, "姓名");
    model->setHeaderData(model->fieldIndex("department"), Qt::Horizontal, "部门");
    model->setHeaderData(model->fieldIndex("payroll_period"), Qt::Horizontal, "薪资周期");
    model->setHeaderData(model->fieldIndex("base_salary"), Qt::Horizontal, "基本工资");
    model->setHeaderData(model->fieldIndex("bonus_attendance"), Qt::Horizontal, "全勤奖");
    model->setHeaderData(model->fieldIndex("bonus_performance"), Qt::Horizontal, "绩效奖");
    model->setHeaderData(model->fieldIndex("allowance"), Qt::Horizontal, "津贴");
    model->setHeaderData(model->fieldIndex("deduction_late"), Qt::Horizontal, "迟到扣款");
    model->setHeaderData(model->fieldIndex("deduction_absence"), Qt::Horizontal, "缺勤扣款");
    model->setHeaderData(model->fieldIndex("deduction_social_security"), Qt::Horizontal, "社保公积金");
    model->setHeaderData(model->fieldIndex("deduction_tax"), Qt::Horizontal, "个税");
    model->setHeaderData(model->fieldIndex("deduction_other"), Qt::Horizontal, "其他扣款");
    model->setHeaderData(model->fieldIndex("gross_salary"), Qt::Horizontal, "应发工资");
    model->setHeaderData(model->fieldIndex("net_salary"), Qt::Horizontal, "实发工资");
    model->setHeaderData(model->fieldIndex("notes"), Qt::Horizontal, "备注");

    ui->salaryTableView->setModel(model);
    // 隐藏不需用户看到的记录主键
    ui->salaryTableView->hideColumn(model->fieldIndex("record_id"));

    // 将我们自定义的委托应用到所有数字输入列
    NoSpinBoxDelegate *noSpinDelegate = new NoSpinBoxDelegate(this);
    const QList<QString> numericColumns = {
        "base_salary", "bonus_attendance", "bonus_performance", "allowance",
        "deduction_late", "deduction_absence", "deduction_social_security",
        "deduction_tax", "deduction_other"
    };

    for (const QString &columnName : numericColumns) {
        ui->salaryTableView->setItemDelegateForColumn(model->fieldIndex(columnName), noSpinDelegate);
    }

    // 设置表格的视觉样式
    ui->salaryTableView->resizeColumnsToContents();
    ui->salaryTableView->resizeRowsToContents();
    ui->salaryTableView->horizontalHeader()->setStretchLastSection(true);
}


/***************************************************************************************************
** 槽函数实现 (Slot Implementations)                                **
***************************************************************************************************/

/**
 * @brief 智能填充槽函数：当用户在“员工ID”列输入数据后触发。
 *
 * - 检查是否为程序自动填充，如果是则直接返回。
 * - 检查被修改的是否是“员工ID”列，如果不是则返回。
 * - 获取用户输入的员工ID。
 * - 使用该ID查询 `EmployeesSalary` 表，获取姓名、部门和基本工资。
 * - 如果查询成功，则用获取到的数据自动填充当前行的相应单元格。
 * - 如果查询失败，弹出警告，并清空相关单元格。
 * - 在填充操作前后，会设置 `isAutoFilling` 标志位以防止无限递归。
 */
void SalaryManagement::on_dataChanged(const QModelIndex &topLeft, const QModelIndex &/*bottomRight*/, const QVector<int> &/*roles*/)
{
    if (isAutoFilling) return;

    if (topLeft.column() != model->fieldIndex("employee_id")) return;

    int row = topLeft.row();
    int employeeId = model->data(topLeft).toInt();
    if (employeeId <= 0) return;

    // *** 重构核心：调用DatabaseManager ***
    QSqlRecord salaryInfo;
    if (DatabaseManager::instance()->getEmployeeSalaryInfo(employeeId, salaryInfo)) {
        QString name = salaryInfo.value("name").toString();
        QString department = salaryInfo.value("department").toString();
        double baseSalary = salaryInfo.value("base_salary").toDouble();

        isAutoFilling = true;
        model->setData(model->index(row, model->fieldIndex("employee_name")), name);
        model->setData(model->index(row, model->fieldIndex("department")), department);
        model->setData(model->index(row, model->fieldIndex("base_salary")), baseSalary);
        isAutoFilling = false;
    } else {
        QMessageBox::warning(this, "查询失败", QString("未找到员工ID为 %1 的员工信息，请确认ID是否正确。").arg(employeeId));
        isAutoFilling = true;
        model->setData(model->index(row, model->fieldIndex("employee_name")), "");
        model->setData(model->index(row, model->fieldIndex("department")), "");
        model->setData(model->index(row, model->fieldIndex("base_salary")), 0.00);
        isAutoFilling = false;
    }

    ui->salaryTableView->resizeColumnsToContents();
}

/**
 * @brief “加载数据”按钮槽函数
 *
 * 根据下拉框中选择的薪资周期，设置模型的过滤器并重新从数据库拉取数据。
 */
void SalaryManagement::on_loadDataButton_clicked()
{
    QString selectedPeriod = ui->periodComboBox->currentText();
    model->setFilter(QString("payroll_period = '%1'").arg(selectedPeriod));
    if(!model->select()){
        QMessageBox::critical(this, "查询失败", "无法加载数据: " + model->lastError().text());
    }
    // 刷新表格尺寸
    ui->salaryTableView->resizeColumnsToContents();
    ui->salaryTableView->resizeRowsToContents();
}

/**
 * @brief “计算工资”按钮槽函数
 *
 * 遍历表格中的每一行，根据薪酬项目的加减，计算出应发工资和实发工资，并更新到模型中。
 */
void SalaryManagement::on_calculateButton_clicked()
{
    if (model->rowCount() == 0) {
        QMessageBox::information(this, "提示", "当前没有数据可供计算。");
        return;
    }

    isAutoFilling = true; // 开始批量计算，设置标志位
    for (int row = 0; row < model->rowCount(); ++row) {
        // 读取所有薪酬项
        double base_salary = model->data(model->index(row, model->fieldIndex("base_salary"))).toDouble();
        double bonus_attendance = model->data(model->index(row, model->fieldIndex("bonus_attendance"))).toDouble();
        double bonus_performance = model->data(model->index(row, model->fieldIndex("bonus_performance"))).toDouble();
        double allowance = model->data(model->index(row, model->fieldIndex("allowance"))).toDouble();
        double deduction_late = model->data(model->index(row, model->fieldIndex("deduction_late"))).toDouble();
        double deduction_absence = model->data(model->index(row, model->fieldIndex("deduction_absence"))).toDouble();
        double deduction_social_security = model->data(model->index(row, model->fieldIndex("deduction_social_security"))).toDouble();
        double deduction_tax = model->data(model->index(row, model->fieldIndex("deduction_tax"))).toDouble();
        double deduction_other = model->data(model->index(row, model->fieldIndex("deduction_other"))).toDouble();

        // 计算总额
        double gross_salary = base_salary + bonus_attendance + bonus_performance + allowance;
        double total_deductions = deduction_late + deduction_absence + deduction_social_security + deduction_tax + deduction_other;
        double net_salary = gross_salary - total_deductions;

        // 将计算结果写回模型
        model->setData(model->index(row, model->fieldIndex("gross_salary")), QString::number(gross_salary, 'f', 2));
        model->setData(model->index(row, model->fieldIndex("net_salary")), QString::number(net_salary, 'f', 2));
    }
    isAutoFilling = false; // 计算结束，重置标志位

    ui->salaryTableView->resizeColumnsToContents();
    QMessageBox::information(this, "计算完成", "所有工资已计算完毕！\n请检查数据，然后点击“保存更改”以写入数据库。");
}

/**
 * @brief “新增记录”按钮槽函数
 *
 * 在模型末尾插入一个空行，并自动将薪资周期设置为当前下拉框选中的周期。
 */
void SalaryManagement::on_addButton_clicked()
{
    int newRow = model->rowCount();
    model->insertRow(newRow);

    // 自动为新行填充薪资周期
    QString currentPeriod = ui->periodComboBox->currentText();
    model->setData(model->index(newRow, model->fieldIndex("payroll_period")), currentPeriod);
    QMessageBox::information(this, "提示", "已新增一行，请在“员工ID”列输入员工编号后按回车或点击其他单元格。");
}

/**
 * @brief “保存更改”按钮槽函数
 *
 * 调用模型的 `submitAll()` 方法，将所有缓存的更改（增、删、改）一次性提交到数据库。
 * 如果提交失败，则弹出错误信息，并调用 `revertAll()` 撤销所有更改，保持界面与数据库一致。
 */
void SalaryManagement::on_saveButton_clicked()
{
    if (model->submitAll()) {
        QMessageBox::information(this, "成功", "所有更改已成功保存到数据库！");
    } else {
        QMessageBox::critical(this, "数据库错误", "保存失败，原因: " + model->lastError().text() + "\n所有更改将被撤销。");
        model->revertAll();
    }
}

/**
 * @brief “删除记录”按钮槽函数
 *
 * 删除表格中所有被选中的行。为确保数据一致性，删除操作完成后会立即调用保存逻辑。
 */
void SalaryManagement::on_deleteButton_clicked()
{
    QModelIndexList selectedRows = ui->salaryTableView->selectionModel()->selectedRows();
    if (selectedRows.isEmpty()) {
        QMessageBox::warning(this, "操作无效", "请先在表格中选择要删除的整行。");
        return;
    }

    // 弹出二次确认对话框，防止误删
    if (QMessageBox::question(this, "确认删除", "确定要删除选中的 " + QString::number(selectedRows.count()) + " 条记录吗？\n此操作将直接从数据库中移除，且不可恢复！",
                              QMessageBox::Yes | QMessageBox::No) == QMessageBox::Yes) {
        // 从后往前删除，避免因行号变化导致索引错误
        for (int i = selectedRows.count() - 1; i >= 0; --i) {
            model->removeRow(selectedRows.at(i).row());
        }
        // 删除后立即尝试保存
        on_saveButton_clicked();
    }
}
// #include "salarymanagement.h"
// #include "ui_salarymanagement.h"
// #include <QSqlDatabase>
// #include <QSqlQuery>
// #include <QMessageBox>
// #include <QSqlError>
// #include <QDate>
// #include <QSqlRelationalDelegate>
// #include <QHeaderView>
// #include <QStyledItemDelegate> // <--- 1. 引入委托基类
// #include <QLineEdit>           // <--- 2. 引入行编辑器
// #include <QDoubleValidator>    // <--- 3. 引入验证器

// // vvv--- 4. 在此处定义一个自定义委托类，用于替换默认的数字编辑器 ---vvv
// class NoSpinBoxDelegate : public QStyledItemDelegate
// {
// public:
//     explicit NoSpinBoxDelegate(QObject *parent = nullptr) : QStyledItemDelegate(parent) {}

//     // 重写 createEditor 函数
//     QWidget *createEditor(QWidget *parent, const QStyleOptionViewItem & /*option*/, const QModelIndex & /*index*/) const override
//     {
//         // 创建一个标准的行编辑器
//         QLineEdit *editor = new QLineEdit(parent);
//         // 为编辑器设置一个验证器，只允许输入浮点数
//         editor->setValidator(new QDoubleValidator(parent));
//         return editor;
//     }
// };
// // ^^^-------------------------------------------------------------------^^^


// SalaryManagement::SalaryManagement(QWidget *parent) :
//     QMainWindow(parent),
//     ui(new Ui::SalaryManagement),
//     isAutoFilling(false)
// {
//     ui->setupUi(this);
//     this->setWindowTitle("薪资管理 (智能表格版)");

//     for (int i = 0; i < 6; ++i) {
//         ui->periodComboBox->addItem(QDate::currentDate().addMonths(-i).toString("yyyy-MM"));
//     }

//     initializeModel();
//     on_loadDataButton_clicked();
// }

// SalaryManagement::~SalaryManagement()
// {
//     delete ui;
// }

// void SalaryManagement::initializeModel()
// {
//     QSqlDatabase db = QSqlDatabase::database();
//     if (!db.isOpen()) {
//         QMessageBox::critical(this, "数据库错误", "数据库未连接！");
//         return;
//     }

//     model = new QSqlTableModel(this, db);
//     model->setTable("PayrollRecords");
//     model->setEditStrategy(QSqlTableModel::OnManualSubmit);

//     connect(model, &QSqlTableModel::dataChanged, this, &SalaryManagement::on_dataChanged);

//     // 设置所有表头的中文
//     model->setHeaderData(model->fieldIndex("employee_id"), Qt::Horizontal, "员工ID");
//     model->setHeaderData(model->fieldIndex("employee_name"), Qt::Horizontal, "姓名");
//     model->setHeaderData(model->fieldIndex("department"), Qt::Horizontal, "部门");
//     model->setHeaderData(model->fieldIndex("payroll_period"), Qt::Horizontal, "薪资周期");
//     model->setHeaderData(model->fieldIndex("base_salary"), Qt::Horizontal, "基本工资");
//     model->setHeaderData(model->fieldIndex("bonus_attendance"), Qt::Horizontal, "全勤奖");
//     model->setHeaderData(model->fieldIndex("bonus_performance"), Qt::Horizontal, "绩效奖");
//     model->setHeaderData(model->fieldIndex("allowance"), Qt::Horizontal, "津贴");
//     model->setHeaderData(model->fieldIndex("deduction_late"), Qt::Horizontal, "迟到扣款");
//     model->setHeaderData(model->fieldIndex("deduction_absence"), Qt::Horizontal, "缺勤扣款");
//     model->setHeaderData(model->fieldIndex("deduction_social_security"), Qt::Horizontal, "社保公积金");
//     model->setHeaderData(model->fieldIndex("deduction_tax"), Qt::Horizontal, "个税");
//     model->setHeaderData(model->fieldIndex("deduction_other"), Qt::Horizontal, "其他扣款");
//     model->setHeaderData(model->fieldIndex("gross_salary"), Qt::Horizontal, "应发工资");
//     model->setHeaderData(model->fieldIndex("net_salary"), Qt::Horizontal, "实发工资");
//     model->setHeaderData(model->fieldIndex("notes"), Qt::Horizontal, "备注");

//     ui->salaryTableView->setModel(model);
//     ui->salaryTableView->hideColumn(model->fieldIndex("record_id"));

//     // 让姓名和部门列不可手动编辑
//     ui->salaryTableView->setItemDelegateForColumn(model->fieldIndex("employee_name"), new QSqlRelationalDelegate(ui->salaryTableView));
//     ui->salaryTableView->setItemDelegateForColumn(model->fieldIndex("department"), new QSqlRelationalDelegate(ui->salaryTableView));

//     // vvv--- 5. 将我们自定义的委托应用到所有数字列 ---vvv
//     NoSpinBoxDelegate *noSpinDelegate = new NoSpinBoxDelegate(this);
//     const QList<QString> numericColumns = {
//         "base_salary", "bonus_attendance", "bonus_performance", "allowance",
//         "deduction_late", "deduction_absence", "deduction_social_security",
//         "deduction_tax", "deduction_other"
//     };

//     for (const QString &columnName : numericColumns) {
//         ui->salaryTableView->setItemDelegateForColumn(model->fieldIndex(columnName), noSpinDelegate);
//     }
//     // ^^^---------------------------------------------------^^^

//     // 设置表格的拉伸和自适应模式
//     ui->salaryTableView->resizeColumnsToContents();
//     ui->salaryTableView->resizeRowsToContents();
//     ui->salaryTableView->horizontalHeader()->setStretchLastSection(true);
// }

// // ... (所有其他函数代码均未改动，为了简洁省略)
// // ... on_dataChanged, on_loadDataButton_clicked, etc.

// void SalaryManagement::on_dataChanged(const QModelIndex &topLeft, const QModelIndex &/*bottomRight*/, const QVector<int> &/*roles*/)
// {
//     if (isAutoFilling) {
//         return;
//     }

//     int employeeIdColumn = model->fieldIndex("employee_id");
//     if (topLeft.column() != employeeIdColumn) {
//         return;
//     }

//     int row = topLeft.row();
//     int employeeId = model->data(topLeft).toInt();

//     if (employeeId <= 0) {
//         return;
//     }

//     QSqlQuery query;
//     query.prepare("SELECT name, department, base_salary FROM EmployeesSalary WHERE employee_id = ?");
//     query.addBindValue(employeeId);

//     if (query.exec() && query.next()) {
//         QString name = query.value("name").toString();
//         QString department = query.value("department").toString();
//         double baseSalary = query.value("base_salary").toDouble();

//         isAutoFilling = true;
//         model->setData(model->index(row, model->fieldIndex("employee_name")), name);
//         model->setData(model->index(row, model->fieldIndex("department")), department);
//         model->setData(model->index(row, model->fieldIndex("base_salary")), baseSalary);
//         isAutoFilling = false;
//     } else {
//         QMessageBox::warning(this, "查询失败", QString("未找到员工ID为 %1 的员工信息，请确认ID是否正确。").arg(employeeId));
//         isAutoFilling = true;
//         model->setData(model->index(row, model->fieldIndex("employee_name")), "");
//         model->setData(model->index(row, model->fieldIndex("department")), "");
//         model->setData(model->index(row, model->fieldIndex("base_salary")), 0.00);
//         isAutoFilling = false;
//     }

//     ui->salaryTableView->resizeColumnsToContents();
// }

// void SalaryManagement::on_loadDataButton_clicked()
// {
//     QString selectedPeriod = ui->periodComboBox->currentText();
//     model->setFilter(QString("payroll_period = '%1'").arg(selectedPeriod));
//     if(!model->select()){
//         QMessageBox::critical(this, "查询失败", "无法加载数据: " + model->lastError().text());
//     }

//     ui->salaryTableView->resizeColumnsToContents();
//     ui->salaryTableView->resizeRowsToContents();
// }

// void SalaryManagement::on_calculateButton_clicked()
// {
//     if (model->rowCount() == 0) {
//         QMessageBox::information(this, "提示", "当前没有数据可供计算。");
//         return;
//     }
//     for (int row = 0; row < model->rowCount(); ++row) {
//         double base_salary = model->data(model->index(row, model->fieldIndex("base_salary"))).toDouble();
//         double bonus_attendance = model->data(model->index(row, model->fieldIndex("bonus_attendance"))).toDouble();
//         double bonus_performance = model->data(model->index(row, model->fieldIndex("bonus_performance"))).toDouble();
//         double allowance = model->data(model->index(row, model->fieldIndex("allowance"))).toDouble();
//         double deduction_late = model->data(model->index(row, model->fieldIndex("deduction_late"))).toDouble();
//         double deduction_absence = model->data(model->index(row, model->fieldIndex("deduction_absence"))).toDouble();
//         double deduction_social_security = model->data(model->index(row, model->fieldIndex("deduction_social_security"))).toDouble();
//         double deduction_tax = model->data(model->index(row, model->fieldIndex("deduction_tax"))).toDouble();
//         double deduction_other = model->data(model->index(row, model->fieldIndex("deduction_other"))).toDouble();
//         double gross_salary = base_salary + bonus_attendance + bonus_performance + allowance;
//         double total_deductions = deduction_late + deduction_absence + deduction_social_security + deduction_tax + deduction_other;
//         double net_salary = gross_salary - total_deductions;
//         model->setData(model->index(row, model->fieldIndex("gross_salary")), QString::number(gross_salary, 'f', 2));
//         model->setData(model->index(row, model->fieldIndex("net_salary")), QString::number(net_salary, 'f', 2));
//     }

//     ui->salaryTableView->resizeColumnsToContents();
//     QMessageBox::information(this, "计算完成", "所有工资已计算完毕！\n请检查数据，然后点击“保存更改”以写入数据库。");
// }

// void SalaryManagement::on_addButton_clicked()
// {
//     int newRow = model->rowCount();
//     model->insertRow(newRow);

//     QString currentPeriod = ui->periodComboBox->currentText();
//     model->setData(model->index(newRow, model->fieldIndex("payroll_period")), currentPeriod);
//     QMessageBox::information(this, "提示", "已新增一行，请在“员工ID”列输入员工编号后按回车或点击其他单元格。");
// }

// void SalaryManagement::on_saveButton_clicked()
// {
//     if (model->submitAll()) {
//         QMessageBox::information(this, "成功", "所有更改已成功保存到数据库！");
//     } else {
//         QMessageBox::critical(this, "数据库错误", "保存失败，原因: " + model->lastError().text() + "\n所有更改将被撤销。");
//         model->revertAll();
//     }
// }

// void SalaryManagement::on_deleteButton_clicked()
// {
//     QModelIndexList selectedRows = ui->salaryTableView->selectionModel()->selectedRows();
//     if (selectedRows.isEmpty()) {
//         QMessageBox::warning(this, "操作无效", "请先在表格中选择要删除的整行。");
//         return;
//     }
//     if (QMessageBox::question(this, "确认删除", "确定要删除选中的 " + QString::number(selectedRows.count()) + " 条记录吗？\n此操作将直接从数据库中移除，且不可恢复！",
//                               QMessageBox::Yes | QMessageBox::No) == QMessageBox::Yes) {
//         for (int i = selectedRows.count() - 1; i >= 0; --i) {
//             model->removeRow(selectedRows.at(i).row());
//         }
//         on_saveButton_clicked();
//     }
// }
