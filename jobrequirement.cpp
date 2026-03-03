#include "jobrequirement.h"
#include "ui_jobrequirement.h"
#include "statusdelegate.h"      // 自定义委托，用于美化面试状态显示
#include "intervieweedialog.h"   // 添加/编辑面试者的对话框

#include <QDebug>
#include <QMessageBox>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QSqlRecord>
#include <QFileDialog>
#include <QDesktopServices>
#include <QDir>
#include <QStandardPaths>

// 仅在 Windows 平台下包含 ActiveQt 模块，用于与 Excel (COM) 交互
#ifdef Q_OS_WIN
#include <ActiveQt/QAxObject>
#endif


/***************************************************************************************************
** 构造函数与析构函数                                            **
***************************************************************************************************/

/**
 * @brief JobRequirement 类的构造函数
 *
 * 初始化UI，并分别调用设置函数来配置 "职位需求" 和 "面试者管理" 两个选项卡。
 * 这种分离使得构造函数本身保持简洁，代码结构更清晰。
 * @param parent 父窗口指针
 */
JobRequirement::JobRequirement(QWidget *parent)
    : QWidget(parent)
    , ui(new Ui::JobRequirement)
{
    ui->setupUi(this);

    // 调用独立的设置函数，让构造函数保持整洁
    setupJobRequirementsTab();
    setupIntervieweesTab();

    // 载入数据时自动调整职位需求的行高列宽
    on_btnAutoAdjustRowHeigth_clicked();
    on_btnAutoAdjustColumnWidth_clicked();
}

/**
 * @brief JobRequirement 类的析构函数
 *
 * 释放UI指针所占用的内存。
 */
JobRequirement::~JobRequirement()
{
    delete ui;
}

// ===============================================================================================
// ======================================== 初始化函数 (Setup Functions) =============================
// ===============================================================================================

/**
 * @brief 初始化 "职位需求管理" 选项卡 (Tab 1)
 *
 * - 创建并配置 QStandardItemModel 用于显示职位需求。
 * - 设置表格的列标题。
 * - 从数据库加载初始数据。
 * - 连接模型的数据变更信号 `itemChanged` 到处理函数 `handleItemChanged`。
 * - 确定并设置用于发布的HTML文件的路径。
 * - 设置UI控件的初始状态（例如，表格默认为不可编辑）。
 */
void JobRequirement::setupJobRequirementsTab()
{
    // 初始化职位需求模型
    m_jobRequirementModel = new QStandardItemModel(0, 7, this);
    m_jobRequirementModel->setHorizontalHeaderLabels({"部门", "需求职位名称", "主要职责", "招聘人数", "性别要求", "学历要求", "任职资格要求"});
    ui->tableViewJobRequirement->setModel(m_jobRequirementModel);

    // 从数据库加载数据
    loadJobRequirementsData();

    // 连接信号和槽：当模型中的任何项发生变化时，调用 handleItemChanged 函数
    connect(m_jobRequirementModel, &QStandardItemModel::itemChanged, this, &JobRequirement::handleItemChanged);

    // HTML 相关设置：确定HTML文件的持久化存储路径
    m_persistentHtmlFilePath = determinePersistentHtmlPath();
    // 根据HTML文件是否存在，来决定“在HTML中查看”按钮是否可用
    //ui->btnCheckInHtml->setEnabled(QFile::exists(m_persistentHtmlFilePath));

    // UI 初始状态：表格默认不可编辑
    ui->tableViewJobRequirement->setEnabled(false);
    ui->checkBox->setChecked(false);
}

/**
 * @brief 初始化 "面试者管理" 选项卡 (Tab 2)
 *
 * - 创建 QSqlTableModel，并将其直接关联到数据库的 "interviewees" 表。
 * 这使得对模型的修改能自动反映到数据库。
 * - 设置清晰的中文列标题。
 * - 设置编辑策略为 `OnFieldChange`，即单元格内容一改变就立刻提交到数据库。
 * - 从数据库拉取数据并填充模型。
 * - 对表格视图进行美化设置（如隐藏ID列、整行选择、隔行变色等）。
 * - 为 "status"（面试状态）列设置自定义的 `StatusDelegate`，以更直观的方式显示状态。
 */
void JobRequirement::setupIntervieweesTab()
{
    // 1. 创建 QSqlTableModel 实例，使用全局数据库连接
    m_intervieweeModel = new QSqlTableModel(this, QSqlDatabase::database());

    // 2. 关联数据库表 "interviewees"
    m_intervieweeModel->setTable("interviewees");

    // 3. 设置表头显示的中文名，提高可读性
    m_intervieweeModel->setHeaderData(m_intervieweeModel->fieldIndex("name"), Qt::Horizontal, "姓名");
    m_intervieweeModel->setHeaderData(m_intervieweeModel->fieldIndex("phone"), Qt::Horizontal, "联系电话");
    m_intervieweeModel->setHeaderData(m_intervieweeModel->fieldIndex("applied_position"), Qt::Horizontal, "应聘岗位");
    m_intervieweeModel->setHeaderData(m_intervieweeModel->fieldIndex("status"), Qt::Horizontal, "面试状态");
    m_intervieweeModel->setHeaderData(m_intervieweeModel->fieldIndex("interview_time"), Qt::Horizontal, "面试时间");
    m_intervieweeModel->setHeaderData(m_intervieweeModel->fieldIndex("interviewer"), Qt::Horizontal, "面试官");
    m_intervieweeModel->setHeaderData(m_intervieweeModel->fieldIndex("notes"), Qt::Horizontal, "备注");

    // 4. 设置编辑策略：OnFieldChange 表示单元格内容改变后立即提交到数据库
    m_intervieweeModel->setEditStrategy(QSqlTableModel::OnFieldChange);

    // 5. 从数据库加载数据到模型，并进行错误检查
    if (!m_intervieweeModel->select()) {
        QMessageBox::critical(this, "数据库错误", "加载面试者数据失败：\n" + m_intervieweeModel->lastError().text());
        return;
    }

    // 6. 将模型设置给 TableView 控件
    ui->intervieweesTableView->setModel(m_intervieweeModel);

    // 7. 优化UI显示：隐藏不必要的列，设置表格属性
    ui->intervieweesTableView->hideColumn(m_intervieweeModel->fieldIndex("id")); // 隐藏ID列
    ui->intervieweesTableView->hideColumn(m_intervieweeModel->fieldIndex("email"));
    ui->intervieweesTableView->hideColumn(m_intervieweeModel->fieldIndex("resume_path"));

    ui->intervieweesTableView->setSelectionBehavior(QAbstractItemView::SelectRows); // 设置为整行选择
    ui->intervieweesTableView->horizontalHeader()->setStretchLastSection(true);   // 最后一列自动拉伸
    ui->intervieweesTableView->setAlternatingRowColors(true);                     // 隔行变色

    // 8. 为 'status' 列设置自定义委托
    int statusColumn = m_intervieweeModel->fieldIndex("status");
    if (statusColumn != -1) {
        // 只有当成功找到 'status' 列时，才设置委托
        ui->intervieweesTableView->setItemDelegateForColumn(statusColumn, new StatusDelegate(this));
        qDebug() << "StatusDelegate successfully set for column" << statusColumn;
    } else {
        // 如果找不到，发出警告，帮助开发者定位问题
        qWarning() << "Could not find the 'status' field in the interviewee model. Delegate not set.";
        QMessageBox::warning(this, "开发警告", "在面试者模型中未能找到 'status' 字段，状态将显示为数字。请检查数据库表结构。");
    }

    ui->statusFilterComboBox->clear();
    ui->statusFilterComboBox->addItem("所有状态", -1); // 添加一个“全部”选项，其关联数据为-1
    ui->statusFilterComboBox->addItem("待筛选", 0);
    ui->statusFilterComboBox->addItem("一面", 1);
    ui->statusFilterComboBox->addItem("二面", 2);
    ui->statusFilterComboBox->addItem("HR面", 3);
    ui->statusFilterComboBox->addItem("待定", 4);
    ui->statusFilterComboBox->addItem("已录用", 5);
    ui->statusFilterComboBox->addItem("未录用", 6);

    // 10. 连接信号和槽以进行统一筛选
    // a. 当文本框内容改变时，调用统一的筛选函数
    connect(ui->searchLineEdit, &QLineEdit::textChanged, this, &JobRequirement::updateIntervieweeFilter);
    // b. 当下拉框选项改变时，也调用统一的筛选函数
    connect(ui->statusFilterComboBox, QOverload<int>::of(&QComboBox::currentIndexChanged), this, &JobRequirement::updateIntervieweeFilter);
}

/**
 * @brief 统一更新面试者表格过滤器的槽函数
 *
 * 这个函数是新的核心筛选逻辑。它会同时考虑搜索框和状态下拉框的内容，
 * 构建一个组合的过滤条件，并应用到模型上。
 */
void JobRequirement::updateIntervieweeFilter()
{
    // 1. 获取各个筛选控件的当前值
    QString searchText = ui->searchLineEdit->text().trimmed();
    int statusValue = ui->statusFilterComboBox->currentData().toInt(); // 获取与选项关联的数字

    // 2. 构建一个条件列表
    QStringList filters;

    // 3. 如果搜索框有文本，添加文本筛选条件
    if (!searchText.isEmpty()) {
        filters.append(QString("(name LIKE '%%1%' OR phone LIKE '%%1%')").arg(searchText));
    }

    // 4. 如果下拉框选择的不是“所有状态”(-1)，则添加状态筛选条件
    if (statusValue != -1) {
        filters.append(QString("status = %1").arg(statusValue));
    }

    // 5. 使用 " AND " 将所有条件连接起来，并应用到模型
    m_intervieweeModel->setFilter(filters.join(" AND "));
}


// ===============================================================================================
// ============================ 面试者管理槽函数 (Tab 2 - Interviewees) =============================
// ===============================================================================================

/**
 * @brief "添加面试者" 按钮的槽函数
 *
 * - 弹出一个 `IntervieweeDialog` 对话框。
 * - 如果用户在对话框中点击 "确定"，则获取对话框中填写的数据。
 * - 使用 `QSqlQuery` 将新记录插入到 "interviewees" 数据库表中。
 * - 插入成功后，刷新表格视图以显示新添加的面试者。
 */
void JobRequirement::on_addIntervieweeButton_clicked()
{
    IntervieweeDialog dlg(this);
    // 传入一个空的 QSqlRecord，让对话框以“添加模式”启动
    dlg.setRecord(m_intervieweeModel->record());

    if (dlg.exec() == QDialog::Accepted) {
        // 从对话框获取用户输入的数据
        QSqlRecord record = dlg.getRecord();

        // 使用 QSqlQuery 执行 INSERT 操作，这种方式更灵活、更安全
        QSqlQuery query;
        query.prepare("INSERT INTO interviewees (name, phone, email, applied_position, status, interview_time, interviewer, notes) "
                      "VALUES (:name, :phone, :email, :applied_position, :status, :interview_time, :interviewer, :notes)");

        // 绑定所有字段的数据
        query.bindValue(":name",             record.value("name"));
        query.bindValue(":phone",            record.value("phone"));
        query.bindValue(":email",            record.value("email"));
        query.bindValue(":applied_position", record.value("applied_position"));
        query.bindValue(":status",           record.value("status"));
        query.bindValue(":interview_time",   record.value("interview_time"));
        query.bindValue(":interviewer",      record.value("interviewer"));
        query.bindValue(":notes",            record.value("notes"));

        if (!query.exec()) {
            QMessageBox::critical(this, "数据库错误", "添加新面试者失败：\n" + query.lastError().text());
        } else {
            QMessageBox::information(this, "成功", "新面试者添加成功！");
            // 刷新模型以显示新数据
            m_intervieweeModel->select();
        }
    }
}

/**
 * @brief "删除面试者" 按钮的槽函数
 *
 * - 检查用户是否已在表格中选中一行。
 * - 弹出确认对话框，防止误操作。
 * - 如果用户确认，则从 `QSqlTableModel` 中删除该行。
 * - `QSqlTableModel` 会自动处理数据库的 DELETE 操作。
 */
void JobRequirement::on_deleteIntervieweeButton_clicked()
{
    QModelIndex currentIndex = ui->intervieweesTableView->currentIndex();
    if (!currentIndex.isValid()) {
        QMessageBox::warning(this, "提示", "请先选择要删除的面试者记录。");
        return;
    }

    auto reply = QMessageBox::question(this, "确认删除", "您确定要删除这条面试者记录吗？", QMessageBox::Yes | QMessageBox::No);

    if (reply == QMessageBox::Yes) {
        // 从模型中删除该行
        m_intervieweeModel->removeRow(currentIndex.row());

        // 提交所有挂起的更改（包括删除）到数据库
        if (!m_intervieweeModel->submitAll()) {
            QMessageBox::critical(this, "删除失败", "从数据库删除记录失败：\n" + m_intervieweeModel->lastError().text());
            m_intervieweeModel->revertAll(); // 如果失败，则撤销更改
        } else {
            // 建议在成功后再次调用 select()，以确保模型与数据库完全同步
            m_intervieweeModel->select();
            QMessageBox::information(this, "成功", "记录已成功删除。");
        }
    }
}

/**
 * @brief "修改面试者" 按钮的槽函数
 *
 * - 检查用户是否已选中要修改的行。
 * - 从模型中获取该行的 `QSqlRecord`，里面包含了所有字段的数据（包括隐藏的ID）。
 * - 使用该记录的数据填充 `IntervieweeDialog` 对话框。
 * - 如果用户确认修改，则获取对话框中更新后的数据。
 * - 使用 `QSqlQuery` 和 `UPDATE` 语句，根据记录的 ID 更新数据库中的相应条目。
 * - 成功后刷新表格。
 */
void JobRequirement::on_editIntervieweeButton_clicked()
{
    QModelIndex currentIndex = ui->intervieweesTableView->currentIndex();
    if (!currentIndex.isValid()) {
        QMessageBox::warning(this, "提示", "请先选择要修改的面试者。");
        return;
    }

    // 从模型中获取选中行的完整记录，这包含了我们需要的 'id'
    QSqlRecord record = m_intervieweeModel->record(currentIndex.row());
    int recordId = record.value("id").toInt(); // 获取记录的ID，用于UPDATE的WHERE子句

    // 创建对话框并用当前记录的数据填充它
    IntervieweeDialog dlg(this);
    dlg.setRecord(record);

    if (dlg.exec() == QDialog::Accepted) {
        // 获取对话框中修改后的数据
        QSqlRecord updatedRecord = dlg.getRecord();

        // 使用 QSqlQuery 执行 UPDATE 操作
        QSqlQuery query;
        query.prepare("UPDATE interviewees SET name=:name, phone=:phone, email=:email, "
                      "applied_position=:applied_position, status=:status, interview_time=:interview_time, "
                      "interviewer=:interviewer, notes=:notes WHERE id=:id");

        // 绑定更新的数据
        query.bindValue(":name", updatedRecord.value("name"));
        query.bindValue(":phone", updatedRecord.value("phone"));
        query.bindValue(":email", updatedRecord.value("email"));
        query.bindValue(":applied_position", updatedRecord.value("applied_position"));
        query.bindValue(":status", updatedRecord.value("status"));
        query.bindValue(":interview_time", updatedRecord.value("interview_time"));
        query.bindValue(":interviewer", updatedRecord.value("interviewer"));
        query.bindValue(":notes", updatedRecord.value("notes"));
        // 绑定WHERE子句的ID
        query.bindValue(":id", recordId);

        if (!query.exec()) {
            QMessageBox::critical(this, "数据库错误", "修改面试者信息失败：\n" + query.lastError().text());
        } else {
            QMessageBox::information(this, "成功", "面试者信息修改成功！");
            // 刷新模型以显示更新后的数据
            m_intervieweeModel->select();
        }
    }
}

// /**
//  * @brief 搜索框文本变化时的槽函数
//  *
//  * - 利用 `QSqlTableModel` 强大的 `setFilter()` 功能实现实时搜索。
//  * - 如果搜索框为空，清除过滤器，显示所有记录。
//  * - 如果有文本，构造一个 SQL `WHERE` 子句（不含 "WHERE" 关键字），
//  * 实现对 "name" 或 "phone" 字段的模糊查询 (LIKE '%...%')。
//  * - 模型会自动应用过滤器并刷新视图，无需手动调用 `select()`。
//  */
// void JobRequirement::on_searchLineEdit_textChanged(const QString &text)
// {
//     if (text.isEmpty()) {
//         // 如果搜索框为空，清除过滤器，显示所有数据
//         m_intervieweeModel->setFilter("");
//     } else {
//         // 根据姓名或电话进行模糊搜索
//         // 注意：这里的字段名 'name' 和 'phone' 必须与数据库表中的列名完全一致
//         QString filterString = QString("name LIKE '%%1%' OR phone LIKE '%%1%'").arg(text);
//         m_intervieweeModel->setFilter(filterString);
//     }
// }


// ===============================================================================================
// ============================= 职位需求槽函数 (Tab 1 - Job Requirements) ===========================
// ===============================================================================================

/**
 * @brief 从数据库加载职位需求数据
 *
 * - 先清空模型中的所有现有行。
 * - 执行 SELECT 查询从 `job_requirements` 表获取所有数据。
 * - 遍历查询结果，将每一行数据填充到 `QStandardItemModel` 中。
 * - **核心技巧**: 将数据库的 `id` 主键存储在第0列的一个自定义角色 `DbIdRole` 中。
 * 这使得ID对用户不可见，但我们可以在程序中随时获取它，用于后续的更新和删除操作。
 */
void JobRequirement::loadJobRequirementsData()
{
    m_jobRequirementModel->removeRows(0, m_jobRequirementModel->rowCount());

    QSqlDatabase db = QSqlDatabase::database();
    if (!db.isOpen()) {
        qWarning() << "错误：数据库未连接，无法加载职位需求数据。";
        return;
    }

    QSqlQuery query(db);
    query.prepare("SELECT id, Department, Position, Duty, NeedNumber, sex, Education, requirement FROM job_requirements ORDER BY id ASC");

    if (!query.exec()) {
        qWarning() << "错误：从数据库加载职位需求数据失败：" << query.lastError().text();
        return;
    }

    int row = 0;
    while (query.next()) {
        m_jobRequirementModel->insertRow(row);
        // 将数据库ID存储在第0列的自定义角色中，它对用户不可见
        m_jobRequirementModel->setData(m_jobRequirementModel->index(row, 0), query.value("id").toInt(), DbIdRole);
        // 设置用户可见的各列数据
        m_jobRequirementModel->setData(m_jobRequirementModel->index(row, 0), query.value("Department").toString());
        m_jobRequirementModel->setData(m_jobRequirementModel->index(row, 1), query.value("Position").toString());
        m_jobRequirementModel->setData(m_jobRequirementModel->index(row, 2), query.value("Duty").toString());
        m_jobRequirementModel->setData(m_jobRequirementModel->index(row, 3), query.value("NeedNumber").toInt());
        m_jobRequirementModel->setData(m_jobRequirementModel->index(row, 4), query.value("sex").toString());
        m_jobRequirementModel->setData(m_jobRequirementModel->index(row, 5), query.value("Education").toString());
        m_jobRequirementModel->setData(m_jobRequirementModel->index(row, 6), query.value("requirement").toString());
        row++;
    }
    // 自动调整行高和列宽以适应内容
    ui->tableViewJobRequirement->resizeRowsToContents();
    ui->tableViewJobRequirement->resizeColumnsToContents();
}

/**
 * @brief 处理 `QStandardItemModel` 中数据项变化的槽函数
 *
 * - 当用户在表格中编辑了任何一个单元格后，此函数被触发。
 * - 它会从被修改项的行中，通过 `DbIdRole` 获取到数据库记录的ID。
 * - 根据被修改的列，确定要更新的数据库字段名。
 * - 执行 `UPDATE` SQL语句，将新的值同步到数据库。
 * - 这种方式实现了 "职位需求" 表格的实时编辑和保存。
 */
void JobRequirement::handleItemChanged(QStandardItem *item)
{
    if (!item) return; // 安全检查

    int row = item->row();
    int column = item->column();
    // 从第0列的自定义角色中获取数据库ID
    QVariant dbIdVariant = m_jobRequirementModel->data(m_jobRequirementModel->index(row, 0), DbIdRole);

    // 如果ID无效（例如，这是一个刚刚添加、尚未存入数据库的行），则不执行任何操作
    if (!dbIdVariant.isValid() || dbIdVariant.toInt() == 0) return;
    int dbId = dbIdVariant.toInt();

    // 根据列索引确定对应的数据库字段名
    QString dbColumnName;
    switch (column) {
    case 0: dbColumnName = "Department";   break;
    case 1: dbColumnName = "Position";     break;
    case 2: dbColumnName = "Duty";         break;
    case 3: dbColumnName = "NeedNumber";   break;
    case 4: dbColumnName = "sex";          break;
    case 5: dbColumnName = "Education";    break;
    case 6: dbColumnName = "requirement";  break;
    default: return; // 无效列
    }

    // 准备并执行 UPDATE 查询
    QSqlQuery query;
    QString queryString = QString("UPDATE job_requirements SET %1 = :newValue WHERE id = :id").arg(dbColumnName);
    query.prepare(queryString);
    query.bindValue(":newValue", item->data(Qt::EditRole));
    query.bindValue(":id", dbId);

    if (!query.exec()) {
        qWarning() << "错误：更新数据库失败：" << query.lastError().text();
    }
}

/**
 * @brief "新增职位" 按钮的槽函数
 *
 * - 首先在 `job_requirements` 数据库表中插入一条空的记录。
 * - 获取新插入记录的自增ID (`lastInsertId`)。
 * - 在 `QStandardItemModel` 的末尾添加一个新行。
 * - 将获取到的新ID存入新行的 `DbIdRole`。
 * - 自动滚动到新行，并使其进入编辑状态，方便用户立即输入。
 */
void JobRequirement::on_btnAddJob_clicked()
{
    QSqlQuery query;
    // 先在数据库中插入一条记录，以获取有效的ID
    query.prepare("INSERT INTO job_requirements (Department, Position) VALUES ('', '')");
    if (!query.exec()) {
        QMessageBox::critical(this, "数据库错误", "新增职位失败！\n" + query.lastError().text());
        return;
    }
    // 获取刚刚插入行的ID
    int newDbId = query.lastInsertId().toInt();

    // 在界面模型中添加新行
    int newRowIndex = m_jobRequirementModel->rowCount();
    m_jobRequirementModel->insertRow(newRowIndex);
    // 将数据库ID与界面行关联起来
    m_jobRequirementModel->setData(m_jobRequirementModel->index(newRowIndex, 0), newDbId, DbIdRole);

    // 提升用户体验：自动滚动到新行、选中并进入编辑模式
    ui->tableViewJobRequirement->scrollToBottom();
    ui->tableViewJobRequirement->selectRow(newRowIndex);
    ui->tableViewJobRequirement->edit(m_jobRequirementModel->index(newRowIndex, 0));
}

/**
 * @brief "删除选中行" 按钮的槽函数
 *
 * - 获取所有被用户选中的行。
 * - 弹出确认对话框。
 * - 使用数据库事务（transaction）来确保操作的原子性：要么所有选中的行都被成功删除，要么都不删除。
 * - 遍历选中的行，从 `DbIdRole` 中获取ID，并执行 `DELETE` SQL语句。
 * - 如果所有删除都成功，则提交事务，并从 `QStandardItemModel` 中移除对应的行。
 * - 如果有任何一条删除失败，则回滚事务，数据库状态回到操作前。
 */
void JobRequirement::on_btnDeleteRow_clicked()
{
    QModelIndexList selectedIndexes = ui->tableViewJobRequirement->selectionModel()->selectedRows();
    if (selectedIndexes.isEmpty()) {
        QMessageBox::information(this, "提示", "请先选择要删除的行。");
        return;
    }
    auto reply = QMessageBox::question(this, "确认删除", "确定要删除选中的记录吗？", QMessageBox::Yes | QMessageBox::No);
    if (reply == QMessageBox::No) return;

    // 使用数据库事务确保操作的原子性
    QSqlDatabase::database().transaction();
    bool success = true;
    QList<int> rowsToRemove; // 存储模型中要删除的行索引

    // 1. 先在数据库中删除
    for (const QModelIndex &index : selectedIndexes) {
        int dbId = m_jobRequirementModel->data(m_jobRequirementModel->index(index.row(), 0), DbIdRole).toInt();
        QSqlQuery query;
        query.prepare("DELETE FROM job_requirements WHERE id = :id");
        query.bindValue(":id", dbId);
        if (query.exec()) {
            rowsToRemove.append(index.row());
        } else {
            success = false;
            qWarning() << "删除失败, ID:" << dbId << ", 错误:" << query.lastError().text();
            break; // 一旦失败，立即停止
        }
    }

    // 2. 根据数据库操作结果，更新UI
    if (success) {
        QSqlDatabase::database().commit(); // 提交事务
        // 从后往前删除，避免因行号变化导致的索引错误
        std::sort(rowsToRemove.begin(), rowsToRemove.end(), std::greater<int>());
        for(int row : rowsToRemove) {
            m_jobRequirementModel->removeRow(row);
        }
        QMessageBox::information(this, "成功", "记录已成功删除。");
    } else {
        QSqlDatabase::database().rollback(); // 回滚事务
        QMessageBox::critical(this, "删除失败", "删除记录时发生错误，操作已回滚。");
    }
}

/**
 * @brief "开启/关闭编辑" 复选框的槽函数
 * @param checked 复选框是否被选中
 */
void JobRequirement::on_checkBox_clicked(bool checked)
{
    ui->tableViewJobRequirement->setEnabled(checked);
}

/**
 * @brief "自动调整行高" 按钮的槽函数
 */
void JobRequirement::on_btnAutoAdjustRowHeigth_clicked()
{
    ui->tableViewJobRequirement->resizeRowsToContents();
}

/**
 * @brief "自动调整列宽" 按钮的槽函数
 */
void JobRequirement::on_btnAutoAdjustColumnWidth_clicked()
{
    ui->tableViewJobRequirement->resizeColumnsToContents();
}

// ===============================================================================================
// ================================= HTML 与 Excel 辅助函数 ======================================
// ===============================================================================================

/**
 * @brief 将职位需求表格数据导出为HTML文件
 *
 * - 动态生成一个包含表格的完整HTML页面字符串。
 * - HTML包含了一些基本的CSS样式，使其更美观。
 * - 将HTML内容写入到由 `determinePersistentHtmlPath()` 决定的文件中。
 */
// void JobRequirement::on_btnRelease_clicked()
// {
//     if (m_jobRequirementModel->rowCount() == 0) {
//         QMessageBox::information(this, "提示", "表格中没有数据可供发布。");
//         ui->btnCheckInHtml->setEnabled(false);
//         return;
//     }

//     // 构建HTML字符串
//     QString htmlContent;
//     htmlContent += "<!DOCTYPE html>\n<html lang=\"zh-CN\">\n<head>\n";
//     htmlContent += "    <meta charset=\"UTF-8\">\n";
//     htmlContent += "    <title>职位需求信息</title>\n";
//     htmlContent += "    <style>\n";
//     htmlContent += "        body { font-family: Arial, 'Microsoft YaHei', sans-serif; margin: 20px; background-color: #f8f9fa; color: #333; }\n";
//     htmlContent += "        h1 { text-align: center; color: #007bff; }\n";
//     htmlContent += "        table { width: 90%; margin: 20px auto; border-collapse: collapse; box-shadow: 0 4px 8px rgba(0,0,0,0.1); background-color: #fff; }\n";
//     htmlContent += "        th, td { border: 1px solid #dee2e6; padding: 12px 15px; text-align: left; vertical-align: top; }\n";
//     htmlContent += "        th { background-color: #007bff; color: white; font-weight: bold; }\n";
//     htmlContent += "        tr:nth-child(even) { background-color: #f2f2f2; }\n";
//     htmlContent += "        tr:hover { background-color: #e9ecef; }\n";
//     htmlContent += "        td { word-wrap: break-word; white-space: pre-wrap; } /* white-space: pre-wrap; 保留换行符和空格 */ \n";
//     htmlContent += "    </style>\n</head>\n<body>\n";
//     htmlContent += "    <h1>职位需求信息表</h1>\n    <table>\n        <thead>\n            <tr>\n";

//     // 添加表头
//     for (int col = 0; col < m_jobRequirementModel->columnCount(); ++col) {
//         htmlContent += "                <th>" + m_jobRequirementModel->headerData(col, Qt::Horizontal).toString().toHtmlEscaped() + "</th>\n";
//     }
//     htmlContent += "            </tr>\n        </thead>\n        <tbody>\n";

//     // 添加数据行
//     for (int row = 0; row < m_jobRequirementModel->rowCount(); ++row) {
//         htmlContent += "            <tr>\n";
//         for (int col = 0; col < m_jobRequirementModel->columnCount(); ++col) {
//             QString cellData = m_jobRequirementModel->data(m_jobRequirementModel->index(row, col)).toString();
//             // toHtmlEscaped 会处理特殊字符如 < > &，CSS的 pre-wrap 会处理换行
//             htmlContent += "                <td>" + cellData.toHtmlEscaped() + "</td>\n";
//         }
//         htmlContent += "            </tr>\n";
//     }
//     htmlContent += "        </tbody>\n    </table>\n</body>\n</html>\n";

//     // 写入文件
//     QFile file(m_persistentHtmlFilePath);
//     if (file.open(QIODevice::WriteOnly | QIODevice::Text | QIODevice::Truncate)) { // Truncate: 覆盖旧文件
//         QTextStream out(&file);
//         out.setEncoding(QStringConverter::Utf8); // 确保使用UTF-8编码
//         out << htmlContent;
//         file.close();
//         QMessageBox::information(this, "成功", "HTML内容已生成并保存。\n文件位置: " + QDir::toNativeSeparators(m_persistentHtmlFilePath));
//         ui->btnCheckInHtml->setEnabled(true);
//     } else {
//         qWarning() << "无法打开持久化HTML文件进行写入：" << m_persistentHtmlFilePath << "错误：" << file.errorString();
//         QMessageBox::critical(this, "错误", "无法保存发布的HTML文件。\n错误: " + file.errorString());
//         ui->btnCheckInHtml->setEnabled(false);
//     }
// }

/**
 * @brief 使用系统默认浏览器打开已发布的HTML文件
 */
// void JobRequirement::on_btnCheckInHtml_clicked()
// {
//     if (!QFile::exists(m_persistentHtmlFilePath)) {
//         QMessageBox::information(this, "提示", "没有已发布的HTML内容可供预览。\n请先点击“发布”按钮生成内容。");
//         ui->btnCheckInHtml->setEnabled(false);
//         return;
//     }

//     // 使用 QDesktopServices 打开本地文件
//     if (!QDesktopServices::openUrl(QUrl::fromLocalFile(m_persistentHtmlFilePath))) {
//         QMessageBox::critical(this, "错误", "无法在浏览器中打开HTML文件。\n文件路径: " + QDir::toNativeSeparators(m_persistentHtmlFilePath));
//     }
// }

/**
 * @brief 确定一个持久、可靠的路径来存储发布的HTML文件
 *
 * 优先使用 AppLocalDataLocation，这是一个标准的、跨平台的应用程序数据存储位置。
 * 如果该位置不可用，则回退到当前工作目录。
 * @return 返回一个包含完整文件名的QString路径
 */
QString JobRequirement::determinePersistentHtmlPath() const
{
    // 使用应用程序的本地数据位置来存储持久化文件，这是推荐做法
    QString dataPath = QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation);
    if (dataPath.isEmpty()) {
        dataPath = QDir::currentPath(); // 回退方案
        qWarning() << "AppLocalDataLocation 不可用，使用当前路径：" << dataPath;
    }

    QDir dir(dataPath);
    // 在数据路径下为我们的应用创建一个子目录，避免文件混乱
    QString subDirName = "HRManagementSystem"; // 选择一个唯一的应用名
    if (!dir.exists(subDirName)) {
        dir.mkpath(subDirName);
    }
    return dataPath + QDir::separator() + subDirName + QDir::separator() + "job_requirements_published.html";
}

/**
 * @brief 导出到Excel的槽函数
 */
void JobRequirement::on_exportButton_clicked()
{
    exportToExcel();
}

/**
 * @brief 从Excel导入的槽函数
 */
void JobRequirement::on_btnFromExcel_clicked()
{
    importFromExcel();
}

/**
 * @brief (仅Windows) 将职位需求表格导出到Excel文件
 *
 * 使用 `QAxObject` (ActiveX) 与本地安装的 Microsoft Excel 进行交互。
 * - 启动Excel应用进程。
 * - 创建新的工作簿和工作表。
 * - 遍历 `QStandardItemModel`，将表头和数据逐一写入到工作表的单元格中。
 * - 弹出文件保存对话框，让用户选择保存位置和文件名。
 * - 保存并关闭Excel。
 */
void JobRequirement::exportToExcel()
{
#ifndef Q_OS_WIN
    QMessageBox::information(this, "功能限制", "此功能仅在Windows平台上可用。");
    return;
#else
    QAxObject excel("Excel.Application");
    if (excel.isNull()) {
        QMessageBox::critical(this, "错误", "无法启动 Excel 应用程序。\n请确保已安装 Microsoft Excel。");
        return;
    }

    excel.setProperty("Visible", false); // 在后台运行
    QAxObject *workbooks = excel.querySubObject("Workbooks");
    QAxObject *workbook = workbooks->querySubObject("Add");
    QAxObject *worksheet = workbook->querySubObject("Worksheets(int)", 1);

    // 写入表头
    for (int col = 0; col < m_jobRequirementModel->columnCount(); ++col) {
        worksheet->querySubObject("Cells(int, int)", 1, col + 1)->setProperty("Value", m_jobRequirementModel->headerData(col, Qt::Horizontal).toString());
    }

    // 写入数据
    for (int row = 0; row < m_jobRequirementModel->rowCount(); ++row) {
        for (int col = 0; col < m_jobRequirementModel->columnCount(); ++col) {
            worksheet->querySubObject("Cells(int, int)", row + 2, col + 1)->setProperty("Value", m_jobRequirementModel->data(m_jobRequirementModel->index(row, col)).toString());
        }
    }

    // 自动调整列宽
    worksheet->querySubObject("Columns")->dynamicCall("AutoFit");

    // 弹出保存对话框
    QString fileName = QFileDialog::getSaveFileName(this, "保存 Excel 文件", "", "Excel 文件 (*.xlsx)");
    if (!fileName.isEmpty()) {
        workbook->dynamicCall("SaveAs(const QString&)", QDir::toNativeSeparators(fileName));
        workbook->dynamicCall("Close()");
        excel.dynamicCall("Quit()");
        QMessageBox::information(this, "成功", "文件已成功导出到 Excel！");
    } else {
        // 用户取消保存
        workbook->dynamicCall("Close(false)"); // 不保存更改并关闭
        excel.dynamicCall("Quit()");
    }
#endif
}

/**
 * @brief (仅Windows) 从Excel文件导入数据
 *
 * 同样使用 `QAxObject`。
 * - 弹出文件打开对话框，让用户选择要导入的Excel文件。
 * - 打开Excel文件并获取第一个工作表。
 * - 获取工作表中已使用的区域（UsedRange）来确定数据范围。
 * - 清空当前 `QStandardItemModel` 中的数据。
 * - 遍历Excel工作表的行和列，将数据读取出来并填充到模型中。
 * - **注意:** 此函数不会将导入的数据立即存入数据库，用户需要手动编辑或通过其他机制触发保存。
 */
void JobRequirement::importFromExcel()
{
#ifndef Q_OS_WIN
    QMessageBox::information(this, "功能限制", "此功能仅在Windows平台上可用。");
    return;
#else
    QString fileName = QFileDialog::getOpenFileName(this, "打开 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)");
    if (fileName.isEmpty()) return;

    QAxObject excel("Excel.Application");
    if (excel.isNull()) {
        QMessageBox::critical(this, "错误", "无法启动 Excel 应用程序。");
        return;
    }

    excel.setProperty("Visible", false);
    excel.setProperty("DisplayAlerts", false); // 禁止Excel弹出任何警告

    QAxObject *workbooks = excel.querySubObject("Workbooks");
    QAxObject *workbook = workbooks->querySubObject("Open(const QString&)", fileName);
    if (!workbook) {
        QMessageBox::critical(this, "错误", "无法打开指定的 Excel 文件。");
        excel.dynamicCall("Quit()");
        return;
    }

    QAxObject *worksheet = workbook->querySubObject("Worksheets(int)", 1);
    QAxObject *usedRange = worksheet->querySubObject("UsedRange");
    if (!usedRange) {
        QMessageBox::warning(this, "警告", "工作表为空或无法读取。");
        workbook->dynamicCall("Close()");
        excel.dynamicCall("Quit()");
        return;
    }

    int rowCount = usedRange->querySubObject("Rows")->property("Count").toInt();
    int colCount = usedRange->querySubObject("Columns")->property("Count").toInt();

    // 清空模型和数据库中的现有数据，这是一个危险操作，最好先警告用户
    auto reply = QMessageBox::question(this, "确认导入", "导入操作将清空当前表格的所有数据，并用Excel文件内容替换。\n确定要继续吗？", QMessageBox::Yes | QMessageBox::No);
    if(reply == QMessageBox::No) {
        workbook->dynamicCall("Close()");
        excel.dynamicCall("Quit()");
        return;
    }

    // TODO: 在清空前最好提供备份当前数据的选项
    m_jobRequirementModel->removeRows(0, m_jobRequirementModel->rowCount());
    // 这里也应该清空数据库表
    QSqlQuery truncateQuery;
    truncateQuery.exec("DELETE FROM job_requirements");

    // 开始读取数据并直接写入数据库和模型
    // 假设Excel第一行是数据，如果第一行是表头，则循环应从 r=2 开始
    for (int r = 1; r <= rowCount; ++r) {
        QList<QVariant> rowData;
        for (int c = 1; c <= m_jobRequirementModel->columnCount(); ++c) {
            if (c > colCount) {
                rowData.append(""); // 如果Excel列数不够，则添加空字符串
            } else {
                QAxObject* cell = worksheet->querySubObject("Cells(int,int)", r, c);
                rowData.append(cell->property("Value"));
                delete cell;
            }
        }

        // 将读取到的一行数据插入数据库
        QSqlQuery insertQuery;
        insertQuery.prepare("INSERT INTO job_requirements (Department, Position, Duty, NeedNumber, sex, Education, requirement) "
                            "VALUES (?, ?, ?, ?, ?, ?, ?)");
        for(const QVariant& val : rowData) {
            insertQuery.addBindValue(val);
        }

        if(insertQuery.exec()) {
            // 插入成功后，再更新UI模型
            int newDbId = insertQuery.lastInsertId().toInt();
            int newRowIndex = m_jobRequirementModel->rowCount();
            m_jobRequirementModel->insertRow(newRowIndex);
            m_jobRequirementModel->setData(m_jobRequirementModel->index(newRowIndex, 0), newDbId, DbIdRole);
            for(int i=0; i<rowData.size(); ++i) {
                m_jobRequirementModel->setData(m_jobRequirementModel->index(newRowIndex, i), rowData[i]);
            }
        }
    }

    // 清理COM对象
    delete usedRange;
    workbook->dynamicCall("Close(false)");
    excel.dynamicCall("Quit()");

    ui->tableViewJobRequirement->resizeRowsToContents();
    ui->tableViewJobRequirement->resizeColumnsToContents();
    QMessageBox::information(this, "成功", QString("成功导入 %1 行数据。").arg(m_jobRequirementModel->rowCount()));
#endif
}


// #include "jobrequirement.h"
// #include "ui_jobrequirement.h"

// #include <QStandardItemModel>
// #include <QTableView>
// #include <QFileDialog>
// #include <QMessageBox>
// #include <QDesktopServices>
// #include <QDir>
// #include <QStandardPaths>
// #include <QDebug>

// // Qt SQL 模块
// #include <QSqlDatabase>
// #include <QSqlQuery>
// #include <QSqlError>
// // 如果需要Excel功能，请包含ActiveQt
// #include <ActiveQt/QAxObject>
// JobRequirement::JobRequirement(QWidget *parent)
//     : QWidget(parent)
//     , ui(new Ui::JobRequirement)
//     , model(new QStandardItemModel(0, 7))// 设置0行7列
// {
//     ui->setupUi(this);


//     // 设置
//     model->setHorizontalHeaderLabels({"部门", "需求职位名称", "主要职责", "招聘人数","性别要求","学历要求","任职资格要求"});
//     ui->tableViewJobRequirement->setModel(model);


//     QSqlDatabase db = QSqlDatabase::database(); // 获取默认连接
//     if (!db.isOpen())
//     {
//         qWarning() << "数据库连接在 JobRequirement 构造时未打开！请确保 main.cpp 中的连接有效。";
//     }

//     loadDataFromDatabase(); // 窗口创建时从数据库加载数据

//     // ui->tableViewJobRequirement->resizeRowsToContents();
//     // ui->tableViewJobRequirement->resizeColumnsToContents();

//     connect(model, &QStandardItemModel::itemChanged, this, &JobRequirement::handleItemChanged);



//     m_persistentHtmlFilePath = determinePersistentHtmlPath();
//     qDebug() << "持久化HTML文件路径：" << m_persistentHtmlFilePath;

//     if (QFile::exists(m_persistentHtmlFilePath)) {
//         ui->btnCheckInHtml->setEnabled(true);
//         qDebug() << "找到已存在的持久化HTML文件。“在HTML中查看”按钮已启用。";
//     } else {
//         ui->btnCheckInHtml->setEnabled(false);
//         qDebug() << "未找到持久化HTML文件。“在HTML中查看”按钮已禁用。";
//     }

//     // 根据您的原始代码，表格初始时不可编辑
//     ui->tableViewJobRequirement->setEnabled(false);
//     ui->checkBox->setChecked(false);

// }



// //
// QString JobRequirement::determinePersistentHtmlPath() const
// {
//     // 使用应用程序的本地数据位置来存储持久化文件
//     QString dataPath = QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation);
//     if (dataPath.isEmpty()) {
//         // 如果 AppLocalDataLocation 不可用，回退到当前目录 (不太理想)
//         dataPath = QDir::currentPath();
//         qWarning() << "AppLocalDataLocation 不可用，使用当前路径：" << dataPath;
//     }

//     QDir dir(dataPath);
//     // 为我们的应用程序文件创建一个子目录（如果它还不存在）
//     QString subDirName = "MyCompanyJobRequirements"; // 选择一个合适的名称
//     if (!dir.exists(subDirName)) {
//         if (!dir.mkpath(subDirName)) {
//             qWarning() << "无法创建目录：" << dir.filePath(subDirName);
//             // 如果子目录创建失败，则直接返回 dataPath 下的文件名
//             return dataPath + QDir::separator() + "job_requirements_published.html";
//         }
//     }
//     return dataPath + QDir::separator() + subDirName + QDir::separator() + "job_requirements_published.html";
// }


// void JobRequirement::loadDataFromDatabase()
// {                                                      // 从数据库加载表格
//     model->removeRows(0, model->rowCount());

//     QSqlDatabase db = QSqlDatabase::database(); // 获取默认连接
//     if (!db.isOpen()) {
//         qWarning() << "错误：数据库未连接，无法加载数据。";
//         QMessageBox::critical(this, "数据库错误", "数据库未连接，无法加载职位需求数据。");
//         return;
//     }

//     QSqlQuery query(db); // 直接使用默认连接创建查询
//     // 注意：您的表名是 job_requirements (根据您之前的描述)
//     query.prepare("SELECT id, Department, Position, Duty, NeedNumber, sex, Education, requirement FROM job_requirements ORDER BY id ASC");

//     if (!query.exec()) {
//         qWarning() << "错误：从数据库加载数据失败：" << query.lastError().text();
//         QMessageBox::critical(this, "查询错误", "加载职位需求数据失败：\n" + query.lastError().text());
//         return;
//     }

//     int row = 0;
//     while (query.next()) {
//         model->insertRow(row);
//         model->setData(model->index(row, 0), query.value("id").toInt(), DbIdRole);
//         model->setData(model->index(row, 0), query.value("Department").toString());
//         model->setData(model->index(row, 1), query.value("Position").toString());
//         model->setData(model->index(row, 2), query.value("Duty").toString());
//         model->setData(model->index(row, 3), query.value("NeedNumber").toInt());
//         model->setData(model->index(row, 4), query.value("sex").toString());
//         model->setData(model->index(row, 5), query.value("Education").toString());
//         model->setData(model->index(row, 6), query.value("requirement").toString());
//         row++;
//     }
//     qDebug() << "从数据库加载了 " << row << " 行职位需求数据。";

//     if (row > 0 || model->columnCount() > 0)
//     { // 确保有数据或至少有表头才调整
//         ui->tableViewJobRequirement->resizeRowsToContents();
//         ui->tableViewJobRequirement->resizeColumnsToContents();
//         qDebug() << "表格行高和列宽已自动调整。";
//     }
// }



// void JobRequirement::handleItemChanged(QStandardItem *item)    // 表格数据加载到数据库
// {
//     if (!item) return;

//     int row = item->row();
//     int column = item->column();

//     QVariant dbIdVariant = model->data(model->index(row, 0), DbIdRole);
//     if (!dbIdVariant.isValid() || dbIdVariant.toInt() == 0) {
//         qWarning() << "警告：在行 " << row << " 中未找到有效的数据库ID，无法更新。";
//         return;
//     }
//     int dbId = dbIdVariant.toInt();

//     QString dbColumnName;
//     QVariant newValue = item->data(Qt::EditRole);

//     switch (column) {
//     case 0: dbColumnName = "Department"; break;
//     case 1: dbColumnName = "Position"; break;
//     case 2: dbColumnName = "Duty"; break;
//     case 3:
//         dbColumnName = "NeedNumber";
//         bool ok;
//         newValue = item->text().toInt(&ok);
//         if (!ok) {
//             qWarning() << "招聘人数 '" << item->text() << "' 转换为整数失败。";
//             QMessageBox::warning(this, "输入错误", "招聘人数必须为有效数字。之前的值将被保留。");
//             // 恢复旧值:
//             // 为了简单，这里我们不直接从数据库读旧值，而是阻止更新
//             // 更好的做法是，在开始编辑前保存旧值，或在失败时从模型中获取之前的值（如果模型没有立即更新的话）
//             // 或者，断开 itemChanged 的连接，设置数据，再连回去，以避免递归
//             model->blockSignals(true); // 阻止信号循环
//             model->setData(item->index(), model->item(row, column)->data(Qt::UserRole + 2)); // 假设旧值存在这里
//             model->blockSignals(false);
//             return;
//         }
//         break;
//     case 4: dbColumnName = "sex"; break;
//     case 5: dbColumnName = "Education"; break;
//     case 6: dbColumnName = "requirement"; break;
//     default:
//         qWarning() << "未知的列索引：" << column;
//         return;
//     }

//     QSqlDatabase db = QSqlDatabase::database(); // 获取默认连接
//     if (!db.isOpen()) {
//         QMessageBox::critical(this, "数据库错误", "数据库未连接，无法更新数据。");
//         return;
//     }

//     QSqlQuery query(db);
//     // 表名 job_requirements
//     QString queryString = QString("UPDATE job_requirements SET %1 = :newValue WHERE id = :id").arg(dbColumnName);
//     query.prepare(queryString);
//     query.bindValue(":newValue", newValue);
//     query.bindValue(":id", dbId);

//     if (!query.exec()) {
//         qWarning() << "错误：更新数据库失败：" << query.lastError().text();
//         QMessageBox::critical(this, "数据库更新失败", "更新记录到数据库时发生错误。\n" + query.lastError().text());
//     } else {
//         qDebug() << "数据库记录 ID:" << dbId << ", 列:" << dbColumnName << " 已更新为:" << newValue.toString();
//         // 如果上面 NeedNumber 转换失败时恢复了旧值，这里可以保存一下当前有效的值到 UserRole+2 以备下次使用
//         // model->item(row, column)->setData(newValue, Qt::UserRole + 2);
//     }
// }



// void JobRequirement::on_btnDeleteRow_clicked()                                                  // 删除表格数据
// {
//     QModelIndexList selectedIndexes = ui->tableViewJobRequirement->selectionModel()->selectedRows();
//     if (selectedIndexes.isEmpty()) {
//         QMessageBox::information(this, "提示", "请先选择要删除的行。");
//         return;
//     }

//     QMessageBox::StandardButton reply;
//     reply = QMessageBox::question(this, "确认删除",
//                                   QString("确定要删除选中的 %1 条记录吗？此操作不可恢复。").arg(selectedIndexes.count()),
//                                   QMessageBox::Yes | QMessageBox::No);
//     if (reply == QMessageBox::No) {
//         return;
//     }

//     QSqlDatabase db = QSqlDatabase::database(); // 获取默认连接
//     if (!db.isOpen()) {
//         QMessageBox::critical(this, "数据库错误", "数据库未连接，无法删除数据。");
//         return;
//     }

//     QList<QPair<int, int>> rowsToDelete;
//     for (const QModelIndex &index : selectedIndexes) {
//         int modelRow = index.row();
//         QVariant dbIdVariant = model->data(model->index(modelRow, 0), DbIdRole);
//         if (dbIdVariant.isValid() && dbIdVariant.toInt() != 0) {
//             rowsToDelete.append(qMakePair(modelRow, dbIdVariant.toInt()));
//         } else {
//             qWarning() << "警告：行 " << modelRow << " 没有有效的数据库ID，将跳过删除。";
//         }
//     }

//     if (rowsToDelete.isEmpty()) {
//         QMessageBox::information(this, "提示", "没有可从数据库删除的有效记录。");
//         return;
//     }

//     std::sort(rowsToDelete.begin(), rowsToDelete.end(), [](const QPair<int, int>& a, const QPair<int, int>& b) {
//         return a.first > b.first;
//     });

//     bool allSuccess = true;
//     if (!db.transaction()) {
//         qWarning() << "无法开始数据库事务：" << db.lastError().text();
//         QMessageBox::critical(this, "数据库错误", "无法开始数据库事务。");
//         return;
//     }

//     for (const auto& rowPair : rowsToDelete) {
//         int dbId = rowPair.second;
//         QSqlQuery query(db);
//         // 表名 job_requirements
//         query.prepare("DELETE FROM job_requirements WHERE id = :id");
//         query.bindValue(":id", dbId);

//         if (!query.exec()) {
//             qWarning() << "错误：从数据库删除记录 ID " << dbId << " 失败：" << query.lastError().text();
//             allSuccess = false;
//             break;
//         } else {
//             qDebug() << "成功从数据库删除记录 ID:" << dbId;
//         }
//     }

//     if (allSuccess) {
//         if (!db.commit()) {
//             qWarning() << "数据库事务提交失败: " << db.lastError().text();
//             QMessageBox::critical(this, "数据库错误", "提交删除操作到数据库失败！");
//             db.rollback();
//         } else {
//             for (const auto& rowPair : rowsToDelete) {
//                 model->removeRow(rowPair.first);
//             }
//             QMessageBox::information(this, "成功", "选中的记录已成功删除。");
//         }
//     } else {
//         db.rollback();
//         QMessageBox::critical(this, "删除失败", "删除部分或全部记录时发生错误，操作已回滚。");
//     }
// }

// JobRequirement::~JobRequirement()
// {
//     delete ui;
// }

// void JobRequirement::on_btnAddJob_clicked()
// {
//     QSqlDatabase db = QSqlDatabase::database(); // 获取默认连接
//     if (!db.isOpen()) {
//         QMessageBox::critical(this, "数据库错误", "数据库未连接，无法添加新职位。");
//         return;
//     }

//     QSqlQuery query(db);
//     // 您的表名是 job_requirements
//     query.prepare("INSERT INTO job_requirements (Department, Position, Duty, NeedNumber, sex, Education, requirement) "
//                   "VALUES (:dept, :pos, :duty, :need, :sex_val, :edu, :req)");
//     // 为新行插入一些默认值或空值到数据库
//     query.bindValue(":dept", "");       // 默认空字符串
//     query.bindValue(":pos", "");       // 默认空字符串
//     query.bindValue(":duty", "");      // 默认空字符串
//     query.bindValue(":need", 0);       // 默认0
//     query.bindValue(":sex_val", "");   // 默认空字符串 (注意MySQL中sex列名)
//     query.bindValue(":edu", "");       // 默认空字符串
//     query.bindValue(":req", "");       // 默认空字符串

//     if (!query.exec()) {
//         qWarning() << "错误：新增职位需求到数据库失败：" << query.lastError().text();
//         QMessageBox::critical(this, "数据库错误", "新增职位需求到数据库失败！\n" + query.lastError().text());
//         return;
//     }

//     // 获取刚插入记录的自增ID
//     QVariant lastId = query.lastInsertId();
//     if (!lastId.isValid()) {
//         qWarning() << "错误：无法获取新插入记录的ID。可能是表结构问题或数据库不支持。";
//         QMessageBox::critical(this, "数据库错误", "无法获取新记录的ID。");
//         // 考虑是否需要回滚或删除刚刚插入的没有有效ID的行（如果允许插入空ID的话）
//         return;
//     }
//     int newDbId = lastId.toInt();
//     qDebug() << "数据库中新插入记录的 ID:" << newDbId;


//     // 在 QStandardItemModel 中添加新行
//     int newRowIndex = model->rowCount();
//     model->insertRow(newRowIndex);

//     // 关键步骤：将数据库ID存储在模型行的 UserRole 中
//     model->setData(model->index(newRowIndex, 0), newDbId, DbIdRole);
//     qDebug() << "已将 DB ID" << newDbId << "存储到模型行" << newRowIndex << "的 DbIdRole 中。";


//     // 用与数据库插入时对应的默认值填充模型中的新行
//     // 这些值用户之后可以在UI上编辑
//     model->setData(model->index(newRowIndex, 0), ""); // 部门
//     model->setData(model->index(newRowIndex, 1), ""); // 需求职位名称
//     model->setData(model->index(newRowIndex, 2), ""); // 主要职责
//     model->setData(model->index(newRowIndex, 3), 0);  // 招聘人数
//     model->setData(model->index(newRowIndex, 4), ""); // 性别要求
//     model->setData(model->index(newRowIndex, 5), ""); // 学历要求
//     model->setData(model->index(newRowIndex, 6), ""); // 任职资格要求

//     qDebug() << "新增职位需求在模型中创建成功，准备UI刷新。";
//     ui->tableViewJobRequirement->scrollToBottom(); // 滚动到新添加的行
//     ui->tableViewJobRequirement->selectRow(newRowIndex); // 选中新行，方便用户编辑
//     ui->tableViewJobRequirement->edit(model->index(newRowIndex, 0)); // 让第一个单元格进入编辑模式
// }


// void JobRequirement::on_checkBox_clicked(bool checked)
// {
//     if(checked)
//     {
//         ui->tableViewJobRequirement->setEnabled(true);
//     }
//     else
//     {
//         ui->tableViewJobRequirement->setEnabled(false);
//     }
// }





// void JobRequirement::on_btnAutoAdjustRowHeigth_clicked()
// {
//     ui->tableViewJobRequirement->resizeRowsToContents(); // 自适应
// }


// void JobRequirement::on_btnAutoAdjustColumnWidth_clicked()
// {
//     ui->tableViewJobRequirement->resizeColumnsToContents(); // 自适应
// }

// void JobRequirement::on_exportButton_clicked()
// {
//     exportToExcel();
// }

// void JobRequirement::exportToExcel()
// {
//     QAxObject excel("Excel.Application");
//     if (excel.isNull()) {
//         QMessageBox::critical(this, "错误", "无法启动 Excel 应用程序");
//         return;
//     }

//     excel.setProperty("Visible", true); // 显示 Excel 应用程序
//     QAxObject *workbooks = excel.querySubObject("Workbooks");
//     QAxObject *workbook = workbooks->querySubObject("Add"); // 创建一个新的工作簿
//     QAxObject *worksheet = workbook->querySubObject("Worksheets(int)", 1); // 获取第一个工作表

//     // 写入表头
//     for (int col = 0; col < model->columnCount(); ++col) {
//         worksheet->querySubObject("Cells(int, int)", 1, col + 1)->setProperty("Value", model->headerData(col, Qt::Horizontal).toString());
//     }

//     // 写入数据
//     for (int row = 0; row < model->rowCount(); ++row) {
//         for (int col = 0; col < model->columnCount(); ++col) {
//             QModelIndex index = model->index(row, col);
//             worksheet->querySubObject("Cells(int, int)", row + 2, col + 1)->setProperty("Value", model->data(index).toString());
//         }
//     }

//     // 保存文件
//     QString fileName = QFileDialog::getSaveFileName(this, "保存 Excel 文件", "", "Excel 文件 (*.xlsx)");
//     if (!fileName.isEmpty()) {
//         workbook->dynamicCall("SaveAs(const QString&)", fileName);
//         QMessageBox::information(this, "成功", "文件已成功导出到 Excel！");
//     }

//     workbook->dynamicCall("Close()"); // 关闭工作簿
//     excel.dynamicCall("Quit()"); // 退出 Excel 应用程序
//     excel.clear(); // 清除对象
// }



// void JobRequirement::on_btnFromExcel_clicked()
// {
//     importFromExcel();
// }

// void JobRequirement::importFromExcel()
// {
//     QString fileName = QFileDialog::getOpenFileName(this, "打开 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)");
//     if (fileName.isEmpty()) {
//         return;
//     }

//     QAxObject excel("Excel.Application");
//     if (excel.isNull()) {
//         QMessageBox::critical(this, "错误", "无法启动 Excel 应用程序。\n请确保已安装 Microsoft Excel。");
//         return;
//     }

//     excel.setProperty("Visible", false);
//     excel.setProperty("DisplayAlerts", false); // Suppress alerts like "file in use"

//     QAxObject *workbooks = excel.querySubObject("Workbooks");
//     if (!workbooks) {
//         QMessageBox::critical(this, "错误", "无法访问 Excel 工作簿集合。");
//         excel.dynamicCall("Quit()");
//         return;
//     }

//     QAxObject *workbook = workbooks->querySubObject("Open(const QString&)", fileName);
//     if (!workbook) {
//         QMessageBox::critical(this, "错误", "无法打开指定的 Excel 文件：\n" + fileName);
//         excel.dynamicCall("Quit()");
//         return;
//     }

//     QAxObject *worksheet = workbook->querySubObject("Worksheets(int)", 1); // Get the first worksheet
//     if (!worksheet) {
//         QMessageBox::critical(this, "错误", "无法访问 Excel 文件中的第一个工作表。");
//         workbook->dynamicCall("Close()");
//         excel.dynamicCall("Quit()");
//         return;
//     }

//     // --- Get the used range in the worksheet ---
//     QAxObject *usedRange = worksheet->querySubObject("UsedRange");
//     if(!usedRange){
//         QMessageBox::warning(this, "警告", "无法获取工作表的使用范围，可能为空。");
//         worksheet->clear(); // Release worksheet
//         workbook->dynamicCall("Close()");
//         excel.dynamicCall("Quit()");
//         return;
//     }

//     QAxObject *rows = usedRange->querySubObject("Rows");
//     QAxObject *columns = usedRange->querySubObject("Columns");

//     if (!rows || !columns) {
//         QMessageBox::critical(this, "错误", "无法获取行列信息。");
//         if(rows) delete rows;
//         if(columns) delete columns;
//         delete usedRange;
//         workbook->dynamicCall("Close()");
//         excel.dynamicCall("Quit()");
//         return;
//     }

//     int rowCount = rows->property("Count").toInt();
//     int colCount = columns->property("Count").toInt();

//     delete rows;    // Clean up COM object wrappers
//     delete columns;
//     delete usedRange;


//     // --- Clear existing data from the model ---
//     model->removeRows(0, model->rowCount());

//     // --- Read data ---
//     // Assuming data starts from the first row in Excel.
//     // If Excel has headers in the first row, start data_row_offset = 1 (or r = 2 in loop)
//     int data_start_row_in_excel = 1; // 1-based index for Excel rows
//     // If your Excel file has a header row that you want to skip:
//     // int data_start_row_in_excel = 2;


//     for (int r = data_start_row_in_excel; r <= rowCount; ++r) {
//         model->insertRow(model->rowCount()); // Add a new row to the QStandardItemModel
//         int newModelRow = model->rowCount() - 1;

//         for (int c = 1; c <= model->columnCount(); ++c) { // Iterate up to the number of columns in your QTableView model
//             if (c > colCount) { // If Excel sheet has fewer columns than our model
//                 break; // Stop reading columns for this row
//             }

//             QAxObject *cell = worksheet->querySubObject("Cells(int, int)", r, c);
//             if (cell) {
//                 QVariant cellValue = cell->property("Value");
//                 model->setData(model->index(newModelRow, c - 1), cellValue.toString());
//                 delete cell; // Clean up cell COM object wrapper
//             }
//         }
//     }

//     // --- Clean up ---
//     workbook->dynamicCall("Close(false)"); // Close the workbook without saving changes
//     excel.dynamicCall("Quit()");           // Quit Excel application

//     // No need to delete worksheet, workbook, workbooks explicitly here if they were
//     // member variables, but since they are local QAxObject*, their destructors
//     // will handle the release. However, for pointers from querySubObject, it's good to delete.
//     // The main `excel` object's destructor will call `clear()` which releases the COM interface.

//     ui->tableViewJobRequirement->resizeRowsToContents();
//     ui->tableViewJobRequirement->resizeColumnsToContents();

//     QMessageBox::information(this, "成功", QString("成功从 %1 导入 %2 行数据。").arg(QFileInfo(fileName).fileName()).arg(model->rowCount()));
// }

// void JobRequirement::on_btnRelease_clicked()
// {
//     if (model->rowCount() == 0) {
//         QMessageBox::information(this, "提示", "表格中没有数据可供发布。");
//         ui->btnCheckInHtml->setEnabled(false);
//         // 可选：如果文件存在，是否删除它或保留它作为最后一次的良好状态
//         // if (QFile::exists(m_persistentHtmlFilePath)) {
//         //     QFile::remove(m_persistentHtmlFilePath);
//         // }
//         return;
//     }

//     QString htmlContent;
//     htmlContent += "<!DOCTYPE html>\n";
//     htmlContent += "<html lang=\"zh-CN\">\n";
//     htmlContent += "<head>\n";
//     htmlContent += "    <meta charset=\"UTF-8\">\n";
//     htmlContent += "    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n";
//     htmlContent += "    <title>职位需求信息</title>\n";
//     htmlContent += "    <style>\n";
//     htmlContent += "        body { font-family: Arial, 'Microsoft YaHei', sans-serif; margin: 20px; background-color: #f8f9fa; color: #333; }\n";
//     htmlContent += "        h1 { text-align: center; color: #007bff; margin-bottom: 30px; }\n";
//     htmlContent += "        table { width: 90%; margin: 20px auto; border-collapse: collapse; box-shadow: 0 4px 8px rgba(0,0,0,0.1); background-color: #fff; }\n";
//     htmlContent += "        th, td { border: 1px solid #dee2e6; padding: 12px 15px; text-align: left; vertical-align: top; }\n";
//     htmlContent += "        th { background-color: #007bff; color: white; font-weight: bold; }\n";
//     htmlContent += "        tr:nth-child(even) { background-color: #f2f2f2; }\n";
//     htmlContent += "        tr:hover { background-color: #e9ecef; }\n";
//     htmlContent += "        td { word-wrap: break-word; max-width: 300px; white-space: pre-wrap; } /* 添加 white-space: pre-wrap; 来处理换行符 */ \n";
//     htmlContent += "    </style>\n";
//     htmlContent += "</head>\n";
//     htmlContent += "<body>\n";
//     htmlContent += "    <h1>职位需求信息表</h1>\n";
//     htmlContent += "    <table>\n";
//     htmlContent += "        <thead>\n";
//     htmlContent += "            <tr>\n";
//     for (int col = 0; col < model->columnCount(); ++col) {
//         htmlContent += "                <th>" + model->headerData(col, Qt::Horizontal).toString().toHtmlEscaped() + "</th>\n";
//     }
//     htmlContent += "            </tr>\n";
//     htmlContent += "        </thead>\n";
//     htmlContent += "        <tbody>\n";
//     for (int row = 0; row < model->rowCount(); ++row) {
//         htmlContent += "            <tr>\n";
//         for (int col = 0; col < model->columnCount(); ++col) {
//             QString cellData = model->data(model->index(row, col)).toString();
//             // 先进行HTML转义，然后如果需要将文本中的换行符转换成<br>，可以这样做：
//             // QString escapedData = cellData.toHtmlEscaped();
//             // escapedData.replace("\n", "<br />\n"); // 如果希望保留文本中的换行并在HTML中也换行
//             // htmlContent += "                <td>" + escapedData + "</td>\n";
//             // 或者，如果使用了 white-space: pre-wrap; CSS属性，则直接转义即可，浏览器会处理换行：
//             htmlContent += "                <td>" + cellData.toHtmlEscaped() + "</td>\n";
//         }
//         htmlContent += "            </tr>\n";
//     }
//     htmlContent += "        </tbody>\n";
//     htmlContent += "    </table>\n";
//     htmlContent += "</body>\n";
//     htmlContent += "</html>\n";

//     // 确保路径已确定 (应该在构造函数中已完成，但以防万一)
//     if (m_persistentHtmlFilePath.isEmpty()) {
//         m_persistentHtmlFilePath = determinePersistentHtmlPath();
//     }

//     QFile file(m_persistentHtmlFilePath);
//     if (file.open(QIODevice::WriteOnly | QIODevice::Text | QIODevice::Truncate)) { // Truncate 会清空已存在的文件
//         QTextStream out(&file);
//         out.setEncoding(QStringConverter::Utf8); // 确保使用UTF-8编码
//         out << htmlContent;
//         file.close();
//         qDebug() << "HTML内容已成功写入：" << m_persistentHtmlFilePath;
//         QMessageBox::information(this, "成功", "HTML内容已生成并保存。\n您可以点击“在HTML中查看”按钮进行预览。\n文件位置: " + QDir::toNativeSeparators(m_persistentHtmlFilePath));
//         ui->btnCheckInHtml->setEnabled(true);
//     } else {
//         qWarning() << "无法打开持久化HTML文件进行写入：" << m_persistentHtmlFilePath << "错误：" << file.errorString();
//         QMessageBox::critical(this, "错误", "无法保存发布的HTML文件。\n错误: " + file.errorString());
//         ui->btnCheckInHtml->setEnabled(false); // 保存失败则禁用查看按钮
//     }
// }


// void JobRequirement::on_btnCheckInHtml_clicked()
// {
//     if (m_persistentHtmlFilePath.isEmpty()) { // 路径未初始化
//         m_persistentHtmlFilePath = determinePersistentHtmlPath(); // 尝试再次获取
//     }

//     if (!QFile::exists(m_persistentHtmlFilePath)) {
//         QMessageBox::information(this, "提示", "没有已发布的HTML内容可供预览。\n请先点击“发布”按钮生成并保存内容。");
//         // 确保按钮状态正确，以防文件被外部删除
//         ui->btnCheckInHtml->setEnabled(false);
//         return;
//     }

//     qDebug() << "尝试打开持久化HTML文件：" << m_persistentHtmlFilePath;
//     if (!QDesktopServices::openUrl(QUrl::fromLocalFile(m_persistentHtmlFilePath))) {
//         QString errorMsg = "无法在浏览器中打开HTML文件。\n文件路径: " + QDir::toNativeSeparators(m_persistentHtmlFilePath);
//         QMessageBox::critical(this, "错误", errorMsg);
//         qCritical() << errorMsg;
//     } else {
//         qDebug() << "QDesktopServices::openUrl 调用成功，文件路径：" << m_persistentHtmlFilePath;
//     }
// }

