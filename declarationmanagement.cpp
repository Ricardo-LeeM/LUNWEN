#include "declarationmanagement.h"
#include "ui_declarationmanagement.h"
#include "databasemanager.h" // <-- 引入DatabaseManager

#include <QStandardItemModel>
#include <QDateTime>
#include <QMessageBox>
#include <QHBoxLayout>
#include <QPushButton>
#include <QInputDialog>
#include <QSqlRecord>

DeclarationManagement::DeclarationManagement(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::DeclarationManagement)
{
    ui->setupUi(this);

    // 设置所有表格的初始结构
    setupTableWidget(ui->tableWidget, true, false);        // 待审批
    setupTableWidget(ui->tableWidgetPass, false, true);     // 已通过
    setupTableWidget(ui->tableWidgetRejected, false, true); // 已驳回
    setupSearchTableWidget();

    // 连接Tab页切换信号，以便动态加载数据
    connect(ui->tabWidget, &QTabWidget::currentChanged, this, &DeclarationManagement::on_tabWidget_currentChanged);

    // 初始加载第一个Tab的内容
    on_tabWidget_currentChanged(ui->tabWidget->currentIndex());
}

DeclarationManagement::~DeclarationManagement()
{
    delete ui;
}

void DeclarationManagement::setupTableWidget(QTableWidget *table, bool includeActionColumn, bool includeApprovalDetails)
{
    if (!table) return;
    table->setRowCount(0);
    table->setColumnCount(0); // 清空所有列，防止重复添加

    QStringList headerText;
    headerText << "ID (隐藏)" << "申报人" << "类型" << "内容" << "提交时间";
    if (includeApprovalDetails) {
        headerText << "状态" << "处理时间" << "审批意见";
    }
    if (includeActionColumn) {
        headerText << "操作";
    }

    table->setColumnCount(headerText.size());
    table->setHorizontalHeaderLabels(headerText);
    table->setEditTriggers(QAbstractItemView::NoEditTriggers);
    table->setColumnHidden(0, true);
    table->horizontalHeader()->setStretchLastSection(!includeActionColumn);
    table->setSelectionBehavior(QAbstractItemView::SelectRows);
}

void DeclarationManagement::on_tabWidget_currentChanged(int index)
{
    if (index == 0) {
        loadRequestsByStatus("待审批", ui->tableWidget, true, false);
    } else if (index == 1) {
        loadRequestsByStatus("已通过", ui->tableWidgetPass, false, true);
    } else if (index == 2) {
        loadRequestsByStatus("已驳回", ui->tableWidgetRejected, false, true);
    }
}

void DeclarationManagement::on_btnUpdateTable_clicked()
{
    on_tabWidget_currentChanged(ui->tabWidget->currentIndex());
}

void DeclarationManagement::loadRequestsByStatus(const QString& status, QTableWidget* targetTable, bool includeActionColumn, bool includeApprovalDetails)
{
    if (!targetTable) return;
    targetTable->setRowCount(0);

    // *** 重构核心：调用DatabaseManager ***
    QList<QSqlRecord> records = DatabaseManager::instance()->getDeclarationsByStatus(status);

    for (const QSqlRecord& record : records) {
        int row = targetTable->rowCount();
        targetTable->insertRow(row);

        int declarationId = record.value("id").toInt();
        QString applicantName = record.value("applicant_name").toString();
        QString declarationType = record.value("declaration_type").toString();

        int currentCol = 0;
        targetTable->setItem(row, currentCol++, new QTableWidgetItem(QString::number(declarationId)));
        targetTable->setItem(row, currentCol++, new QTableWidgetItem(applicantName));
        targetTable->setItem(row, currentCol++, new QTableWidgetItem(declarationType));
        targetTable->setItem(row, currentCol++, new QTableWidgetItem(record.value("declaration_content").toString()));
        targetTable->setItem(row, currentCol++, new QTableWidgetItem(record.value("submission_timestamp").toDateTime().toString("yyyy-MM-dd hh:mm:ss")));

        if (includeApprovalDetails) {
            targetTable->setItem(row, currentCol++, new QTableWidgetItem(record.value("status").toString()));
            QDateTime approvalTime = record.value("approval_timestamp").toDateTime();
            targetTable->setItem(row, currentCol++, new QTableWidgetItem(approvalTime.isNull() ? "" : approvalTime.toString("yyyy-MM-dd hh:mm:ss")));
            targetTable->setItem(row, currentCol++, new QTableWidgetItem(record.value("approval_comment").toString()));
        }

        if (includeActionColumn) {
            int actionCol = targetTable->columnCount() - 1;
            QWidget *actionWidget = new QWidget();
            QHBoxLayout *layout = new QHBoxLayout(actionWidget);
            layout->setContentsMargins(5, 2, 5, 2);
            layout->setSpacing(5);
            QPushButton *btnApproveRow = new QPushButton("通过");
            QPushButton *btnRejectRow = new QPushButton("驳回");
            layout->addWidget(btnApproveRow);
            layout->addWidget(btnRejectRow);
            targetTable->setCellWidget(row, actionCol, actionWidget);

            connect(btnApproveRow, &QPushButton::clicked, this, [this, declarationId, applicantName, declarationType]() {
                if (QMessageBox::question(this, "确认通过", QString("确定要通过【%1】的【%2】吗？").arg(applicantName).arg(declarationType), QMessageBox::Yes | QMessageBox::No) == QMessageBox::Yes) {
                    // *** 重构核心：调用DatabaseManager ***
                    if (DatabaseManager::instance()->processDeclaration(declarationId, applicantName, true)) {
                        QMessageBox::information(this, "成功", "请求已批准。");
                        on_btnUpdateTable_clicked(); // 刷新表格
                    } else {
                        QMessageBox::critical(this, "失败", "操作失败，该申请可能已被处理或数据库出错。");
                    }
                }
            });

            connect(btnRejectRow, &QPushButton::clicked, this, [this, declarationId, applicantName, declarationType]() {
                bool ok;
                QString comment = QInputDialog::getText(this, "驳回意见", "请输入驳回理由：", QLineEdit::Normal, "", &ok);
                if (ok) {
                    // *** 重构核心：调用DatabaseManager ***
                    if (DatabaseManager::instance()->processDeclaration(declarationId, applicantName, false, comment)) {
                        QMessageBox::information(this, "成功", "请求已驳回。");
                        on_btnUpdateTable_clicked(); // 刷新表格
                    } else {
                        QMessageBox::critical(this, "失败", "操作失败，该申请可能已被处理或数据库出错。");
                    }
                }
            });
        }
    }
    targetTable->resizeColumnsToContents();
}


void DeclarationManagement::setupSearchTableWidget()
{
    setupTableWidget(ui->tableWidgetSearch, false, true);
}


void DeclarationManagement::on_btnSearch_clicked()
{
    QString applicantName = ui->lineEditSearch->text().trimmed();
    QString selectedType = ui->comboBoxDeclarationType->currentText();
    if(ui->comboBoxDeclarationType->currentIndex() == 0) {
        selectedType.clear(); // 如果是“无”，则设为空字符串
    }

    if (applicantName.isEmpty() && selectedType.isEmpty()) {
        QMessageBox::information(this, "提示", "请输入员工姓名或选择申报类型进行搜索。");
        ui->tableWidgetSearch->setRowCount(0);
        return;
    }

    // *** 重构核心：调用DatabaseManager ***
    QList<QSqlRecord> records = DatabaseManager::instance()->searchDeclarations(applicantName, selectedType);
    ui->tableWidgetSearch->setRowCount(0); // 清空旧结果

    if(records.isEmpty()){
        QMessageBox::information(this, "搜索结果", "未找到匹配的申报记录。");
        return;
    }

    // 填充搜索结果表格的逻辑与 loadRequestsByStatus 类似
    for(const QSqlRecord& record : records)
    {
        int row = ui->tableWidgetSearch->rowCount();
        ui->tableWidgetSearch->insertRow(row);
        int currentCol = 0;
        ui->tableWidgetSearch->setItem(row, currentCol++, new QTableWidgetItem(record.value("id").toString()));
        ui->tableWidgetSearch->setItem(row, currentCol++, new QTableWidgetItem(record.value("applicant_name").toString()));
        ui->tableWidgetSearch->setItem(row, currentCol++, new QTableWidgetItem(record.value("declaration_type").toString()));
        ui->tableWidgetSearch->setItem(row, currentCol++, new QTableWidgetItem(record.value("declaration_content").toString()));
        ui->tableWidgetSearch->setItem(row, currentCol++, new QTableWidgetItem(record.value("submission_timestamp").toDateTime().toString("yyyy-MM-dd hh:mm:ss")));
        ui->tableWidgetSearch->setItem(row, currentCol++, new QTableWidgetItem(record.value("status").toString()));
        QDateTime approvalTime = record.value("approval_timestamp").toDateTime();
        ui->tableWidgetSearch->setItem(row, currentCol++, new QTableWidgetItem(approvalTime.isNull() ? "" : approvalTime.toString("yyyy-MM-dd hh:mm:ss")));
        ui->tableWidgetSearch->setItem(row, currentCol++, new QTableWidgetItem(record.value("approval_comment").toString()));
    }
    ui->tableWidgetSearch->resizeColumnsToContents();
}

// 以下两个函数已被移入DatabaseManager，不再需要
// bool DeclarationManagement::updateApprovalStats(const QString& applicantName, bool isApproved) { return false; }
// bool DeclarationManagement::processDeclaration(int declarationId, const QString& applicantName, const QString& declarationType, bool approveAction, const QString& comment){ return false; }




