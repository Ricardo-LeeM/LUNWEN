#ifndef DECLARATIONMANAGEMENT_H
#define DECLARATIONMANAGEMENT_H

#include <QWidget>
#include <QStandardItemModel>
#include <QTableWidget>
namespace Ui {
class DeclarationManagement;
}

class DeclarationManagement : public QWidget
{
    Q_OBJECT

public:
    explicit DeclarationManagement(QWidget *parent = nullptr);
    ~DeclarationManagement();



private slots:

    void on_btnUpdateTable_clicked();
    void on_tabWidget_currentChanged(int index); // Tab页切换

    // 行内按钮的槽函数不再直接定义，由Lambda处理，
    // 但如果需要弹出输入审批意见的对话框，可以考虑辅助函数

    void on_btnSearch_clicked();




private:
    Ui::DeclarationManagement *ui;

    // 审批申报等函数
    void setupTableWidget(QTableWidget *table, bool includeActionColumn, bool includeApprovalDetails);// 数据加载
    void loadRequestsByStatus(const QString& status, QTableWidget* targetTable, bool includeActionColumn, bool includeApprovalDetails);
    bool updateApprovalStats(const QString& applicantName, bool isApproved); // 更新 personal_application 表中对应员工的通过/驳回次数统计


    // 搜索功能函数
    void setupSearchTableWidget(); // 为搜索结果表格设置表头等
    void loadSearchResults(const QString& applicantName, const QString& declarationType); //加载搜索结果
    bool processDeclaration(int declarationId, const QString& applicantName, const QString& declarationType, bool approveAction, const QString& comment = "");// 核心处理逻辑

};

#endif // DECLARATIONMANAGEMENT_H


