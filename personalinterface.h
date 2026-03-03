#ifndef PERSONALINTERFACE_H
#define PERSONALINTERFACE_H

/***************************************************************************************************
** 头文件依赖 (Header Dependencies)                                  **
***************************************************************************************************/
#include <QMainWindow>
#include "taskcardwidget.h"
#include <QStandardItemModel> // <-- **重要修改：将 QSqlTableModel 替换为 QStandardItemModel**

// 前向声明
class QLayout;
class QSqlRecord;

// UI类的标准前向声明
QT_BEGIN_NAMESPACE
namespace Ui { class PersonalInterface; }
QT_END_NAMESPACE

/***************************************************************************************************
** PersonalInterface 类定义                                    **
***************************************************************************************************/
class PersonalInterface : public QMainWindow
{
    Q_OBJECT

public:
    explicit PersonalInterface(QWidget *parent = nullptr);
    ~PersonalInterface();

private slots:
    // 保留UI自动生成的槽函数和我们需要连接的槽函数
    void on_tbtnFinancial_clicked();
    void on_tbtnVacation_clicked();
    void on_tbtnProjection_clicked();
    void moveTaskCard(TaskCardWidget* card, const QString &newStatus);
    void showTaskDetails(const QSqlRecord &record);
    void on_btnFromExcel_clicked();
    void on_btnToExcel_clicked();
    void on_shuaxinButton_clicked();

private:
    Ui::PersonalInterface *ui; // 指向UI界面的指针

    // --- 数据模型 ---
    // *** 重要修改：将数据模型类型从 QSqlTableModel* 改为 QStandardItemModel* ***
    QStandardItemModel *salaryHistoryModel;

    // --- 私有函数 ---

    // **新增的函数声明**
    void updateAllCounts();
    void setupTaskboard();

    // **保留和修改后的函数声明**
    void loadCurrentUserData();
    void initializeSalaryTable();
    void loadTaskCards();
    void clearLayout(QLayout* layout);

    // **不再需要的旧函数声明（已移除）**
    // void updateFinancialNumSlot();
    // void updateVacationNumslot();
    // void updateProjectionNumslot();
    // bool addDeclarationToDb(const QString &applicant, const QString &type, const QString &content);
    // bool updatePersonalApplicationCount(const QString &declarationType);
};

#endif // PERSONALINTERFACE_H
