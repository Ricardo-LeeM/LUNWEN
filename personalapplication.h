// #ifndef PERSONALAPPLICATION_H
// #define PERSONALAPPLICATION_H
// #include <QtSql>
// #include <QWidget>
// #include "dialogreimbursement.h"
// #include "dialogsalaryadjustment.h"
// #include "dialogloan.h"
// #include "dialogvacation.h"
// #include "dialogprojection.h"
// namespace Ui {
// class PersonalApplication;
// }

// class PersonalApplication : public QWidget
// {
//     Q_OBJECT

// public:
//     explicit PersonalApplication(QWidget *parent = nullptr);
//     ~PersonalApplication();

// private slots:
//     void on_actReimbursementApplication_triggered(); // 报销申请

//     void on_actLoanApplication_triggered(); // 借款申请

//     void on_actSalaryAdjustmentApplication_triggered(); // 薪资调整申请

//     void on_tbtnVacation_clicked();  // 休假申请



//     void updateFinancialNumSlot(); // 更新报销次数
//     void updateVacationNumslot(); // 更新请假申请次数
//     void updateProjectionNumslot(); // 更新立项申请次数

//     void on_tbtnProjection_clicked();

// private:
//     Ui::PersonalApplication *ui;
//     int NumFinancial = 0; // 报销申请次数初始化
//     int NumVacation = 0; // 休假申请次数初始化
//     int NumProjection = 0; // 立项申请次数初始化

//     DialogReimbursement *dialogReimbursement; // 报销申请界面
//     DialogSalaryAdjustment *dialogsalaryadjustment; // 薪资调整界面
//     DialogLoan *dialogloan; // 借款申请界面
//     DialogVacation *dialogvacation; // 休假申请界面
//     DialogProjection *dialogprojection; // 立项申请界面

// };

// #endif // PERSONALAPPLICATION_H

