// #include "personalapplication.h"
// #include "dialogreimbursement.h"
// #include "ui_personalapplication.h"
// #include "dialogloan.h"
// #include "dialogsalaryadjustment.h"
// #include <QMenu>
// PersonalApplication::PersonalApplication(QWidget *parent)
//     : QWidget(parent)
//     , ui(new Ui::PersonalApplication)
//     , dialogReimbursement(nullptr) // 初始化 DialogReimbursement 指针
//     , dialogsalaryadjustment(nullptr)
//     , dialogloan(nullptr)
//     , dialogvacation(nullptr)
//     , dialogprojection(nullptr)
// {
//     ui->setupUi(this);
//     QMenu *menuSel = new QMenu(this);
//     menuSel->addAction(ui->actReimbursementApplication);
//     menuSel->addAction(ui->actLoanApplication);
//     menuSel->addAction(ui->actSalaryAdjustmentApplication);

//     ui->tbtnFinancial->setMenu(menuSel);



//     // QStringList headerText;
//     // headerText << ;
//     // ui->tabWidgetHistory->setColumnCount(headerText.size());
//     // for(int i = 0; i < ui->tabWidgetHistory->columnCount(); i ++)
//     // {
//     //     QTableWidgetItem *item = new QTableWidgetItem(headerText.at(i));
//     //     QFont font = item->font();
//     //     font.setBold(true);
//     //     font.setPointSize(11);

//     //     item->setFont(font);
//     //     item->setForeground(QBrush(Qt::black));
//     //     ui->tabWidgetHistory->setHorizontalHeaderItem(i,item);
//     // }

// }

// PersonalApplication::~PersonalApplication()
// {
//     delete ui;
//     if(dialogReimbursement) // 如果指针不为空，删除对象
//     {
//         delete dialogReimbursement;
//     }
// }

// void PersonalApplication::on_actReimbursementApplication_triggered() // 报销窗口
// {

//     if(dialogReimbursement == nullptr) // 如果指针为空，才创建新的窗口对象
//     {
//         dialogReimbursement = new DialogReimbursement(this);
//         connect(dialogReimbursement, &DialogReimbursement::updateFinancialNumSignal, this, &PersonalApplication::updateFinancialNumSlot); // 连接信号与槽
//     }
//     dialogReimbursement->exec(); // 以模态显示该窗口
//     dialogReimbursement->setAttribute(Qt::WA_DeleteOnClose);
// }


// void PersonalApplication::on_actLoanApplication_triggered() // 借款申请窗口
// {
//     if(dialogloan == nullptr)
//     {
//         dialogloan = new DialogLoan(this);
//         connect(dialogloan, &DialogLoan::updateFinancialNumSignal, this, &PersonalApplication::updateFinancialNumSlot); // 连接信号与槽
//     }
//     dialogloan->exec();
//     dialogloan->setAttribute(Qt::WA_DeleteOnClose);
// }


// void PersonalApplication::on_actSalaryAdjustmentApplication_triggered() // 薪资调整窗口
// {

//     if(dialogsalaryadjustment == nullptr)
//     {
//         dialogsalaryadjustment = new DialogSalaryAdjustment(this);
//         connect(dialogsalaryadjustment, &DialogSalaryAdjustment::updateFinancialNumSignal, this, &PersonalApplication::updateFinancialNumSlot);
//     }
//     dialogsalaryadjustment->exec(); // 以模态显示该窗口
//     dialogsalaryadjustment->setAttribute(Qt::WA_DeleteOnClose);
// }



// void PersonalApplication::updateFinancialNumSlot() // 更新 labelFinancialNum 的槽函数
// {
//     NumFinancial++;
//     ui->labelFinancialNum->setText("我的申报次数：" + QString::number(NumFinancial)); // 更新 label 的文本
// }
// void PersonalApplication::updateVacationNumslot()
// {
//     NumVacation++;
//     ui->labelVacationNum->setText("我的申报次数：" + QString::number(NumVacation)); // 更新 label 的文本
// }
// void PersonalApplication::updateProjectionNumslot()
// {
//     NumProjection++;
//     ui->labelProjectionNum->setText("我的申报次数：" + QString::number(NumProjection)); // 更新 label 的文本
// }





// void PersonalApplication::on_tbtnVacation_clicked() // 休假申请窗口
// {
//     if(dialogvacation == nullptr)
//     {
//         dialogvacation = new DialogVacation(this);
//         connect(dialogvacation, &DialogVacation::updateVacationlNumSignal, this, &PersonalApplication::updateVacationNumslot);

//     }
//     dialogvacation->exec();
//     dialogvacation->setAttribute(Qt::WA_DeleteOnClose);
// }


// void PersonalApplication::on_tbtnProjection_clicked() // 休假申请窗口
// {
//     if(dialogprojection == nullptr)
//     {
//         dialogprojection = new DialogProjection(this);
//         connect(dialogprojection, &DialogProjection::updateProjectionNumSignal, this, &PersonalApplication::updateProjectionNumslot);
//     }
//     dialogprojection->exec();
//     dialogprojection->setAttribute(Qt::WA_DeleteOnClose);
// }

