#ifndef INTERVIEWEEDIALOG_H
#define INTERVIEWEEDIALOG_H

#include <QDialog>
#include <QSqlRecord> // 包含头文件，用于传递数据

namespace Ui {
class IntervieweeDialog;
}

class IntervieweeDialog : public QDialog
{
    Q_OBJECT

public:
    explicit IntervieweeDialog(QWidget *parent = nullptr);
    ~IntervieweeDialog();

    // 公有函数：用于设置对话框的数据（修改模式）
    void setRecord(const QSqlRecord &record);

    // 公有函数：用于获取对话框中填写好的数据
    QSqlRecord getRecord(); // 注意这里返回的是一个值，而不是引用

private slots:
    void on_okButton_clicked();
    void on_cancelButton_clicked();

    void on_enableTimeCheckBox_toggled(bool checked);

private:
    Ui::IntervieweeDialog *ui;
    QSqlRecord m_record; // 用一个成员变量来暂存记录

    void setupUI(); // 辅助函数，用于初始化控件
};

#endif // INTERVIEWEEDIALOG_H
