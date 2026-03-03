#ifndef EMPLOYEEEDITDIALOG_H
#define EMPLOYEEEDITDIALOG_H

#include <QDialog>
#include <QSqlRecord>

enum EditMode { Add, Edit };

namespace Ui {
class EmployeeEditDialog;
}

class EmployeeEditDialog : public QDialog
{
    Q_OBJECT

public:
    explicit EmployeeEditDialog(const QSqlRecord &record, EditMode mode, QWidget *parent = nullptr);
    ~EmployeeEditDialog();

    QSqlRecord getRecordData() const;

protected:
    void accept() override;

private:
    void initComboBoxes();
    void fillData();

private:
    Ui::EmployeeEditDialog *ui;
    QSqlRecord m_record;
};

#endif // EMPLOYEEEDITDIALOG_H
