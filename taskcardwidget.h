
#ifndef TASKCARDWIDGET_H
#define TASKCARDWIDGET_H

#include <QWidget>
#include <QSqlRecord> // 包含 QSqlRecord 以便在类中使用

namespace Ui {
class TaskCardWidget;
}

/**
 * @brief 任务卡片控件，用于在看板上显示单个任务信息。
 *
 * 该控件负责展示任务的关键信息，并提供一个按钮来触发状态变更。
 * 当需要变更状态时，它会发出 requestMove 信号。
 */
class TaskCardWidget : public QWidget
{
    Q_OBJECT

public:
    explicit TaskCardWidget(QWidget *parent = nullptr);
    ~TaskCardWidget();

    // --- 公共接口函数 ---
    /**
     * @brief 使用从数据库获取的一行记录来填充卡片内容。
     * @param taskRecord 包含任务所有信息的QSqlRecord对象。
     */
    void setTaskData(const QSqlRecord &taskRecord);

    /**
     * @brief 更新卡片的状态并刷新UI（主要是按钮）。
     * @param newStatus 新的状态字符串，如 "进行中" 或 "已完成"。
     */
    void updateStatus(const QString &newStatus);

    // --- Getter函数 ---
    int getTaskId() const;
    QSqlRecord getRecord() const;


signals:
    /**
     * @brief 当用户点击移动按钮时发出的信号。
     * @param card 指向当前卡片实例的指针。
     * @param newStatus 卡片请求移动到的新状态。
     */
    void requestMove(TaskCardWidget* card, const QString &newStatus);
    void viewDetailsRequested(QSqlRecord record); // 新增信号，直接传递记录
protected:
    void mouseDoubleClickEvent(QMouseEvent *event) override; // 重写鼠标双击事件


private slots:
    /**
     * @brief 响应“移动按钮”点击事件的槽函数（由Qt自动连接）。
     */
    void on_moveButton_clicked();


private:
    Ui::TaskCardWidget *ui; // 指向UI设计器创建的控件

    // --- 内部数据成员 ---
    int m_taskId = -1;              // 存储当前任务的ID
    QString m_currentStatus;        // 存储当前任务的状态字符串
    QSqlRecord m_record;            // 存储与此卡片关联的完整数据库记录
};

#endif // TASKCARDWIDGET_H

