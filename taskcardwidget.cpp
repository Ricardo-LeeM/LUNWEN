#include "taskcardwidget.h"
#include "ui_taskcardwidget.h"

#include <QSqlQuery>
#include <QSqlRecord>
#include <QDate>
#include <QDebug>
#include "databasemanager.h"


/**
 * @brief TaskCardWidget 构造函数
 * @param parent 父控件指针
 */
TaskCardWidget::TaskCardWidget(QWidget *parent)
    : QWidget(parent)
    , ui(new Ui::TaskCardWidget)
{
    // 初始化UI，这是Qt自动连接信号和槽（如on_moveButton_clicked）的前提
    ui->setupUi(this);
}

/**
 * @brief TaskCardWidget 析构函数
 */
TaskCardWidget::~TaskCardWidget()
{
    delete ui;
}

/**
 * @brief 使用数据库记录来填充卡片的UI和内部数据
 * @param taskRecord 包含单条任务信息的一行QSqlRecord
 */
void TaskCardWidget::setTaskData(const QSqlRecord &taskRecord)
{
    m_record = taskRecord;
    m_taskId = m_record.value("TaskID").toInt();
    ui->taskNameLabel->setText(m_record.value("TaskName").toString());
    ui->dueDateLabel->setText("截止日期: " + m_record.value("DueDate").toDate().toString("yyyy-MM-dd"));

    QString taskType = m_record.value("TaskType").toString();
    ui->taskTypeLabel->setText(taskType);
    if (taskType == "出差任务") {
        ui->taskTypeLabel->setStyleSheet("background-color: #a5d6a7; color: #1b5e20; padding: 2px 5px; border-radius: 4px; font-weight: bold;");
    } else {
        ui->taskTypeLabel->setStyleSheet("background-color: #90caf9; color: #0d47a1; padding: 2px 5px; border-radius: 4px; font-weight: bold;");
    }

    // *** 重构核心：调用DatabaseManager ***
    int assignerId = m_record.value("AssignerID").toInt();
    QString assignerName;
    if(DatabaseManager::instance()->getAssignerNameById(assignerId, assignerName)) {
        ui->assignerLabel->setText("分派人: " + assignerName);
    } else {
        ui->assignerLabel->setText("分派人: 未知");
    }

    updateStatus(m_record.value("Status").toString());
}


/**
 * @brief 更新卡片的状态，并相应地调整UI（主要是移动按钮）
 * @param newStatus 新的状态字符串
 */


void TaskCardWidget::updateStatus(const QString &newStatus)
{
    m_currentStatus = newStatus;
    m_record.setValue("Status", newStatus);

    if (m_currentStatus == "未开始") {
        ui->moveButton->setText("开始任务 →");
        ui->moveButton->setVisible(true);
    } else if (m_currentStatus == "待审批") {
        ui->moveButton->setText("等待审批...");
        ui->moveButton->setEnabled(false); // 禁用按钮，直到管理员审批
        ui->moveButton->setVisible(true);
    } else if (m_currentStatus == "进行中") {
        ui->moveButton->setText("完成任务 ✔");
        ui->moveButton->setEnabled(true);
        ui->moveButton->setVisible(true);
    } else {
        ui->moveButton->setVisible(false);
    }
}

/**
 * @brief 获取当前任务的ID
 * @return 任务ID
 */
int TaskCardWidget::getTaskId() const
{
    return m_taskId;
}

/**
 * @brief 获取当前卡片存储的完整数据库记录
 * @return QSqlRecord对象
 */
QSqlRecord TaskCardWidget::getRecord() const
{
    return m_record;
}


/**
 * @brief 响应“移动按钮”的点击事件的槽函数
 *
 * 该函数根据卡片的当前状态，决定下一个状态是什么，
 * 然后发出requestMove信号，通知父窗口(PersonalInterface)来处理实际的移动逻辑。
 */
void TaskCardWidget::on_moveButton_clicked()
{
    QString newStatus;
    // 根据当前状态确定目标状态
    if (m_currentStatus == "未开始") {
        newStatus = "进行中";
    } else if (m_currentStatus == "进行中") {
        newStatus = "已完成";
    } else {
        // 如果当前状态不是“未开始”或“进行中”，则不执行任何操作
        return;
    }

    // 发出信号，将自身指针和目标新状态传递出去
    emit requestMove(this, newStatus);
}

/**
 * @brief 响应鼠标双击事件，发出查看详情的请求
 * @param event 鼠标事件指针
 */
void TaskCardWidget::mouseDoubleClickEvent(QMouseEvent *event)
{
    // 发出“请求查看详情”的信号，并把自己存储的完整任务记录传递出去
    emit viewDetailsRequested(m_record);
    // 调用基类实现
    QWidget::mouseDoubleClickEvent(event);
}
