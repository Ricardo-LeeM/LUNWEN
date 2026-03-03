#ifndef DATABASEMANAGER_H
#define DATABASEMANAGER_H

#include <QObject>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QSqlRecord>
#include <QVariant>
#include <QList>
#include <QPixmap>
#include <QDate>

// 前向声明，可以减少不必要的头文件包含，加快编译速度
class UserSession;

class DatabaseManager : public QObject
{
    Q_OBJECT
public:
    // 获取单例实例的静态方法
    static DatabaseManager* instance();

    // 数据库连接与关闭
    bool openDatabase(const QString& host, int port, const QString& dbName, const QString& user, const QString& pass);
    void closeDatabase();

    // --- 用户认证与员工管理 ---
    bool validateUser(const QString& username, const QString& password, const QString& role, QSqlRecord& userInfo);
    bool fetchEmployeeDetails(const QString& username, UserSession* session);
    bool getEmployeeNameById(int employeeId, QString& name);
    QList<QSqlRecord> getAllActiveEmployees(int currentUserId);
    QList<QString> getAllDepartments();
    bool addNewEmployee(const QSqlRecord& record);
    bool checkEmployeeIdExists(int employeeId);
    bool updateEmployee(const QSqlRecord& record);
    bool setEmployeeStatusToResigned(int employeeId);
    QList<QSqlRecord> getFilteredEmployees(const QString& name, const QString& department);
    QSqlRecord getEmployeeRecordById(int employeeId); // 用于编辑员工信息

    // --- 申报与审批 ---
    bool getApplicationCounts(int employeeId, int& financialCount, int& vacationCount, int& projectionCount);
    bool addDeclaration(const QString& applicantName, const QString& type, const QString& content);
    bool processDeclaration(int declarationId, const QString& applicantName, bool isApproved, const QString& comment = "");
    QList<QSqlRecord> getDeclarationsByStatus(const QString& status);
    QList<QSqlRecord> searchDeclarations(const QString& applicantName, const QString& declarationType);

    // --- 薪酬管理 ---
    QList<QSqlRecord> getSalaryHistory(int employeeId);
    bool getEmployeeSalaryInfo(int employeeId, QSqlRecord& info);

    // --- 任务管理 ---
    QList<QSqlRecord> getTasksForEmployee(int employeeId);
    QList<QSqlRecord> getFilteredTasks(int assigneeId, const QString& taskType, const QString& status);
    bool updateTaskStatus(int taskId, const QString& newStatus);
    bool getAssignerNameById(int assignerId, QString& name);
    bool addTask(QSqlRecord& record);
    bool updateTask(const QSqlRecord& record);
    bool setTaskStatusToCancelled(int taskId);
    QSqlRecord getTaskRecordById(int taskId);

    // --- 招聘管理 ---
    QList<QSqlRecord> getJobRequirements();
    bool updateJobRequirement(int dbId, int column, const QVariant& value);
    bool addJobRequirement(int& newId);
    bool deleteJobRequirements(const QList<int>& ids);
    bool addInterviewee(const QSqlRecord& record, int& newId);
    bool updateInterviewee(const QSqlRecord& record);
    bool deleteInterviewee(int id);


private:
    // 私有构造函数和析构函数，确保是单例
    explicit DatabaseManager(QObject *parent = nullptr);
    ~DatabaseManager();

    // 禁止拷贝和赋值，防止意外创建多个实例
    DatabaseManager(const DatabaseManager&) = delete;
    DatabaseManager& operator=(const DatabaseManager&) = delete;

    // 内部辅助函数
    bool updateApplicationCount(const QString &applicantName, const QString &declarationType);
    bool updateApprovalStats(const QString& applicantName, bool isApproved);

    // 单例实例指针和数据库连接对象
    static DatabaseManager* m_instance;
    QSqlDatabase m_db;
};

#endif // DATABASEMANAGER_H
