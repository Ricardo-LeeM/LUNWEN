#include "databasemanager.h"
#include "usersession.h"
#include <QDebug>

// 初始化静态单例指针
DatabaseManager* DatabaseManager::m_instance = nullptr;

// 获取单例实例
DatabaseManager* DatabaseManager::instance()
{
    if (!m_instance) {
        m_instance = new DatabaseManager();
    }
    return m_instance;
}

// 私有构造函数
DatabaseManager::DatabaseManager(QObject *parent) : QObject(parent)
{
    if (QSqlDatabase::contains("qt_sql_default_connection")) {
        m_db = QSqlDatabase::database("qt_sql_default_connection");
    } else {
        m_db = QSqlDatabase::addDatabase("QMYSQL");
    }
}

// 析构函数
DatabaseManager::~DatabaseManager()
{
    if (m_db.isOpen()) {
        m_db.close();
    }
}

// 打开数据库连接
bool DatabaseManager::openDatabase(const QString& host, int port, const QString& dbName, const QString& user, const QString& pass)
{
    m_db.setHostName(host);
    m_db.setPort(port);
    m_db.setDatabaseName(dbName);
    m_db.setUserName(user);
    m_db.setPassword(pass);
    if (!m_db.open()) {
        qCritical() << "数据库连接失败:" << m_db.lastError().text();
        return false;
    }
    qDebug() << "数据库连接成功!";
    return true;
}

// 关闭数据库连接
void DatabaseManager::closeDatabase()
{
    m_db.close();
}


/***************************************************************************************************
** 用户认证与员工管理模块
***************************************************************************************************/

bool DatabaseManager::validateUser(const QString& username, const QString& password, const QString& role, QSqlRecord& userInfo)
{
    QSqlQuery query(m_db);
    query.prepare("SELECT password FROM users WHERE username = :username");
    query.bindValue(":username", username);
    if (!query.exec() || !query.next()) return false;

    if (query.value(0).toString() == password) {
        query.prepare("SELECT * FROM BasicInfo WHERE Name = :loginName AND EmployeeType = :role");
        query.bindValue(":loginName", username);
        query.bindValue(":role", role);
        if (query.exec() && query.next()) {
            userInfo = query.record();
            return true;
        }
    }
    return false;
}

bool DatabaseManager::fetchEmployeeDetails(const QString& username, UserSession* session)
{
    QSqlQuery query(m_db);
    query.prepare("SELECT * FROM BasicInfo WHERE Name = :username");
    query.bindValue(":username", username);
    if (!query.exec() || !query.next()) {
        qWarning() << "登录后获取员工详细信息失败: " << query.lastError().text();
        return false;
    }

    QByteArray imageBytes = query.value("Pic").toByteArray();
    QPixmap pixmap;
    if (!imageBytes.isNull() && !imageBytes.isEmpty()) {
        pixmap.loadFromData(imageBytes);
    }

    session->setCurrentUser(
        query.value("EmployeeID").toInt(),
        query.value("Name").toString(),
        query.value("Department").toString(),
        query.value("EmployeeStatus").toString(),
        query.value("IsPartyMember").toBool(),
        query.value("WorkNature").toString(),
        query.value("Position").toString(),
        query.value("JoinDate").toDate(),
        pixmap,
        query.value("EmployeeType").toString()
        );
    return true;
}

bool DatabaseManager::getEmployeeNameById(int employeeId, QString& name)
{
    QSqlQuery query(m_db);
    query.prepare("SELECT Name FROM basicinfo WHERE EmployeeID = :id");
    query.bindValue(":id", employeeId);
    if (query.exec() && query.next()) {
        name = query.value(0).toString();
        return true;
    }
    name = "未知员工";
    return false;
}

QList<QSqlRecord> DatabaseManager::getAllActiveEmployees(int currentUserId)
{
    QList<QSqlRecord> records;
    QSqlQuery query(m_db);
    if (currentUserId > 0) {
        query.prepare("SELECT EmployeeID, Name FROM basicinfo WHERE EmployeeStatus = '在职' AND EmployeeID != :currentUserId");
        query.bindValue(":currentUserId", currentUserId);
    } else {
        query.prepare("SELECT EmployeeID, Name FROM basicinfo WHERE EmployeeStatus = '在职'");
    }

    if (query.exec()) {
        while(query.next()) {
            records.append(query.record());
        }
    } else {
        qCritical() << "获取在职员工列表失败:" << query.lastError().text();
    }
    return records;
}

QList<QString> DatabaseManager::getAllDepartments()
{
    QList<QString> departments;
    QSqlQuery query("SELECT DISTINCT Department FROM basicinfo WHERE Department IS NOT NULL AND Department != '高管'", m_db);
    if(query.exec()){
        while(query.next()) {
            departments.append(query.value(0).toString());
        }
    } else {
        qCritical() << "获取部门列表失败: " << query.lastError().text();
    }
    return departments;
}

bool DatabaseManager::addNewEmployee(const QSqlRecord& record)
{
    if (!m_db.transaction()) {
        qCritical() << "数据库事务错误：无法启动 (addNewEmployee): " << m_db.lastError().text();
        return false;
    }

    // 插入 basicinfo 表
    QSqlQuery basicInfoQuery(m_db);
    basicInfoQuery.prepare("INSERT INTO basicinfo (EmployeeID, Name, Department, EmployeeStatus, IsPartyMember, WorkNature, Position, JoinDate, EmployeeType, Pic) "
                           "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
    basicInfoQuery.addBindValue(record.value("EmployeeID"));
    basicInfoQuery.addBindValue(record.value("Name"));
    basicInfoQuery.addBindValue(record.value("Department"));
    basicInfoQuery.addBindValue(record.value("EmployeeStatus"));
    basicInfoQuery.addBindValue(record.value("IsPartyMember"));
    basicInfoQuery.addBindValue(record.value("WorkNature"));
    basicInfoQuery.addBindValue(record.value("Position"));
    basicInfoQuery.addBindValue(record.value("JoinDate"));
    basicInfoQuery.addBindValue(record.value("EmployeeType"));
    basicInfoQuery.addBindValue(QByteArray()); // 为 Pic 提供一个默认的空值

    if (!basicInfoQuery.exec()) {
        qCritical() << "数据库执行错误 (basicinfo insert): " << basicInfoQuery.lastError().text();
        m_db.rollback();
        return false;
    }

    // 插入 users 表
    QSqlQuery userQuery(m_db);
    userQuery.prepare("INSERT INTO users (EmployeeID, username, password) VALUES (?, ?, ?)");
    userQuery.addBindValue(record.value("EmployeeID").toInt());
    userQuery.addBindValue(record.value("Name").toString());
    userQuery.addBindValue(record.value("EmployeeType").toString() == "管理员" ? "gggggg" : "yyyyyy");

    if (!userQuery.exec()) {
        qCritical() << "数据库执行错误 (users insert): " << userQuery.lastError().text();
        m_db.rollback();
        return false;
    }

    // 提交事务
    if (m_db.commit()) {
        return true;
    } else {
        qCritical() << "数据库提交错误 (addNewEmployee): " << m_db.lastError().text();
        m_db.rollback();
        return false;
    }
}

bool DatabaseManager::checkEmployeeIdExists(int employeeId)
{
    QSqlQuery query(m_db);
    query.prepare("SELECT COUNT(*) FROM basicinfo WHERE EmployeeID = ?");
    query.addBindValue(employeeId);
    if (query.exec() && query.next()) {
        return query.value(0).toInt() > 0;
    }
    return false;
}

bool DatabaseManager::updateEmployee(const QSqlRecord& record)
{
    if (!m_db.transaction()) {
        qCritical() << "无法启动数据库事务 (updateEmployee): " << m_db.lastError().text();
        return false;
    }
    QSqlQuery query(m_db);
    query.prepare("UPDATE basicinfo SET Name=:Name, Department=:Department, Position=:Position, "
                  "EmployeeStatus=:EmployeeStatus, WorkNature=:WorkNature, EmployeeType=:EmployeeType, "
                  "JoinDate=:JoinDate, IsPartyMember=:IsPartyMember WHERE EmployeeID=:EmployeeID");

    for (int i = 0; i < record.count(); ++i) {
        query.bindValue(":" + record.fieldName(i), record.value(i));
    }

    if (!query.exec()) {
        qCritical() << "更新员工信息失败: " << query.lastError().text();
        m_db.rollback();
        return false;
    }

    return m_db.commit();
}

bool DatabaseManager::setEmployeeStatusToResigned(int employeeId)
{
    QSqlQuery query(m_db);
    query.prepare("UPDATE basicinfo SET EmployeeStatus = '离职' WHERE EmployeeID = ?");
    query.addBindValue(employeeId);
    if (!query.exec()) {
        qCritical() << "设置员工离职状态失败: " << query.lastError().text();
        return false;
    }
    return query.numRowsAffected() > 0;
}

QList<QSqlRecord> DatabaseManager::getFilteredEmployees(const QString& name, const QString& department)
{
    QList<QSqlRecord> records;
    QString sql = "SELECT * FROM basicinfo WHERE 1=1";
    if (!name.isEmpty()) {
        sql += " AND Name LIKE :name";
    }
    if (!department.isEmpty() && department != "所有部门") {
        sql += " AND Department = :department";
    }

    QSqlQuery query(m_db);
    query.prepare(sql);
    if (!name.isEmpty()) {
        query.bindValue(":name", "%" + name + "%");
    }
    if (!department.isEmpty() && department != "所有部门") {
        query.bindValue(":department", department);
    }

    if (query.exec()) {
        while (query.next()) {
            records.append(query.record());
        }
    } else {
        qCritical() << "筛选员工失败: " << query.lastError().text();
    }
    return records;
}

QSqlRecord DatabaseManager::getEmployeeRecordById(int employeeId)
{
    QSqlQuery query(m_db);
    query.prepare("SELECT * FROM basicinfo WHERE EmployeeID = ?");
    query.addBindValue(employeeId);
    if (query.exec() && query.next()) {
        return query.record();
    }
    return QSqlRecord();
}

/***************************************************************************************************
** 申报与审批模块
***************************************************************************************************/

bool DatabaseManager::getApplicationCounts(int employeeId, int& financialCount, int& vacationCount, int& projectionCount)
{
    financialCount = 0;
    vacationCount = 0;
    projectionCount = 0;
    QSqlQuery query(m_db);
    query.prepare("SELECT FinancialNum, VacationNum, ProjectionNum FROM personal_application WHERE EmployeeID = :id");
    query.bindValue(":id", employeeId);
    if (query.exec() && query.next()) {
        financialCount = query.value("FinancialNum").toInt();
        vacationCount = query.value("VacationNum").toInt();
        projectionCount = query.value("ProjectionNum").toInt();
    }
    return true;
}

bool DatabaseManager::addDeclaration(const QString& applicantName, const QString& type, const QString& content)
{
    if (!m_db.transaction()) {
        qCritical() << "无法启动数据库事务 (addDeclaration):" << m_db.lastError().text();
        return false;
    }

    QSqlQuery query(m_db);
    query.prepare("INSERT INTO declarations (applicant_name, declaration_type, declaration_content, status, submission_timestamp) "
                  "VALUES (:applicant, :type, :content, '待审批', CURRENT_TIMESTAMP)");
    query.bindValue(":applicant", applicantName);
    query.bindValue(":type", type);
    query.bindValue(":content", content);
    if (!query.exec()) {
        qCritical() << "添加申报记录失败:" << query.lastError().text();
        m_db.rollback();
        return false;
    }

    if (!updateApplicationCount(applicantName, type)) {
        m_db.rollback();
        return false;
    }

    return m_db.commit();
}

bool DatabaseManager::processDeclaration(int declarationId, const QString& applicantName, bool isApproved, const QString& comment)
{
    if (!m_db.transaction()) {
        qCritical() << "无法启动数据库事务 (processDeclaration):" << m_db.lastError().text();
        return false;
    }

    bool success = true;
    QSqlQuery queryUpdate(m_db);
    QString newStatus = isApproved ? "已通过" : "已驳回";
    queryUpdate.prepare("UPDATE declarations SET status = :status, approval_timestamp = CURRENT_TIMESTAMP, approval_comment = :comment "
                        "WHERE id = :id AND status = '待审批'");
    queryUpdate.bindValue(":status", newStatus);
    queryUpdate.bindValue(":comment", comment);
    queryUpdate.bindValue(":id", declarationId);
    if (!queryUpdate.exec() || queryUpdate.numRowsAffected() == 0) {
        success = false;
        qWarning() << "更新申报状态失败或该申报已被处理。";
    }

    if (success) {
        if (!updateApprovalStats(applicantName, isApproved)) {
            success = false;
        }
    }

    if (success) {
        m_db.commit();
    } else {
        m_db.rollback();
        qWarning() << "处理申报事务失败，操作已回滚。";
    }
    return success;
}

QList<QSqlRecord> DatabaseManager::getDeclarationsByStatus(const QString& status)
{
    QList<QSqlRecord> records;
    QSqlQuery query(m_db);
    QString sql = "SELECT id, applicant_name, declaration_type, declaration_content, submission_timestamp, status, approval_timestamp, approval_comment "
                  "FROM declarations WHERE status = :status ORDER BY submission_timestamp DESC";
    query.prepare(sql);
    query.bindValue(":status", status);

    if (query.exec()) {
        while (query.next()) {
            records.append(query.record());
        }
    } else {
        qCritical() << QString("获取'%1'列表失败: %2").arg(status).arg(query.lastError().text());
    }
    return records;
}

QList<QSqlRecord> DatabaseManager::searchDeclarations(const QString& applicantName, const QString& declarationType)
{
    QList<QSqlRecord> records;
    QSqlQuery query(m_db);
    QString sql = "SELECT * FROM declarations WHERE 1=1";
    if (!applicantName.isEmpty()) {
        sql += " AND applicant_name LIKE :applicantName";
    }
    if (!declarationType.isEmpty()) {
        sql += " AND declaration_type = :declarationType";
    }
    sql += " ORDER BY submission_timestamp DESC";

    query.prepare(sql);

    if (!applicantName.isEmpty()) {
        query.bindValue(":applicantName", "%" + applicantName + "%");
    }
    if (!declarationType.isEmpty()) {
        query.bindValue(":declarationType", declarationType);
    }

    if (query.exec()) {
        while (query.next()) {
            records.append(query.record());
        }
    } else {
        qCritical() << "搜索历史申报失败: " << query.lastError().text();
    }
    return records;
}

/***************************************************************************************************
** 薪酬管理模块
***************************************************************************************************/

QList<QSqlRecord> DatabaseManager::getSalaryHistory(int employeeId)
{
    QList<QSqlRecord> records;
    QSqlQuery query(m_db);
    query.prepare("SELECT * FROM PayrollRecords WHERE employee_id = :id ORDER BY payroll_period DESC");
    query.bindValue(":id", employeeId);
    if(query.exec()){
        while(query.next()){
            records.append(query.record());
        }
    } else {
        qCritical() << "获取薪资历史失败:" << query.lastError().text();
    }
    return records;
}

bool DatabaseManager::getEmployeeSalaryInfo(int employeeId, QSqlRecord& info)
{
    QSqlQuery query(m_db);
    query.prepare("SELECT name, department, base_salary FROM EmployeesSalary WHERE employee_id = :id");
    query.bindValue(":id", employeeId);
    if (query.exec() && query.next()) {
        info = query.record();
        return true;
    }
    return false;
}

/***************************************************************************************************
** 任务管理模块
***************************************************************************************************/

QList<QSqlRecord> DatabaseManager::getTasksForEmployee(int employeeId)
{
    QList<QSqlRecord> records;
    QSqlQuery query(m_db);
    query.prepare("SELECT * FROM tasks WHERE AssigneeID = :employeeId ORDER BY DueDate ASC");
    query.bindValue(":employeeId", employeeId);
    if (query.exec()) {
        while (query.next()) {
            records.append(query.record());
        }
    } else {
        qCritical() << "加载任务失败:" << query.lastError().text();
    }
    return records;
}

QList<QSqlRecord> DatabaseManager::getFilteredTasks(int assigneeId, const QString& taskType, const QString& status)
{
    QList<QSqlRecord> records;
    QString sql = "SELECT * FROM Tasks WHERE 1=1";
    if (assigneeId > 0) {
        sql += " AND AssigneeID = :assigneeId";
    }
    if (!taskType.isEmpty() && taskType != "所有类型") {
        sql += " AND TaskType = :taskType";
    }
    if (!status.isEmpty() && status != "所有状态") {
        sql += " AND Status = :status";
    }
    sql += " ORDER BY DueDate DESC";

    QSqlQuery query(m_db);
    query.prepare(sql);

    if (assigneeId > 0) query.bindValue(":assigneeId", assigneeId);
    if (!taskType.isEmpty() && taskType != "所有类型") query.bindValue(":taskType", taskType);
    if (!status.isEmpty() && status != "所有状态") query.bindValue(":status", status);

    if (query.exec()) {
        while (query.next()) {
            records.append(query.record());
        }
    } else {
        qCritical() << "筛选任务失败: " << query.lastError().text();
    }
    return records;
}

bool DatabaseManager::updateTaskStatus(int taskId, const QString& newStatus)
{
    QSqlQuery query(m_db);
    query.prepare("UPDATE tasks SET Status = :newStatus WHERE TaskID = :taskId");
    query.bindValue(":newStatus", newStatus);
    query.bindValue(":taskId", taskId);
    if (!query.exec()) {
        qCritical() << "更新任务状态失败:" << query.lastError().text();
        return false;
    }
    return true;
}

bool DatabaseManager::getAssignerNameById(int assignerId, QString& name)
{
    return getEmployeeNameById(assignerId, name);
}

bool DatabaseManager::addTask(QSqlRecord& record)
{
    if (!m_db.transaction()) {
        qCritical() << "数据库事务错误：无法启动 (addTask): " << m_db.lastError().text();
        return false;
    }

    QSqlQuery query(m_db);
    query.prepare("INSERT INTO Tasks (TaskName, TaskType, AssignerID, AssigneeID, Status, DueDate, TaskDescription, Evaluation, Destination, StartDate) "
                  "VALUES (:TaskName, :TaskType, :AssignerID, :AssigneeID, :Status, :DueDate, :TaskDescription, :Evaluation, :Destination, :StartDate)");

    // 显式绑定所有值
    query.bindValue(":TaskName", record.value("TaskName"));
    query.bindValue(":TaskType", record.value("TaskType"));
    query.bindValue(":AssignerID", record.value("AssignerID"));
    query.bindValue(":AssigneeID", record.value("AssigneeID"));
    query.bindValue(":Status", record.value("Status"));
    query.bindValue(":DueDate", record.value("DueDate"));
    query.bindValue(":TaskDescription", record.value("TaskDescription"));
    query.bindValue(":Evaluation", record.value("Evaluation"));
    query.bindValue(":Destination", record.value("Destination"));
    query.bindValue(":StartDate", record.value("StartDate"));

    if (!query.exec()) {
        qCritical() << "数据库执行错误 (addTask): " << query.lastError().text();
        m_db.rollback(); // 失败时回滚
        return false;
    }

    record.setValue("TaskID", query.lastInsertId());

    if (m_db.commit()) {
        return true;
    } else {
        qCritical() << "数据库提交错误 (addTask): " << m_db.lastError().text();
        m_db.rollback(); // 提交失败也要回滚
        return false;
    }
}

bool DatabaseManager::updateTask(const QSqlRecord& record)
{
    QSqlQuery query(m_db);
    query.prepare("UPDATE Tasks SET TaskName=:TaskName, TaskType=:TaskType, AssigneeID=:AssigneeID, Status=:Status, "
                  "DueDate=:DueDate, TaskDescription=:TaskDescription, Evaluation=:Evaluation, Destination=:Destination, StartDate=:StartDate "
                  "WHERE TaskID=:TaskID");

    for (int i = 0; i < record.count(); ++i) {
        query.bindValue(":" + record.fieldName(i), record.value(i));
    }

    if (!query.exec()) {
        qCritical() << "更新任务失败: " << query.lastError().text();
        return false;
    }
    return true;
}

bool DatabaseManager::setTaskStatusToCancelled(int taskId)
{
    QSqlQuery query(m_db);
    query.prepare("UPDATE Tasks SET Status = '已作废' WHERE TaskID = ?");
    query.addBindValue(taskId);
    if (!query.exec()) {
        qCritical() << "作废任务失败: " << query.lastError().text();
        return false;
    }
    return query.numRowsAffected() > 0;
}

QSqlRecord DatabaseManager::getTaskRecordById(int taskId)
{
    QSqlQuery query(m_db);
    query.prepare("SELECT * FROM Tasks WHERE TaskID = ?");
    query.addBindValue(taskId);
    if (query.exec() && query.next()) {
        return query.record();
    }
    // 如果没有找到或执行失败，返回一个空的 QSqlRecord
    return QSqlRecord();
}


/***************************************************************************************************
** 招聘管理模块
***************************************************************************************************/

QList<QSqlRecord> DatabaseManager::getJobRequirements()
{
    QList<QSqlRecord> records;
    QSqlQuery query("SELECT id, Department, Position, Duty, NeedNumber, sex, Education, requirement FROM job_requirements ORDER BY id ASC", m_db);
    if (query.exec()) {
        while (query.next()) {
            records.append(query.record());
        }
    } else {
        qCritical() << "加载职位需求数据失败: " << query.lastError().text();
    }
    return records;
}

bool DatabaseManager::updateJobRequirement(int dbId, int column, const QVariant& value)
{
    QString dbColumnName;
    switch (column) {
    case 0: dbColumnName = "Department";   break;
    case 1: dbColumnName = "Position";     break;
    case 2: dbColumnName = "Duty";         break;
    case 3: dbColumnName = "NeedNumber";   break;
    case 4: dbColumnName = "sex";          break;
    case 5: dbColumnName = "Education";    break;
    case 6: dbColumnName = "requirement";  break;
    default: return false;
    }

    QSqlQuery query(m_db);
    query.prepare(QString("UPDATE job_requirements SET %1 = :newValue WHERE id = :id").arg(dbColumnName));
    query.bindValue(":newValue", value);
    query.bindValue(":id", dbId);
    if (!query.exec()) {
        qWarning() << "更新职位需求失败: " << query.lastError().text();
        return false;
    }
    return true;
}

bool DatabaseManager::addJobRequirement(int& newId)
{
    QSqlQuery query(m_db);
    query.prepare("INSERT INTO job_requirements (Department, Position) VALUES ('', '')");
    if (!query.exec()) {
        qCritical() << "新增职位需求失败: " << query.lastError().text();
        return false;
    }
    newId = query.lastInsertId().toInt();
    return true;
}

bool DatabaseManager::deleteJobRequirements(const QList<int>& ids)
{
    if (!m_db.transaction()) {
        qCritical() << "无法启动数据库事务 (deleteJobRequirements): " << m_db.lastError().text();
        return false;
    }
    QSqlQuery query(m_db);
    query.prepare("DELETE FROM job_requirements WHERE id = ?");
    for (int id : ids) {
        query.addBindValue(id);
        if (!query.exec()) {
            qCritical() << "删除职位需求ID " << id << " 失败: " << query.lastError().text();
            m_db.rollback();
            return false;
        }
    }
    return m_db.commit();
}

bool DatabaseManager::addInterviewee(const QSqlRecord& record, int& newId)
{
    QSqlQuery query(m_db);
    query.prepare("INSERT INTO interviewees (name, phone, email, applied_position, status, interview_time, interviewer, notes) "
                  "VALUES (:name, :phone, :email, :applied_position, :status, :interview_time, :interviewer, :notes)");
    query.bindValue(":name", record.value("name"));
    query.bindValue(":phone", record.value("phone"));
    query.bindValue(":email", record.value("email"));
    query.bindValue(":applied_position", record.value("applied_position"));
    query.bindValue(":status", record.value("status"));
    query.bindValue(":interview_time", record.value("interview_time"));
    query.bindValue(":interviewer", record.value("interviewer"));
    query.bindValue(":notes", record.value("notes"));

    if (!query.exec()) {
        qCritical() << "添加新面试者失败: " << query.lastError().text() << query.lastQuery();
        return false;
    }
    newId = query.lastInsertId().toInt();
    return true;
}

bool DatabaseManager::updateInterviewee(const QSqlRecord& record)
{
    QSqlQuery query(m_db);
    query.prepare("UPDATE interviewees SET name=:name, phone=:phone, email=:email, "
                  "applied_position=:applied_position, status=:status, interview_time=:interview_time, "
                  "interviewer=:interviewer, notes=:notes WHERE id=:id");

    for (int i = 0; i < record.count(); ++i) {
        query.bindValue(":" + record.fieldName(i), record.value(i));
    }

    if (!query.exec()) {
        qCritical() << "修改面试者信息失败: " << query.lastError().text();
        return false;
    }
    return true;
}

bool DatabaseManager::deleteInterviewee(int id)
{
    QSqlQuery query(m_db);
    query.prepare("DELETE FROM interviewees WHERE id = ?");
    query.addBindValue(id);
    if (!query.exec()) {
        qCritical() << "删除面试者失败: " << query.lastError().text();
        return false;
    }
    return query.numRowsAffected() > 0;
}


/***************************************************************************************************
** 私有辅助函数
***************************************************************************************************/

bool DatabaseManager::updateApplicationCount(const QString &applicantName, const QString &declarationType)
{
    QSqlQuery queryGetId(m_db);
    queryGetId.prepare("SELECT EmployeeID FROM basicinfo WHERE Name = :name");
    queryGetId.bindValue(":name", applicantName);
    if (!queryGetId.exec() || !queryGetId.next()) {
        qWarning() << "更新申报次数时未找到员工:" << applicantName;
        return false;
    }
    int employeeID = queryGetId.value(0).toInt();

    QString countColumnName;
    if (declarationType == "财务申报") countColumnName = "FinancialNum";
    else if (declarationType == "休假申报") countColumnName = "VacationNum";
    else if (declarationType == "立项申报") countColumnName = "ProjectionNum";
    else return false;

    QSqlQuery queryUpdateCount(m_db);
    queryUpdateCount.prepare(QString("UPDATE personal_application SET %1 = IFNULL(%1, 0) + 1 WHERE EmployeeID = :employeeID").arg(countColumnName));
    queryUpdateCount.bindValue(":employeeID", employeeID);

    if (queryUpdateCount.exec() && queryUpdateCount.numRowsAffected() > 0) {
        return true;
    } else {
        QSqlQuery queryInsert(m_db);
        queryInsert.prepare(QString("INSERT INTO personal_application (EmployeeID, Name, %1) VALUES (:employeeID, :name, 1)").arg(countColumnName));
        queryInsert.bindValue(":employeeID", employeeID);
        queryInsert.bindValue(":name", applicantName);
        if (!queryInsert.exec()) {
            qCritical() << "插入新的个人申请计数失败:" << queryInsert.lastError().text();
            return false;
        }
    }
    return true;
}

bool DatabaseManager::updateApprovalStats(const QString& applicantName, bool isApproved)
{
    QString countColumn = isApproved ? "PassNum" : "RejectNum";

    QSqlQuery queryGetId(m_db);
    queryGetId.prepare("SELECT EmployeeID FROM basicinfo WHERE Name = :name");
    queryGetId.bindValue(":name", applicantName);
    if (!queryGetId.exec() || !queryGetId.next()) {
        qWarning() << "更新审批统计时未找到员工:" << applicantName;
        return false;
    }
    int employeeID = queryGetId.value(0).toInt();

    QSqlQuery queryUpdateStats(m_db);
    queryUpdateStats.prepare(QString("UPDATE personal_application SET %1 = IFNULL(%1, 0) + 1 WHERE EmployeeID = :id").arg(countColumn));
    queryUpdateStats.bindValue(":id", employeeID);

    if (queryUpdateStats.exec() && queryUpdateStats.numRowsAffected() > 0) {
        return true;
    } else {
        QSqlQuery queryInsertStats(m_db);
        queryInsertStats.prepare(QString("INSERT INTO personal_application (EmployeeID, Name, %1) VALUES (:id, :name, 1)").arg(countColumn));
        queryInsertStats.bindValue(":id", employeeID);
        queryInsertStats.bindValue(":name", applicantName);
        if (!queryInsertStats.exec()) {
            qCritical() << "插入新的审批统计失败:" << queryInsertStats.lastError().text();
            return false;
        }
    }
    return true;
}


bool DatabaseManager::insertChatMessage(int senderId, int receiverId, const QString &content)
{
    QSqlQuery query(m_db); // 假设你的数据库连接对象叫 m_database，请根据你实际代码调整
    query.prepare("INSERT INTO chat_messages (sender_id, receiver_id, content) "
                  "VALUES (:sender, :receiver, :content)");
    query.bindValue(":sender", senderId);
    query.bindValue(":receiver", receiverId);
    query.bindValue(":content", content);

    if (!query.exec()) {
        qDebug() << "插入聊天记录失败: " << query.lastError().text();
        return false;
    }
    return true;
}

QSqlQuery DatabaseManager::getChatHistory(int user1Id, int user2Id)
{
    QSqlQuery query(m_db);
    query.prepare("SELECT sender_id, content FROM chat_messages "
                  "WHERE (sender_id = :u1 AND receiver_id = :u2) "
                  "   OR (sender_id = :u3 AND receiver_id = :u4) "
                  "ORDER BY send_time ASC");
    query.bindValue(":u1", user1Id);
    query.bindValue(":u2", user2Id);
    query.bindValue(":u3", user2Id);
    query.bindValue(":u4", user1Id);
    query.exec();
    return query;
}


QSqlQuery DatabaseManager::getUnreadMessages(int myUserId)
{
    QSqlQuery query(m_db);
    query.prepare("SELECT msg_id, sender_id, content FROM chat_messages "
                  "WHERE receiver_id = :myId AND is_read = 0 ORDER BY send_time ASC");
    query.bindValue(":myId", myUserId);
    query.exec();
     return query;
}


bool DatabaseManager::markMessageAsRead(int msgId)
{
    QSqlQuery query(m_db);
    query.prepare("UPDATE chat_messages SET is_read = 1 WHERE msg_id = :msgId");
    query.bindValue(":msgId", msgId);
    return query.exec();
}

QSqlQuery DatabaseManager::getOtherEmployees(int myUserId)
{
    QSqlQuery query(m_db);
    query.prepare("SELECT EmployeeID, Name, Department, Position, Pic FROM basicinfo WHERE EmployeeID != :myId");
    query.bindValue(":myId", myUserId);
    query.exec();
    return query;
}
