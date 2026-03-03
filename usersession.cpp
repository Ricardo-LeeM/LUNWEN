// usersession.cpp
#include "usersession.h"

UserSession* UserSession::m_instance = nullptr;

UserSession* UserSession::instance()
{
    if (!m_instance) {
        // 通常在多线程环境下，这里需要加锁保护
        // 但对于典型的Qt GUI应用，如果 UserSession 在主线程中首次创建和使用，可以简化
        m_instance = new UserSession;
    }
    return m_instance;
}

UserSession::UserSession(QObject *parent)
    : QObject(parent)
{
    // 初始化，如果需要的话
}

UserSession::~UserSession()
{
    // 在单例模式下，这个析构函数通常不会被外部直接调用
    // m_instance 的清理可以在程序退出时处理，或者依赖操作系统
    // 如果需要显式清理，可以添加一个 static 方法来 delete m_instance
}

void UserSession::setCurrentUser(int id,
                                 const QString &name,
                                 const QString &department,
                                 const QString &employeeStatus,
                                 bool isPartyMember,
                                 const QString &workNature,
                                 const QString &position,
                                 const QDate &joinDate,
                                 const QPixmap &pixmap,
                                 const QString &employeeType)
{
    m_employeeID = id;
    m_name = name;
    m_department = department;
    m_employeeStatus = employeeStatus;
    m_isPartyMember = isPartyMember;
    m_workNature = workNature;
    m_position = position;
    m_joinDate = joinDate;
    m_userPixmap = pixmap; // 直接存储 QPixmap
    m_employeeType = employeeType;
}

void UserSession::clearSession()
{
    m_employeeID = -1;
    m_name.clear();
    m_department.clear();
    m_employeeStatus.clear();
    m_isPartyMember = false;
    m_workNature.clear();
    m_position.clear();
    m_joinDate = QDate(); // 重置为无效日期
    m_userPixmap = QPixmap(); // 清空图片
    m_employeeType.clear();
}

