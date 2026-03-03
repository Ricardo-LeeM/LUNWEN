// usersession.h
#ifndef USERSESSION_H
#define USERSESSION_H

#include <QObject>
#include <QString>
#include <QDate>
#include <QPixmap> // 用于存储图片

class UserSession : public QObject
{
    Q_OBJECT
public:
    static UserSession* instance();

    // Setter for all current user information (called after successful login)
    void setCurrentUser(int id,
                        const QString &name,
                        const QString &department,
                        const QString &employeeStatus,
                        bool isPartyMember,
                        const QString &workNature,
                        const QString &position,
                        const QDate &joinDate,
                        const QPixmap &pixmap, // 存储处理后的 QPixmap
                        const QString &employeeType);

    // Getters for individual pieces of information
    int employeeID() const { return m_employeeID; }
    QString name() const { return m_name; }
    QString department() const { return m_department; }
    QString employeeStatus() const { return m_employeeStatus; }
    bool isPartyMember() const { return m_isPartyMember; }
    QString workNature() const { return m_workNature; }
    QString position() const { return m_position; }
    QDate joinDate() const { return m_joinDate; }
    QPixmap userPixmap() const { return m_userPixmap; } // 获取用户头像
    QString employeeType() const { return m_employeeType; }

    // Optional: A method to clear session data on logout
    void clearSession();

private:
    explicit UserSession(QObject *parent = nullptr);
    ~UserSession(); // 添加析构函数声明，虽然在这个单例中可能不直接调用delete

    static UserSession *m_instance;

    // Member variables to store user information
    int m_employeeID = -1;
    QString m_name;
    QString m_department;
    QString m_employeeStatus;
    bool m_isPartyMember = false;
    QString m_workNature;
    QString m_position;
    QDate m_joinDate;
    QPixmap m_userPixmap; // 存储头像
    QString m_employeeType;
};

#endif // USERSESSION_H
