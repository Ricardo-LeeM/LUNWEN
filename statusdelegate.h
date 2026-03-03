#ifndef STATUSDELEGATE_H
#define STATUSDELEGATE_H

/***************************************************************************************************
** 头文件依赖 (Header Dependencies)                                  **
***************************************************************************************************/
#include <QStyledItemDelegate>

/***************************************************************************************************
** StatusDelegate 类定义                                       **
***************************************************************************************************/
/**
 * @class StatusDelegate
 * @brief 状态显示委托
 *
 * 在Qt的Model/View架构中，委托（Delegate）负责在视图（View）中渲染和编辑由模型（Model）提供的数据项。
 * 这个 `StatusDelegate` 类是一个自定义的委托，它的核心任务是将模型中存储的整数类型的状态码
 * (例如 0, 1, 2...) 转换成用户友好的、可读的文本字符串 (例如 "待筛选", "一面", "二面"...)
 * 并显示在表格中。
 *
 * 这样做的主要好处是实现了数据存储与数据表现的分离。数据库可以高效地存储数字，而UI则能
 * 以一种更直观的方式向用户展示信息。
 */
class StatusDelegate : public QStyledItemDelegate
{
    Q_OBJECT
public:
    /**
     * @brief 构造函数
     * @param parent 父对象指针，默认为 nullptr。
     */
    explicit StatusDelegate(QObject *parent = nullptr);

    /**
     * @brief 重写 displayText 函数
     *
     * 这是 QStyledItemDelegate 中的一个虚函数。当视图需要显示一个数据项的文本时，
     * 它会调用这个函数。我们通过重写它，来提供自定义的显示逻辑，即“翻译”状态码。
     *
     * @param value 模型传递过来的原始数据，是一个 QVariant 类型，在这里我们预期它包含一个整数。
     * @param locale 区域设置信息，用于支持国际化，本例中未使用。
     * @return 返回最终要显示在单元格中的文本字符串。
     */
    QString displayText(const QVariant &value, const QLocale &locale) const override;
};

#endif // STATUSDELEGATE_H
// #ifndef STATUSDELEGATE_H
// #define STATUSDELEGATE_H

// #include <QStyledItemDelegate>

// class StatusDelegate : public QStyledItemDelegate
// {
//     Q_OBJECT
// public:
//     explicit StatusDelegate(QObject *parent = nullptr);

//     // 重写这个函数，来自定义如何显示数据
//     QString displayText(const QVariant &value, const QLocale &locale) const override;
// };

// #endif // STATUSDELEGATE_H
