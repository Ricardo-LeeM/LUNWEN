#include "statusdelegate.h"

    /***************************************************************************************************
** 构造函数实现 (Constructor Implementation)                         **
***************************************************************************************************/

    /**
 * @brief StatusDelegate 构造函数
 * @param parent 父对象指针
 */
    StatusDelegate::StatusDelegate(QObject *parent) : QStyledItemDelegate(parent)
{
}

/***************************************************************************************************
** 核心功能实现 (Core Functionality Implementation)                  **
***************************************************************************************************/

/**
 * @brief 实现状态码到显示文本的转换
 *
 * 这是 `StatusDelegate` 的核心。每当表格需要绘制一个单元格时，如果该列被设置了此委托，
 * 那么这个函数就会被调用。
 *
 * @param value 从模型中获取的原始数据（QVariant类型）。
 * @param locale 区域设置（本函数中未使用）。
 * @return 根据 value 转换后的、要在UI上显示的字符串。
 */
QString StatusDelegate::displayText(const QVariant &value, const QLocale &locale) const
{
    // Q_UNUSED 是一个宏，用来告诉编译器我们知道 locale 这个参数存在，但是我们故意没有使用它。
    // 这可以避免编译器产生“未使用参数”的警告。
    Q_UNUSED(locale);

    // 1. 从 QVariant 中提取整数状态码。
    // toInt() 会尝试将 QVariant 转换为整数。如果转换失败，通常会返回0。
    int status = value.toInt();

    // 2. 使用 switch 语句进行“翻译”。
    // 这是将数字状态映射到描述性文本的核心逻辑。
    switch (status) {
    case 0: return "待筛选";
    case 1: return "一面";
    case 2: return "二面";
    case 3: return "HR面";
    case 4: return "待定";
    case 5: return "已录用";
    case 6: return "未录用";
    default: return "未知状态"; // 提供一个默认返回值，以处理意外的状态码。
    }
}
// #include "statusdelegate.h"

// StatusDelegate::StatusDelegate(QObject *parent) : QStyledItemDelegate(parent)
// {
// }

// QString StatusDelegate::displayText(const QVariant &value, const QLocale &locale) const
// {
//     Q_UNUSED(locale); // 告诉编译器我们没用到这个参数

//     int status = value.toInt(); // 从模型获取的值 (0, 1, 2...)

//     // 在这里进行“翻译”
//     switch (status) {
//     case 0: return "待筛选";
//     case 1: return "一面";
//     case 2: return "二面";
//     case 3: return "HR面";
//     case 4: return "待定";
//     case 5: return "已录用";
//     case 6: return "未录用";
//     default: return "未知状态";
//     }
// }
