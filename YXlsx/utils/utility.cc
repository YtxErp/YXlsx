/*
 * MIT License
 *
 * Copyright (c) 2024 YTX
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

#include "utility.h"

#include <QDebug>
#include <QRegularExpression>

QT_BEGIN_NAMESPACE_YXLSX

QStringList Utility::SplitPath(const QString& path)
{
    int idx = path.lastIndexOf(QLatin1Char('/'));
    if (idx == -1)
        return { QStringLiteral("."), path };

    return { path.left(idx), path.mid(idx + 1) };
}

/*
 * Return the .rel file path based on filePath
 */
QString Utility::GetRelFilePath(const QString& filePath)
{
    QString ret;

    int idx = filePath.lastIndexOf(QLatin1Char('/'));
    if (idx == -1) // not found
    {
        // return QString();

        // dev34
        ret = QLatin1String("_rels/") + QStringLiteral("%0.rels").arg(filePath);
        return ret;
    }

    ret = QString(filePath.left(idx) + QLatin1String("/_rels/") + filePath.mid(idx + 1) + QLatin1String(".rels"));
    return ret;
}

/*
  Creates a valid and unique sheet name.
    - Minimum length: 1
    - Maximum length: 31
    - Invalid characters (/ \ ? * ] [ :) are replaced by a single space (' ').
    - Sheet names must not begin or end with an apostrophe ('').
    - Ensures the name is unique by appending a number if necessary.
    - Generates a default name ("Sheet <N>") if the proposed name is empty or invalid.
 */
QString Utility::GenerateSheetName(const QStringList& sheet_names, const QString& name_proposal, int& last_sheet_index)
{
    QString ret { name_proposal };

    // If name is empty, generate a default name
    if (ret.isEmpty()) {
        do {
            ++last_sheet_index;
            ret = QStringLiteral("Sheet %1").arg(last_sheet_index);
        } while (sheet_names.contains(ret));

        return ret;
    }

    // Unescape sheet name if it starts and ends with an apostrophe
    if (ret.length() >= 3 && ret.startsWith(QLatin1Char('\'')) && ret.endsWith(QLatin1Char('\''))) {
        ret = UnescapeSheetName(ret);
    }

    // Replace invalid characters with a space
    static const QRegularExpression invalid_chars(QStringLiteral("[/\\\\?*\\][:]"));
    ret.replace(invalid_chars, QStringLiteral(" "));

    // Ensure the name doesn't start or end with an apostrophe
    if (ret.startsWith(QLatin1Char('\'')))
        ret[0] = QLatin1Char(' ');
    if (ret.endsWith(QLatin1Char('\'')))
        ret[ret.size() - 1] = QLatin1Char(' ');

    // Truncate to the maximum allowed length
    if (ret.size() >= 32)
        ret = ret.left(31);

    // Ensure uniqueness by appending a number if necessary
    QString unique_name = ret;
    int counter = 1;
    while (sheet_names.contains(unique_name)) {
        QString suffix = QStringLiteral(" (%1)").arg(counter++);
        if (ret.size() + suffix.size() >= 32)
            unique_name = ret.left(31 - suffix.size()) + suffix;
        else
            unique_name = ret + suffix;
    }

    return unique_name;
}

/*
 */
QString Utility::UnescapeSheetName(const QString& sheetName)
{
    Q_ASSERT(sheetName.length() > 2 && sheetName.startsWith(QLatin1Char('\'')) && sheetName.endsWith(QLatin1Char('\'')));

    QString name = sheetName.mid(1, sheetName.length() - 2);
    name.replace(QLatin1String("\'\'"), QLatin1String("\'"));
    return name;
}

/*
 * whether the string s starts or ends with space
 */
bool Utility::IsSpaceReserveNeeded(const QString& s)
{
    QString spaces(QStringLiteral(" \t\n\r"));
    return !s.isEmpty() && (spaces.contains(s.at(0)) || spaces.contains(s.at(s.length() - 1)));
}

std::pair<int, int> Utility::ParseCoordinate(const QString& cell)
{
    static const QRegularExpression re(QRegularExpression::anchoredPattern(QStringLiteral(R"(\$?([A-Za-z]{1,3})\$?(\d+))")));

    int row {};
    int column {};

    const QRegularExpressionMatch match = re.match(cell);
    if (!match.hasMatch()) {
        return { kInvalidInt, kInvalidInt };
    }

    const QString col_str = match.captured(1);
    const QString row_str = match.captured(2);

    row = row_str.toInt();
    column = ParseColumn(col_str);

    if (row <= 0 || column <= 0) {
        return { kInvalidInt, kInvalidInt };
    }

    return { row, column };
}

int Utility::ParseColumn(const QString& column)
{
    if (column.isEmpty()) {
        return kInvalidInt;
    }

    int col = 0;
    for (QChar ch : column) {
        if (!ch.isLetter()) {
            return kInvalidInt;
        }

        col = col * 26 + (ch.toUpper().unicode() - 'A' + 1);
    }

    return col;
}

QString Utility::ComposeCoordinate(int row, int column, bool row_abs, bool col_abs)
{
    if (row <= 0 || column <= 0)
        return {};

    return QString("%1%2%3%4").arg(col_abs ? "$" : "").arg(ComposeColumn(column)).arg(row_abs ? "$" : "").arg(row);
}

QString Utility::ComposeColumn(int column)
{
    QString col_str {};

    while (column > 0) {
        int remainder = (column - 1) % 26;
        col_str.prepend(QChar('A' + remainder));
        column = (column - 1) / 26;
    }

    return col_str;
}

QT_END_NAMESPACE_YXLSX
