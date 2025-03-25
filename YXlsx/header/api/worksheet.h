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

#ifndef YXLSX_WORKSHEET_H
#define YXLSX_WORKSHEET_H

#include <QXmlStreamReader>
#include <QXmlStreamWriter>

#include "abstractsheet.h"
#include "cell.h"
#include "coordinate.h"
#include "dimension.h"
#include "sharedstring.h"
#include "sheetformatprops.h"

QT_BEGIN_NAMESPACE_YXLSX

template <typename T>
concept Container = std::ranges::range<T> && requires(T t) {
    typename std::ranges::range_value_t<T>;
    requires std::convertible_to<std::ranges::range_value_t<T>, QVariant>;
};

class Worksheet final : public AbstractSheet {
public:
    Worksheet(const QString& sheet_name, int sheet_id, QSharedPointer<SharedString> shared_strings, SheetType sheet_type);
    ~Worksheet();

public:
    bool Write(const Coordinate& coordinate, const QVariant& data);
    bool Write(int row, int column, const QVariant& data);

    template <Container T> inline bool WriteColumn(int row, int column, const T& data);
    template <Container T> inline bool WriteRow(int row, int column, const T& data);

    QVariant Read(const Coordinate& coordinate) const;
    QVariant Read(int row, int column) const;

private:
    void ComposeXml(QIODevice* device) const override;
    bool ParseXml(QIODevice* device) override;
    void ProcessCell(QXmlStreamReader& reader, int row, int& column);
    QVariant ParseCellValue(const QString& value, CellType cell_type, int row, int column);

    bool UpdateDimension(int row, int col);
    QString ComposeDimension() const;
    void CalculateSpans() const;

    void ComposeSheet(QXmlStreamWriter& writer) const;
    void ComposeCell(QXmlStreamWriter& writer, int row, int col, const QSharedPointer<Cell>& cell) const;

    void ParseSheet(QXmlStreamReader& reader);

    inline void WriteMatrix(int row, int column, const QSharedPointer<Cell>& cell) { matrix_.insert({ row, column }, cell); }
    inline QSharedPointer<Cell> ReadMatrix(int row, int column) const { return matrix_.value({ row, column }); }

    bool WriteBlank(int row, int column);

    inline bool Contains(int row, int column) const { return matrix_.contains({ row, column }); }
    CellType DetermineCellType(const QVariant& value) const;

private:
    Dimension dimension_ {};
    QSharedPointer<SharedString> shared_string_ {};
    mutable QHash<int, QString> row_spans_hash_ {};
    SheetFormatProps sheet_format_props_ {};
    QMap<std::pair<int, int>, QSharedPointer<Cell>> matrix_ {};

    static constexpr std::array<std::pair<int, CellType>, 9> type_map = { {
        { QMetaType::QString, CellType::kSharedString },
        { QMetaType::Int, CellType::kNumber },
        { QMetaType::UInt, CellType::kNumber },
        { QMetaType::LongLong, CellType::kNumber },
        { QMetaType::ULongLong, CellType::kNumber },
        { QMetaType::Double, CellType::kNumber },
        { QMetaType::Float, CellType::kNumber },
        { QMetaType::Bool, CellType::kBoolean },
        { QMetaType::QDateTime, CellType::kDate },
    } };
};

QT_END_NAMESPACE_YXLSX

// Include template implementations
#include "worksheet_impl.h"

#endif // YXLSX_WORKSHEET_H
