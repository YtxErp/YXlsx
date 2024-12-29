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

// Concept to check if a type is a container where all elements are QVariant.
template <typename T>
concept Container = requires(T t) {
    // Check if the type provides a size() method that returns an integral type.
    { t.size() } -> std::integral;
    // Check if the type provides an indexing method (operator[] or at) that returns a const QVariant&.
    { t.at(0) } -> std::convertible_to<const QVariant&>;
    // Ensure T provides iteration and its elements are QVariant-compatible.
    requires std::same_as<std::remove_cvref_t<decltype(*std::begin(t))>, QVariant>;
};

class Worksheet final : public AbstractSheet {
public:
    Worksheet(const QString& sheet_name, int sheet_id, QSharedPointer<SharedString> shared_strings, SheetType sheet_type);
    ~Worksheet();

public:
    bool Write(const Coordinate& coordinate, const QVariant& data);
    bool Write(int row, int column, const QVariant& data);

    template <Container T> inline bool WriteColumn(int start_row, int column, const T& data);
    template <Container T> inline bool WriteRow(int row, int start_column, const T& data);

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

    void WriteMatrix(int row, int column, const QSharedPointer<Cell>& cell) { matrix_.insert({ row, column }, cell); }
    QSharedPointer<Cell> ReadMatrix(int row, int column) const { return matrix_.value({ row, column }); }

    bool WriteBlank(int row, int column);

    bool Contains(int row, int column) const { return matrix_.contains({ row, column }); }
    CellType DetermineCellType(const QVariant& value) const;

private:
    Dimension dimension_ {};
    QSharedPointer<SharedString> shared_string_ {};
    mutable QHash<int, QString> row_spans_hash_ {};
    SheetFormatProps sheet_format_props_ {};
    QMap<std::pair<int, int>, QSharedPointer<Cell>> matrix_ {};
};

QT_END_NAMESPACE_YXLSX

// Include template implementations
#include "worksheet_impl.h"

#endif // YXLSX_WORKSHEET_H
