#pragma once

#include <QDebug>

#include "namespace.h"
#include "worksheet.h"

QT_BEGIN_NAMESPACE_YXLSX

template <Container T> bool Worksheet::WriteColumn(int start_row, int column, const T& data)
{
    if (data.size() == 0) {
        qWarning() << "Data is empty for column write.";
        return false;
    }

    int end_row = start_row + data.size() - 1;
    if (!UpdateDimension(start_row, column) || !UpdateDimension(end_row, column)) {
        qWarning() << "Failed to update dimensions for column write.";
        return false;
    }

    CellType cell_type = DetermineCellType(data.at(0));
    if (cell_type == CellType::kUnknown) {
        return false;
    }

    for (int i = 0; i != data.size(); ++i) {
        int current_row = start_row + i;
        const auto& value = data.at(i);

        auto cell = QSharedPointer<Cell>::create(value, cell_type);
        if (cell_type == CellType::kSharedString) {
            shared_string_->SetSharedString(value.toString(), current_row, column);
        }

        WriteMatrix(current_row, column, cell);
    }

    return true;
}

template <Container T> bool Worksheet::WriteRow(int row, int start_column, const T& data)
{
    if (data.size() == 0) {
        qWarning() << "Data is empty for row write.";
        return false;
    }

    int end_column = start_column + data.size() - 1;
    if (!UpdateDimension(row, start_column) || !UpdateDimension(row, end_column)) {
        qWarning() << "Failed to update dimensions for row write.";
        return false;
    }

    for (int i = 0; i != data.size(); ++i) {
        int current_column = start_column + i;
        const auto& value = data.at(i);

        CellType cell_type = DetermineCellType(value);
        if (cell_type == CellType::kUnknown) {
            qWarning() << "Unsupported data type at row" << row << ", column" << current_column;
            continue;
        }

        auto cell = QSharedPointer<Cell>::create(value, cell_type);
        if (cell_type == CellType::kSharedString) {
            shared_string_->SetSharedString(value.toString(), row, current_column);
        }

        WriteMatrix(row, current_column, cell);
    }

    return true;
}

QT_END_NAMESPACE_YXLSX
