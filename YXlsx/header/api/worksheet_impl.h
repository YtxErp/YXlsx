#pragma once

#include <QDebug>

#include "namespace.h"
#include "utility.h"
#include "worksheet.h"

QT_BEGIN_NAMESPACE_YXLSX

template <Container T> bool Worksheet::WriteColumn(int row, int column, const T& container)
{
    if (container.size() == 0 || !Utility::CheckCoordinateValid(row, column)) {
        qWarning() << "Data is empty for column write.";
        return false;
    }

    int end_row = row + container.size() - 1;
    if (!UpdateDimension(row, column) || !UpdateDimension(end_row, column)) {
        qWarning() << "Failed to update dimensions for column write.";
        return false;
    }

    int current_row { row };
    for (const auto& value : container) {
        QVariant qvalue(value);

        const CellType cell_type { DetermineCellType(qvalue) };

        if (cell_type != CellType::kUnknown) {
            auto cell { QSharedPointer<Cell>::create(qvalue, cell_type) };

            if (cell_type == CellType::kSharedString) {
                shared_string_->SetSharedString(qvalue.toString(), current_row, column);
            }

            WriteMatrix(current_row, column, cell);
        }

        ++current_row;
    }

    return true;
}

template <Container T> bool Worksheet::WriteRow(int row, int column, const T& container)
{
    if (container.size() == 0 || !Utility::CheckCoordinateValid(row, column)) {
        qWarning() << "Data is empty for row write.";
        return false;
    }

    int end_column = column + container.size() - 1;
    if (!UpdateDimension(row, column) || !UpdateDimension(row, end_column)) {
        qWarning() << "Failed to update dimensions for row write.";
        return false;
    }

    int current_column { column };
    for (const auto& value : container) {
        QVariant qvalue(value);

        const CellType cell_type { DetermineCellType(qvalue) };

        if (cell_type != CellType::kUnknown) {
            auto cell { QSharedPointer<Cell>::create(qvalue, cell_type) };

            if (cell_type == CellType::kSharedString) {
                shared_string_->SetSharedString(qvalue.toString(), row, current_column);
            }

            WriteMatrix(row, current_column, cell);
        }

        ++current_column;
    }

    return true;
}

QT_END_NAMESPACE_YXLSX
