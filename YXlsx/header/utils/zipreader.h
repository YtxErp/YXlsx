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

#ifndef YXLSX_ZIPREADER_H
#define YXLSX_ZIPREADER_H

#include <private/qzipreader_p.h>

#include <QScopedPointer>

#include "namespace.h"

QT_BEGIN_NAMESPACE_YXLSX

class ZipReader {
    Q_DISABLE_COPY(ZipReader)

public:
    explicit ZipReader(const QString& file_path);
    explicit ZipReader(QIODevice* device);
    ~ZipReader() = default;

    inline const QStringList& GetFilePath() const { return file_path_; }
    inline QByteArray GetFileData(const QString& file_path) const { return reader_->fileData(file_path); }

private:
    void Init();

private:
    QScopedPointer<QZipReader> reader_ {};
    QStringList file_path_ {};
};

QT_END_NAMESPACE_YXLSX

#endif // YXLSX_ZIPREADER_H
