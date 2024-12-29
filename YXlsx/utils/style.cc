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

#include "style.h"

#include <QXmlStreamReader>
#include <QXmlStreamWriter>

QT_BEGIN_NAMESPACE_YXLSX

Style::Style(XmlMode mode)
    : AbstractOOXmlFile { mode }
{
}

Style::~Style() { }

void Style::ComposeXml(QIODevice* device) const
{
    QXmlStreamWriter writer(device);
    writer.writeStartDocument();

    // Start the styleSheet element with the necessary namespace
    writer.writeStartElement(QStringLiteral("styleSheet"));
    writer.writeDefaultNamespace(QStringLiteral("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));

    // Write <numFmts> element (Optional: Only include if you need custom number formats)
    writer.writeStartElement(QStringLiteral("numFmts"));
    writer.writeAttribute(QStringLiteral("count"), QStringLiteral("1"));
    writer.writeStartElement(QStringLiteral("numFmt"));
    writer.writeAttribute(QStringLiteral("numFmtId"), QStringLiteral("164")); // General format
    writer.writeAttribute(QStringLiteral("formatCode"), QStringLiteral("General"));
    writer.writeEndElement(); // numFmt
    writer.writeEndElement(); // numFmts

    // Write <borders> element (Defines a border style, here an empty one)
    writer.writeStartElement(QStringLiteral("borders"));
    writer.writeAttribute(QStringLiteral("count"), QStringLiteral("1"));
    writer.writeEmptyElement(QStringLiteral("border")); // Empty border definition
    writer.writeEndElement(); // borders

    // Write <cellXfs> element (Defines the text style without font and size)
    writer.writeStartElement(QStringLiteral("cellXfs"));
    writer.writeAttribute(QStringLiteral("count"), QStringLiteral("1"));
    writer.writeStartElement(QStringLiteral("xf"));
    writer.writeAttribute(QStringLiteral("numFmtId"), QStringLiteral("0")); // Default number format (General)
    writer.writeAttribute(QStringLiteral("borderId"), QStringLiteral("0")); // First (and only) border
    writer.writeEndElement(); // xf
    writer.writeEndElement(); // cellXfs

    // End the styleSheet element
    writer.writeEndElement(); // styleSheet

    writer.writeEndDocument();
}

bool Style::ParseXml(QIODevice* device)
{
    Q_UNUSED(device)

    // QXmlStreamReader reader(device);
    // while (!reader.atEnd() && !reader.hasError()) {
    //     if (reader.readNextStartElement() && reader.name() == QLatin1String("styleSheet")) {
    //         // 样式解析逻辑为空，因为无需处理任何样式
    //     }
    // }
    // return !reader.hasError();

    return true;
}

QT_END_NAMESPACE_YXLSX
