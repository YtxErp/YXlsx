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

#include "docpropscore.h"

#include <QDateTime>
#include <QDebug>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

QT_BEGIN_NAMESPACE_YXLSX

inline const QString kCorePropertyies { QStringLiteral("http://schemas.openxmlformats.org/package/2006/metadata/core-properties") };
inline const QString kPLElements { QStringLiteral("http://purl.org/dc/elements/1.1/") };
inline const QString kPLTerms { QStringLiteral("http://purl.org/dc/terms/") };
inline const QString kPLDcmiType = QStringLiteral("http://purl.org/dc/dcmitype/");
inline const QString kW3SchemaInstance = QStringLiteral("http://www.w3.org/2001/XMLSchema-instance");

const QHash<QString, QString> DocPropsCore::kElementNamespaceHash { { QStringLiteral("creator"), kPLElements },
    { QStringLiteral("lastModifiedBy"), kCorePropertyies }, { QStringLiteral("created"), kPLTerms }, { QStringLiteral("modified"), kPLTerms } };

DocPropsCore::DocPropsCore(XmlMode mode)
    : AbstractOOXmlFile { mode }
{
}

bool DocPropsCore::SetProperty(const QString& name, const QString& value)
{
    static const QSet<QString> valid_keys { QStringLiteral("creator"), QStringLiteral("lastModifiedBy"), QStringLiteral("created"),
        QStringLiteral("modified") };

    if (!valid_keys.contains(name)) {
        return false;
    }

    if (value.isEmpty()) {
        return false;
    }

    property_hash_[name] = value;
    return true;
}

QString DocPropsCore::GetProperty(const QString& name) const { return property_hash_.value(name, QString()); }

QStringList DocPropsCore::GetProperty() const { return property_hash_.keys(); }

void DocPropsCore::ComposeXml(QIODevice* device) const
{
    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("cp:coreProperties"));

    writer.writeNamespace(kCorePropertyies, QStringLiteral("cp"));
    writer.writeNamespace(kPLElements, QStringLiteral("dc"));
    writer.writeNamespace(kPLTerms, QStringLiteral("dcterms"));
    writer.writeNamespace(kPLDcmiType, QStringLiteral("dcmitype"));
    writer.writeNamespace(kW3SchemaInstance, QStringLiteral("xsi"));

    auto it = property_hash_.constFind(QStringLiteral("creator"));
    writer.writeTextElement(kPLElements, QStringLiteral("creator"), it != property_hash_.constEnd() ? it.value() : QStringLiteral("YXlsx Library"));

    it = property_hash_.constFind(QStringLiteral("lastModifiedBy"));
    writer.writeTextElement(kCorePropertyies, QStringLiteral("lastModifiedBy"), it != property_hash_.constEnd() ? it.value() : QStringLiteral("YXlsx Library"));

    auto WriteTimeElement = [&](const QString& name, const QString& value) {
        writer.writeStartElement(kPLTerms, name);
        writer.writeAttribute(kW3SchemaInstance, QStringLiteral("type"), QStringLiteral("dcterms:W3CDTF"));
        writer.writeCharacters(value);
        writer.writeEndElement();
    };

    const auto created { property_hash_.constFind(QStringLiteral("created")) };
    WriteTimeElement(QStringLiteral("created"), created != property_hash_.constEnd() ? created.value() : QDateTime::currentDateTime().toString(Qt::ISODate));
    WriteTimeElement(QStringLiteral("modified"), QDateTime::currentDateTime().toString(Qt::ISODate));

    writer.writeEndElement(); // cp:coreProperties
    writer.writeEndDocument();
}

bool DocPropsCore::ParseXml(QIODevice* device)
{
    QXmlStreamReader reader(device);

    while (!reader.atEnd() && !reader.hasError()) {
        if (!reader.readNextStartElement())
            continue;

        const auto namespace_uri { reader.namespaceUri().toString() };
        const auto name { reader.name().toString() };

        auto it = kElementNamespaceHash.constFind(name);
        if (it != kElementNamespaceHash.constEnd() && it.value() == namespace_uri) {
            SetProperty(name, reader.readElementText());
        }
    }

    if (reader.hasError()) {
        qDebug() << "DocPropsCore Error reading doc props app file:" << reader.errorString();
        return false;
    }

    return true;
}

QT_END_NAMESPACE_YXLSX
