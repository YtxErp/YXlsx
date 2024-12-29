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

#include "docpropsapp.h"

#include <QtCore/qdebug.h>

#include <QBuffer>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

QT_BEGIN_NAMESPACE_YXLSX

DocPropsApp::DocPropsApp(XmlMode mode)
    : AbstractOOXmlFile { mode }
{
}

void DocPropsApp::AddTitle(const QString& title) { title_list_.append(title); }

void DocPropsApp::AddHeading(const QString& name, int value) { heading_list_.append({ name, value }); }

bool DocPropsApp::SetProperty(const QString& name, const QString& value)
{
    static const QSet<QString> kValidKey { QStringLiteral("Manager"), QStringLiteral("Company"), QStringLiteral("Application"), QStringLiteral("DocSecurity"),
        QStringLiteral("ScaleCrop"), QStringLiteral("LinksUpToDate"), QStringLiteral("SharedDoc"), QStringLiteral("HyperlinksChanged"),
        QStringLiteral("AppVersion") };

    // Check if the property name is valid
    if (!kValidKey.contains(name)) {
        return false;
    }

    // Check if the value is empty
    if (value.isEmpty()) {
        return false;
    }

    // Handle special cases for boolean properties
    if (name == QStringLiteral("ScaleCrop") || name == QStringLiteral("LinksUpToDate") || name == QStringLiteral("SharedDoc")
        || name == QStringLiteral("HyperlinksChanged")) {
        // If the value is not "true" or "false", return false
        if (value != QStringLiteral("true") && value != QStringLiteral("false")) {
            qWarning() << "DocPropsApp Invalid boolean value for property:" << name << value;
            return false;
        }
    }

    // Store the property name and value in the hash map
    property_hash_[name] = value;

    return true;
}

/*!
 * Retrieves the value of the property with the given \a name.
 *
 * \param name The name of the property to retrieve.
 * \return The value of the property if it exists; otherwise, an empty QString.
 */
QString DocPropsApp::GetProperty(const QString& name) const { return property_hash_.value(name, QString()); }

QStringList DocPropsApp::GetProperty() const { return property_hash_.keys(); }

void DocPropsApp::ComposeXml(QIODevice* device) const
{
    if (!device || !device->isWritable()) {
        qWarning() << "Invalid or unwritable device.";
        return;
    }

    QXmlStreamWriter writer(device);
    QString vt = QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

    writer.writeStartDocument(QStringLiteral("1.0"), true);
    writer.writeStartElement(QStringLiteral("Properties"));
    writer.writeDefaultNamespace(QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"));
    writer.writeNamespace(vt, QStringLiteral("vt"));

    // Write basic properties
    auto writeProperty = [&writer, this](const QString& key, const QString& default_value) {
        auto it = property_hash_.constFind(key);
        writer.writeTextElement(key, it != property_hash_.constEnd() ? it.value() : default_value);
    };

    writeProperty(QStringLiteral("Application"), QStringLiteral("Microsoft Excel"));
    writeProperty(QStringLiteral("DocSecurity"), QStringLiteral("0"));
    writeProperty(QStringLiteral("ScaleCrop"), QStringLiteral("false"));

    // Write other properties
    writeProperty(QStringLiteral("Manager"), QStringLiteral());
    writeProperty(QStringLiteral("Company"), QStringLiteral());

    // Write fixed properties
    writeProperty(QStringLiteral("LinksUpToDate"), QStringLiteral("false"));
    writeProperty(QStringLiteral("SharedDoc"), QStringLiteral("false"));
    writeProperty(QStringLiteral("HyperlinksChanged"), QStringLiteral("false"));
    writeProperty(QStringLiteral("AppVersion"), QStringLiteral("12.0000"));

    writer.writeStartElement(QStringLiteral("HeadingPairs"));
    writer.writeStartElement(vt, QStringLiteral("vector"));
    writer.writeAttribute(QStringLiteral("size"), QString::number(heading_list_.size() * 2));
    writer.writeAttribute(QStringLiteral("baseType"), QStringLiteral("variant"));

    for (const auto& pair : heading_list_) {
        writer.writeStartElement(vt, QStringLiteral("variant"));
        writer.writeTextElement(vt, QStringLiteral("lpstr"), pair.first);
        writer.writeEndElement(); // vt:variant
        writer.writeStartElement(vt, QStringLiteral("variant"));
        writer.writeTextElement(vt, QStringLiteral("i4"), QString::number(pair.second));
        writer.writeEndElement(); // vt:variant
    }
    writer.writeEndElement(); // vt:vector
    writer.writeEndElement(); // HeadingPairs

    writer.writeStartElement(QStringLiteral("TitlesOfParts"));
    writer.writeStartElement(vt, QStringLiteral("vector"));
    writer.writeAttribute(QStringLiteral("size"), QString::number(title_list_.size()));
    writer.writeAttribute(QStringLiteral("baseType"), QStringLiteral("lpstr"));
    for (const QString& title : title_list_)
        writer.writeTextElement(vt, QStringLiteral("lpstr"), title);
    writer.writeEndElement(); // vt:vector
    writer.writeEndElement(); // TitlesOfParts

    writer.writeEndElement(); // Properties
    writer.writeEndDocument();
}

bool DocPropsApp::ParseXml(QIODevice* device)
{
    QXmlStreamReader reader(device);

    while (!reader.atEnd() && !reader.hasError()) {
        if (!reader.readNextStartElement())
            continue;

        const auto name { reader.name() };

        if (name == QStringLiteral("Manager")) {
            SetProperty(QStringLiteral("manager"), reader.readElementText());
        } else if (name == QStringLiteral("Company")) {
            SetProperty(QStringLiteral("company"), reader.readElementText());
        } else if (name == QStringLiteral("Application")) {
            SetProperty(QStringLiteral("Application"), reader.readElementText());
        } else if (name == QStringLiteral("DocSecurity")) {
            SetProperty(QStringLiteral("DocSecurity"), reader.readElementText());
        } else if (name == QStringLiteral("ScaleCrop")) {
            SetProperty(QStringLiteral("ScaleCrop"), reader.readElementText());
        } else if (name == QStringLiteral("LinksUpToDate")) {
            SetProperty(QStringLiteral("LinksUpToDate"), reader.readElementText());
        } else if (name == QStringLiteral("SharedDoc")) {
            SetProperty(QStringLiteral("SharedDoc"), reader.readElementText());
        } else if (name == QStringLiteral("HyperlinksChanged")) {
            SetProperty(QStringLiteral("HyperlinksChanged"), reader.readElementText());
        } else if (name == QStringLiteral("AppVersion")) {
            SetProperty(QStringLiteral("AppVersion"), reader.readElementText());
        }
    }

    if (reader.hasError()) {
        qDebug() << "DocPropsApp Error reading doc props app file:" << reader.errorString();
        return false;
    }

    return true;
}

QT_END_NAMESPACE_YXLSX
