#ifndef NAMELISTREADER_H
#define NAMELISTREADER_H

#include <QFile>
#include <QDebug>

class nameListReader
{
public:
    nameListReader() = delete;
    nameListReader(const nameListReader &) = delete;
    ~nameListReader() = delete;
    static const QStringList readNameList(const QString& path);
};

#endif // NAMELISTREADER_H
