#include "nameListReader.h"

const QStringList nameListReader::readNameList(const QString& Path)
{
    QStringList nameList;
    if(!Path.isEmpty())
    {
        QFile file(Path);
        if (file.open(QIODevice::ReadOnly | QIODevice::Text))
        {
            while (!file.atEnd())
            {
                QByteArray line = file.readLine().trimmed();
                nameList.append(QString(line));
            }
            file.close();
        }
    }
    return nameList;
}
