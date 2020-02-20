#include "excelHelper.h"

excelHelper::excelHelper(QString _fileName,QString _sheetName,bool _isVisible,QObject *parent)
    : QObject(parent)
{
    visibility = _isVisible;
    fileName = _fileName;
    sheetName = _sheetName;
}

excelHelper::~excelHelper()
{
    close_and_quit();
}

void excelHelper::openApp()
{
    excel = new QAxObject("Excel.Application",this);
    excel->setProperty("Visible", visibility);
}

bool excelHelper::openSheet()
{
    QAxObject *work_books = excel->querySubObject("WorkBooks");
    work_books->dynamicCall("Open (const QString&)", QString(fileName));
    work_book = excel->querySubObject("ActiveWorkBook");
    work_sheet = work_book->querySubObject("Worksheets(string)",sheetName);
    delete work_books;
    if(work_sheet == nullptr)
    {
        return false;
    }
    return true;
}

void excelHelper::setVisible(bool _isVisible)
{
    excel->setProperty("Visible", _isVisible);
}

QStringList& excelHelper::get_header_info()
{
        QAxObject *used_range = work_sheet->querySubObject("UsedRange");

        QAxObject *rows= used_range->querySubObject("Rows");
        QAxObject *columns= used_range->querySubObject("Columns");
        QVariant row_count = rows->property("Count");
        QVariant column_count = columns->property("Count");
        int row_start = used_range->property("Row").toInt();  //index of the first row of a "range"
        //int column_start = used_range->property("Column").toInt();
        //qDebug() << "rows: " << row_count.toInt();
        //qDebug() << "columns: " << column_count.toInt();
        //qDebug() << "row_start" << row_start;
        //qDebug() << "column_start" << column_start;
        QAxObject *header_raw = rows->querySubObject("Rows(int)",row_start);  //get a pointer to the first row
        QVariant var = header_raw->dynamicCall("Value");  //get data

        /* actually the 'var' objecct stores the data like this:
         * QVariant(QVariant(QVariant)))
         * the outer layers are lists
         * so, in order to convert 'var' to data type that can be handled , the "toList" function need to be call twice
         * first convert 'var' to QVariantList, each element represent a row in the sheet
         * then convert every element we got in last step to a QVariantList, which contains the data of each cell
         * after thease jobs we got a list like this:
         *
         * var =
         * datasheet(list){
         * row1(list){ cell1(data), cell2(data), ...},
         * row2(list){ cells(data), ......},
         * .
         * .
         * .
         * }
         *
         *
         */
        //deal with it!
        QVariantList varRows = var.toList();
        QStringList *header = new QStringList;
        header->append(row_count.toString());
        header->append(column_count.toString());
        header->append(used_range->property("Row").toString());
        header->append(used_range->property("Column").toString());
        foreach(QVariant val,varRows[0].toList())
        {
            header->append(val.toString());
        }
        return *header;
}

void excelHelper::save()
{
    work_book->dynamicCall("Save()");
}

void excelHelper::save(QString path)
{
    // '/' in path must transform to "\\"
    work_book->dynamicCall("SaveAs(const QString&)",path.replace("/","\\"));
}

void excelHelper::close(bool is_need_save)
{
    if(work_sheet != nullptr)
    {
        delete work_sheet;
        work_sheet = nullptr;
    }
    if(work_book != nullptr)
    {
        work_book->dynamicCall("Close(Boolean)",is_need_save);
        delete work_book;
        work_book = nullptr;
    }
}

void excelHelper::close_and_quit(bool is_need_save)
{
    close(is_need_save);
    if(excel != nullptr)
    {
        excel->dynamicCall("Quit()");
        delete excel;
        excel = nullptr;
    }
}

QList<QStringList> &excelHelper::get_range_data(int row_start, int column_start, int row_end, int column_end)
{
    QAxObject *range = work_sheet->querySubObject("Range(const QString&)",QString("%1%2:%3%4").arg(QChar(column_start-1+'A')).arg(row_start).arg(QChar(column_end-1+'A')).arg(row_end));
    QVariant var = range->dynamicCall("Value");  //get data
    QVariantList varRows = var.toList();
    QList<QStringList> *result = new QList<QStringList>;
    foreach(QVariant row,varRows)
    {
        QStringList rows_buf;
        foreach(QVariant cell,row.toList())
        {
            rows_buf.append(cell.toString());
        }
        result->append(rows_buf);
    }
     //we only need the data, so delete the pointer to the range to reduce memory usage;
    delete range;
    return *result;
}

QAxObject *excelHelper::loadRange(int row_start, int column_start, int row_end, int column_end)
{
    QAxObject *range = new QAxObject(this);
    range = work_sheet->querySubObject("Range(const QString&)",QString("%1%2:%3%4").arg(QChar(column_start-1+'A')).arg(row_start).arg(QChar(column_end-1+'A')).arg(row_end));
    return range;
}

QAxObject *excelHelper::loadRange()
{
    QAxObject *range = new QAxObject(this);
    range = work_sheet->querySubObject("UsedRange");
    return range;
}

QAxObject *excelHelper::get_cell_pointer(QAxObject *loadedRange, int row, int column)
{
    QAxObject * cell = new QAxObject(this);
    cell = loadedRange->querySubObject("Cells(int,int)",row,column);
    return cell;
}

QString excelHelper::get_cell_data(QAxObject *loadedRange, int row, int column)
{
    QAxObject * cell = new QAxObject(this);
    cell = loadedRange->querySubObject("Cells(int,int)",row,column);
    QString data = cell->dynamicCall("Value").toString();
    //we only need the data, so delete the pointer to the cell to reduce memory usage;
    delete cell;
    return data;
}

QAxObject *excelHelper::getFront(QAxObject *loadedRange)
{
    QAxObject *front = new QAxObject;
    front = loadedRange->querySubObject("Font");
    return front;
}

void excelHelper::setValue(QAxObject *dst, QVariant value)
{
    dst->setProperty("Value",value);
}

void excelHelper::setProperty(QAxObject *dst,char* Property, QVariant value)
{
    dst->setProperty(Property,value);
}

void excelHelper::sort(QAxObject* sort_range,QAxObject* key1)
{
    sort_range->dynamicCall("Sort(QAxObject*)",key1->asVariant());
}




