#ifndef EXCELHELPER_H
#define EXCELHELPER_H

#include <QObject>
#include <ActiveQt/QAxObject>
#include <QDebug>
#include <QScopedPointer>

// 特别注意！！excel的行和列下标是从1开始的，对结果的List使用内置的indexOF（）时注意转换！！

/*  一个二次封装的用于处理excel文档的库
 *  基本的单元格操作和QAxObject的特性差不多搞懂了点
 *  用到setProperty的操作，具体属性还是要去翻翻vbx的文档，这里只用了一小部分属性
 *  库还有很多不完善的地方，例如多数据表，多文档管理
 *  添加和删除表还是得调用dynamicCall，等待进一步封装
 *  vbx wit qt的资料真的不是很全。。微软官网也没翻到，改天再说
 *
 * -by cloudicen 2020.2.19
 */

class excelHelper:public QObject
{
Q_OBJECT
private:
    QAxObject* excel;  //excel instance
    QAxObject* work_book;//workbook instance--the .xslx file
    QAxObject* work_sheet;//data sheet that currently being edit
    QString fileName;//store the .xslx file name
    QString sheetName;//store the sheet name
    bool visibility;
public:
    //pass on the file name and the sheet name,and excel's visibility
    //constructor will open excel app,then you need run openSheet() manually, to open the workbook and data sheet.
    //more work needs to be done to support features such as following:

    /*  - open multiple workbooks
     *  - open multiple worksheets
     *  - no need to pass file name and sheet name at first, workbooks and worksheets can be opend later using other functions
     */

    excelHelper(QString _fileName,QString _sheetName,bool _isVisible = false,QObject* paerent = nullptr);
    excelHelper(const excelHelper &) = delete;
    ~excelHelper();//will close the opened workbook and exit excel
public:
    bool openSheet();//open the given file and sheet
    void setVisible(bool _isVisible);//set the visibility of excel app

    QStringList& get_header_info();//it will be useful if there are miningful headers in the first row.
    //return data like this: "num of rows","num of columns","header1","header2","..."

    void save();//save
    void save(QString path);//save as
    void close(bool is_need_save = false);
    void close_and_quit(bool is_need_save = false);//quit with/without saving


    //note that a 'range' object in excel can be a single cell, or a combination of multiple cells
    //the entire sheet is also a 'range' object
    QList<QStringList> & get_range_data(int row_start,int column_start,int row_end,int column_end);//return data from a range
    QAxObject* loadRange(int row_start,int column_start,int row_end,int column_end);//load given range from file
    QAxObject* loadRange();//load all used cells

    QAxObject* get_cell_pointer(QAxObject *loadedRange,int row,int column);//from loaded range,select a single cell
    QString get_cell_data(QAxObject *loadedRange,int row,int column);//from loaded range,get a single cell's data
    QAxObject* getFront(QAxObject *loadedRange);//from range loaded,get the range's front

    static void setValue(QAxObject * dst,QVariant value);//set range's value
    static void setProperty(QAxObject * dst,char * Property,QVariant value);//set range's property
    void sort(QAxObject* sort_range,QAxObject* key1);

    //for multi threads
public slots:
    void openApp();

};

#endif // EXCELHELPER_H
