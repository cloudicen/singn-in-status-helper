#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QFileDialog>
#include <QMessageBox>
#include <QCloseEvent>
#include <QtConcurrent>
#include "excelHelper.h"
#include "nameListReader.h"

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT
private:
    //basic paths
    QString excelPath="";
    QString txtPath="";
    QString saveasPath="";

    //basic object
    QStringList nameList={};
    excelHelper * excelFile=nullptr;

    //multi threads
    QFuture<bool> future;
    QFutureWatcher<bool> watcher;
    bool isRunning = false;

    //options:
    QString target_worksheet_name = "签到详情";
    QString match_target_header = "QQ昵称";
    QString sort_key_header = "学号";
    bool is_run_frontend = false;
    bool is_need_save_as = false;
    bool is_need_mark_out = true;
    bool do_not_save = false;

    //worksheet info:
    int rows = 0;
    int columns = 0;
    int startRow = 0;
    int startColumn = 0;
    int match_target_index = 0;
    int sort_key_index = 0;

    //work step
    enum step{
        step_open_excel_and_worksheet,
        step_open_txt_and_read,
        step_read_worksheet_info,
        step_do_worksheet_sort,
        step_do_signed_in_match,
        step_do_file_save
    };

    bool open_excel_and_worksheet();

    bool open_txt_and_read();

    bool read_worksheet_info();

    bool do_worksheet_sort();

    bool do_signed_in_match();

    bool do_file_save();

    enum step stepFlag;


    //other function
    void checkButton();
    void init();
    void setAllEnabled(bool isEnabled);
    void clean_and_exit();

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    void closeEvent(QCloseEvent *event);

private slots:
    void on_pushButton_excel_clicked();

    void on_pushButton_txt_clicked();

    void on_lineEdit_excel_textChanged(const QString &arg1);

    void on_lineEdit_txt_textChanged(const QString &arg1);

    void on_pushButton_exit_clicked();

    void on_pushButton_go_clicked();

    void handleFinished();

private:
    Ui::MainWindow *ui;

};
#endif // MAINWINDOW_H
