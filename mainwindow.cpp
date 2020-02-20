#include "mainwindow.h"
#include "ui_mainwindow.h"

void MainWindow::checkButton()
{
    if(!excelPath.isEmpty() && !txtPath.isEmpty())
    {
        ui->pushButton_go->setEnabled(true);
    }
    else
    {
        ui->pushButton_go->setEnabled(false);
    }
}

void MainWindow::init()
{
    setAllEnabled(true);

    ui->lineEdit_txt->setText("");
    ui->lineEdit_excel->setText("");
    ui->pushButton_go->setText("走你！！");

    if(excelFile!=nullptr)
    {
        delete excelFile;
        excelFile = nullptr;
    }

    QString excelPath = "";
    QString txtPath = "";
    QString saveasPath = "";
    ui->pushButton_go->setEnabled(false);
    isRunning = false;
}

void MainWindow::setAllEnabled(bool isEnabled)
{
    ui->pushButton_go->setEnabled(isEnabled);
    ui->pushButton_exit->setEnabled(isEnabled);
    ui->pushButton_txt->setEnabled(isEnabled);
    ui->pushButton_excel->setEnabled(isEnabled);
    ui->checkBox_mark->setEnabled(isEnabled);
    ui->checkBox_saveAs->setEnabled(isEnabled);
    ui->checkBox_runFrontend->setEnabled(isEnabled);
}

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ui->pushButton_go->setEnabled(false);
    connect(&watcher, SIGNAL(finished()), this, SLOT(handleFinished()));
}

MainWindow::~MainWindow()
{
    delete ui;
}

bool MainWindow::open_excel_and_worksheet()
{
    stepFlag = step_open_excel_and_worksheet;
    excelFile = new excelHelper(excelPath,target_worksheet_name,is_run_frontend);
    excelFile->openApp();

    return excelFile->openSheet();
}

bool MainWindow::open_txt_and_read()
{
    stepFlag = step_open_txt_and_read;
    nameList = nameListReader::readNameList(txtPath);

    return !nameList.isEmpty();
}

bool MainWindow::read_worksheet_info()
{
    stepFlag = step_read_worksheet_info;
    QStringList info = excelFile->get_header_info();
    rows = info.front().toInt();
    info.pop_front();
    columns = info.front().toInt();
    info.pop_front();
    startRow = info.front().toInt();
    info.pop_front();
    startColumn = info.front().toInt();
    info.pop_front();
    match_target_index = info.indexOf(match_target_header);
    sort_key_index = info.indexOf(sort_key_header);

    return (sort_key_index != -1 && match_target_index != -1);
}

bool MainWindow::do_worksheet_sort()
{
    stepFlag = step_do_worksheet_sort;
    QAxObject* sort_range = excelFile->loadRange(startRow+1,startColumn,rows,columns);
    QAxObject* key_range = excelFile->loadRange(startRow+1,sort_key_index+1,rows,sort_key_index+1);
    excelFile->sort(sort_range,key_range);

    //delete manually to reduce memory usage
    delete sort_range;
    delete key_range;
    return true;
}

bool MainWindow::do_signed_in_match()
{
    stepFlag = step_do_signed_in_match;
    //load range first to reduce IO
    QAxObject* range = excelFile->loadRange(startRow+1,match_target_index+1,rows,match_target_index+1);
    QAxObject* newcol = excelFile->loadRange(startRow,columns+1,rows,columns+1);

    excelFile->setValue(excelFile->get_cell_pointer(newcol,1,1),"签到情况");

    for (int i = 1;i<=rows-startRow;i++) {
        QAxObject* currentCell = excelFile->get_cell_pointer(newcol,i+1,1);
        if(!nameList.isEmpty() && nameList.removeOne(excelFile->get_cell_data(range,i,1)))
        {
            excelFile->setValue(currentCell,"已签到");
        }
        else
        {
            excelFile->setValue(currentCell,"未签到");
            if(is_need_mark_out)
            {
                excelFile->setProperty(excelFile->getFront(currentCell),(char*)"Color",QColor(255,0,0));
            }
        }
        //delete manually to reduce memory usage
        delete currentCell;
    }
    delete range;
    delete newcol;
    return true;
}

bool MainWindow::do_file_save()
{
    stepFlag = step_do_file_save;
    if(do_not_save)
    {
        excelFile->close_and_quit(false);
        return false;
    }
    else
    {
        if(is_need_save_as)
        {
            excelFile->save(saveasPath);
            excelFile->close_and_quit();
        }
        else
        {
            excelFile->close_and_quit(true);
        }
        return true;
    }
}
void MainWindow::clean_and_exit()
{
    if(excelFile!=nullptr)
    {
        delete excelFile;
        excelFile = nullptr;
    }
    close();
}

void MainWindow::on_pushButton_excel_clicked()
{
    QString fileName = QFileDialog::getOpenFileName(this,QString("打开excel表格"),"","excel工作簿 (*.xlsx *.lsx)");
    if(!fileName.isEmpty())
    {
        excelPath = fileName;
    }
    ui->lineEdit_excel->setText(excelPath);
}

void MainWindow::on_pushButton_txt_clicked()
{
    QString fileName = QFileDialog::getOpenFileName(this,QString("打开签到名单"),"","txt文本文档 (*.txt)");
    if(!fileName.isEmpty())
    {
        txtPath = fileName;
    }
    ui->lineEdit_txt->setText(txtPath);
}

void MainWindow::on_lineEdit_excel_textChanged(const QString &arg1)
{
    Q_UNUSED(arg1);
    checkButton();
}

void MainWindow::on_lineEdit_txt_textChanged(const QString &arg1)
{
    Q_UNUSED(arg1);
    checkButton();
}

void MainWindow::on_pushButton_exit_clicked()
{
    clean_and_exit();
}

void MainWindow::on_pushButton_go_clicked()
{
    isRunning = true;
    is_need_save_as = ui->checkBox_saveAs->isChecked();
    is_run_frontend = ui->checkBox_runFrontend->isChecked();
    is_need_mark_out = ui->checkBox_mark->isChecked();
    ui->pushButton_go->setText("正在打开excel...");
    setAllEnabled(false);
    future = QtConcurrent::run(this,&MainWindow::open_excel_and_worksheet);
    watcher.setFuture(future);
}

void MainWindow::handleFinished()
{
    switch (stepFlag) {
    case step_open_excel_and_worksheet:
    {
        if(future.result())
        {
            ui->pushButton_go->setText("正在读取txt名单...");
            future = QtConcurrent::run(this,&MainWindow::open_txt_and_read);
            watcher.setFuture(future);
        }
        else
        {
            QMessageBox::warning(this, "警告", "文档读取失败！\n文档格式或内容错误！", QMessageBox::Ok, QMessageBox::NoButton);
            init();
        }
        break;
    }
    case step_open_txt_and_read:
    {
        if(future.result())
        {
            ui->pushButton_go->setText("正在读取工作簿...");
            future = QtConcurrent::run(this,&MainWindow::read_worksheet_info);
            watcher.setFuture(future);
        }
        else
        {
            QMessageBox::warning(this, "警告", "文档读取失败！\n文档格式错误或内容为空！", QMessageBox::Ok, QMessageBox::NoButton);
            init();
        }
        break;
    }
    case step_read_worksheet_info:
    {
        if(future.result())
        {
            ui->pushButton_go->setText("执行排序...");
            future = QtConcurrent::run(this,&MainWindow::do_worksheet_sort);
            watcher.setFuture(future);
        }
        else
        {
            QMessageBox::warning(this, "警告", "文档读取失败！\n文档格式或内容错误！", QMessageBox::Ok, QMessageBox::NoButton);
            init();
        }
        break;
    }
    case step_do_worksheet_sort:
    {
        ui->pushButton_go->setText("执行匹配...");
        future = QtConcurrent::run(this,&MainWindow::do_signed_in_match);
        watcher.setFuture(future);
        break;
    }
    case step_do_signed_in_match:
    {
        ui->pushButton_go->setText("操作完成,等待保存...");
        if(is_need_save_as)
        {
            while(true)
            {
                saveasPath = QFileDialog::getSaveFileName(this,QString("操作完成，选择保存位置"),"","excel工作簿 (*.xlsx *.lsx)");
                if(!saveasPath.isEmpty())
                {
                    future = QtConcurrent::run(this,&MainWindow::do_file_save);
                    watcher.setFuture(future);
                    break;
                }
                else
                {
                    if(QMessageBox::warning(this, "警告", "取消保存则所有更改都不会生效\n是否继续？", QMessageBox::Ok, QMessageBox::No) == QMessageBox::Ok)
                    {
                        do_not_save = true;
                        future = QtConcurrent::run(this,&MainWindow::do_file_save);
                        watcher.setFuture(future);
                        break;
                    }
                }
            }
        }
        else
        {
            future = QtConcurrent::run(this,&MainWindow::do_file_save);
            watcher.setFuture(future);
        }
        break;
    }
    case step_do_file_save:
    {
        if(future.result() == true)
        {
            QMessageBox::warning(this, "提示", "所有操作完成", QMessageBox::Ok, QMessageBox::NoButton);
        }
        else
        {
            QMessageBox::warning(this, "提示", "操作已取消", QMessageBox::Ok, QMessageBox::NoButton);
        }
        init();
    }
    }
}

void MainWindow::closeEvent(QCloseEvent *event)
{
    if(isRunning)
    {
        QMessageBox::StandardButton button;
        button = QMessageBox::question(this, QString("退出程序"),QString("警告：程序有一个任务正在运行中，是否强行退出?"),QMessageBox::Yes | QMessageBox::No);
        if (button == QMessageBox::No)
        {
            event->ignore();  //忽略退出信号，程序继续运行
        }
        else if (button == QMessageBox::Yes)
        {
            event->accept();  //接受退出信号，程序退出
        }
    }
    else
    {
        event->accept();
    }

}
