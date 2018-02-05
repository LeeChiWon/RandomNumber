#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    Setting=new QSettings("EachOne","RandomNumber",this);
    ui->lineEdit_LineNumber->setText(Setting->value("LineNumber").toString());
    ui->lineEdit_RangeHigh->setText(Setting->value("RangeHigh").toString());
    ui->lineEdit_RangeLow->setText(Setting->value("RangeLow").toString());
    ui->lineEdit_Number->setText(Setting->value("Number").toString());
    ui->comboBox_Type->addItems(QStringList()<<tr("Row")<<tr("Column"));
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_pushButton_ExcelSave_clicked()
{
    QString FileName = QFileDialog::getSaveFileName(this,tr("Save File"),Setting->value("FilePath").toString(),tr("Excel File (*.xlsx)"));

    QFile File(FileName);
    if(File.exists())
    {
        File.remove();
    }
    QXlsx::Document xlsx(FileName);   
    qsrand(QTime::currentTime().msecsSinceStartOfDay());

    for(int i=1; i<=ui->lineEdit_LineNumber->text().toInt(); i++)
    {
        for(int j=1; j<=ui->lineEdit_Number->text().toInt(); j++)
        {
            switch(ui->comboBox_Type->currentIndex())
            {
            case ROW:
                xlsx.write(i,j,IsRandomNumberCheck(qrand()%ui->lineEdit_RangeHigh->text().toInt()+ui->lineEdit_RangeLow->text().toInt()));
                break;
            case COLUMN:
                xlsx.write(j,i,IsRandomNumberCheck(qrand()%ui->lineEdit_RangeHigh->text().toInt()+ui->lineEdit_RangeLow->text().toInt()));
                break;
            }
        }
        RandomList.clear();
    }

    xlsx.saveAs(FileName);
    xlsx.deleteLater();

    Setting->setValue("LineNumber",ui->lineEdit_LineNumber->text());
    Setting->setValue("RangeHigh",ui->lineEdit_RangeHigh->text());
    Setting->setValue("RangeLow",ui->lineEdit_RangeLow->text());
    Setting->setValue("Number",ui->lineEdit_Number->text());
    Setting->setValue("FilePath",FileName);
    QMessageBox::information(this,"Success","Excel Saved.",QMessageBox::Ok);
}

int MainWindow::IsRandomNumberCheck(int Number)
{
    while(1)
    {
        if(!RandomList.contains(Number))
        {
            RandomList.append(Number);
            break;
        }
        else
        {
            Number=qrand()%ui->lineEdit_RangeHigh->text().toInt()+ui->lineEdit_RangeLow->text().toInt();
        }
    }
    return Number;
}
