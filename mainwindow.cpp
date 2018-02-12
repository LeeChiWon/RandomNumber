#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    Setting=new QSettings("EachOne","RandomNumber",this);
    ui->lineEdit_LineNumber->setText(Setting->value("LineNumber").toString());
    ui->lineEdit_Number->setText(Setting->value("Number").toString());
    ui->lineEdit_CellNumber->setText(Setting->value("CellNumber").toString());
    ui->comboBox_Type->addItems(QStringList()<<tr("Row")<<tr("Column"));
    ui->comboBox_Type->setCurrentIndex(Setting->value("ComboType").toInt());
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_pushButton_ExcelSave_clicked()
{
    QString FileName = QFileDialog::getSaveFileName(this,tr("Save File"),Setting->value("FilePath").toString(),tr("Excel File (*.xlsx)"));
    QString Str;
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
            for(int k=0; k<ui->lineEdit_CellNumber->text().toInt(); k++)
            {
                switch((qrand()%100+0)%2)
                {
                case ODD:
                    Str.append(QString::number(IsRandomNumberCheck(qrand()%9+0)));
                    break;
                case EVEN:
                    Str.append(IsRandomAsciiCheck(qrand()%26+65));
                    break;
                }
            }

            switch(ui->comboBox_Type->currentIndex())
            {
            case ROW:
                xlsx.write(i,j,Str);
                break;
            case COLUMN:
                xlsx.write(j,i,Str);
                break;
            }
            Str.clear();
            RandomNumberList.clear();
            RandomCharList.clear();
        }
    }

    xlsx.saveAs(FileName);
    xlsx.deleteLater();

    Setting->setValue("LineNumber",ui->lineEdit_LineNumber->text());
    Setting->setValue("Number",ui->lineEdit_Number->text());
    Setting->setValue("CellNumber",ui->lineEdit_CellNumber->text());
    Setting->setValue("FilePath",FileName);
    Setting->setValue("ComboType",ui->comboBox_Type->currentIndex());
    QMessageBox::information(this,"Success","Excel Saved.",QMessageBox::Ok);
}

int MainWindow::IsRandomNumberCheck(int Number)
{
    while(1)
    {
        if(!RandomNumberList.contains(Number))
        {
            RandomNumberList.append(Number);
            //qDebug()<<Number;
            break;
        }
        else
        {
            Number=qrand()%9+0;
        }
    }
    return Number;
}

char MainWindow::IsRandomAsciiCheck(char Alphabet)
{
    while(1)
    {
        if(!RandomCharList.contains(Alphabet))
        {
            RandomCharList.append(Alphabet);
            //qDebug()<<Alphabet;
            break;
        }
        else
        {
            Alphabet=qrand()%26+65;
        }
    }
    return Alphabet;
}
