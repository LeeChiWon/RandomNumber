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
    QList<QString> RandomList;
    QString FileName = QFileDialog::getSaveFileName(this,tr("Save File"),Setting->value("FilePath").toString(),tr("Excel File (*.xlsx)"));
    QString Str;
    if(FileName.isEmpty())
    {
        return;
    }
    QFile File(FileName);
    QChar Test;
    QProgressDialog *ProgressDlg=new QProgressDialog(tr("Waiting..."),NULL,0,100,this);
    ProgressDlg->setWindowTitle(tr("RandomNumber"));
    ProgressDlg->setWindowModality(Qt::WindowModal);
    ProgressDlg->show();

    int Count=1,PreCount=0,Total=ui->lineEdit_LineNumber->text().toInt();

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
            while(1)
            {
                for(int k=0; k<ui->lineEdit_CellNumber->text().toInt(); k++)
                {
                    if((qrand()%100+0)%2==1 ||(qrand()%100+0)%2==2 || (qrand()%100+0)%2==3)
                    {
                        //Str.append(QString::number(IsRandomNumberCheck(qrand()%9+0)));
                        Str.append(QString::number(qrand()%9+0));
                    }
                    else
                    {
                        //Str.append(IsRandomAsciiCheck(qrand()%26+65));
                        Test=qrand()%26+65;
                        Str.append(Test);
                    }
                    /*switch((qrand()%100+0)%2)
                    {
                    case ODD:
                        Str.append(QString::number(IsRandomNumberCheck(qrand()%9+0)));
                        break;
                    case EVEN:
                        Str.append(IsRandomAsciiCheck(qrand()%26+65));
                        break;
                    }*/
                }

                if(!RandomList.contains(Str))
                {
                    RandomList.append(Str);
                    break;
                }
                else
                {
                    Str.clear();
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
           // qDebug()<<Total<<Count<<PreCount<<(float)Count/Total*100;
            //if(Count>0 && PreCount!=(float)Count/Total*100)
           // {
                PreCount=(float)Count/Total*100;
                ProgressDlg->setValue(PreCount);
            //}
            Count++;
        }        
    }
    xlsx.saveAs(FileName);
    xlsx.deleteLater();
    ProgressDlg->deleteLater();

    Setting->setValue("LineNumber",ui->lineEdit_LineNumber->text());
    Setting->setValue("Number",ui->lineEdit_Number->text());
    Setting->setValue("CellNumber",ui->lineEdit_CellNumber->text());
    Setting->setValue("FilePath",FileName);
    Setting->setValue("ComboType",ui->comboBox_Type->currentIndex());

    QMessageBox::information(this,tr("information"),tr("Excel Saved."),QMessageBox::Ok);
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
