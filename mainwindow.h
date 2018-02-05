#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QtWidgets>
#include <QtXlsx>
#define ROW 0
#define COLUMN 1
#define ODD 0
#define EVEN 1

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void on_pushButton_ExcelSave_clicked();

private:
    Ui::MainWindow *ui;
    QSettings *Setting;
    QList<int> RandomNumberList;
    QList<char> RandomCharList;

    int IsRandomNumberCheck(int Number);
    char IsRandomAsciiCheck(char Alphabet);
};

#endif // MAINWINDOW_H
