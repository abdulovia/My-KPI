#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QStandardPaths>
#include <QCoreApplication>
#include <QTableWidget>
#include <QListWidget>
#include <QTextEdit>
#include <QComboBox>
#include <QPrinter>

QT_BEGIN_NAMESPACE
namespace Ui {
class MainWindow;
}
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    void importDepList(QListWidget *list, QString &path);
    void importDepComboBox(QComboBox *box, QString &path);
    void searchButtonClicked();
    void importCSV(QTableWidget *tw, QString &path);
    void saveCSV(QTableWidget *tw, QString &path);
    void importTxt(QTextEdit *te, QString &path);
    void saveTxt(QTextEdit *te, QString &path);
    void calculateCells(QTableWidgetItem *item);
    void insertCalcItem(QString val, QString points, int row, int col);
    void refreshTableWidget();
    void boldHeaders();
    void filterByDep();
    void saveToPdf(QTableWidget *tableWidget);
    void saveToHtml(QTableWidget *tableWidget);


private:
    Ui::MainWindow *ui;
    QString PATH = QCoreApplication::applicationDirPath();
};
#endif // MAINWINDOW_H
