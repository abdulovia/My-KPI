#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFile>
#include <QMessageBox>
#include <QFileInfo>
#include <QFileDialog>
#include <QDesktopServices>
#include <QPainter>

QMap<QString, QString> departmentMap, monthMap = {{"январь", "01"},{"февраль", "02"}, {"март", "03"}, {"апрель", "04"},
                                                  {"май", "05"}, {"июнь", "06"}, {"июль", "07"}, {"август", "08"},
                                                  {"сентябрь", "09"}, {"октябрь", "10"}, {"ноябрь", "11"}, {"декабрь", "12"}};

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    ui->rightMenu->setHidden(true);

    QString depsPath = PATH + "/СписокУчреждений.txt";
    importDepList(ui->departmentListWidget, depsPath);
    QString istPath = PATH + "/СписокИсточников.txt";
    importDepComboBox(ui->comboDepBox, istPath);

    ui->stackedWidget->setCurrentIndex(0);

    connect(ui->searchBtn, &QPushButton::clicked, this, &MainWindow::searchButtonClicked);
    connect(ui->txtSearchLine, &QLineEdit::returnPressed, this, &MainWindow::searchButtonClicked);

    connect(ui->departmentListWidget, &QListWidget::itemClicked, this, [=](QListWidgetItem* item) {
        ui->stackedWidget->setCurrentIndex(1);
        ui->txtDepartment->setText(departmentMap[item->text()]);
        refreshTableWidget();
    });
    connect(ui->comboDepBox, &QComboBox::currentIndexChanged, this, [=]() {
        refreshTableWidget();
    });
    connect(ui->comboMonthBox, &QComboBox::currentIndexChanged, this, [=]() {
        refreshTableWidget();
    });
    connect(ui->yearLineEdit, &QLineEdit::textChanged, this, [=]() {
        refreshTableWidget();
    });

    connect(ui->homeBtn, &QPushButton::clicked, this, [=]() {
        ui->stackedWidget->setCurrentIndex(0);
        ui->departmentListWidget->clearSelection();
        ui->leftMenu->setVisible(true);
        ui->leftMenuBtn->setChecked(false);
    });

    connect(ui->tableWidget, &QTableWidget::itemChanged, this, [=](QTableWidgetItem *item) {
        ui->tableWidget->blockSignals(true);
        calculateCells(item);
        ui->tableWidget->blockSignals(false);
        QString department = ui->departmentListWidget->currentItem()->text();
        QString appendix = "/" + ui->yearLineEdit->text() + "-" + monthMap[ui->comboMonthBox->currentText()] + "_" + department + ".csv";
        qDebug() << appendix << "\n";
        QString file_path = PATH + appendix;
        saveCSV(ui->tableWidget, file_path);
    });

    connect(ui->downloadBtn, &QPushButton::clicked, this, [=]() {
        QString department = ui->departmentListWidget->currentItem()->text();
        QString fileName = ui->yearLineEdit->text() + "-" + monthMap[ui->comboMonthBox->currentText()] + "_" + department;
        QString destinationFilePath = QFileDialog::getSaveFileName(this, tr("Save File"), fileName, tr("Text Files (*.csv);"));
        if (!destinationFilePath.isEmpty()) {
            QString appendix = "/" + ui->yearLineEdit->text() + "-" + monthMap[ui->comboMonthBox->currentText()] + "_" + department + ".csv";
            QString sourceFilePath = PATH + appendix;
            saveCSV(ui->tableWidget, sourceFilePath);
            QFile sourceFile(sourceFilePath);
            if (sourceFile.exists() && QFile::copy(sourceFilePath, destinationFilePath)) {
                QMessageBox::information(nullptr, QObject::tr("Success"), QObject::tr("File copied successfully."));
            }
        }
    });
    connect(ui->exportSettingsBtn, &QPushButton::clicked, this, [=]() {
        QString destDirPath = QFileDialog::getExistingDirectory(nullptr, "Выберите папку для экспорта");
        if (destDirPath.isEmpty()) {
            QMessageBox::information(nullptr, QObject::tr("Ошибка"), QObject::tr("Такого пути не существует"));
            return;
        }
        QString newDirName = "Настройки_Мой_KPI";
        QDir destDir(destDirPath);
        if (!destDir.exists(newDirName)) {
            if (!destDir.mkdir(newDirName)) {
                QMessageBox::information(nullptr, QObject::tr("Ошибка"), QObject::tr("Папка Настройки_Мой_KPI уже существует"));
                return;
            }
        }
        destDirPath = destDir.filePath(newDirName);
        QStringList sourceFilePaths = {"/СписокУчреждений.txt", "/СписокИсточников.txt", "/Шаблон.csv"};
        foreach (QString sourceFilePath, sourceFilePaths) {
            QString sourceFileName = QFileInfo(sourceFilePath).fileName();
            QString destFilePath = QDir(destDirPath).filePath(sourceFileName);
            if (QFile::exists(destFilePath))
                QFile::remove(destFilePath);
            QString sourceFileFullPath = PATH + sourceFilePath;
            if (QFile::copy(sourceFileFullPath, destFilePath))
                qDebug() << "File" << sourceFileName << "exported successfully.";
            else
                QMessageBox::information(nullptr, QObject::tr("Ошибка"), sourceFileName + " уже существует. Удалите имеющийся для экспорта.");
        }
    });
    connect(ui->importSettingsBtn, &QPushButton::clicked, this, [=]() {
        QString sourceDirPath = QFileDialog::getExistingDirectory(nullptr, "Выберите папку с файлами для импорта");
        if (sourceDirPath.isEmpty()) {
            QMessageBox::information(nullptr, QObject::tr("Ошибка"), "Не выбрана папка для импорта.");
            return;
        }
        QDir sourceDir(sourceDirPath), destDir(PATH);
        QStringList files = sourceDir.entryList(QDir::Files);
        foreach (QString fileName, files) {
            QString sourceFilePath = sourceDir.filePath(fileName);
            QString destFilePath = destDir.filePath(fileName);
            if (QFile::exists(destFilePath))
                QFile::remove(destFilePath);
            if (QFile::copy(sourceFilePath, destFilePath))
                qDebug() << "Файл " + fileName + " успешно импортирован.";
            else
                QMessageBox::information(nullptr, QObject::tr("Ошибка"), "Не удалось импортировать файл " + fileName);
        }
        QCoreApplication::quit();
    });
    connect(ui->resumeBtn, &QPushButton::toggled, this, [=]() {
        ui->tableWidget->blockSignals(true);
        if (ui->resumeBtn->isChecked()) {
            for (int col = 2; col < 7; col++)
                ui->tableWidget->setColumnHidden(col, true);
        } else {
            for (int col = 2; col < 7; col++)
                ui->tableWidget->setColumnHidden(col, false);
        }
        ui->tableWidget->blockSignals(false);
    });
    connect(ui->printBtn, &QPushButton::clicked, this, [=]() {
        saveToPdf(ui->tableWidget);
    });
    connect(ui->htmlBtn, &QPushButton::clicked, this, [=]() {
        saveToHtml(ui->tableWidget);
    });

    connect(ui->docsBtn, &QPushButton::clicked, this, [=]() {
        ui->stackedWidget->setCurrentIndex(3);
        ui->docsListWidget->clear();
        QDir docsDir(QCoreApplication::applicationDirPath());
        QStringList files = docsDir.entryList(QStringList() << "*.csv", QDir::Files);
        if (files.size() > 0 && files[files.size()-1] == "Шаблон.csv")
            files.pop_back();
        std::reverse(files.begin(), files.end());
        ui->docsListWidget->addItems(files);
    });
    connect(ui->docsListWidget, &QListWidget::itemClicked, this, [=](QListWidgetItem* item) {
        QStringList data = item->text().replace(".csv", "").split("_");
        QString year = data[0].split("-")[0], month = data[0].split("-")[1];
        for (int row = 0; row < ui->departmentListWidget->count(); row++) {
            QListWidgetItem *item2 = ui->departmentListWidget->item(row);
            if (item2->text() == data[1]) {
                ui->departmentListWidget->setCurrentRow(row);
                ui->stackedWidget->setCurrentIndex(1);
                ui->txtDepartment->setText(departmentMap[data[1]]);
                break;
            }
        }
        ui->yearLineEdit->setText(year);
        ui->comboMonthBox->setCurrentIndex(month.toInt()-1);
        ui->comboDepBox->setCurrentIndex(0);
        QString filepath = PATH + "/" + item->text();
        importCSV(ui->tableWidget, filepath);
        ui->tableWidget->blockSignals(true);
        boldHeaders();
        filterByDep();
        ui->tableWidget->blockSignals(false);
    });
    connect(ui->settingsBtn_1, &QPushButton::clicked, this, [=]() {
        ui->stackedWidget->setCurrentIndex(4);
        ui->leftMenu->setHidden(true);
        ui->leftMenuBtn->setChecked(true);
        QString templatePath = PATH + "/Шаблон.csv";
        importCSV(ui->tableWidget_2, templatePath);
    });
    connect(ui->addCol, &QPushButton::clicked, this, [=]() {
        if (ui->tableWidget_2->currentColumn() != -1)
            ui->tableWidget_2->insertColumn(ui->tableWidget_2->currentColumn() + 1);
        else
            ui->tableWidget_2->insertColumn(ui->tableWidget_2->columnCount());
    });
    connect(ui->addRow, &QPushButton::clicked, this, [=]() {
        if (ui->tableWidget_2->currentRow() != -1)
            ui->tableWidget_2->insertRow(ui->tableWidget_2->currentRow() + 1);
        else
            ui->tableWidget_2->insertRow(ui->tableWidget_2->rowCount());
    });
    connect(ui->rmCol, &QPushButton::clicked, this, [=]() {
        if (ui->tableWidget_2->currentColumn() != -1)
            ui->tableWidget_2->removeColumn(ui->tableWidget_2->currentColumn());
        else
            ui->tableWidget_2->removeColumn(ui->tableWidget_2->columnCount() - 1);
    });
    connect(ui->rmRow, &QPushButton::clicked, this, [=]() {
        if (ui->tableWidget_2->currentRow() != -1)
            ui->tableWidget_2->removeRow(ui->tableWidget_2->currentRow());
        else
            ui->tableWidget_2->removeRow(ui->tableWidget_2->rowCount() - 1);
    });
    connect(ui->saveBtn_2, &QPushButton::clicked, this, [=]() {
        ui->tableWidget_2->setCurrentItem(nullptr);
        QString templatePath = PATH + "/Шаблон.csv";
        saveCSV(ui->tableWidget_2, templatePath);
    });
    connect(ui->discardBtn_2, &QPushButton::clicked, this, [=]() {
        ui->tableWidget_2->setCurrentItem(nullptr);
        QString templatePath = PATH + "/Шаблон.csv";
        importCSV(ui->tableWidget_2, templatePath);
    });
    connect(ui->settingsBtn_2, &QPushButton::clicked, this, [=]() {
        ui->stackedWidget->setCurrentIndex(5);
        QString path1 = PATH + "/СписокУчреждений.txt";
        importTxt(ui->depTextEdit, path1);
        path1 = PATH + "/СписокИсточников.txt";
        importTxt(ui->depTextEdit_2, path1);
    });
    connect(ui->saveBtn_3, &QPushButton::clicked, this, [=]() {
        QString path1 = PATH + "/СписокУчреждений.txt";
        saveTxt(ui->depTextEdit, path1);
        QCoreApplication::quit();
    });
    connect(ui->saveBtn_4, &QPushButton::clicked, this, [=]() {
        QString path1 = PATH + "/СписокИсточников.txt";
        saveTxt(ui->depTextEdit_2, path1);
        QCoreApplication::quit();
    });
    connect(ui->settingsBtn, &QPushButton::clicked, this, [=]() {
        ui->stackedWidget->setCurrentIndex(2);
    });
    connect(ui->helpBtn, &QPushButton::clicked, this, [=]() {
        QString filePath = PATH + "/Порядок.docx";
        if (!QDesktopServices::openUrl(QUrl::fromLocalFile(filePath))) {
            QMessageBox::warning(this, tr("Warning"), tr("Could not open the file."));
        }
    });
    connect(ui->aboutBtn, &QPushButton::clicked, this, [=]() {
        QMessageBox::information(this, tr("Информация о приложении"),
                                 tr("Единая Информационная Система Мой KPI является инструментом для внесения данных КПЭ по примеру таблиц Excel. "
                                    "Была разработана в 2024г. в Проектном Офисе АГ СПб."));
    });
    connect(ui->exitBtn, &QPushButton::clicked, this, [=]() {
        QCoreApplication::quit();
    });
}

MainWindow::~MainWindow() {
    delete ui;
}

void MainWindow::importDepList(QListWidget *list, QString &path) {
    QFile file(path);
    if(!file.open(QIODevice::ReadWrite)) {
        QMessageBox::information(0, "error!", file.errorString());
    }
    QTextStream* in = new QTextStream(&file);
    while(!in->atEnd()) {
        QStringList department = in->readLine().split(";");
        if (department.size() == 2) {
            QListWidgetItem* item = new QListWidgetItem(department[0], ui->departmentListWidget);
            list->addItem(item);
            departmentMap[department[0]] = department[1];
        } else {
            QMessageBox::information(0, "Error reading file!", "Файл СписокУчреждений.txt с ошибками!");
            break;
        }
    }
    delete in;
    file.close();
}

void MainWindow::importDepComboBox(QComboBox *box, QString &path) {
    QFile file(path);
    if(!file.open(QIODevice::ReadWrite)) {
        QMessageBox::information(0, "error!", file.errorString());
    }
    QTextStream* in = new QTextStream(&file);
    while(!in->atEnd()) {
        QString dep = in->readLine();
        box->addItem(dep);
    }
    delete in;
    file.close();
}

void MainWindow::searchButtonClicked() {
    QString searchString = ui->txtSearchLine->text();
    ui->departmentListWidget->clearSelection();
    for (int i = 0; i < ui->departmentListWidget->count(); ++i) {
        QListWidgetItem *item = ui->departmentListWidget->item(i);
        bool match = item->text().contains(searchString, Qt::CaseInsensitive);
        ui->departmentListWidget->setRowHidden(i, !match);
    }
}

void MainWindow::refreshTableWidget() {
    QString department = ui->departmentListWidget->currentItem()->text();
    QString appendix = "/" + ui->yearLineEdit->text() + "-" + monthMap[ui->comboMonthBox->currentText()] + "_" + department + ".csv";
    QString file_path = PATH + appendix;
    if (!(QFileInfo::exists(file_path) && QFileInfo(file_path).isFile()))
        appendix = "/Шаблон.csv";
    QString path = PATH + appendix;
    importCSV(ui->tableWidget, path);
    ui->tableWidget->blockSignals(true);
    boldHeaders();
    filterByDep();
    ui->tableWidget->blockSignals(false);
}

void MainWindow::boldHeaders() {
    for (int col = 0; col < ui->tableWidget->columnCount(); ++col)
        ui->tableWidget->item(0, col)->setFlags(ui->tableWidget->item(0, col)->flags() ^ Qt::ItemIsEditable),
            ui->tableWidget->item(0, col)->setFont(QFont(ui->tableWidget->font().family(), -1, QFont::Bold));
}

void MainWindow::filterByDep() {
    int col = -1;
    for (int i = 0; i < ui->tableWidget->columnCount(); i++)
        if (ui->tableWidget->item(0, i)->text().startsWith("источник", Qt::CaseInsensitive))
            col = i;
    if (col == -1 || ui->comboDepBox->currentIndex() == 0) return;
    for (int row = 1; row < ui->tableWidget->rowCount(); row++)
        if (ui->tableWidget->item(row, col)->text() != ui->comboDepBox->currentText())
            ui->tableWidget->setRowHidden(row, true);
}

void MainWindow::importCSV(QTableWidget *tw, QString &path) {
    tw->blockSignals(true);
    tw->clear();
    tw->setRowCount(0);
    tw->setColumnCount(0);
    QFile CSVFile(path);
    if (CSVFile.open(QIODevice::ReadWrite)) {
        QTextStream* Stream;
        Stream = new QTextStream(&CSVFile);
        int row = 0;
        while (!Stream->atEnd()) {
            QString LineData = Stream->readLine();
            QStringList Data = LineData.split(";");
            tw->insertRow(row);
            tw->setColumnCount(Data.length());
            for (int col = 0; col < Data.length(); col++) {
                QTableWidgetItem *tableItem = new QTableWidgetItem(Data.at(col));
                tw->setItem(row, col, tableItem);
            }
            row++;
        }
        delete Stream;
    }
    CSVFile.close();
    tw->resizeRowsToContents();
    tw->blockSignals(false);
    // tw->horizontalHeader()->resizeSections(QHeaderView::ResizeToContents);
}

void MainWindow::saveCSV(QTableWidget *tw, QString &path) {
    QFile CSVFile(path);
    if (CSVFile.open(QIODevice::WriteOnly | QIODevice::Text)) {
        QTextStream *out = new QTextStream(&CSVFile);
        for (int row = 0; row < tw->rowCount(); ++row) {
            for (int col = 0; col < tw->columnCount(); ++col) {
                QTableWidgetItem *item = tw->item(row, col);
                QString text = item->text(), ch = (col == tw->columnCount() - 1) ? "" : ";";
                (*out) << text + ch;
            }
            (*out) << "\n";
        }
        delete out;
    }
    CSVFile.close();
}

void MainWindow::importTxt(QTextEdit *te, QString &path) {
    QFile CSVFile(path);
    if (CSVFile.open(QIODevice::ReadWrite)) {
        QTextStream* Stream;
        Stream = new QTextStream(&CSVFile);
        te->setText(Stream->readAll());
        delete Stream;
    }
    CSVFile.close();
}

void MainWindow::saveTxt(QTextEdit *te, QString &path) {
    QFile CSVFile(path);
    if (CSVFile.open(QIODevice::WriteOnly | QIODevice::Text)) {
        QTextStream *out = new QTextStream(&CSVFile);
        (*out) << te->toPlainText();
        delete out;
    }
    CSVFile.close();
}

void MainWindow::saveToPdf(QTableWidget *tableWidget) {
    QString filePath = QFileDialog::getSaveFileName(nullptr, "Save PDF", "table_content.pdf", "PDF files (*.pdf)");
    if (filePath.isEmpty())
        return;
    QPrinter printer(QPrinter::PrinterResolution);
    printer.setOutputFormat(QPrinter::PdfFormat);
    printer.setOutputFileName(filePath);
    printer.setPageSize(QPageSize::A4);
    printer.setPageMargins(QMargins(30, 30, 30, 30));
    printer.setPageOrientation(QPageLayout::Landscape);
    QTextDocument doc;
    doc.setPageSize(printer.pageRect(QPrinter::DevicePixel).size());
    QPainter painter;
    if (!painter.begin(&printer)) {
        // Handle the error if unable to start the printer
        return;
    }
    painter.setPen(Qt::black);
    painter.setFont(QFont("Time", 60));
    // int tableWidth = printer.pageRect(QPrinter::DevicePixel).width() - 30; // Subtract margins
    // int tableHeight = printer.pageRect(QPrinter::DevicePixel).height() - 30; // Subtract margins
    // tableWidget->setFixedSize(tableWidth, tableHeight);
    tableWidget->render(&painter);
    painter.end();
    QMessageBox::information(nullptr, "Success", "PDF file saved successfully.");
}

void MainWindow::saveToHtml(QTableWidget *tableWidget) {
    QString filePath = QFileDialog::getSaveFileName(nullptr, "Save HTML", "table_content.html", "HTML files (*.html)");
    if (filePath.isEmpty())
        return;

    QFile file(filePath);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Text)) {
        // Handle error opening file
        return;
    }

    QTextStream out(&file);
    out << "<html>\n";
    out << "<head>\n";
    out << "<style>\n";
    out << "table, th, td {\n";
    out << "  border: 1px solid black;\n";
    out << "  border-collapse: collapse;\n";
    out << "}\n";
    out << "</style>\n";
    out << "</head>\n";
    out << "<body>\n";

    out << "<table>\n";
    for (int row = 0; row < tableWidget->rowCount(); ++row) {
        if (tableWidget->isRowHidden(row)) continue;
        out << "  <tr>\n";
        for (int col = 0; col < tableWidget->columnCount(); ++col) {
            if (tableWidget->isColumnHidden(col)) continue;
            QTableWidgetItem *item = tableWidget->item(row, col);
            if (item) {
                out << "    <td>" << item->text() << "</td>\n";
            } else {
                out << "    <td></td>\n";
            }
        }
        out << "  </tr>\n";
    }

    out << "</table>\n";
    out << "</body>\n";
    out << "</html>\n";

    file.close();

    QMessageBox::information(nullptr, "Success", "HTML file saved successfully.");
}

void MainWindow::calculateCells(QTableWidgetItem *item) {
    int row = item->row(), col = item->column();
    if (ui->tableWidget->item(0, col)->text().toLower() != "поле ввода") {
        qDebug() << "Не поле ввода данных";
        return;
    }
    QVector<int> numbers;
    QStringList list = item->text().split(" ", Qt::SkipEmptyParts);
    bool conversionOK;
    for(const QString &str : list) {
        int num = str.toDouble(&conversionOK);
        if(conversionOK)
            numbers.push_back(num);
        else {
            insertCalcItem("Ошибка ввода", "", row, col);
            return;
        }
    }
    int f_col = -1;
    for (int column = 0; column < ui->tableWidget->columnCount(); column++)
        if (ui->tableWidget->item(0, column)->text().startsWith("Формула", Qt::CaseInsensitive))
            f_col = column;
    if (f_col == -1) {
        insertCalcItem("Не найдена формула", "", row, col);
        return;
    }
    QString formula = ui->tableWidget->item(row, f_col)->text();
    if (numbers.size() == 2 && formula.startsWith("M/N*100%", Qt::CaseInsensitive)) {
        double res = (double)numbers[0] / numbers[1];
        QString N = ui->tableWidget->item(row, 0)->text();
        if (N.startsWith("1") || N.startsWith("3")) {
            if (res >= 1.0) insertCalcItem("100%", "5", row, col);
            else if (res >= 0.85) insertCalcItem(">85%", "4", row, col);
            else if (res >= 0.7) insertCalcItem(">70%", "3", row, col);
            else if (res >= 0.55) insertCalcItem(">55%", "2", row, col);
            else insertCalcItem("<55%", "1", row, col);
        } else if (N.startsWith("2")) {
            if (res >= 1.0) insertCalcItem("100%", "5", row, col);
            else insertCalcItem("<100%", "1", row, col);
        } else if (N.startsWith("5")) {
            if (res >= 0.9) insertCalcItem(">90%", "5", row, col);
            else insertCalcItem("<90%", "1", row, col);
        } else if (N.startsWith("6.2")) {
            if (res >= 1.0) insertCalcItem("100%", "5", row, col);
            else if (res >= 0.8) insertCalcItem(">80%", "4", row, col);
            else if (res >= 0.75) insertCalcItem(">75%", "3", row, col);
            else if (res >= 0.7) insertCalcItem(">70%", "2", row, col);
            else insertCalcItem("<70%", "1", row, col);
        } else if (N.startsWith("6.3.3")) {
            if (res >= 0.95) insertCalcItem(">95%", "5", row, col);
            else if (res >= 0.85) insertCalcItem(">85%", "4", row, col);
            else if (res >= 0.75) insertCalcItem(">75%", "3", row, col);
            else if (res >= 0.65) insertCalcItem(">65%", "2", row, col);
            else insertCalcItem("<65%", "1", row, col);
        } else if (N.startsWith("6.3.4") || N.startsWith("6.4.3") || N.startsWith("6.5")) {
            if (res >= 1.0) insertCalcItem("100%", "5", row, col);
            else if (res >= 0.85) insertCalcItem(">85%", "3", row, col);
            else insertCalcItem("<85%", "1", row, col);
        } else
            insertCalcItem("Ошибка", "", row, col);
    }
    else if (numbers.size() == 2 && formula.startsWith("Чфакт/Чштат*100%", Qt::CaseInsensitive)) {
        double res = (double)numbers[0] / numbers[1];
        if (res >= 0.9) insertCalcItem(">90%", "5", row, col);
        else if (res >= 0.8) insertCalcItem(">80%", "3", row, col);
        else insertCalcItem("<80%", "1", row, col);
    }
    else if (numbers.size() == 2 && formula.startsWith("Чув/Чс*100%", Qt::CaseInsensitive)) {
        double res = (double)numbers[0] / numbers[1];
        if (res >= 0.1) insertCalcItem(">10%", "1", row, col);
        else if (res < 0.1) insertCalcItem("<10%", "5", row, col);
        else insertCalcItem("Ошибка", "", row, col);
    }
    else if (numbers.size() == 3 && formula.startsWith("M/N/Fср*100%", Qt::CaseInsensitive)) {
        double res = (double)numbers[0] / numbers[1] / numbers[2];
        if (res >= 1.0) insertCalcItem("100%", "5", row, col);
        else if (res >= 0.85) insertCalcItem(">85%", "4", row, col);
        else if (res >= 0.7) insertCalcItem(">70%", "3", row, col);
        else if (res >= 0.55) insertCalcItem(">55%", "2", row, col);
        else insertCalcItem("<55%", "1", row, col);
    }
    else if (numbers.size() == 3 && formula.startsWith("(M1+M2)/N*100%", Qt::CaseInsensitive)) {
        double res = (double)(numbers[0] + numbers[1]) / numbers[2];
        if (res >= 1.0) insertCalcItem("100%", "5", row, col);
        else if (res >= 0.99) insertCalcItem(">99%", "4", row, col);
        else if (res >= 0.98) insertCalcItem(">98%", "3", row, col);
        else if (res >= 0.97) insertCalcItem(">97%", "2", row, col);
        else insertCalcItem("<97%", "1", row, col);
    }
    else if (numbers.size() == 3 && formula.startsWith("(M/N*100%)/F*100%", Qt::CaseInsensitive)) {
        double res = (double)numbers[0] / numbers[1] * 100 / numbers[2];
        if (res >= 1.0) insertCalcItem("100%", "5", row, col);
        else insertCalcItem("<100%", "1", row, col);
    }
    else if (numbers.size() == 5 && formula.startsWith("M/(N1-N2-N3-N4)*100%", Qt::CaseInsensitive)) {
        double res = (double)numbers[0] / (numbers[1] - numbers[2] - numbers[3] - numbers[4]);
        if (res >= 1.0) insertCalcItem("100%", "5", row, col);
        else if (res >= 0.85) insertCalcItem(">85%", "4", row, col);
        else if (res >= 0.7) insertCalcItem(">70%", "3", row, col);
        else if (res >= 0.55) insertCalcItem(">55%", "2", row, col);
        else insertCalcItem("<55%", "1", row, col);
    }
    else if (numbers.size() == 1) {
        QString N = ui->tableWidget->item(row, 0)->text();
        if (N.startsWith("6.4.2") || N.startsWith("6.5.2")) {
            if (numbers[0] == 0.0) insertCalcItem("0", "5", row, col);
            else if (numbers[0] == 1.0) insertCalcItem("1", "3", row, col);
            else if (numbers[0] == 2.0) insertCalcItem("2", "1", row, col);
            else insertCalcItem("Ошибка работы с double", "", row, col);
        } else {
            if (numbers[0] >= 0.95) insertCalcItem("соблюдение", "5", row, col);
            else insertCalcItem("несоблюдение", "1", row, col);
        }
    } else
        insertCalcItem("", "", row, col);
}

void MainWindow::insertCalcItem(QString val, QString points, int row, int col) {
    QTableWidgetItem *tableItemVal = new QTableWidgetItem(val);
    QTableWidgetItem *tableItemPoints = new QTableWidgetItem(points);
    ui->tableWidget->setItem(row, col + 1, tableItemVal);
    ui->tableWidget->setItem(row, col + 2, tableItemPoints);
}
