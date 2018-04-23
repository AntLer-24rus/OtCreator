#include "widget.h"
#include "ui_widget.h"


Widget::Widget(QAxObject *wordIn, QWidget *parent) : QWidget(parent), ui(new Ui::Widget)
{
    qDebug() << "Create Widget instanse";
    ui->setupUi(this);
    word = wordIn;
        // Control loaded OK; connecting to catch exceptions from control
        word->setProperty("Visible", true);
        word->blockSignals(false);
        connect(word, SIGNAL(exception(int, const QString &, const QString &, const QString &)),
                this,
                SLOT(onWordException(int, const QString &, const QString &, const QString &)));
}

Widget::~Widget()
{
    delete ui;
    word->dynamicCall("Quit(Boolean)", false);
    word->clear();
    delete word;
    word = NULL;
}

void Widget::onWordException(int code, const QString &source, const QString &desc, const QString &help) {
    QMessageBox::warning(this, tr("OtCreator"),
                         "Критическая ошибка\n"
                         "Код: " + QString::number(code) + "\n"
                         "Источник: " + source + "\n"
                         "desc: " + desc + "\n"
                         "Помощь: " + help);
    word->dynamicCall("Quit(Boolean)", false);
    word->clear();
    delete word;
    word = NULL;
}

bool Widget::newDocument()
{
    if(templateWord.isEmpty()) {
        QMessageBox msgBox;
        QMessageBox::warning(this, tr("OtCreator"),
                             tr("Не выбран шаблон, заполнение невозможно"));

        return false;
    }
    QAxObject *documents = word->querySubObject("Documents");
    if (documents) {
        documents->blockSignals(false);
        connect(documents, SIGNAL(exception(int, const QString &, const QString &, const QString &)),
                this,
                SLOT(onWordException(int, const QString &, const QString &, const QString &)));
        document = documents->querySubObject("Add(const QString&)", templateWord);
        if (document) {
            document->blockSignals(false);
            connect(document, SIGNAL(exception(int, const QString &, const QString &, const QString &)),
                    this,
                    SLOT(onWordException(int, const QString &, const QString &, const QString &)));
            return true;
        }
    }
    return false;
}

void Widget::closeDocument()
{
    if(document) {
        document->dynamicCall("Close(Boolean)", false);
        document = NULL;
    }
}

void Widget::on_getTemplate_clicked()
{
    templateWord = QFileDialog::getOpenFileName(this,
                                                "Открыть шаблон отзыва/рецензии",
                                                QDir::currentPath(),
                                                "All files (*.*); Template doc files (*.dot *.dotx);");
    //    if(fileName.isEmpty()) return;
}

void Widget::on_loadFromFile_clicked()
{
    QString fileName = QFileDialog::getOpenFileName(this,
                                                    "Открыть файл со студентами",
                                                    QDir::currentPath(),
                                                    "Student list (*.slst)");
    qDebug() << "Select file" << fileName;
    QFile fileStudents(fileName);
    if(fileStudents.open(QIODevice::ReadWrite | QIODevice::Text)) {
        int i =0;
        ui->tableWidget->clear();
        ui->tableWidget->setRowCount(0);
        //        ui->tableWidget->setColumnCount(4);
        ui->tableWidget->setHorizontalHeaderLabels(QStringList() << "Ф.И.О" << "Тема" << "Дата" << "Оценка");
        while (!fileStudents.atEnd()) {
            QString line = fileStudents.readLine();
            line = line.trimmed();
            if (line.left(2) == "//") {
                continue;
            }
            QStringList lst = line.split(";");
            // Строка в файле содержит:
            //              0 - ФИО
            //              1 - Тема
            //              2 - Дата сдачи
            //              3 - Оценка
            ui->tableWidget->insertRow(i);

            for (int j = 0; j < lst.count(); j++) {
                if(ui->tableWidget->columnCount() <= j){
                    ui->tableWidget->insertColumn(j);

                }
                ui->tableWidget->setItem(i,j,new QTableWidgetItem(lst.at(j)));
                switch (j) {
                case 0:
                case 1:
                    ui->tableWidget->item(i,j)->setTextAlignment(Qt::AlignLeft | Qt::AlignVCenter);
                    break;
                default:
                    ui->tableWidget->item(i,j)->setTextAlignment(Qt::AlignHCenter | Qt::AlignVCenter);
                    break;
                }
            }
            i++;
        }
        // Ресайзим колонки по содержимому
        ui->tableWidget->resizeColumnsToContents();
    }

    fileStudents.close();
}

void Widget::on_setAutoColumnSize_clicked()
{
    // Ресайзим колонки по содержимому
    ui->tableWidget->resizeColumnsToContents();
}

void Widget::on_pushButton_clicked()
{
    closeDocument();
    //
    //    QCalendarWidget *cal = new QCalendarWidget(this);
    ////    QVBoxLayout *vbox = new QVBoxLayout(this);
    ////    vbox->addWidget(cal);
    //    cal->setGeometry(390,75,20,20);
}

void Widget::on_clearTable_clicked()
{
    ui->tableWidget->clear();
    ui->tableWidget->setRowCount(0);
    ui->tableWidget->setHorizontalHeaderLabels(QStringList() << "Ф.И.О" << "Тема" << "Дата" << "Оценка");
}

void Widget::on_saveToFile_clicked()
{
    QString fileName = QFileDialog::getSaveFileName(
                this,
                "Выберте файл для сохранения",
                QDir::currentPath(),
                "Student list (*.slst)");
    QFile fileStudents(fileName);
    if(fileStudents.open(QIODevice::WriteOnly | QIODevice::Text)) {
        for (int i = 0; i < ui->tableWidget->rowCount(); i++) {
            QString line;
            for (int j = 0; j < ui->tableWidget->columnCount(); j++) {
                line = line.append(ui->tableWidget->item(i,j)->text().trimmed());
                if (j != ui->tableWidget->columnCount() - 1) {
                    line = line.append(";");
                }
            }
            fileStudents.write(line.append("\n").toUtf8());
            qDebug() << line;
        }
    }
    fileStudents.close();
}

void Widget::on_addRow_clicked()
{
    ui->tableWidget->insertRow(ui->tableWidget->rowCount());
}

void Widget::on_removeRow_clicked()
{
    ui->tableWidget->removeRow(ui->tableWidget->currentRow());
}

void Widget::on_pushButton_2_clicked()
{
    for (int i = 0; i < ui->tableWidget->rowCount(); i++) {
        if(newDocument()) {
            QAxObject *bookmark = document->querySubObject("Bookmarks(const QString&)","ТестоваяЗакладка");
            QAxObject *range = bookmark->querySubObject("Range");
            range->setProperty("Text", ui->tableWidget->item(i,0)->text().trimmed());

        }
    }
}
