#ifndef WIDGET_H
#define WIDGET_H

#include <QWidget>
#include <QFileDialog>
#include <QDebug>
#include <QCalendarWidget>
#include <QVBoxLayout>
#include <QAxObject>
#include <QMessageBox>

namespace Ui {
class Widget;
}

class Widget : public QWidget
{
    Q_OBJECT

public:
    explicit Widget(QWidget *parent = 0);
    ~Widget();

public Q_SLOTS:
    void onWordException(int, const QString &, const QString &, const QString &);

private slots:
    void on_getTemplate_clicked();

    void on_loadFromFile_clicked();

    void on_setAutoColumnSize_clicked();

    void on_pushButton_clicked();

    void on_clearTable_clicked();

    void on_saveToFile_clicked();

    void on_addRow_clicked();

    void on_removeRow_clicked();

    void on_pushButton_2_clicked();

public:
    QAxObject *word;
private:
    Ui::Widget *ui;
    QAxObject *document = NULL;
    QString templateWord;

    bool newDocument();
    void closeDocument();
};

#endif // WIDGET_H
