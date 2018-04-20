#include "widget.h"
#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    Widget w;
    if (strcmp(w.word->metaObject()->className(),"QAxObject")) {
        qDebug() << "!word" << w.word->control();
        return -1;
    }
    w.show();

    return a.exec();
}
