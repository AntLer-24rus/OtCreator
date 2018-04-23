#include "word.h"

Word::Word(QString msg, QObject *parent) : QObject(parent)
{
    qDebug() << "Create Word instanse";
    qDebug() << msg;
//    setControl("Word.Application");
//    return Q_NULLPTR;
//    exit(-1);
}

Word::~Word()
{
//    dynamicCall("Quit(Boolean)", false);
//    clear();
}
