#ifndef WORD_H
#define WORD_H
#include <QAxObject>
#include <QDebug>

class Word : public QObject
{
    Q_OBJECT
public:
    explicit Word(QString msg, QObject *parent = Q_NULLPTR);
    ~Word();
};

#endif // WORD_H
