#include "widget.h"
#include "word.h"
#include <QApplication>

int main(int argc, char *argv[])
{
  QApplication a(argc, argv);

  Word *wrd = new Word("test");
  QAxObject *word = new QAxObject();
  bool controlLoaded = word->setControl("Word.Application");
  if (!controlLoaded)
  {
      QMessageBox msgBox;
      msgBox.setText("Критическая ошибка!!!\n"
                     "Не удалось подключиться к программе Word....");
      msgBox.setIcon(QMessageBox::Critical);
      msgBox.exec();
      return -1;
      // Message about control didn't load
  }

  Widget w(word);
  w.show();
  int result = a.exec();
  delete wrd;
  return result;
}
