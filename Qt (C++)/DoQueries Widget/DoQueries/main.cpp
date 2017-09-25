#include "mainwindow.h"
#include <QApplication>
#include <mail.h>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    MainWindow w;
    w.show();

    CSendFileTo sendTo;
    sendTo.(m_hWnd, _T("C:\test.docx"),
                    _T("Here's the lunch menu"));

    return a.exec();
}
