#include <QApplication>
#include "MainWindow.h"

int main(int argc, char* argv[])
{
    QApplication app(argc, argv);
    app.setApplicationName("FatturaView");
    app.setApplicationVersion("1.0.0");
    app.setOrganizationName("FatturaView");
    app.setOrganizationDomain("fatturaview.local");

    MainWindow w;
    w.show();

    return app.exec();
}
