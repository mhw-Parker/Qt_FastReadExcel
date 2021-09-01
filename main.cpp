#include "mainwindow.h"

#include <QApplication>
#include "readExcelData.h"

using namespace std;

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    MainWindow w;
    MatrixXf m;
    QString Car_Parameter_Path = QDir::currentPath() + "/车辆参数.xlsx";
    ExcelRead test;
    test.readExcelData(Car_Parameter_Path,m);
    cout << m;
    w.show();
    return a.exec();
}
