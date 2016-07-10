#ifndef DATAOUT_H
#define DATAOUT_H

#include <QDialog>
#include <QAxObject>
#include <QSqlQuery>

#include <qdebug.h>
#include <iostream>
#include <fstream>
#include <iomanip>
#include <string.h>

namespace Ui {
class Dataout;
}

class Dataout : public QDialog
{
    Q_OBJECT

public:
    explicit Dataout(QWidget *parent = 0);
    ~Dataout();

private:
    Ui::Dataout *ui;
private slots:
//    void MS();
//    void MSC();
//    void BSC();
//    void BTS();
//    void Cell();
//    void Freq();
//    void Tianxian();
//    void Neighbor();
//    void Datas();
//    void Locate();
    void data_out();
};

#endif // DATAOUT_H
