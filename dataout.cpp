#include "dataout.h"
#include "ui_dataout.h"

Dataout::Dataout(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::Dataout)
{
    ui->setupUi(this);
}

Dataout::~Dataout()
{
    delete ui;
}
