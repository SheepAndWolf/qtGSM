#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    connect(ui->pushButton,SIGNAL(clicked()),this,SLOT(toDataIn()));
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::toDataIn()
{
    Dataout* out = new Dataout();
    DataIn* in = new DataIn();
    in->exec();
}
