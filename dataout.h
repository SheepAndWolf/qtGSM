#ifndef DATAOUT_H
#define DATAOUT_H

#include <QDialog>

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
};

#endif // DATAOUT_H
