#ifndef OPENDB_H
#define OPENDB_H
#include <QSqlDatabase>

class OpenDB
{
public:
    OpenDB();
    ~OpenDB();
    QSqlDatabase createDB();
    QSqlDatabase getDB();
    QSqlDatabase DB;
    QString DBtest;
};

extern OpenDB* odb;

#endif // OPENDB_H
