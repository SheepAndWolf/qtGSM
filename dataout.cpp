#include "dataout.h"
#include "ui_dataout.h"
#include <QDir>
#include <QMessageBox>

using namespace std;

Dataout::Dataout(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::Dataout)
{
    ui->setupUi(this);
    QButtonGroup *type = new QButtonGroup;
    type->addButton(ui->txt);
    type->addButton(ui->excel);
//    connect(ui->MS,SIGNAL(clicked()),this,SLOT(MS()));
//    connect(ui->MSC,SIGNAL(clicked()),this,SLOT(MSC()));
//    connect(ui->BSC,SIGNAL(clicked()),this,SLOT(BSC()));
//    connect(ui->BTS,SIGNAL(clicked()),this,SLOT(BTS()));
//    connect(ui->Cell,SIGNAL(clicked()),this,SLOT(Cell()));
//    connect(ui->Freq,SIGNAL(clicked()),this,SLOT(Freq()));
//    connect(ui->Tianxian,SIGNAL(clicked()),this,SLOT(Tianxian()));
//    connect(ui->Neighbor,SIGNAL(clicked()),this,SLOT(Neighbor()));
//    connect(ui->Datas,SIGNAL(clicked()),this,SLOT(Datas()));
//    connect(ui->Locate,SIGNAL(clicked()),this,SLOT(Locate()));

    connect(ui->confirm,SIGNAL(clicked()),this,SLOT(data_out()));
}

Dataout::~Dataout()
{
    delete ui;
}

void Dataout::data_out()
{
    ui->confirm->setEnabled(false);
    QString where = ui->comboBox->currentText();
    qDebug()<<where;
    if(!ui->txt->isChecked()&&!ui->excel->isChecked())
    {
        QMessageBox* message = new QMessageBox();
        QString str;
        str += "请选择导出文件类型！";
        message->setWindowTitle("错误");
        message->setText(str);
        message->exec();
        ui->confirm->setEnabled(true);
        return;
    }
    if(ui->excel->isChecked())
    {
        QString filePath = "C:\\Users\\luckydog\\Desktop\\DB\\GSM_"+where+".xlsx";
        if(!filePath.isEmpty()){
            QAxObject *excel = new QAxObject(this);
            excel->setControl("Excel.Application");//连接Excel控件
            excel->dynamicCall("SetVisible (bool Visible)","false");//不显示窗体
            excel->setProperty("DisplayAlerts", false);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
            QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
            workbooks->dynamicCall("Add");//新建一个工作簿
            QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
            QAxObject *worksheets = workbook->querySubObject("Sheets");//获取工作表集合
            QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1

            QSqlQuery query;
            int count=0;
            query.exec("Select name from syscolumns Where ID=OBJECT_ID('"+where+"') order by colid");
            for(;query.next();count++)
            {
                QString v = (char)(65+count) + QString::number(1);
                QAxObject* cell = worksheet->querySubObject("Range(QVariant, QVariant)",v);
                cell->dynamicCall("SetValue(const QVariant&)",QVariant(query.value(0).toString()));
            }
            query.exec("select * from "+where);
            int i = 1;
            while(query.next())
            {
                for(int j=0;j<count;j++)
                {
                    QString v = (char)(65+j) + QString::number(i+1);
                    QAxObject* cell = worksheet->querySubObject("Range(QVariant, QVariant)",v);
                    cell->dynamicCall("SetValue(const QVariant&)",QVariant(query.value(j).toString()));
    //                qDebug()<<query.value(j).toString();
                }
                    i++;
            }

            workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filePath));//保存至filePath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
            workbook->dynamicCall("Close()");//关闭工作簿
            excel->dynamicCall("Quit()");//关闭excel
            delete excel;
            excel=NULL;
        }
    }
    else
    {
        QSqlQuery query;
        query.exec("Select count(name) from syscolumns Where ID=OBJECT_ID('"+where+"')");
        int count;
        if(query.next())
            count = query.value(0).toInt();
        string filePath = "C:\\Users\\luckydog\\Desktop\\DB\\"+where.toStdString()+".txt";
        char filefile[60];
        filePath.copy(filefile,filePath.length(),0);
        filefile[filePath.length()]='\0';
        ofstream fs(filefile,ios::out|ios::trunc);
        string a[count];
        query.exec("Select name from syscolumns Where ID=OBJECT_ID('"+where+"') order by colid");
        for(int j=0;query.next();j++)
        {
            a[j] = query.value(0).toString().toStdString();
            fs<<setiosflags(ios::left)<<setw(20)<<a[j];
        }
        fs<<endl;
        query.exec("select * from "+where);
        while(query.next())
        {
            for(int j=0;j<count;j++)
            {
                a[j] = query.value(j).toString().toStdString();
                fs<<setiosflags(ios::left)<<setw(20)<<a[j];
            }
            fs<<endl;
        }
        fs.close();
    }
    qDebug()<<"导出"+where+"成功！";
    ui->confirm->setEnabled(true);
}

//void Dataout::MS()
//{
//    QString filePath = "C:\\Users\\luckydog\\Desktop\\DB\\GSM_MS.xlsx";
//    if(!filePath.isEmpty()){
//        qDebug()<<(char)65;
//        QAxObject *excel = new QAxObject(this);
//        excel->setControl("Excel.Application");//连接Excel控件
//        excel->dynamicCall("SetVisible (bool Visible)","false");//不显示窗体
//        excel->setProperty("DisplayAlerts", false);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
//        QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
//        workbooks->dynamicCall("Add");//新建一个工作簿
//        QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
//        QAxObject *worksheets = workbook->querySubObject("Sheets");//获取工作表集合
//        QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1
////            QAxObject *cellX,*cellY;
////            for(int i=0;i<10;i++){
////                QString X=(char)65+QString::number(i+1);//设置要操作的单元格，如A1
////                QString Y="B"+QString::number(i+1);
////                cellX = worksheet->querySubObject("Range(QVariant, QVariant)",X);//获取单元格
////                cellY = worksheet->querySubObject("Range(QVariant, QVariant)",Y);
////                cellX->dynamicCall("SetValue(const QVariant&)",QVariant("qwe"));//设置单元格的值
////                cellY->dynamicCall("SetValue(const QVariant&)",QVariant(520));
////            }
//        QSqlQuery query;
//        int count=0;
//        query.exec("Select name from syscolumns Where ID=OBJECT_ID('MS') order by colid");
//        for(;query.next();count++)
//        {
//            QString v = (char)(65+count) + QString::number(1);
//            QAxObject* cell = worksheet->querySubObject("Range(QVariant, QVariant)",v);
//            cell->dynamicCall("SetValue(const QVariant&)",QVariant(query.value(0).toString()));
//        }
//        query.exec("select * from MS");
//        int i = 1;
//        while(query.next())
//        {
//            for(int j=0;j<count;j++)
//            {
//                QString v = (char)(65+j) + QString::number(i+1);
//                QAxObject* cell = worksheet->querySubObject("Range(QVariant, QVariant)",v);
//                cell->dynamicCall("SetValue(const QVariant&)",QVariant(query.value(j).toString()));
//                qDebug()<<query.value(j).toString();
//            }
//                i++;
//        }

//        workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filePath));//保存至filePath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
//        workbook->dynamicCall("Close()");//关闭工作簿
//        excel->dynamicCall("Quit()");//关闭excel
//        delete excel;
//        excel=NULL;
//    }
////    QSqlQuery query;
////    query.exec("select * from MS");
////    string filePath = "C:\\Users\\luckydog\\Desktop\\DB\\MS.txt";
////    char filefile[60];
////    filePath.copy(filefile,filePath.length(),0);
////    filefile[filePath.length()]='\0';
////    ofstream fs(filefile,ios::out|ios::trunc);
////    string a[8];
////    while(query.next())
////    {
////        for(int j=0;j<8;j++)
////        {
////            a[j] = query.value(j).toString().toStdString();
////            fs<<setiosflags(ios::left)<<setw(20)<<a[j];
////        }
////        fs<<endl;
////    }
////    fs.close();
//    qDebug("导出MS成功！");
//}
//void Dataout::MSC()
//{
//    QString filePath = "C:\\Users\\luckydog\\Desktop\\DB\\GSM_MSC.xlsx";
//    if(!filePath.isEmpty()){
//        qDebug()<<(char)65;
//        QAxObject *excel = new QAxObject(this);
//        excel->setControl("Excel.Application");//连接Excel控件
//        excel->dynamicCall("SetVisible (bool Visible)","false");//不显示窗体
//        excel->setProperty("DisplayAlerts", false);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
//        QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
//        workbooks->dynamicCall("Add");//新建一个工作簿
//        QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
//        QAxObject *worksheets = workbook->querySubObject("Sheets");//获取工作表集合
//        QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1

//        QSqlQuery query;
//        int count=0;
//        query.exec("Select name from syscolumns Where ID=OBJECT_ID('MSC') order by colid");
//        for(;query.next();count++)
//        {
//            QString v = (char)(65+count) + QString::number(1);
//            QAxObject* cell = worksheet->querySubObject("Range(QVariant, QVariant)",v);
//            cell->dynamicCall("SetValue(const QVariant&)",QVariant(query.value(0).toString()));
//    //        qDebug()<<query.value(j).toString();
//        }
//        qDebug()<<count;
//        query.exec("select * from MSC");
//        int i = 1;
//        while(query.next())
//        {
//            for(int j=0;j<count;j++)
//            {
//                QString v = (char)(65+j) + QString::number(i+1);
//                QAxObject* cell = worksheet->querySubObject("Range(QVariant, QVariant)",v);
//                cell->dynamicCall("SetValue(const QVariant&)",QVariant(query.value(j).toString()));
//                qDebug()<<query.value(j).toString();
//            }
//            i++;
//        }
//        workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filePath));//保存至filePath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
//        workbook->dynamicCall("Close()");//关闭工作簿
//        excel->dynamicCall("Quit()");//关闭excel
//        delete excel;
//        excel=NULL;
//    }
////    QSqlQuery query;
////    query.exec("select * from MSC");
////    string filePath = "C:\\Users\\luckydog\\Desktop\\DB\\MSC.txt";
////    char filefile[60];
////    filePath.copy(filefile,filePath.length(),0);
////    filefile[filePath.length()]='\0';
////    ofstream fs(filefile,ios::out|ios::trunc);
////    string a[6];
////    while(query.next())
////    {
////        for(int j=0;j<6;j++)
////        {
////            a[j] = query.value(j).toString().toStdString();
////            fs<<setiosflags(ios::left)<<setw(20)<<a[j];
////        }
////        fs<<endl;
////    }
////    fs.close();
//    qDebug("导出MSC成功！");
//}

//void Dataout::BSC()
//{
//    QString filePath = "C:\\Users\\luckydog\\Desktop\\DB\\GSM_BSC.xlsx";
//    if(!filePath.isEmpty()){
//        qDebug()<<(char)65;
//        QAxObject *excel = new QAxObject(this);
//        excel->setControl("Excel.Application");//连接Excel控件
//        excel->dynamicCall("SetVisible (bool Visible)","false");//不显示窗体
//        excel->setProperty("DisplayAlerts", false);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示
//        QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
//        workbooks->dynamicCall("Add");//新建一个工作簿
//        QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
//        QAxObject *worksheets = workbook->querySubObject("Sheets");//获取工作表集合
//        QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1

//        QSqlQuery query;
//        int count=0;
//        query.exec("Select name from syscolumns Where ID=OBJECT_ID('BSC') order by colid");
//        for(;query.next();count++)
//        {
//            QString v = (char)(65+count) + QString::number(1);
//            QAxObject* cell = worksheet->querySubObject("Range(QVariant, QVariant)",v);
//            cell->dynamicCall("SetValue(const QVariant&)",QVariant(query.value(0).toString()));
//    //        qDebug()<<query.value(j).toString();
//        }
//        qDebug()<<count;
//        query.exec("select * from BSC");
//        int i = 1;
//        while(query.next())
//        {
//            for(int j=0;j<count;j++)
//            {
//                QString v = (char)(65+j) + QString::number(i+1);
//                QAxObject* cell = worksheet->querySubObject("Range(QVariant, QVariant)",v);
//                cell->dynamicCall("SetValue(const QVariant&)",QVariant(query.value(j).toString()));
//                qDebug()<<query.value(j).toString();
//            }
//            i++;
//        }
//        workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filePath));//保存至filePath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
//        workbook->dynamicCall("Close()");//关闭工作簿
//        excel->dynamicCall("Quit()");//关闭excel
//        delete excel;
//        excel=NULL;
//    }
////    QSqlQuery query;
////    query.exec("select * from BSC");
////    string filePath = "C:\\Users\\luckydog\\Desktop\\DB\\BSC.txt";
////    char filefile[60];
////    filePath.copy(filefile,filePath.length(),0);
////    filefile[filePath.length()]='\0';
////    ofstream fs(filefile,ios::out|ios::trunc);
////    string a[6];
////    while(query.next())
////    {
////        for(int j=0;j<6;j++)
////        {
////            a[j] = query.value(j).toString().toStdString();
////            fs<<setiosflags(ios::left)<<setw(20)<<a[j];
////        }
////        fs<<endl;
////    }
////    fs.close();
//    qDebug("导出BSC成功！");
//}

//void Dataout::BTS()
//{
//    QSqlQuery query;
//    query.exec("select * from BTS");
//    string filePath = "C:\\Users\\luckydog\\Desktop\\DB\\BTS.txt";
//    char filefile[60];
//    filePath.copy(filefile,filePath.length(),0);
//    filefile[filePath.length()]='\0';
//    ofstream fs(filefile,ios::out|ios::trunc);
//    string a[7];
//    while(query.next())
//    {
//        for(int j=0;j<7;j++)
//        {
//            a[j] = query.value(j).toString().toStdString();
//            fs<<setiosflags(ios::left)<<setw(20)<<a[j];
//        }
//        fs<<endl;
//    }
//    fs.close();
//    qDebug()<<"BTS导出成功";
//}

//void Dataout::Cell()
//{
//    QSqlQuery query;
//    query.exec("select * from CELL");
//    string filePath = "C:\\Users\\luckydog\\Desktop\\DB\\CELL.txt";
//    char filefile[60];
//    filePath.copy(filefile,filePath.length(),0);
//    filefile[filePath.length()]='\0';
//    ofstream fs(filefile,ios::out|ios::trunc);
//    string a[9];
//    while(query.next())
//    {
//        for(int j=0;j<9;j++)
//        {
//            a[j] = query.value(j).toString().toStdString();
//            fs<<setiosflags(ios::left)<<setw(20)<<a[j];
//        }
//        fs<<endl;
//    }
//    fs.close();
//    qDebug()<<"小区基本信息导出成功";
//}

//void Dataout::Freq()
//{
//    QSqlQuery query;
//    query.exec("select * from FREQ");
//    string filePath = "C:\\Users\\luckydog\\Desktop\\DB\\FREQ.txt";
//    char filefile[60];
//    filePath.copy(filefile,filePath.length(),0);
//    filefile[filePath.length()]='\0';
//    ofstream fs(filefile,ios::out|ios::trunc);
//    string a[2];
//    while(query.next())
//    {
//        for(int j=0;j<2;j++)
//        {
//            a[j] = query.value(j).toString().toStdString();
//            fs<<setiosflags(ios::left)<<setw(20)<<a[j];
//        }
//        fs<<endl;
//    }
//    fs.close();
//    qDebug()<<"小区频点信息导出成功";
//}

//void Dataout::Tianxian()
//{
//    QSqlQuery query;
//    query.exec("select * from Antenna");
//    string filePath = "C:\\Users\\luckydog\\Desktop\\DB\\Antenna.txt";
//    char filefile[60];
//    filePath.copy(filefile,filePath.length(),0);
//    filefile[filePath.length()]='\0';
//    ofstream fs(filefile,ios::out|ios::trunc);
//    string a[8];
//    while(query.next())
//    {
//        for(int j=0;j<8;j++)
//        {
//            a[j] = query.value(j).toString().toStdString();
//            fs<<setiosflags(ios::left)<<setw(20)<<a[j];
//        }
//        fs<<endl;
//    }
//    fs.close();
//    qDebug()<<"天线信息导出成功";
//}

//void Dataout::Neighbor()
//{
//    QSqlQuery query;
//    query.exec("select * from neighbor");
//    string filePath = "C:\\Users\\luckydog\\Desktop\\DB\\neighbor.txt";
//    char filefile[60];
//    filePath.copy(filefile,filePath.length(),0);
//    filefile[filePath.length()]='\0';
//    ofstream fs(filefile,ios::out|ios::trunc);
//    string a[4];
//    while(query.next())
//    {
//        for(int j=0;j<4;j++)
//        {
//            a[j] = query.value(j).toString().toStdString();
//            fs<<setiosflags(ios::left)<<setw(20)<<a[j];
//        }
//        fs<<endl;
//    }
//    fs.close();
//    qDebug()<<"邻区信息导出成功";
//}

//void Dataout::Datas()
//{
//    QSqlQuery query;
//    query.exec("select * from DATAS");
//    string filePath = "C:\\Users\\luckydog\\Desktop\\DB\\DATA.txt";
//    char filefile[60];
//    filePath.copy(filefile,filePath.length(),0);
//    filefile[filePath.length()]='\0';
//    ofstream fs(filefile,ios::out|ios::trunc);
//    string a[10];
//    while(query.next())
//    {
//        for(int j=0;j<10;j++)
//        {
//            a[j] = query.value(j).toString().toStdString();
//            fs<<setiosflags(ios::left)<<setw(20)<<a[j];
//        }
//        fs<<endl;
//    }
//    fs.close();
//    qDebug()<<"话务数据导出成功";
//}

//void Dataout::Locate()
//{
//    QSqlQuery query;
//    query.exec("select * from detectInfo");
//    string filePath = "C:\\Users\\luckydog\\Desktop\\DB\\detectInfo.txt";
//    char filefile[60];
//    filePath.copy(filefile,filePath.length(),0);
//    filefile[filePath.length()]='\0';
//    ofstream fs(filefile,ios::out|ios::trunc);
//    string a[5];
//    while(query.next())
//    {
//        for(int j=0;j<5;j++)
//        {
//            a[j] = query.value(j).toString().toStdString();
//            fs<<setiosflags(ios::left)<<setw(20)<<a[j];
//        }
//        fs<<endl;
//    }
//    fs.close();
//   qDebug()<<"路测信息导出成功";
//}
