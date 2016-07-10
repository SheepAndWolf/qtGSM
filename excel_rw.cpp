#include "excel_rw.h"

namespace JSK {
excel_RW::excel_RW()
    : m_excel(NULL)
    , m_books(NULL)
    , m_book(NULL)
    , m_sheets(NULL)
    , m_sheet(NULL)
    , m_sheetName("")
{

}

excel_RW::~excel_RW()
{
    close();
}

bool excel_RW::create(const QString& filename)
{
    close();
    m_excel = new QAxWidget("Excel.Application");
    m_books = m_excel->querySubObject("WorkBooks");
    m_book  = m_books->querySubObject("ActiveWorkBook");

    m_filename = filename;

    return false;
}

bool excel_RW::open(const QString& filename)
{
    close();
    m_excel  = new QAxWidget("Excel.Application");
    m_books  = m_excel->querySubObject("WorkBooks");
    m_book   = m_books->querySubObject("Open(const QString&)", filename);
    m_sheets = m_book ->querySubObject("WorkSheets");

    bool ret = m_book != NULL;
    if ( ret )//说明成功打开了
    {
        m_filename = filename;
    }
    return ret;
}

void excel_RW::save(const QString& filename)
{
    if ( ! filename.isEmpty() )
    {
        m_filename = filename;
    }
    m_books->dynamicCall("SaveAs(const QString&)", m_filename);
}

void excel_RW::close()
{
    m_sheet  = NULL;
    m_sheets = NULL;
    if ( m_book != NULL )
    {
        m_book->dynamicCall("Close(Boolean)", true);
        m_book = NULL;
    }
    m_books = NULL;
    if ( m_excel != NULL )
    {
        m_excel->dynamicCall("Quit(void)");
        m_excel = NULL;
    }
}

QStringList excel_RW::sheetNames()//获取sheet名列表
{
    QStringList ret;
    if ( m_sheets != NULL )
    {
        int sheetCount = m_sheets->property("Count").toInt();
        for( int i=1;i<=sheetCount;i++ )
        {
            QAxObject* sheet = m_sheets->querySubObject("Item(int)", i);
            ret.append(sheet->property("Name").toString());
        }
    }
    return ret;
}

void excel_RW::setVisible(bool value)//是否显示窗体
{
    m_excel->setProperty("Visible", value);
}

void excel_RW::setCaption(const QString& value)
{
    m_excel->setProperty("Caption", value);
}

QAxObject* excel_RW::addBook()
{
    return m_excel->querySubObject("WorkBooks");
}

QAxObject* excel_RW::currentSheet()
{
    return m_excel->querySubObject("ActiveWorkBook");
}

int excel_RW::sheetCount()
{
    int ret = 0;
    if ( m_sheets != NULL )
    {
        ret = m_sheets->property("Count").toInt();
    }
    return ret;
}

QAxObject* excel_RW::sheet(int index)
{
    m_sheet = NULL;
    if ( m_sheets != NULL )
    {
        m_sheet = m_sheets->querySubObject("Item(int)", index);
        m_sheetName = m_sheet->property("Name").toString();
    }
    return m_sheet;
}

QVariant excel_RW::read(int row, int col)
{
    QVariant ret;
    if ( m_sheet != NULL )
    {
        QAxObject* range = m_sheet->querySubObject("Cells(int, int)", row, col);
        ret = range->property("Value");
    }
    return ret;
}

bool excel_RW::usedRange()
{
    bool ret = false;
    if ( m_sheet != NULL )
    {
        QAxObject* urange  = m_sheet->querySubObject("UsedRange");
        m_rowStart = urange->property("Row").toInt();
        m_colStart = urange->property("Column").toInt();

        QAxObject* rows    = urange->querySubObject("Rows"   );
        QAxObject* columns = urange->querySubObject("Columns");
        m_rowEnd = rows   ->property("Count").toInt();
        m_colEnd = columns->property("Count").toInt();
        ret = true;
    }
    return ret;
}

}
