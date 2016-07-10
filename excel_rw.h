#ifndef EXCEL_RW_H
#define EXCEL_RW_H

#include <QString>

#include <ActiveQt/QAxWidget>
#include <ActiveQt/QAxObject>

namespace JSK
{

class excel_RW
{
public:
    excel_RW();
    ~excel_RW();

private:
    QAxWidget*  m_excel;
    QAxObject*  m_books;
    QAxObject*  m_book;
    QAxObject*  m_sheets;
    QAxObject*  m_sheet;
    QString     m_filename;
    QString     m_sheetName;

    int m_rowStart;
    int m_colStart;
    int m_rowEnd;
    int m_colEnd;

public:
    bool        create(const QString& filename="");
    bool        open(const QString& filename="");
    void        save(const QString& filename="");
    void        close();

    void        setVisible(bool value);
    void        setCaption(const QString& value);

    QAxObject*  addBook();

    QVariant    read(int row, int col);
    void        write(int row, int col, const QVariant& value);

    int         sheetCount();
    QAxObject*  currentSheet();

    QAxObject*  sheet(int index);

    bool        usedRange();

    QStringList sheetNames();
    QString     sheetName() const { return m_sheetName; }

    inline int rowStart() const { return m_rowStart; }
    inline int colStart() const { return m_colStart; }
    inline int rowEnd() const { return m_rowEnd; }
    inline int colEnd() const { return m_colEnd; }
};

} // namespace JSK

#endif // EXCEL_RW_H
