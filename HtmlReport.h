//---------------------------------------------------------------------------
#ifndef HtmlReportH
#define HtmlReportH
//---------------------------------------------------------------------------
#include <Classes.hpp>  // ��� ������������� ����������� �++ - ���� ������ (���� AnsiString)
#include <ADODB.hpp>    //
#include <DB.hpp>       // ����� ����� ���� �������� � TADOQuery
//---------------------------------------------------------------------------
// �����, ������������ HTML-��������
//---------------------------------------------------------------------------
class HtmlReport
{
protected:

private:

        // ����������
        TStringList *Text;      // ���������� ����� ������       
        
public:
        HtmlReport(void);       // ���������� ������
        virtual ~HtmlReport();  // ���������� ������

        // ����������


        // �������
        bool LoadFile(AnsiString Path);         // �������� ��������� �� �������� ����� ������������� �����
        bool SaveFile(AnsiString Path);         // ���������� �������� ������
        void ReplaceZaklad(AnsiString Name, AnsiString Txt);    // ������ �������� � ������ Name �� ����� Txt
        bool ReplaceZakladWithADOQuery(AnsiString Name, TADOQuery* Data);       // ������ �������� � ������ Name �� ��� ��������, ������������ � ������� Array
        void ReplaceZakladWithArray(AnsiString Name, AnsiString Array[], int ArraySize); // ������ �������� Name �� ��� �������� �� ����������� ������� Arr �������� ArraySize
        void FreeZaklad(AnsiString Name);       // �������� �������� ���������� ����� (�.�. �������� ���������, ����� ��������)
        void DeleteZaklad(AnsiString Name);     // ������ �������� �������� � ������������� � ��� ������

};
//---------------------------------------------------------------------------
#endif

