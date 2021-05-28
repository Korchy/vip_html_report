//---------------------------------------------------------------------------
#pragma hdrstop
//---------------------------------------------------------------------------
#include "HtmlReport.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
//---------------------------------------------------------------------------
__fastcall HtmlReport::HtmlReport()
{
        // ���������� ������
Path = GetCurrentDir()+"\\";
Name = "Report";
Ext = "htm";
Text = new TStringList();
}
//---------------------------------------------------------------------------
__fastcall HtmlReport::~HtmlReport()
{
        // ���������� ������
delete Text;
}
//---------------------------------------------------------------------------
//                                ������� ������
//---------------------------------------------------------------------------
bool HtmlReport::LoadFromFile(AnsiString SourceFile)
{
        // ������� ��������� HTML-���� �� �������� ����� ����� �������� �����.
bool ReturnParam = true;
        // ��������� ����� ������
try {
        Text->LoadFromFile(SourceFile);
        }
catch(...) {
        MessageBox(NULL,"���������� ������� ���� ������!","ERROR",MB_OK);
        ReturnParam = false;
        }
return ReturnParam;
}
//---------------------------------------------------------------------------
bool HtmlReport::Save()
{
        // ���������� �������� ������
bool ReturnParam = true;
        // ��������� ����� ������
try {
        Text->SaveToFile(Path+Name+"."+Ext);
        }
catch(...) {
        MessageBox(NULL,"���������� ��������� �����!","ERROR",MB_OK);
        ReturnParam = false;
        }
return ReturnParam;
}
//---------------------------------------------------------------------------
bool HtmlReport::Show()
{
        // ����� ���������� ������
Save(); // ��������� ����� �������
ShellExecute(NULL,NULL,(Path+Name+"."+Ext).c_str(),NULL,NULL,SW_MAXIMIZE);
return true;
}
//---------------------------------------------------------------------------
//                      ������� ��� ������ � ����� �������
//---------------------------------------------------------------------------
bool HtmlReport::New()
{
        // �������� ������ ������
bool Rez = true;
if(Text->Count>0) Save();       // ��������� ������ �����
Text->Clear();  // �������� ��������
AnsiString Head = "";
Head += "<HTML>\n";
Head += "<HEAD>\n";
Head += "<TITLE><!--/Title/--></TITLE>\n";
Head += "</HEAD>\n";
Head += "<BODY>\n";
Text->Add(Head);
AnsiString Bottom = "";
Bottom += "</BODY>\n";
Bottom += "</HTML>\n";
Text->Add(Bottom);
return Rez;
}
//---------------------------------------------------------------------------
bool HtmlReport::Add(AnsiString Txt)
{
        // ���������� ������ � ����� ������
bool Rez = true;
Text->Text = StringReplace(Text->Text,"</BODY>",Txt+"\n</BODY>",TReplaceFlags()<<rfReplaceAll);
return Rez;
}
//---------------------------------------------------------------------------
//             ������� ��� ������ � ���������� � �������������� ������
//---------------------------------------------------------------------------
void HtmlReport::ReplaceZaklad(AnsiString Name, AnsiString Txt)
{
        // ������ �������� � ������ Name �� ����� Txt
AnsiString FullName = "<!--/" + Name + "/-->";
Text->Text = StringReplace(Text->Text,FullName,Txt,TReplaceFlags()<<rfReplaceAll);
}
//---------------------------------------------------------------------------
bool HtmlReport::ReplaceZakladWithADOQuery(AnsiString Name, TADOQuery* Data)
{
        // ������ �������� � ������ Name �� ��� ��������, ���������� ����� ���������
        // TADOQuery. �� ������ ���������� TADOQuery ��������� HTML-������� �
        // ����������� �� ����� ��������.
        // ��������� ������� �������
int LastActiveRecord;
try {
        LastActiveRecord = Data->RecNo - 1;
        }
catch(...) {
        return false;
        }
AnsiString OutputData = "\n<TABLE WIDTH=90% BORDER=1 CELLPADDING=0 CELLSPACING=0>\n";
        // ������� �� Query HTML-�������
if(Data->RecordCount>0) {
        // ���� ���� ������ � Query
        Data->First();
        while(!Data->Eof) {
                OutputData += "<TR ALIGN=CENTER>";
                // �� ���� ������� ��������������� ��� ����
                for(int i=0; i<Data->FieldCount; i++) {
                        AnsiString FieldData = Trim(Data->Fields->Fields[i]->AsString);
                        if(FieldData=="") FieldData = "&nbsp";
                        OutputData += "<TD><FONT SIZE=2>" + FieldData + "</FONT></TD>";
                        }
                OutputData += "</TR>\n";
                Data->Next();
                }
        }
else {
        // ������� ������ Query
        OutputData += "<TR ALIGN=CENTER><TD><I><B>������ ���</B></I></TD></TR>"; 
        }
OutputData += "</TABLE>\n";
        // ������� ������ �� ��������� �������
Data->First();
Data->MoveBy(LastActiveRecord);
        // �������� ����� � ������
AnsiString FullName = "<!--/" + Name + "/-->";
Text->Text = StringReplace(Text->Text,FullName,OutputData,TReplaceFlags()<<rfReplaceAll);
return true;
}
//---------------------------------------------------------------------------
void HtmlReport::ReplaceZakladWithArray(AnsiString Name, AnsiString Array[], int ArraySize)
{
        // ������ �������� � ������ Name �� ����� ������ �� ����������� �������
        // Array �������� ArraySize. �� ������ ������� ��������� HTML-������� �
        // �������� ��������.
        // ������� �������
AnsiString Txt = "\n<TABLE BORDER=1 CELLPADDING=0 CELLSPACING=0>\n";
for(int i=0; i<ArraySize; i++) {
        Txt += "<TR ALIGN=CENTER><TD><FONT SIZE=2>" + Array[i] + "</FONT></TD></TR>\n";
        }
Txt += "</TABLE>\n";
        // �������� �������� �� �����
AnsiString FullName = "<!--/" + Name + "/-->";
Text->Text = StringReplace(Text->Text,FullName,Txt,TReplaceFlags()<<rfReplaceAll);
}
//---------------------------------------------------------------------------
void HtmlReport::FreeZaklad(AnsiString Name)
{
        // ������������ �������� ���������� ����� (�.�. �������� ���������, ����� ��������)
AnsiString StartZaklad = "<!--/" + Name + "/";
Text->Text = StringReplace(Text->Text,StartZaklad,"",TReplaceFlags()<<rfReplaceAll);
AnsiString EndZaklad = "/" + Name + "/-->";
Text->Text = StringReplace(Text->Text,EndZaklad,"",TReplaceFlags()<<rfReplaceAll);
}
//---------------------------------------------------------------------------
void HtmlReport::DeleteZaklad(AnsiString Name)
{
        // ������ �������� �������� � ������������� � ��� ������
int StartPoint = (Text->Text).Pos ("<!--/" + Name + "/");
int StrCount = (Text->Text).Pos("/" + Name + "/-->") - StartPoint + 1 + Name.Length() + 4;
AnsiString FullZaklad = (Text->Text).SubString(StartPoint,StrCount);
Text->Text = StringReplace(Text->Text,FullZaklad,"",TReplaceFlags()<<rfReplaceAll);
}
//---------------------------------------------------------------------------

