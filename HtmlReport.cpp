//---------------------------------------------------------------------------
#pragma hdrstop

#include "HtmlReport.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
//---------------------------------------------------------------------------
HtmlReport::HtmlReport()
{
        // Конструктр класса
Text = new TStringList();

}
//---------------------------------------------------------------------------
HtmlReport::~HtmlReport()
{
        // Деструктор класса
delete Text;

}
//---------------------------------------------------------------------------
//                                ФУНКЦИИ КЛАССА
//---------------------------------------------------------------------------
bool HtmlReport::LoadFile(AnsiString Path)
{
        // Функция загружает HTML-файл из которого далее будет делаться отчет.
bool ReturnParam = true;
        // Загрузить текст отчета
try {
        Text->LoadFromFile(Path);
        }
catch(...) {
        MessageBox(NULL,"Невозможно открыть файл отчета!","ERROR",MB_OK);
        ReturnParam = false;
        }

return ReturnParam;
}
//---------------------------------------------------------------------------
bool HtmlReport::SaveFile(AnsiString Path)
{
        // Сохранение готового отчета
bool ReturnParam = true;
        // Загрузить текст отчета
try {
        Text->SaveToFile(Path);
        }
catch(...) {
        MessageBox(NULL,"Невозможно сохранить отчет!","ERROR",MB_OK);
        ReturnParam = false;
        }
return ReturnParam;
}
//---------------------------------------------------------------------------
void HtmlReport::ReplaceZaklad(AnsiString Name, AnsiString Txt)
{
        // Замена закладки с именем Name на текст Txt
AnsiString FullName = "<!--/" + Name + "/-->";
Text->Text = StringReplace(Text->Text,FullName,Txt,TReplaceFlags()<<rfReplaceAll);
}
//---------------------------------------------------------------------------
bool HtmlReport::ReplaceZakladWithADOQuery(AnsiString Name, TADOQuery* Data)
{
        // Замена закладки с именем Name на ряд значений, переданных через компонент
        // TADOQuery. Из данных компонента TADOQuery создается HTML-таблица и
        // вставляется на место закладки.
        // Сохранить позицию курсора
int LastActiveRecord;
try {
        LastActiveRecord = Data->RecNo - 1;
        }
catch(...) {
        return false;
        }
AnsiString OutputData = "\n<TABLE WIDTH=650 BORDER=1 CELLPADDING=0 CELLSPACING=0>\n";
        // Создать из Query HTML-таблицу
if(Data->RecordCount>0) {
        // Если есть записи в Query
        Data->First();
        while(!Data->Eof) {
                OutputData += "<TR ALIGN=CENTER>";
                // По всем записям проссуммировать все поля
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
        // Передан пустой Query
        OutputData += "<TR ALIGN=CENTER><TD><I><B>ДАННЫХ НЕТ</B></I></TD></TR>"; 
        }
OutputData += "</TABLE>\n";
        // Вернуть курсор на начальную позицию
Data->First();
Data->MoveBy(LastActiveRecord);
        // Заменить текст в отчете
AnsiString FullName = "<!--/" + Name + "/-->";
Text->Text = StringReplace(Text->Text,FullName,OutputData,TReplaceFlags()<<rfReplaceAll);
return true;
}
//---------------------------------------------------------------------------
void HtmlReport::ReplaceZakladWithArray(AnsiString Name, AnsiString Array[], int ArraySize)
{
        // Замена закладки с именем Name на набор данных из одномерного массива
        // Array размером ArraySize. Из данных массива создается HTML-таблица и
        // заменяет закладку.
        // Создать таблицу
AnsiString Txt = "\n<TABLE BORDER=1 CELLPADDING=0 CELLSPACING=0>\n";
for(int i=0; i<ArraySize; i++) {
        Txt += "<TR ALIGN=CENTER><TD><FONT SIZE=2>" + Array[i] + "</FONT></TD></TR>\n";
        }
Txt += "</TABLE>\n";
        // Заменить закладку на текст
AnsiString FullName = "<!--/" + Name + "/-->";
Text->Text = StringReplace(Text->Text,FullName,Txt,TReplaceFlags()<<rfReplaceAll);
}
//---------------------------------------------------------------------------
void HtmlReport::FreeZaklad(AnsiString Name)
{
        // Освобождение закладки содержащей текст (т.е. закладка удаляется, текст остается)
AnsiString StartZaklad = "<!--/" + Name + "/";
Text->Text = StringReplace(Text->Text,StartZaklad,"",TReplaceFlags()<<rfReplaceAll);
AnsiString EndZaklad = "/" + Name + "/-->";
Text->Text = StringReplace(Text->Text,EndZaklad,"",TReplaceFlags()<<rfReplaceAll);
}
//---------------------------------------------------------------------------
void HtmlReport::DeleteZaklad(AnsiString Name)
{
        // Полное удаление закладки и содержащегося в ней текста
int StartPoint = (Text->Text).Pos ("<!--/" + Name + "/");
int StrCount = (Text->Text).Pos("/" + Name + "/-->") - StartPoint + 1 + Name.Length() + 4;
AnsiString FullZaklad = (Text->Text).SubString(StartPoint,StrCount);
Text->Text = StringReplace(Text->Text,FullZaklad,"",TReplaceFlags()<<rfReplaceAll);
}
//---------------------------------------------------------------------------

