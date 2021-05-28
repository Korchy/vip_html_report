//---------------------------------------------------------------------------
#ifndef HtmlReportH
#define HtmlReportH
//---------------------------------------------------------------------------
#include <Classes.hpp>  // Для использования стандартных С++ - тпов данных (типа AnsiString)
#include <ADODB.hpp>    //
#include <DB.hpp>       // Чтобы можно было работать с TADOQuery
//---------------------------------------------------------------------------
// Класс, генерирующий HTML-документ
//---------------------------------------------------------------------------
class HtmlReport
{
protected:

private:

        // Переменные
        TStringList *Text;      // Собственно текст отчета       
        
public:
        HtmlReport(void);       // Констуктор класса
        virtual ~HtmlReport();  // Деструктор класса

        // Переменные


        // Функции
        bool LoadFile(AnsiString Path);         // Загрузка документа из которого будет формироваться отчет
        bool SaveFile(AnsiString Path);         // Сохранение готового отчета
        void ReplaceZaklad(AnsiString Name, AnsiString Txt);    // Замена закладки с именем Name на текст Txt
        bool ReplaceZakladWithADOQuery(AnsiString Name, TADOQuery* Data);       // Замена закладки с именем Name на ряд значений, содержащихся в массиве Array
        void ReplaceZakladWithArray(AnsiString Name, AnsiString Array[], int ArraySize); // Замена закладки Name на ряд значений из одномерного массива Arr размером ArraySize
        void FreeZaklad(AnsiString Name);       // Удаление закладки содержащей текст (т.е. закладка удаляется, текст остается)
        void DeleteZaklad(AnsiString Name);     // Полное удаление закладки и содержащегося в ней текста

};
//---------------------------------------------------------------------------
#endif

