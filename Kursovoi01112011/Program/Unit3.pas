unit Unit3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, ComObj, DB, IBDatabase, IBCustomDataSet, IBTable,
  ComCtrls, IBQuery;

type
  TForm3 = class(TForm)
    GroupBox1: TGroupBox;
    StringGrid1: TStringGrid;
    GroupBox2: TGroupBox;
    Label1: TLabel;
    ComboBox1: TComboBox;
    Label2: TLabel;
    ComboBox2: TComboBox;
    Button1: TButton;
    GroupBox3: TGroupBox;
    Label3: TLabel;
    Edit1: TEdit;
    Button2: TButton;
    Button3: TButton;
    OpenDialog1: TOpenDialog;
    GroupBox4: TGroupBox;
    Label4: TLabel;
    Edit2: TEdit;
    Button4: TButton;
    Button5: TButton;
    GroupBox5: TGroupBox;
    Label5: TLabel;
    Label6: TLabel;
    ComboBox3: TComboBox;
    ComboBox4: TComboBox;
    Button6: TButton;
    IBDatabase1: TIBDatabase;
    IBTable1: TIBTable;
    IBTransaction1: TIBTransaction;
    Label7: TLabel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    OpenDialog2: TOpenDialog;
    IBQuery1: TIBQuery;
    IBQuery2: TIBQuery;
    IBQuery3: TIBQuery;
    IBTable2: TIBTable;
    Label8: TLabel;
    ComboBox5: TComboBox;
    procedure Button2Click(Sender: TObject);    //Выбор файла для импорта
    procedure FormCreate(Sender: TObject);      //Событие при создании формы
    procedure Import(var importFileName:String);  //Импорт данных
    function Clean_String(var str:String):String; //Чистка строки от посторонних символов
    procedure Button3Click(Sender: TObject);   //Событие нажатия кнопки для импорта
    procedure FormClose(Sender: TObject; var Action: TCloseAction); //Событие при закрытии формы
    procedure Button6Click(Sender: TObject);      //Процедура сравнения
    procedure Button5Click(Sender: TObject);    //Экспорт сводного отчета
    procedure Button4Click(Sender: TObject);  //Выбор файла для экспорта
    procedure RefreshReportsCombo;             //Обновление полей со списком отчетный месяц
    procedure Button1Click(Sender: TObject);   //Поиск и заполнение сетки
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;
  VA :array[1..27] of Integer;
  MA :array[1..27] of Integer;
  FA :array[1..27] of Integer;
  TVA:array[1..27] of Integer;
  TMA:array[1..27] of Integer;
  TFA:array[1..27] of Integer;

implementation

{$R *.dfm}

//Поиск и заполнение сетки
procedure TForm3.Button1Click(Sender: TObject);
var
 sqlString:String;
 i:Integer;
begin
 if (ComboBox1.ItemIndex>-1) and (ComboBox2.ItemIndex>-1) then    //Блок №2 Процедура поиска
  begin
   IBQuery1.Close;
   IBQuery1.SQL.Clear;
   sqlString:='SELECT * FROM BASE WHERE LPU_KEY ='+IntToStr(ComboBox1.ItemIndex+1)+
              ' AND KIND_KEY=1 '+
              ' AND REPORT_DATE='''+ ComboBox2.Text+'''';
   IBQuery1.SQL.Add(sqlString);
   IBQuery1.Open;
   IBQuery1.Last;
   IBQuery1.First;
   if IBQuery1.RecordCount>0 then   //Блок №3 Процедура поиска
    begin
     for i := 1 to 27 do    //Блок №4 Процедура поиска
      VA[i]:=IBQuery1.Fields[i+3].AsInteger;   //Блок №5 Процедура поиска
    end;

   IBQuery1.Close;
   IBQuery1.SQL.Clear;
   sqlString:='SELECT * FROM BASE WHERE LPU_KEY ='+IntToStr(ComboBox1.ItemIndex+1)+
              ' AND KIND_KEY=2 '+
              ' AND REPORT_DATE='''+ ComboBox2.Text+'''';
   IBQuery1.SQL.Add(sqlString);
   IBQuery1.Open;
   IBQuery1.Last;
   IBQuery1.First;
   if IBQuery1.RecordCount>0 then  //Блок №6 Процедура поиска
    begin
     for i := 1 to 27 do    //Блок №7 Процедура поиска
      MA[i]:=IBQuery1.Fields[i+3].AsInteger;  //Блок №8 Процедура поиска
    end;

   IBQuery1.Close;
   IBQuery1.SQL.Clear;
   sqlString:='SELECT * FROM BASE WHERE LPU_KEY ='+IntToStr(ComboBox1.ItemIndex+1)+
              ' AND KIND_KEY=3 '+
              ' AND REPORT_DATE='''+ ComboBox2.Text+'''';
   IBQuery1.SQL.Add(sqlString);
   IBQuery1.Open;
   IBQuery1.Last;
   IBQuery1.First;
   if IBQuery1.RecordCount>0 then   //Блок №9 Процедура поиска
    begin
     for i := 1 to 27 do   //Блок №10 Процедура поиска
      FA[i]:=IBQuery1.Fields[i+3].AsInteger; //Блок №11 Процедура поиска
    end;

  //Выдача в сетку
  //Блок №12 Процедура поиска
  for i:=1  to 27 do
   begin
    StringGrid1.Cells[1,i]:=IntToStr(VA[i]);
    StringGrid1.Cells[2,i]:=IntToStr(MA[i]);
    StringGrid1.Cells[3,i]:=IntToStr(FA[i]);
   end;
  end;
end;

//Выбор файла для импорта
procedure TForm3.Button2Click(Sender: TObject);
begin
 OpenDialog1.InitialDir:=ExtractFilePath(Application.ExeName); //Путь к программе
 if OpenDialog1.Execute then
  Edit1.Text:=OpenDialog1.FileName;
end;

//Событие нажатия кнопки для импорта
procedure TForm3.Button3Click(Sender: TObject);
var
 importFileName:String;
begin
 importFileName:=Edit1.Text;
 Import(importFileName);
end;

//Выбор файла для экспорта
procedure TForm3.Button4Click(Sender: TObject);
begin
 OpenDialog2.InitialDir:=ExtractFilePath(Application.ExeName); //Путь к программе
 if OpenDialog2.Execute then
  Edit2.Text:=OpenDialog2.FileName;
end;

//Экспорт сводного отчета
procedure TForm3.Button5Click(Sender: TObject);
var
  i,j: Integer;
  App: Variant;
  sqlString:String;
  LPUKey:Integer;
  PERORT_DATE:String;
begin
 if Edit2.Text<>'' then //Блок №2 процедура экспорта данных
  begin
  //Чистим временный массив
   for i := 1 to 27 do  //Блок №3 процедура экспорта данных
    begin
    //Блок №4 процедура экспорта данных
     TVA[i]:=0;
     TMA[i]:=0;
     TFA[i]:=0;
    end;
   //Чтение из базы всех отчетов за этот месяц
   if (ComboBox5.ItemIndex>-1) then   //Блок №5 процедура экспорта данных
   begin
    IBQuery1.Close;
    IBQuery1.SQL.Clear;
    sqlString:='SELECT BASE.LPU_KEY, BASE.REPORT_DATE FROM BASE WHERE LPU_KEY ='+IntToStr(ComboBox5.ItemIndex+1)+
               ' AND REPORT_DATE1>'''+ DateToStr(DateTimePicker1.DateTime)+''''+
               ' AND REPORT_DATE1<'''+ DateToStr(DateTimePicker2.DateTime)+''''+
               ' GROUP BY BASE.LPU_KEY, BASE.REPORT_DATE';
    IBQuery1.SQL.Add(sqlString);
    IBQuery1.Open;
    IBQuery1.Last;
    IBQuery1.First;
    for i := 1 to IBQuery1.RecordCount do //Блок №6 процедура экспорта данных
     begin
      //Блок №7 процедура экспорта данных
      LPUKey:=IBQuery1.Fields[0].AsInteger;
      PERORT_DATE:=IBQuery1.Fields[1].AsString;
      IBQuery2.Close;
      IBQuery2.SQL.Clear;
      sqlString:='SELECT * FROM BASE WHERE LPU_KEY ='+IntToStr(LPUKey)+
                 ' AND KIND_KEY=1'+
                 ' AND REPORT_DATE='''+PERORT_DATE +'''';
      IBQuery2.SQL.Add(sqlString);
      IBQuery2.Open;
      IBQuery2.Last;
      IBQuery2.First;
      if IBQuery2.RecordCount>0 then  //Блок №8 процедура экспорта данных
       begin
        for j := 1 to 27 do  //Блок №9 процедура экспорта данных
         TVA[j]:=TVA[j]+IBQuery2.Fields[j+3].AsInteger;  //Блок №10 процедура экспорта данных
       end;

      IBQuery2.Close;
      IBQuery2.SQL.Clear;
      sqlString:='SELECT * FROM BASE WHERE LPU_KEY ='+IntToStr(LPUKey)+
                 ' AND KIND_KEY=2'+
                 ' AND REPORT_DATE='''+ PERORT_DATE +'''';
      IBQuery2.SQL.Add(sqlString);
      IBQuery2.Open;
      IBQuery2.Last;
      IBQuery2.First;
      if IBQuery2.RecordCount>0 then   //Блок №11 процедура экспорта данных
       begin
        for j := 1 to 27 do //Блок №12 процедура экспорта данных
         TMA[j]:=TMA[j]+IBQuery2.Fields[j+3].AsInteger; //Блок №13 процедура экспорта данных
       end;

      IBQuery2.Close;
      IBQuery2.SQL.Clear;
      sqlString:='SELECT * FROM BASE WHERE LPU_KEY ='+IntToStr(LPUKey)+
                 ' AND KIND_KEY=2'+
                 ' AND REPORT_DATE='''+PERORT_DATE +'''';
      IBQuery2.SQL.Add(sqlString);
      IBQuery2.Open;
      IBQuery2.Last;
      IBQuery2.First;
      if IBQuery2.RecordCount>0 then   //Блок №14 процедура экспорта данных
       begin
        for j := 1 to 27 do //Блок №15 процедура экспорта данных
         TFA[j]:=TFA[j]+IBQuery2.Fields[j+3].AsInteger; //Блок №16 процедура экспорта данных
       end;
       IBQuery1.Next;  //Блок №17 процедура экспорта данных
     end;

    //Блок №18 процедура экспорта данных
    //Выводим данные в Word
    App := CreateOleObject('Word.Application');
    App.Documents.Open(Edit2.Text);
    App.Visible:=true;
    for i:=1 to 27 do
     begin
      App.ActiveDocument.Tables.Item(1).Cell(i+2,3).Range.Text:=IntToStr(TVA[i]);
      App.ActiveDocument.Tables.Item(1).Cell(i+2,4).Range.Text:=IntToStr(TMA[i]);
      App.ActiveDocument.Tables.Item(1).Cell(i+2,5).Range.Text:=IntToStr(TFA[i]);
     end;
    App:= UnAssigned;
   end;
  end;
end;

//Процедура сравнения
procedure TForm3.Button6Click(Sender: TObject);
var
 sqlString:String;
 i:Integer;
begin

 if (ComboBox1.ItemIndex>-1) and (ComboBox2.ItemIndex>-1) and     //Блок №2 (Процедура сравнения)
    (ComboBox3.ItemIndex>-1) and (ComboBox4.ItemIndex>-1) then
  begin
    //Чтения из базы первого отчета
    Button1Click(nil);   //Блок №3 (Процедура сравнения)
   //Чтения из базы второго отчета
   IBQuery1.Close;
   IBQuery1.SQL.Clear;
   sqlString:='SELECT * FROM BASE WHERE LPU_KEY ='+IntToStr(ComboBox3.ItemIndex+1)+
              ' AND KIND_KEY=1 '+
              ' AND REPORT_DATE='''+ ComboBox4.Text+'''';
   IBQuery1.SQL.Add(sqlString);
   IBQuery1.Open;
   IBQuery1.Last;
   IBQuery1.First;
   if IBQuery1.RecordCount>0 then    //Блок №4 (Процедура сравнения)
    begin
     for i := 1 to 27 do   //Блок №5 (Процедура сравнения)
      TVA[i]:=IBQuery1.Fields[i+3].AsInteger;  //Блок №6 (Процедура сравнения)
    end;

   IBQuery1.Close;
   IBQuery1.SQL.Clear;
   sqlString:='SELECT * FROM BASE WHERE LPU_KEY ='+IntToStr(ComboBox3.ItemIndex+1)+
              ' AND KIND_KEY=2 '+
              ' AND REPORT_DATE='''+ ComboBox4.Text+'''';
   IBQuery1.SQL.Add(sqlString);
   IBQuery1.Open;
   IBQuery1.Last;
   IBQuery1.First;
   if IBQuery1.RecordCount>0 then    //Блок №7 (Процедура сравнения)
    begin
     for i := 1 to 27 do      //Блок №8 (Процедура сравнения)
      TMA[i]:=IBQuery1.Fields[i+3].AsInteger;   //Блок №9 (Процедура сравнения)
    end;

   IBQuery1.Close;
   IBQuery1.SQL.Clear;
   sqlString:='SELECT * FROM BASE WHERE LPU_KEY ='+IntToStr(ComboBox3.ItemIndex+1)+
              ' AND KIND_KEY=3 '+
              ' AND REPORT_DATE='''+ ComboBox4.Text+'''';
   IBQuery1.SQL.Add(sqlString);
   IBQuery1.Open;
   IBQuery1.Last;
   IBQuery1.First;
   if IBQuery1.RecordCount>0 then    //Блок №10 (Процедура сравнения)
    begin
     for i := 1 to 27 do     //Блок №11 (Процедура сравнения)
      TFA[i]:=IBQuery1.Fields[i+3].AsInteger;  //Блок №12 (Процедура сравнения)
    end;

    //Выдача в сетку
    //Блок №13 (Процедура сравнения)
   for i:=1  to 27 do
    begin
     StringGrid1.Cells[1,i]:=IntToStr(Abs(VA[i]-TVA[i]));
     StringGrid1.Cells[2,i]:=IntToStr(Abs(MA[i]-TMA[i]));
     StringGrid1.Cells[3,i]:=IntToStr(Abs(FA[i]-TFA[i]));
    end;
  end;
end;

//Событие при закрытии формы
procedure TForm3.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 if IBDatabase1.Connected then
  IBDatabase1.Connected:=false;
end;

//Событие при создании формы
procedure TForm3.FormCreate(Sender: TObject);
var
 i:integer;
begin
 //Заполняем заглавие строк сетки
 StringGrid1.Cells[0,1]:='Подлежат диспансеризации дети  в возрасте 14 лет, (всего детей)';
 StringGrid1.Cells[0,2]:='Прошли диспансеризацию  дети  в возрасте 14 лет   (нарастающим итогом)';
 StringGrid1.Cells[0,3]:='Не прошли диспансеризацию  дети  в возрасте 14 лет, не охвачены по причинам (всего):';
 StringGrid1.Cells[0,4]:='В том числе:';
 StringGrid1.Cells[0,5]:='Отсутствие на момент проведения диспансе-ризации';
 StringGrid1.Cells[0,6]:='Другие*';
 StringGrid1.Cells[0,7]:='Количество детей, у которых выявлены заболевания в ходе диспансеризации (всего детей)';
 StringGrid1.Cells[0,8]:='в том числе: количество детей с заболеваниями,  зарегистрированные впервые в ходе проведения диспансеризации (всего детей)';
 StringGrid1.Cells[0,9]:='Количество детей с нарушениями  полового развития (всего детей)';
 StringGrid1.Cells[0,10]:='в том числе: количество детей с нарушениями  полового развития, зарегистрированными впервые  (всего детей)';
 StringGrid1.Cells[0,11]:='Количество девочек с нарушениями менструального цикла';
 StringGrid1.Cells[0,12]:='Количество детей без нарушений полового развития (всего детей)';
 StringGrid1.Cells[0,13]:='Всего зарегистрировано заболеваний(всего заболеваний)';
 StringGrid1.Cells[0,14]:='Из числа зарегистрированных заболеваний выявлено впервые (всего заболеваний)';
 StringGrid1.Cells[0,15]:='Количество детей, получивших по результатам диспансеризации лечение(всего детей)';
 StringGrid1.Cells[0,16]:='Количество детей-инвалидов из числа прошедших диспансеризацию (всего детей)';
 StringGrid1.Cells[0,17]:='Количество детей, у которых выявлены заболевания в ходе диспансеризации (всего детей)';
 StringGrid1.Cells[0,18]:='Рекомендовано:';
 StringGrid1.Cells[0,19]:='- дополнительных консультаций';
 StringGrid1.Cells[0,20]:='- дополнительных обследований';
 StringGrid1.Cells[0,21]:='Распределение детей в возрасте 14 лет, прошедших диспансеризацию, по группам здоровья:';
 StringGrid1.Cells[0,22]:='1 группа здоровья (абс. число)';
 StringGrid1.Cells[0,23]:='2 группа здоровья (абс. число)';
 StringGrid1.Cells[0,24]:='3 группа здоровья (абс. число)';
 StringGrid1.Cells[0,25]:='4 группа здоровья (абс. число)';
 StringGrid1.Cells[0,26]:='5 группа здоровья (абс. число)';
 //Заполняем заглавие столбцов сетки
 StringGrid1.Cells[0,0]:='Название';
 StringGrid1.Cells[1,0]:='Всего:';
 StringGrid1.Cells[2,0]:='Мальчиков:';
 StringGrid1.Cells[3,0]:='Девочек:';
 //Подключение к базе данных
 IBDatabase1.Params.Clear();
 IBDatabase1.DatabaseName:=ExtractFilePath(Application.ExeName)+'BASE.FDB';
 IBDatabase1.Params.Add('User_name=SYSDBA');
 IBDatabase1.Params.Add('Password=masterkey');
 IBDatabase1.Params.Add('lc_ctype=WIN1251');
 IBDatabase1.Connected:=true;
 //Заполнение списка ЛПУ
 IBTable1.TableName:='LPU';
 IBTable1.Active:=True;
 IBTable1.Last;
 IBTable1.First;
 for i:=1 to IBTable1.RecordCount do
  begin
   ComboBox1.Items.Add(IBTable1.Fields[2].AsString);
   ComboBox3.Items.Add(IBTable1.Fields[2].AsString);
   ComboBox5.Items.Add(IBTable1.Fields[2].AsString);
   IBTable1.Next;
  end;
 IBTable1.Active:=False;
 RefreshReportsCombo; //Обновление полей со списком отчетный месяц
end;


//Процедура импорта с MS Word
procedure TForm3.Import(var importFileName:String);
var
 tempNum:String;
 App: Variant;
 i:integer;
 j:integer;
 LPUCode:String;
 LPUKey:Integer;
 tempStr: String;
 Error: Integer;
 f:TextFile;
 textFileName:String;
 sqlString:String;
begin
  //Достаем код ЛПУ  Блок №2 (Процедура Import)
  tempStr:=ExtractFileName(importFileName);
  LPUCode:=tempStr[8]+tempStr[9]+tempStr[10]+tempStr[11];
  Error:=0;
  //Открываем документ
  App := CreateOleObject('Word.Application');
  App.Documents.Open(importFileName);
  App.Visible:=false;
  for i := 1 to 27 do     //Блок №3 (Процедура Import)
   begin
    tempNum:=App.ActiveDocument.Tables.Item(1).Cell(i+2,3).Range.Text; //Блок №4 (Процедура Import)
    if Clean_String(tempNum)<>'' then //Блок №5 (Процедура Import)
     VA[i]:=Trunc(StrToInt(Clean_String(tempNum))) //Блок №6 (Процедура Import)
    else
     VA[i]:=0; //Блок №7 (Процедура Import)

    tempNum:=App.ActiveDocument.Tables.Item(1).Cell(i+2,4).Range.Text; //Блок №8 (Процедура Import)
    if Clean_String(tempNum)<>'' then    //Блок №9 (Процедура Import)
     MA[i]:=Trunc(StrToInt(Clean_String(tempNum))) //Блок №10 (Процедура Import)
    else
     MA[i]:=0; //Блок №11 (Процедура Import)

    tempNum:=App.ActiveDocument.Tables.Item(1).Cell(i+2,5).Range.Text; //Блок №12 (Процедура Import)
    if Clean_String(tempNum)<>'' then //Блок №13 (Процедура Import)
     FA[i]:=Trunc(StrToInt(Clean_String(tempNum))) //Блок №14 (Процедура Import)
    else
     FA[i]:=0;  //Блок №15 (Процедура Import)
   end;
  App:= UnAssigned;
  //Открываем лог файл для записи
  //Блок №16 (Процедура Import)
  textFileName:=ExtractFilePath(Application.ExeName)+DateToStr(Date)+'.txt';
  AssignFile(f,textFileName);
  //Блок №17 (Процедура Import)
  if not FileExists(textFileName) then
   begin
    //Блок №18 (Процедура Import)
    Rewrite(f);
    CloseFile(f);
   end;
  //Блок №19 (Процедура Import)
  Append(f);
  //Блок №20 (Процедура Import)
  for i:=7 to 15 do
   begin
    //Блок №21 (Процедура Import)
    if (VA[2]<VA[i]) or (MA[2]<MA[i]) or (FA[2]<FA[i])  then
     begin
      Error:=Error+1;     //Блок №22 (Процедура Import)
      writeln(f,'Количеcтво ошибок:'+InttoStr(Error));  //Блок №23 (Процедура Import)
    end;
   end;

  //Блок №24 (Процедура Import)
   if VA[22]<>VA[23]+VA[24]+VA[25]+VA[26]+VA[27] then
    begin
      Error:=Error+1; //Блок №25 (Процедура Import)
      writeln(f,'Количеcтво ошибок:'+InttoStr(Error)); //Блок №26 (Процедура Import)
    end;

    //Блок №27 (Процедура Import)
    if MA[22]<>MA[23]+MA[24]+MA[25]+MA[26]+MA[27] then
    begin
      Error:=Error+1; //Блок №28 (Процедура Import)
      writeln(f,'Количеcтво ошибок:'+InttoStr(Error)); //Блок №29 (Процедура Import)
    end;

    //Блок №30 (Процедура Import)
    if FA[22]<>FA[23]+FA[24]+FA[25]+FA[26]+FA[27] then
    begin
      Error:=Error+1; //Блок №31 (Процедура Import)
      writeln(f,'Количеcтво ошибок:'+IntToStr(Error));//Блок №32 (Процедура Import)
    end;

    //Блок №33 (Процедура Import)
    if VA[1]=FA[1]+MA[1] then
     begin
      Error:=Error+1; //Блок №34 (Процедура Import)
      writeln(f,'Количеcтво ошибок:'+IntToStr(Error));  //Блок №35 (Процедура Import)
     end;

     //Блок №36 (Процедура Import)
     if VA[3]=FA[3]+MA[3] then
     begin
      Error:=Error+1; //Блок №37 (Процедура Import)
      writeln(f,'Количеcтво ошибок:'+IntToStr(Error)); //Блок №38 (Процедура Import)
     end;

     //Блок №39 (Процедура Import)
     if VA[4]=FA[4]+MA[4] then
     begin
      Error:=Error+1; //Блок №40 (Процедура Import)
      writeln(f,'Количеcтво ошибок:'+IntToStr(Error)); //Блок №41 (Процедура Import)
     end;

      //Блок №42 (Процедура Import)
      if VA[5]=FA[5]+MA[5] then
     begin
      Error:=Error+1; //Блок №43 (Процедура Import)
      writeln(f,'Количеcтво ошибок:'+IntToStr(Error)); //Блок №44 (Процедура Import)
     end;

     //Блок №45 (Процедура Import)
     if VA[6]=FA[6]+MA[6] then
     begin
      Error:=Error+1; //Блок №46 (Процедура Import)
      writeln(f,'Количеcтво ошибок:'+IntToStr(Error)); //Блок №47 (Процедура Import)
     end;

    //Блок №48 (Процедура Import)
    if VA[11]=FA[11]+MA[11] then
     begin
      Error:=Error+1; //Блок №49 (Процедура Import)
      writeln(f,'Количеcтво ошибок:'+IntToStr(Error)); //Блок №50 (Процедура Import)
     end;
     CloseFile(f); //Блок №51 (Процедура Import)

  //Запись в базу
  if (Error=0) then   //Блок №52 (Процедура Import)
   begin
    //Блок №53 (Процедура Import)
    //Вытаскиваем ключ ЛПУ
    IBQuery1.Close;
    IBQuery1.SQL.Clear;
    sqlString:='SELECT LPU_KEY FROM LPU WHERE LPU_CODE ='''+LPUCode+'''';
    IBQuery1.SQL.Add(sqlString);
    IBQuery1.Open;
    if IBQuery1.RecordCount>0 then
     LPUKey:=IBQuery1.Fields[0].AsInteger
    else
     LPUKey:=1;
    //Удаляем дубликат отчета (перезаписывание)
    IBQuery1.Close;
    IBQuery1.SQL.Clear;
    sqlString:='DELETE  FROM BASE WHERE LPU_KEY ='+IntToStr(LPUKey)+
               ' and REPORT_DATE='''+ DateToStr(Date)+'''';
    IBQuery1.SQL.Add(sqlString);
    IBQuery1.ExecSQL;
    //Сохраняем записи в таблицу
    IBTable1.TableName:='BASE';
    IBTable1.Active:=True;
    for i:= 1 to 3 do
      begin
       IBTable1.Insert;
       IBTable1.Fields[1].AsInteger:=LPUKey;
       IBTable1.Fields[2].AsInteger:=i;
       IBTable1.Fields[3].AsString:=DateToStr(Date);
       for j := 1 to 27 do
        begin
          if i=1 then
           IBTable1.Fields[j+3].AsInteger:=VA[j]
          else if i=2 then
           IBTable1.Fields[j+3].AsInteger:=MA[j]
          else if i=3 then
           IBTable1.Fields[j+3].AsInteger:=FA[j];
        end;
       IBTable1.Fields[31].AsDateTime:=Date;
       IBTable1.Post;
      end;
      RefreshReportsCombo; //Обновление полей со списком отчетный месяц
      ShowMessage('Сохранение в базу отчета '+DateToStr(Date)+' ЛПУ Код:'+LPUCode);
   end
  else
   ShowMessage('При импорте возникли ошибки! Лог файл:'+textFileName);   //Блок №54 (Процедура Import)
end;

//Чистка строки от посторонних символов
function TForm3.Clean_String(var str:String):String;
var
 tmpStr:String;
 i:Integer;
begin
 for i:= 1 to Length(str) do
  if (ord(str[i])>48) and (ord(str[i])<58) then
   tmpStr:=tmpStr+str[i];
 Clean_String:=tmpStr;
end;


//Обновление полей со списком отчетный месяц
procedure TForm3.RefreshReportsCombo;
var
 sqlString:String;
 i:Integer;
begin
  IBQuery1.Close;
  IBQuery1.SQL.Clear;
  sqlString:='SELECT BASE.REPORT_DATE FROM BASE GROUP BY BASE.REPORT_DATE';
  IBQuery1.SQL.Add(sqlString);
  IBQuery1.Open;
  IBQuery1.Last;
  IBQuery1.First;
  ComboBox2.Items.Clear;
  ComboBox4.Items.Clear;
  for i := 1 to IBQuery1.RecordCount do
   begin
    ComboBox2.Items.Add(IBQuery1.Fields[0].AsString);
    ComboBox4.Items.Add(IBQuery1.Fields[0].AsString);
   end;
end;

end.
