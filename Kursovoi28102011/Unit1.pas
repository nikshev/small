unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, IBDatabase, DB, IBCustomDataSet, IBTable, IBQuery, ComObj,
  DBTables;

type
  TForm3 = class(TForm)
    GroupBox1: TGroupBox;
    ComboBox1: TComboBox;
    ListBox1: TListBox;
    Button1: TButton;
    GroupBox2: TGroupBox;
    Label2: TLabel;
    ComboBox2: TComboBox;
    Label3: TLabel;
    Edit1: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Edit3: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit7: TEdit;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Label12: TLabel;
    Edit8: TEdit;
    Label13: TLabel;
    Edit9: TEdit;
    Button2: TButton;
    Button3: TButton;
    GroupBox3: TGroupBox;
    Label14: TLabel;
    Edit10: TEdit;
    Button4: TButton;
    Button5: TButton;
    GroupBox4: TGroupBox;
    CheckBox1: TCheckBox;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    Button9: TButton;
    OpenDialog1: TOpenDialog;
    IBDatabase1: TIBDatabase;
    IBTableLPU: TIBTable;
    IBTransaction1: TIBTransaction;
    IBTableBase: TIBTable;
    IBQuery1: TIBQuery;
    DateTimePicker3: TDateTimePicker;
    Label1: TLabel;
    DateTimePicker4: TDateTimePicker;
    Label15: TLabel;
    Edit2: TEdit;
    Button10: TButton;
    SaveDialog1: TSaveDialog;
    IBQuery2: TIBQuery;
    Label16: TLabel;
    DateTimePicker5: TDateTimePicker;
    DateTimePicker6: TDateTimePicker;
    procedure Button4Click(Sender: TObject); //Диалог открытия файла для загрузки данных
    procedure FormCreate(Sender: TObject);  //Создание формы
    procedure FormClose(Sender: TObject; var Action: TCloseAction);  //Закрытие формы
    procedure Button7Click(Sender: TObject); //Новая запись
    procedure Button6Click(Sender: TObject); //Сохранение записи
    procedure Button8Click(Sender: TObject); //Удаление записи
    procedure Button1Click(Sender: TObject); //Поиск по ЛПУ
    procedure ListBox1DblClick(Sender: TObject); //По двойному клику на список, выводим карточку врача
    procedure Button2Click(Sender: TObject); //Поиск по фамилии
    procedure Button3Click(Sender: TObject);  //Поиск по категории
    procedure Import(var fileName:String);   //Импорт данных
    procedure Button5Click(Sender: TObject);   //По нажатию кнопки делаем импорт
    procedure Button9Click(Sender: TObject);   //По нажатию кнопки делаем экспорт
    procedure ExportData(var ExpTable1,ExpTable2,ExpTable3:boolean; dbfFileName:String);
    procedure Button10Click(Sender: TObject); //Экспорт данных
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}

//Открытие диалога сохранения
procedure TForm3.Button10Click(Sender: TObject);
begin
 SaveDialog1.InitialDir:=ExtractFilePath(Application.ExeName);
 if SaveDialog1.Execute then
  Edit2.Text:=SaveDialog1.FileName;
end;

//Поиск по ЛПУ
procedure TForm3.Button1Click(Sender: TObject);
var
sqlString:String;
i:Integer;
begin
 if ComboBox1.ItemIndex>-1 then
  begin
    IBQuery1.Close;
    IBQuery1.SQL.Clear;
    sqlString:='SELECT FIO FROM BASE WHERE KOD='+IntToStr(ComboBox1.ItemIndex+1);
    IBQuery1.SQL.Add(sqlString);
    IBQuery1.Open; //Вытаскиваем колличество изменений
    IBQuery1.Last;
    IBQuery1.First;
    ListBox1.Clear;
    for i := 1 to IBQuery1.RecordCount do
     begin
      ListBox1.Items.Add(IBQuery1.Fields[0].AsString);
      IBQuery1.Next;
     end;
  end
   else
    ShowMessage('Заполните поле "ЛПУ"');
end;

//Поиск по фамилии
procedure TForm3.Button2Click(Sender: TObject);
var
sqlString:String;
i:Integer;
begin
 if Edit8.Text<>'' then
  begin
   IBQuery1.Close;
   IBQuery1.SQL.Clear;
   sqlString:='SELECT FIO FROM BASE WHERE FIO Like('''+'%'+Edit8.Text+'%'+''''+')';
   IBQuery1.SQL.Add(sqlString);
   IBQuery1.Open; //Вытаскиваем колличество изменений
   IBQuery1.Last;
   IBQuery1.First;
   ListBox1.Clear;
   for i := 1 to IBQuery1.RecordCount do
    begin
     ListBox1.Items.Add(IBQuery1.Fields[0].AsString);
     IBQuery1.Next;
    end;
  end
 else
  ShowMessage('Заполните поле "ФИО" в поиске');
end;

//Поиск по категории
procedure TForm3.Button3Click(Sender: TObject);
var
 sqlString:String;
 i:Integer;
begin
 if Edit9.Text<>'' then
  begin
   IBQuery1.Close;
   IBQuery1.SQL.Clear;
   sqlString:='SELECT FIO FROM BASE WHERE KAT Like('''+'%'+Edit9.Text+'%'+''''+')';
   IBQuery1.SQL.Add(sqlString);
   IBQuery1.Open; //Вытаскиваем колличество изменений
   IBQuery1.Last;
   IBQuery1.First;
   ListBox1.Clear;
   for i := 1 to IBQuery1.RecordCount do
    begin
     ListBox1.Items.Add(IBQuery1.Fields[0].AsString);
     IBQuery1.Next;
    end;
  end
 else
  ShowMessage('Заполните поле "ФИО" в поиске');
end;

//Диалог открытия файла
procedure TForm3.Button4Click(Sender: TObject);
begin
 OpenDialog1.InitialDir:=ExtractFilePath(Application.ExeName);
 if OpenDialog1.Execute then
  Edit10.Text:=OpenDialog1.FileName;
end;

//По нажатию кнопки вызываем процедуру импорта
procedure TForm3.Button5Click(Sender: TObject);
var
 fileName:String;
begin
 fileName:=Edit10.Text;   //Блок №2 (Главной программы)
 Import(fileName);       //Блок №3 (Главной программы)
end;

//Сохранение записи
procedure TForm3.Button6Click(Sender: TObject);
var
 sqlString:String;
begin
 //Сохраняем данные
 if ComboBox2.ItemIndex>-1 then
  begin
   if Edit1.Text<>'' then
     begin
      if DateTimePicker3.DateTime<>Date then
       begin
        if Edit5.Text<>'' then
         begin
          if Edit6.Text<>'' then
           begin
            if Edit7.Text<>'' then
             begin
              //Запрос на дубль работника
              IBQuery1.Close;
              IBQuery1.SQL.Clear;
              sqlString:='SELECT NIZ FROM BASE WHERE KOD='+IntToStr(ComboBox2.ItemIndex+1)+
                         ' AND FIO ='''+Edit1.Text+'''';
              IBQuery1.SQL.Add(sqlString);
              IBQuery1.Open; //Вытаскиваем колличество изменений
              if IBQuery1.RecordCount>0 then
                Edit3.Text:=FloatToStr(IBQuery1.Fields[0].AsFloat+1); //Добавляем единицу к изменениям
              //Удаляем запись
              IBQuery1.Close;
              IBQuery1.SQL.Clear;
              sqlString:='DELETE FROM BASE WHERE KOD='+IntToStr(ComboBox2.ItemIndex+1)+
                         ' AND FIO ='''+Edit1.Text+'''';
              IBQuery1.SQL.Add(sqlString);
              IBQuery1.ExecSQL; //Удаляем такого работниука если он есть
              //Вставляем карточку по новой
              IBTableBase.Active:=True;
              IBTableBase.Insert;
              IBTableBase.Fields[0].AsInteger:=ComboBox2.ItemIndex+1; //Код ЛПУ
              IBTableBase.Fields[1].AsString:=Edit1.Text; //ФИО врача
              IBTableBase.Fields[2].AsDateTime:=DateTimePicker2.DateTime;//Дата рождения
              IBTableBase.Fields[3].AsFloat:=StrToFloat(Edit3.Text); //Номер изменения записи в сведениях о враче
              IBTableBase.Fields[4].AsDateTime:=Date; //Дата изменения в сведениях о враче
              IBTableBase.Fields[5].AsString:=Edit5.Text;//Должность
              IBTableBase.Fields[6].AsString:=Edit6.Text; //Специальность
              IBTableBase.Fields[7].AsString:=Edit7.Text;; //Категория
              IBTableBase.Fields[8].AsDateTime:=DateTimePicker1.DateTime; //Дата возникновения права на выписку льготных рецептов
              IBTableBase.Fields[9].AsDateTime:=DateTimePicker2.DateTime; //Дата прекращения права на выписку льготных рецептов
              IBTableBase.Post;
              ShowMessage('Сохранение карточки врача:'+Edit1.Text+' прошло успешно!');
             end
              else
               ShowMessage('Заполните поле "Категория"');
           end
           else
            ShowMessage('Заполните поле "Специальность"');
         end
         else
          ShowMessage('Заполните поле "Должность"');
       end

       else
        ShowMessage('Заполните поле "Дата рождения"');
     end
    else
      ShowMessage('Заполните поле "ФИО"');
  end
   else
     ShowMessage('Заполните поле "ЛПУ"');
end;

//Новая запись
procedure TForm3.Button7Click(Sender: TObject);
begin
 //Чистим все поля
 ComboBox2.ItemIndex:=-1;
 Edit1.Text:='';
 Edit3.Text:='1';
 Edit5.Text:='';
 Edit6.Text:='';
 Edit7.Text:='';
 DateTimePicker1.DateTime:=Date;
 DateTimePicker2.DateTime:=Date;
 DateTimePicker3.DateTime:=Date;
 DateTimePicker4.DateTime:=Date;
end;

//Удаление записи
procedure TForm3.Button8Click(Sender: TObject);
var
 sqlString:String;
begin
 //Удаляем запись
 if ComboBox2.ItemIndex>-1 then
  begin
   if Edit1.Text<>'' then
    begin
     IBQuery1.Close;
     IBQuery1.SQL.Clear;
     sqlString:='DELETE FROM BASE WHERE KOD='+IntToStr(ComboBox2.ItemIndex+1)+
                ' AND FIO ='''+Edit1.Text+'''';
     IBQuery1.SQL.Add(sqlString);
     IBQuery1.ExecSQL; //Удаляем работниука
    end
     else
      ShowMessage('Заполните поле "ФИО"');
  end
   else
    ShowMessage('Заполните поле "ЛПУ"');
end;

//По нажатию кнопки делаем экспорт
procedure TForm3.Button9Click(Sender: TObject);
 var
 ExpTable1,ExpTable2,ExpTable3:boolean;
begin
 //Блок №4 (Главной программы)
 ExpTable1:=CheckBox1.Checked;
 ExpTable2:=CheckBox2.Checked;
 ExpTable3:=CheckBox3.Checked;
 ExportData(ExpTable1,ExpTable2,ExpTable3, Edit2.Text);    //Блок №5 (Главной программы)
end;

//Закрытие формы
procedure TForm3.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if IBDatabase1.Connected then
   IBDatabase1.Connected:=false
end;

//Создание формы
procedure TForm3.FormCreate(Sender: TObject);
var
 i:integer;
begin
 //Подключение к базе данных
 IBDatabase1.Params.Clear();
 IBDatabase1.DatabaseName:=ExtractFilePath(Application.ExeName)+'BASE.FDB';
 IBDatabase1.Params.Add('User_name=SYSDBA');
 IBDatabase1.Params.Add('Password=masterkey');
 IBDatabase1.Params.Add('lc_ctype=WIN1251');
 IBDatabase1.Connected:=true;
 //Заполнение списка ЛПУ
 IBTableLPU.Active:=True;
 IBTableLPU.Last;
 IBTableLPU.First;
 for i := 1 to IBTableLPU.RecordCount do
  begin
   //Поле со списком на поиске
   ComboBox1.Items.Add(IBTableLPU.Fields[1].AsString+
                      '(Код:'+IBTableLPU.Fields[2].AsString+')');
   //Поле со списком на карточке врача
   ComboBox2.Items.Add(IBTableLPU.Fields[1].AsString+
                      '(Код:'+IBTableLPU.Fields[2].AsString+')');
   IBTableLPU.Next;
  end;
 IBTableLPU.Active:=False;

end;

//По двойному клику на список, выводим карточку врача
procedure TForm3.ListBox1DblClick(Sender: TObject);
var
 sqlString:String;
begin
 if (ListBox1.ItemIndex>-1) then
  begin
   ListBox1.Items[ListBox1.ItemIndex];
   IBQuery1.Close;
   IBQuery1.SQL.Clear;
   sqlString:='SELECT * FROM BASE WHERE FIO ='''+ListBox1.Items[ListBox1.ItemIndex]+'''';
   IBQuery1.SQL.Add(sqlString);
   IBQuery1.Open; //Удаляем такого работниука если он есть
   if IBQuery1.RecordCount>0 then
    begin
      ComboBox2.ItemIndex:=IBQuery1.Fields[0].AsInteger-1; //Код ЛПУ
      Edit1.Text:=IBQuery1.Fields[1].AsString; //ФИО врача
      DateTimePicker3.DateTime:=IBQuery1.Fields[2].AsDateTime;//Дата рождения
      Edit3.Text:=FloatToStr(IBQuery1.Fields[3].AsFloat); //Номер изменения записи в сведениях о враче
      DateTimePicker4.DateTime:=IBQuery1.Fields[4].AsDateTime; //Дата изменения в сведениях о враче
      Edit5.Text:=IBQuery1.Fields[5].AsString;//Должность
      Edit6.Text:=IBQuery1.Fields[6].AsString; //Специальность
      Edit7.Text:=IBQuery1.Fields[7].AsString; //Категория
      DateTimePicker1.DateTime:=IBQuery1.Fields[8].AsDateTime; //Дата возникновения права на выписку льготных рецептов
      DateTimePicker2.DateTime:=IBQuery1.Fields[9].AsDateTime; //Дата прекращения права на выписку льготных рецептов
    end;
  end;
end;

//Импорт данных
procedure TForm3.Import(var fileName:String); //Импорт данных
var
 sqlString:String;
 XL,Sheet:Variant;
 i:Integer;
 EndImport:boolean;
 KOD:Integer;
 FIO:String;
 DB:TDateTime;
 NIZ:Real;
 NAME_DL:String;
 NAME_SP:String;
 KAT:String;
 DBEG:TDateTime;
 DEND:TDateTime;
 LPUCODE:Integer;
 tmpStr:String;
begin
 if fileName<>'' then   //Блок №2  (Процедура Import)
  begin
   //Открываем Excel файл
   XL:=CreateOLEObject('Excel.Application');
   XL.WorkBooks.Open(fileName);
   Sheet:=XL.ActiveWorkBook.Sheets[1];
   //Находи ЛПУ
   //Блок №3  (Процедура Import)
   LPUCODE:=-1;
   //Блок №4  (Процедура Import)
   IBQuery1.Close;
   IBQuery1.SQL.Clear;
   sqlString:='SELECT LPU_KEY FROM LPU WHERE LPU_NAME ='''+Sheet.Cells[4,3].Value+'''';
   IBQuery1.SQL.Add(sqlString);
   IBQuery1.Open; //Вытаскиваем колличество изменений
   IBQuery1.Last;
   IBQuery1.First;
   if IBQuery1.RecordCount>0 then
    LPUCODE:=  IBQuery1.Fields[0].AsInteger;
   //Начальные значения
   //Блок №5   (Процедура Import)
   i:=7;
   EndImport:=false;
   if LPUCODE>-1 then      //Блок №6  (Процедура Import)
    begin
     while not EndImport do  //Блок №7  (Процедура Import)
      begin
       tmpStr:=Sheet.Cells[i,1].Value;    //Блок №8  (Процедура Import)
       if tmpStr<>'' then          //Блок №9  (Процедура Import)
        begin
         //Блок №10   (Процедура Import)
         KOD:= LPUCODE;
         FIO:=Sheet.Cells[i,2].Value;
         DB:=StrToDate(Sheet.Cells[i,3].Value);
         NAME_DL:=Sheet.Cells[i,4].Value;
         NAME_SP:=Sheet.Cells[i,5].Value;
         KAT:=Sheet.Cells[i,6].Value;
         DBEG:=StrToDate(Sheet.Cells[i,7].Value);
         DEND:=StrToDate(Sheet.Cells[i,8].Value);

         //Запрос на дубль работника
         IBQuery1.Close;
         IBQuery1.SQL.Clear;
         sqlString:='SELECT NIZ FROM BASE WHERE KOD='+IntToStr(LPUCODE)+
                    ' AND FIO ='''+FIO+'''';
         IBQuery1.SQL.Add(sqlString);
         IBQuery1.Open; //Вытаскиваем колличество изменений
         if IBQuery1.RecordCount>0 then
           NIZ:=IBQuery1.Fields[0].AsFloat+1 //Добавляем единицу к изменениям
         else
          NIZ:=1;

         //Удаляем запись
         IBQuery1.Close;
         IBQuery1.SQL.Clear;
         sqlString:='DELETE FROM BASE WHERE KOD='+IntToStr(LPUCODE)+
                    ' AND FIO ='''+FIO+'''';
         IBQuery1.SQL.Add(sqlString);
         IBQuery1.ExecSQL; //Удаляем такого работниука если он есть

          //Блок №11   (Процедура Import)
         //Вставляем карточку по новой
         IBTableBase.Active:=True;
         IBTableBase.Insert;
         IBTableBase.Fields[0].AsInteger:=KOD; //Код ЛПУ
         IBTableBase.Fields[1].AsString:=FIO; //ФИО врача
         IBTableBase.Fields[2].AsDateTime:=DB;//Дата рождения
         IBTableBase.Fields[3].AsFloat:=NIZ; //Номер изменения записи в сведениях о враче
         IBTableBase.Fields[4].AsDateTime:=Date; //Дата изменения в сведениях о враче
         IBTableBase.Fields[5].AsString:=NAME_DL;//Должность
         IBTableBase.Fields[6].AsString:=NAME_SP; //Специальность
         IBTableBase.Fields[7].AsString:=KAT; //Категория
         IBTableBase.Fields[8].AsDateTime:=DBEG; //Дата возникновения права на выписку льготных рецептов
         IBTableBase.Fields[9].AsDateTime:=DEND; //Дата прекращения права на выписку льготных рецептов
         IBTableBase.Post;
         i:=i+1;      //Блок №12   (Процедура Import)
        end
        else
         begin
          EndImport:=True;   //Блок №13  (Процедура Import)
          ShowMessage('Сохранение карточеки врачей по ЛПУ:'+Sheet.Cells[4,3].Value+' прошло успешно!');
          //Ставим что врачи утратили право выписки
          IBQuery1.Close;
          IBQuery1.SQL.Clear;
          sqlString:='UPDATE BASE SET DEND='''+DateToStr(Date)+''''+
                     ' WHERE DIZ<>'''+DateToStr(Date)+''''+
                     ' AND KOD='+IntToStr(LPUCODE);
          IBQuery1.SQL.Add(sqlString);
          IBQuery1.ExecSQL; //Удаляем такого работниука если он есть
          IBTransaction1.Commit;
         end;
      end;
    end
     else
      ShowMessage('Неправильное название ЛПУ в файле: '+fileName);

   //Закрываем Excel файл
   Sheet:= UnAssigned;
   XL:= UnAssigned;
  end
 else
  ShowMessage('Не заполнено имя файла для импорта!');
end;

//Экспорт данных
procedure TForm3.ExportData(var ExpTable1,ExpTable2,ExpTable3:boolean; dbfFileName:String);
var
 sqlString:String;
 XL,Sheet:Variant;
 i,j:Integer;
 RowIndex:Integer;
 LPUName:String;
 TDBFTable:TTable;
begin
   //Экспорт в Excel таблицы 1 (Количество врачей имеющих право выписки льготных лекарственных препаратов на )
  if ExpTable1 then    //Блок №2  (Процедура ExportData)
   begin
    //Открываем Excel файл
    XL:=CreateOLEObject('Excel.Application');
    XL.WorkBooks.Add;
    XL.Visible:=True;
    Sheet:=XL.ActiveWorkBook.Sheets[1];
    Sheet.Cells[1,1]:='Количество врачей имеющих право выписки льготных лекарственных препаратов на '+DateToStr(Date);
    Sheet.Cells[2,1]:='Наименование ЛПУ ';
    Sheet.Cells[2,2]:='Количество врачей';
    IBQuery1.Close;
    IBQuery1.SQL.Clear;
    sqlString:='SELECT KOD FROM BASE GROUP BY KOD';
    IBQuery1.SQL.Add(sqlString);
    IBQuery1.Open; //Вытаскиваем колличество изменений
    IBQuery1.Last;
    IBQuery1.First;
    for i := 1 to IBQuery1.RecordCount do //Блок №3 (Процедура ExportData)
      begin
       //Находи ЛПУ
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT LPU_NAME FROM LPU WHERE LPU_KEY ='+IntToStr(IBQuery1.Fields[0].AsInteger);
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        if IBQuery2.RecordCount>0 then    //Блок №4  (Процедура ExportData)
         Sheet.Cells[i+2,1]:= IBQuery2.Fields[0].AsString; //Блок №5   (Процедура ExportData)

        //Количество врачей иеющих право выписки
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT * FROM BASE WHERE KOD ='+IntToStr(IBQuery1.Fields[0].AsInteger)+
                   ' AND DEND>'''+DateToStr(Date)+'''';
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        IBQuery2.Last;
        IbQuery2.First;
        if IBQuery2.RecordCount>0 then        //Блок №6 (Процедура ExportData)
         Sheet.Cells[i+2,2]:= IntToStr(IBQuery2.RecordCount);    //Блок №7  (Процедура ExportData)
      IBQuery1.Next; //Блок №8 (Процедура ExportData)
      end;
    //Закрываем Excel файл
   Sheet:= UnAssigned;
   XL:= UnAssigned;
   end;

  //Экспорт в Excel таблицы 2 (Изменение персонала осуществляющего выписку льготных лекарственных препаратов)
  if ExpTable2 then //Блок №9  (Процедура ExportData)
   begin
    //Открываем Excel файл
    XL:=CreateOLEObject('Excel.Application');
    XL.WorkBooks.Add;
    XL.Visible:=True;
    Sheet:=XL.ActiveWorkBook.Sheets[1];
    Sheet.Cells[1,1]:='Изменение персонала осуществляющего '+
                      'выписку льготных лекарственных препаратов с '+DateToStr(DateTimePicker5.DateTime)+
                      ' по '+DateToStr(DateTimePicker6.DateTime);
    Sheet.Cells[2,1]:='Наименование ЛПУ ';
    Sheet.Cells[2,2]:='Количество новых врачей за период';
    Sheet.Cells[2,3]:='Количество врачей утративших право на выписку льготных лекарственных препаратов';
    Sheet.Cells[2,4]:='Среднее количество врачей, осуществляющих выписку льготных лекарственных препаратов';
    IBQuery1.Close;
    IBQuery1.SQL.Clear;
    sqlString:='SELECT KOD FROM BASE GROUP BY KOD';
    IBQuery1.SQL.Add(sqlString);
    IBQuery1.Open; //Вытаскиваем колличество изменений
    IBQuery1.Last;
    IBQuery1.First;
    for i := 1 to IBQuery1.RecordCount do   //Блок №10 (Процедура ExportData)
       begin
        //Находи ЛПУ
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT LPU_NAME FROM LPU WHERE LPU_KEY ='+IntToStr(IBQuery1.Fields[0].AsInteger);
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        if IBQuery2.RecordCount>0 then   //Блок №11  (Процедура ExportData)
         Sheet.Cells[i+2,1]:= IBQuery2.Fields[0].AsString; //Блок №12 (Процедура ExportData)

        //Количество новых врачей
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT * FROM BASE WHERE KOD ='+IntToStr(IBQuery1.Fields[0].AsInteger)+
                   ' AND NIZ=1';
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        IBQuery2.Last;
        IbQuery2.First;
        if IBQuery2.RecordCount>0 then    //Блок №13 (Процедура ExportData)
         Sheet.Cells[i+2,2]:= IntToStr(IBQuery2.RecordCount);    //Блок №14 (Процедура ExportData)


         //Количество врачей утративших право выписки
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT * FROM BASE WHERE KOD ='+IntToStr(IBQuery1.Fields[0].AsInteger)+
                   ' AND (DEND<'''+DateToStr(DateTimePicker5.DateTime)+''''+
                   ' OR DEND='''+DateToStr(DateTimePicker5.DateTime)+''''+')';
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        IBQuery2.Last;
        IbQuery2.First;
        if IBQuery2.RecordCount>0 then    //Блок №15  (Процедура ExportData)
         Sheet.Cells[i+2,3]:= IntToStr(IBQuery2.RecordCount); //Блок №16   (Процедура ExportData)

        //Количество врачей иеющих право выписки
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT * FROM BASE WHERE KOD ='+IntToStr(IBQuery1.Fields[0].AsInteger)+
                   ' AND DEND>'''+DateToStr(DateTimePicker6.DateTime)+'''';
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        IBQuery2.Last;
        IbQuery2.First;
        if IBQuery2.RecordCount>0 then        //Блок №17 (Процедура ExportData)
         Sheet.Cells[i+2,4]:= IntToStr(IBQuery2.RecordCount); //Блок №18 (Процедура ExportData)
      IBQuery1.Next;     //Блок №19 (Процедура ExportData)
     end;
    //Закрываем Excel файл
   Sheet:= UnAssigned;
   XL:= UnAssigned;
   end;

  //Экспорт в Excel таблицы 3 (Реестр врачей осуществляющих выписку льготных лекарственных препаратов)
  if ExpTable3 then      //Блок №20   (Процедура ExportData)
   begin
   //Открываем Excel файл
    XL:=CreateOLEObject('Excel.Application');
    XL.WorkBooks.Add;
    XL.Visible:=True;
    Sheet:=XL.ActiveWorkBook.Sheets[1];
    Sheet.Cells[1,1]:='Реестр врачей осуществляющих выписку льготных лекарственных препаратов на '+DateToStr(Date);
    Sheet.Cells[2,1]:='Наименование ЛПУ ';
    Sheet.Cells[2,2]:='ФИО врача';
    IBQuery1.Close;
    IBQuery1.SQL.Clear;
    sqlString:='SELECT KOD FROM BASE GROUP BY KOD';
    IBQuery1.SQL.Add(sqlString);
    IBQuery1.Open; //Вытаскиваем колличество изменений
    IBQuery1.Last;
    IBQuery1.First;
    RowIndex:=3;      //Блок №21 (Процедура ExportData)
    for i := 1 to IBQuery1.RecordCount do //Блок №22    (Процедура ExportData)
     begin
        //Находи ЛПУ
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT LPU_NAME FROM LPU WHERE LPU_KEY ='+IntToStr(IBQuery1.Fields[0].AsInteger);
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        if IBQuery2.RecordCount>0 then        //Блок №23  (Процедура ExportData)
         LPUName:= IBQuery2.Fields[0].AsString;  //Блок №24 (Процедура ExportData)

        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT FIO FROM BASE WHERE KOD ='+IntToStr(IBQuery1.Fields[0].AsInteger)+
                   ' AND DEND>'''+DateToStr(Date)+'''';
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        IBQuery2.Last;
        IbQuery2.First;
        for j := 1 to IbQuery2.RecordCount do     //Блок №25 (Процедура ExportData)
         begin
          //Блок №26   (Процедура ExportData)
          Sheet.Cells[RowIndex,1]:=LPUName;
          Sheet.Cells[RowIndex,2]:=IBQuery2.Fields[0].AsString;
          RowIndex:=RowIndex+1; //Блок №27  (Процедура ExportData)
          IbQuery2.Next; //Блок №28  (Процедура ExportData)
         end;
      IBQuery1.Next;  //Блок №29  (Процедура ExportData)
     end;
    //Закрываем Excel файл
   Sheet:= UnAssigned;
   XL:= UnAssigned;
   end;

   //Выгрузка в dbf файл
   if dbfFileName<>'' then    //Блок №30 (Процедура ExportData)
    begin
     TDBFTable:=TTable.Create(nil);
     TDBFTable.TableType:=ttDBase;
     TDBFTable.TableLevel:=4;
     TDBFTable.DatabaseName:=ExtractFilePath(dbfFileName);
     TDBFTable.TableName:=ExtractFileName(dbfFileName);
     TDBFTable.FieldDefs.Add('KOD',ftFloat,0,false);
     TDBFTable.FieldDefs.Add('FIO',ftString,120,false);
     TDBFTable.FieldDefs.Add('DB',ftString,10,false);
     TDBFTable.FieldDefs.Add('NIZ',ftFloat,0,false);
     TDBFTable.FieldDefs.Add('DIZ',ftString,10,false);
     TDBFTable.FieldDefs.Add('NAME_DL',ftString,80,false);
     TDBFTable.FieldDefs.Add('NAME_SP',ftString,80,false);
     TDBFTable.FieldDefs.Add('KAT',ftString,30,false);
     TDBFTable.FieldDefs.Add('DBEG',ftString,10,false);
     TDBFTable.FieldDefs.Add('DEND',ftString,10,false);
     TDBFTable.CreateTable;
     IBQuery1.Close;
     IBQuery1.SQL.Clear;
     sqlString:='SELECT * FROM BASE';
     IBQuery1.SQL.Add(sqlString);
     IBQuery1.Open; //Вытаскиваем колличество изменений
     IBQuery1.Last;
     IBQuery1.First;
     TDBFTable.Active:=true;
     for i := 1 to IBQuery1.RecordCount do     //Блок №31  (Процедура ExportData)
      begin
       TDBFTable.Insert;
       //Блок №32   (Процедура ExportData)
       TDBFTable.Fields[0].AsFloat:=IBQuery1.Fields[0].AsFloat;
       TDBFTable.Fields[1].AsString:=IBQuery1.Fields[1].AsString;
       TDBFTable.Fields[2].AsString:=FormatDateTime('yyyy.mm.dd',IBQuery1.Fields[2].AsDateTime);
       TDBFTable.Fields[3].AsFloat:=IBQuery1.Fields[3].AsFloat;
       TDBFTable.Fields[4].AsString:=DateToStr(IBQuery1.Fields[4].AsDateTime);
       //Блок №33   (Процедура ExportData)
       TDBFTable.Fields[5].AsString:=IBQuery1.Fields[5].AsString;
       TDBFTable.Fields[6].AsString:=IBQuery1.Fields[6].AsString;
       TDBFTable.Fields[7].AsString:=IBQuery1.Fields[7].AsString;
       TDBFTable.Fields[8].AsString:=DateToStr(IBQuery1.Fields[8].AsDateTime);
       TDBFTable.Fields[9].AsString:=DateToStr(IBQuery1.Fields[9].AsDateTime);
       TDBFTable.Post;
       IBQuery1.Next; //Блок №34   (Процедура ExportData)
      end;
     TDBFTable.Active:=false;
     ShowMessage('Выгрузка данных в файл: '+dbfFileName+' завершена успешно!');
    end;
end;
end.
