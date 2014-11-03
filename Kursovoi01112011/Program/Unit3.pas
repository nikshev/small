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
    procedure Button2Click(Sender: TObject);    //����� ����� ��� �������
    procedure FormCreate(Sender: TObject);      //������� ��� �������� �����
    procedure Import(var importFileName:String);  //������ ������
    function Clean_String(var str:String):String; //������ ������ �� ����������� ��������
    procedure Button3Click(Sender: TObject);   //������� ������� ������ ��� �������
    procedure FormClose(Sender: TObject; var Action: TCloseAction); //������� ��� �������� �����
    procedure Button6Click(Sender: TObject);      //��������� ���������
    procedure Button5Click(Sender: TObject);    //������� �������� ������
    procedure Button4Click(Sender: TObject);  //����� ����� ��� ��������
    procedure RefreshReportsCombo;             //���������� ����� �� ������� �������� �����
    procedure Button1Click(Sender: TObject);   //����� � ���������� �����
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

//����� � ���������� �����
procedure TForm3.Button1Click(Sender: TObject);
var
 sqlString:String;
 i:Integer;
begin
 if (ComboBox1.ItemIndex>-1) and (ComboBox2.ItemIndex>-1) then    //���� �2 ��������� ������
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
   if IBQuery1.RecordCount>0 then   //���� �3 ��������� ������
    begin
     for i := 1 to 27 do    //���� �4 ��������� ������
      VA[i]:=IBQuery1.Fields[i+3].AsInteger;   //���� �5 ��������� ������
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
   if IBQuery1.RecordCount>0 then  //���� �6 ��������� ������
    begin
     for i := 1 to 27 do    //���� �7 ��������� ������
      MA[i]:=IBQuery1.Fields[i+3].AsInteger;  //���� �8 ��������� ������
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
   if IBQuery1.RecordCount>0 then   //���� �9 ��������� ������
    begin
     for i := 1 to 27 do   //���� �10 ��������� ������
      FA[i]:=IBQuery1.Fields[i+3].AsInteger; //���� �11 ��������� ������
    end;

  //������ � �����
  //���� �12 ��������� ������
  for i:=1  to 27 do
   begin
    StringGrid1.Cells[1,i]:=IntToStr(VA[i]);
    StringGrid1.Cells[2,i]:=IntToStr(MA[i]);
    StringGrid1.Cells[3,i]:=IntToStr(FA[i]);
   end;
  end;
end;

//����� ����� ��� �������
procedure TForm3.Button2Click(Sender: TObject);
begin
 OpenDialog1.InitialDir:=ExtractFilePath(Application.ExeName); //���� � ���������
 if OpenDialog1.Execute then
  Edit1.Text:=OpenDialog1.FileName;
end;

//������� ������� ������ ��� �������
procedure TForm3.Button3Click(Sender: TObject);
var
 importFileName:String;
begin
 importFileName:=Edit1.Text;
 Import(importFileName);
end;

//����� ����� ��� ��������
procedure TForm3.Button4Click(Sender: TObject);
begin
 OpenDialog2.InitialDir:=ExtractFilePath(Application.ExeName); //���� � ���������
 if OpenDialog2.Execute then
  Edit2.Text:=OpenDialog2.FileName;
end;

//������� �������� ������
procedure TForm3.Button5Click(Sender: TObject);
var
  i,j: Integer;
  App: Variant;
  sqlString:String;
  LPUKey:Integer;
  PERORT_DATE:String;
begin
 if Edit2.Text<>'' then //���� �2 ��������� �������� ������
  begin
  //������ ��������� ������
   for i := 1 to 27 do  //���� �3 ��������� �������� ������
    begin
    //���� �4 ��������� �������� ������
     TVA[i]:=0;
     TMA[i]:=0;
     TFA[i]:=0;
    end;
   //������ �� ���� ���� ������� �� ���� �����
   if (ComboBox5.ItemIndex>-1) then   //���� �5 ��������� �������� ������
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
    for i := 1 to IBQuery1.RecordCount do //���� �6 ��������� �������� ������
     begin
      //���� �7 ��������� �������� ������
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
      if IBQuery2.RecordCount>0 then  //���� �8 ��������� �������� ������
       begin
        for j := 1 to 27 do  //���� �9 ��������� �������� ������
         TVA[j]:=TVA[j]+IBQuery2.Fields[j+3].AsInteger;  //���� �10 ��������� �������� ������
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
      if IBQuery2.RecordCount>0 then   //���� �11 ��������� �������� ������
       begin
        for j := 1 to 27 do //���� �12 ��������� �������� ������
         TMA[j]:=TMA[j]+IBQuery2.Fields[j+3].AsInteger; //���� �13 ��������� �������� ������
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
      if IBQuery2.RecordCount>0 then   //���� �14 ��������� �������� ������
       begin
        for j := 1 to 27 do //���� �15 ��������� �������� ������
         TFA[j]:=TFA[j]+IBQuery2.Fields[j+3].AsInteger; //���� �16 ��������� �������� ������
       end;
       IBQuery1.Next;  //���� �17 ��������� �������� ������
     end;

    //���� �18 ��������� �������� ������
    //������� ������ � Word
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

//��������� ���������
procedure TForm3.Button6Click(Sender: TObject);
var
 sqlString:String;
 i:Integer;
begin

 if (ComboBox1.ItemIndex>-1) and (ComboBox2.ItemIndex>-1) and     //���� �2 (��������� ���������)
    (ComboBox3.ItemIndex>-1) and (ComboBox4.ItemIndex>-1) then
  begin
    //������ �� ���� ������� ������
    Button1Click(nil);   //���� �3 (��������� ���������)
   //������ �� ���� ������� ������
   IBQuery1.Close;
   IBQuery1.SQL.Clear;
   sqlString:='SELECT * FROM BASE WHERE LPU_KEY ='+IntToStr(ComboBox3.ItemIndex+1)+
              ' AND KIND_KEY=1 '+
              ' AND REPORT_DATE='''+ ComboBox4.Text+'''';
   IBQuery1.SQL.Add(sqlString);
   IBQuery1.Open;
   IBQuery1.Last;
   IBQuery1.First;
   if IBQuery1.RecordCount>0 then    //���� �4 (��������� ���������)
    begin
     for i := 1 to 27 do   //���� �5 (��������� ���������)
      TVA[i]:=IBQuery1.Fields[i+3].AsInteger;  //���� �6 (��������� ���������)
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
   if IBQuery1.RecordCount>0 then    //���� �7 (��������� ���������)
    begin
     for i := 1 to 27 do      //���� �8 (��������� ���������)
      TMA[i]:=IBQuery1.Fields[i+3].AsInteger;   //���� �9 (��������� ���������)
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
   if IBQuery1.RecordCount>0 then    //���� �10 (��������� ���������)
    begin
     for i := 1 to 27 do     //���� �11 (��������� ���������)
      TFA[i]:=IBQuery1.Fields[i+3].AsInteger;  //���� �12 (��������� ���������)
    end;

    //������ � �����
    //���� �13 (��������� ���������)
   for i:=1  to 27 do
    begin
     StringGrid1.Cells[1,i]:=IntToStr(Abs(VA[i]-TVA[i]));
     StringGrid1.Cells[2,i]:=IntToStr(Abs(MA[i]-TMA[i]));
     StringGrid1.Cells[3,i]:=IntToStr(Abs(FA[i]-TFA[i]));
    end;
  end;
end;

//������� ��� �������� �����
procedure TForm3.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 if IBDatabase1.Connected then
  IBDatabase1.Connected:=false;
end;

//������� ��� �������� �����
procedure TForm3.FormCreate(Sender: TObject);
var
 i:integer;
begin
 //��������� �������� ����� �����
 StringGrid1.Cells[0,1]:='�������� ��������������� ����  � �������� 14 ���, (����� �����)';
 StringGrid1.Cells[0,2]:='������ ���������������  ����  � �������� 14 ���   (����������� ������)';
 StringGrid1.Cells[0,3]:='�� ������ ���������������  ����  � �������� 14 ���, �� �������� �� �������� (�����):';
 StringGrid1.Cells[0,4]:='� ��� �����:';
 StringGrid1.Cells[0,5]:='���������� �� ������ ���������� ��������-�������';
 StringGrid1.Cells[0,6]:='������*';
 StringGrid1.Cells[0,7]:='���������� �����, � ������� �������� ����������� � ���� ��������������� (����� �����)';
 StringGrid1.Cells[0,8]:='� ��� �����: ���������� ����� � �������������,  ������������������ ������� � ���� ���������� ��������������� (����� �����)';
 StringGrid1.Cells[0,9]:='���������� ����� � �����������  �������� �������� (����� �����)';
 StringGrid1.Cells[0,10]:='� ��� �����: ���������� ����� � �����������  �������� ��������, ������������������� �������  (����� �����)';
 StringGrid1.Cells[0,11]:='���������� ������� � ����������� �������������� �����';
 StringGrid1.Cells[0,12]:='���������� ����� ��� ��������� �������� �������� (����� �����)';
 StringGrid1.Cells[0,13]:='����� ���������������� �����������(����� �����������)';
 StringGrid1.Cells[0,14]:='�� ����� ������������������ ����������� �������� ������� (����� �����������)';
 StringGrid1.Cells[0,15]:='���������� �����, ���������� �� ����������� ��������������� �������(����� �����)';
 StringGrid1.Cells[0,16]:='���������� �����-��������� �� ����� ��������� ��������������� (����� �����)';
 StringGrid1.Cells[0,17]:='���������� �����, � ������� �������� ����������� � ���� ��������������� (����� �����)';
 StringGrid1.Cells[0,18]:='�������������:';
 StringGrid1.Cells[0,19]:='- �������������� ������������';
 StringGrid1.Cells[0,20]:='- �������������� ������������';
 StringGrid1.Cells[0,21]:='������������� ����� � �������� 14 ���, ��������� ���������������, �� ������� ��������:';
 StringGrid1.Cells[0,22]:='1 ������ �������� (���. �����)';
 StringGrid1.Cells[0,23]:='2 ������ �������� (���. �����)';
 StringGrid1.Cells[0,24]:='3 ������ �������� (���. �����)';
 StringGrid1.Cells[0,25]:='4 ������ �������� (���. �����)';
 StringGrid1.Cells[0,26]:='5 ������ �������� (���. �����)';
 //��������� �������� �������� �����
 StringGrid1.Cells[0,0]:='��������';
 StringGrid1.Cells[1,0]:='�����:';
 StringGrid1.Cells[2,0]:='���������:';
 StringGrid1.Cells[3,0]:='�������:';
 //����������� � ���� ������
 IBDatabase1.Params.Clear();
 IBDatabase1.DatabaseName:=ExtractFilePath(Application.ExeName)+'BASE.FDB';
 IBDatabase1.Params.Add('User_name=SYSDBA');
 IBDatabase1.Params.Add('Password=masterkey');
 IBDatabase1.Params.Add('lc_ctype=WIN1251');
 IBDatabase1.Connected:=true;
 //���������� ������ ���
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
 RefreshReportsCombo; //���������� ����� �� ������� �������� �����
end;


//��������� ������� � MS Word
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
  //������� ��� ���  ���� �2 (��������� Import)
  tempStr:=ExtractFileName(importFileName);
  LPUCode:=tempStr[8]+tempStr[9]+tempStr[10]+tempStr[11];
  Error:=0;
  //��������� ��������
  App := CreateOleObject('Word.Application');
  App.Documents.Open(importFileName);
  App.Visible:=false;
  for i := 1 to 27 do     //���� �3 (��������� Import)
   begin
    tempNum:=App.ActiveDocument.Tables.Item(1).Cell(i+2,3).Range.Text; //���� �4 (��������� Import)
    if Clean_String(tempNum)<>'' then //���� �5 (��������� Import)
     VA[i]:=Trunc(StrToInt(Clean_String(tempNum))) //���� �6 (��������� Import)
    else
     VA[i]:=0; //���� �7 (��������� Import)

    tempNum:=App.ActiveDocument.Tables.Item(1).Cell(i+2,4).Range.Text; //���� �8 (��������� Import)
    if Clean_String(tempNum)<>'' then    //���� �9 (��������� Import)
     MA[i]:=Trunc(StrToInt(Clean_String(tempNum))) //���� �10 (��������� Import)
    else
     MA[i]:=0; //���� �11 (��������� Import)

    tempNum:=App.ActiveDocument.Tables.Item(1).Cell(i+2,5).Range.Text; //���� �12 (��������� Import)
    if Clean_String(tempNum)<>'' then //���� �13 (��������� Import)
     FA[i]:=Trunc(StrToInt(Clean_String(tempNum))) //���� �14 (��������� Import)
    else
     FA[i]:=0;  //���� �15 (��������� Import)
   end;
  App:= UnAssigned;
  //��������� ��� ���� ��� ������
  //���� �16 (��������� Import)
  textFileName:=ExtractFilePath(Application.ExeName)+DateToStr(Date)+'.txt';
  AssignFile(f,textFileName);
  //���� �17 (��������� Import)
  if not FileExists(textFileName) then
   begin
    //���� �18 (��������� Import)
    Rewrite(f);
    CloseFile(f);
   end;
  //���� �19 (��������� Import)
  Append(f);
  //���� �20 (��������� Import)
  for i:=7 to 15 do
   begin
    //���� �21 (��������� Import)
    if (VA[2]<VA[i]) or (MA[2]<MA[i]) or (FA[2]<FA[i])  then
     begin
      Error:=Error+1;     //���� �22 (��������� Import)
      writeln(f,'������c��� ������:'+InttoStr(Error));  //���� �23 (��������� Import)
    end;
   end;

  //���� �24 (��������� Import)
   if VA[22]<>VA[23]+VA[24]+VA[25]+VA[26]+VA[27] then
    begin
      Error:=Error+1; //���� �25 (��������� Import)
      writeln(f,'������c��� ������:'+InttoStr(Error)); //���� �26 (��������� Import)
    end;

    //���� �27 (��������� Import)
    if MA[22]<>MA[23]+MA[24]+MA[25]+MA[26]+MA[27] then
    begin
      Error:=Error+1; //���� �28 (��������� Import)
      writeln(f,'������c��� ������:'+InttoStr(Error)); //���� �29 (��������� Import)
    end;

    //���� �30 (��������� Import)
    if FA[22]<>FA[23]+FA[24]+FA[25]+FA[26]+FA[27] then
    begin
      Error:=Error+1; //���� �31 (��������� Import)
      writeln(f,'������c��� ������:'+IntToStr(Error));//���� �32 (��������� Import)
    end;

    //���� �33 (��������� Import)
    if VA[1]=FA[1]+MA[1] then
     begin
      Error:=Error+1; //���� �34 (��������� Import)
      writeln(f,'������c��� ������:'+IntToStr(Error));  //���� �35 (��������� Import)
     end;

     //���� �36 (��������� Import)
     if VA[3]=FA[3]+MA[3] then
     begin
      Error:=Error+1; //���� �37 (��������� Import)
      writeln(f,'������c��� ������:'+IntToStr(Error)); //���� �38 (��������� Import)
     end;

     //���� �39 (��������� Import)
     if VA[4]=FA[4]+MA[4] then
     begin
      Error:=Error+1; //���� �40 (��������� Import)
      writeln(f,'������c��� ������:'+IntToStr(Error)); //���� �41 (��������� Import)
     end;

      //���� �42 (��������� Import)
      if VA[5]=FA[5]+MA[5] then
     begin
      Error:=Error+1; //���� �43 (��������� Import)
      writeln(f,'������c��� ������:'+IntToStr(Error)); //���� �44 (��������� Import)
     end;

     //���� �45 (��������� Import)
     if VA[6]=FA[6]+MA[6] then
     begin
      Error:=Error+1; //���� �46 (��������� Import)
      writeln(f,'������c��� ������:'+IntToStr(Error)); //���� �47 (��������� Import)
     end;

    //���� �48 (��������� Import)
    if VA[11]=FA[11]+MA[11] then
     begin
      Error:=Error+1; //���� �49 (��������� Import)
      writeln(f,'������c��� ������:'+IntToStr(Error)); //���� �50 (��������� Import)
     end;
     CloseFile(f); //���� �51 (��������� Import)

  //������ � ����
  if (Error=0) then   //���� �52 (��������� Import)
   begin
    //���� �53 (��������� Import)
    //����������� ���� ���
    IBQuery1.Close;
    IBQuery1.SQL.Clear;
    sqlString:='SELECT LPU_KEY FROM LPU WHERE LPU_CODE ='''+LPUCode+'''';
    IBQuery1.SQL.Add(sqlString);
    IBQuery1.Open;
    if IBQuery1.RecordCount>0 then
     LPUKey:=IBQuery1.Fields[0].AsInteger
    else
     LPUKey:=1;
    //������� �������� ������ (���������������)
    IBQuery1.Close;
    IBQuery1.SQL.Clear;
    sqlString:='DELETE  FROM BASE WHERE LPU_KEY ='+IntToStr(LPUKey)+
               ' and REPORT_DATE='''+ DateToStr(Date)+'''';
    IBQuery1.SQL.Add(sqlString);
    IBQuery1.ExecSQL;
    //��������� ������ � �������
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
      RefreshReportsCombo; //���������� ����� �� ������� �������� �����
      ShowMessage('���������� � ���� ������ '+DateToStr(Date)+' ��� ���:'+LPUCode);
   end
  else
   ShowMessage('��� ������� �������� ������! ��� ����:'+textFileName);   //���� �54 (��������� Import)
end;

//������ ������ �� ����������� ��������
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


//���������� ����� �� ������� �������� �����
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
