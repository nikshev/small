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
    procedure Button4Click(Sender: TObject); //������ �������� ����� ��� �������� ������
    procedure FormCreate(Sender: TObject);  //�������� �����
    procedure FormClose(Sender: TObject; var Action: TCloseAction);  //�������� �����
    procedure Button7Click(Sender: TObject); //����� ������
    procedure Button6Click(Sender: TObject); //���������� ������
    procedure Button8Click(Sender: TObject); //�������� ������
    procedure Button1Click(Sender: TObject); //����� �� ���
    procedure ListBox1DblClick(Sender: TObject); //�� �������� ����� �� ������, ������� �������� �����
    procedure Button2Click(Sender: TObject); //����� �� �������
    procedure Button3Click(Sender: TObject);  //����� �� ���������
    procedure Import(var fileName:String);   //������ ������
    procedure Button5Click(Sender: TObject);   //�� ������� ������ ������ ������
    procedure Button9Click(Sender: TObject);   //�� ������� ������ ������ �������
    procedure ExportData(var ExpTable1,ExpTable2,ExpTable3:boolean; dbfFileName:String);
    procedure Button10Click(Sender: TObject); //������� ������
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}

//�������� ������� ����������
procedure TForm3.Button10Click(Sender: TObject);
begin
 SaveDialog1.InitialDir:=ExtractFilePath(Application.ExeName);
 if SaveDialog1.Execute then
  Edit2.Text:=SaveDialog1.FileName;
end;

//����� �� ���
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
    IBQuery1.Open; //����������� ����������� ���������
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
    ShowMessage('��������� ���� "���"');
end;

//����� �� �������
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
   IBQuery1.Open; //����������� ����������� ���������
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
  ShowMessage('��������� ���� "���" � ������');
end;

//����� �� ���������
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
   IBQuery1.Open; //����������� ����������� ���������
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
  ShowMessage('��������� ���� "���" � ������');
end;

//������ �������� �����
procedure TForm3.Button4Click(Sender: TObject);
begin
 OpenDialog1.InitialDir:=ExtractFilePath(Application.ExeName);
 if OpenDialog1.Execute then
  Edit10.Text:=OpenDialog1.FileName;
end;

//�� ������� ������ �������� ��������� �������
procedure TForm3.Button5Click(Sender: TObject);
var
 fileName:String;
begin
 fileName:=Edit10.Text;   //���� �2 (������� ���������)
 Import(fileName);       //���� �3 (������� ���������)
end;

//���������� ������
procedure TForm3.Button6Click(Sender: TObject);
var
 sqlString:String;
begin
 //��������� ������
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
              //������ �� ����� ���������
              IBQuery1.Close;
              IBQuery1.SQL.Clear;
              sqlString:='SELECT NIZ FROM BASE WHERE KOD='+IntToStr(ComboBox2.ItemIndex+1)+
                         ' AND FIO ='''+Edit1.Text+'''';
              IBQuery1.SQL.Add(sqlString);
              IBQuery1.Open; //����������� ����������� ���������
              if IBQuery1.RecordCount>0 then
                Edit3.Text:=FloatToStr(IBQuery1.Fields[0].AsFloat+1); //��������� ������� � ����������
              //������� ������
              IBQuery1.Close;
              IBQuery1.SQL.Clear;
              sqlString:='DELETE FROM BASE WHERE KOD='+IntToStr(ComboBox2.ItemIndex+1)+
                         ' AND FIO ='''+Edit1.Text+'''';
              IBQuery1.SQL.Add(sqlString);
              IBQuery1.ExecSQL; //������� ������ ���������� ���� �� ����
              //��������� �������� �� �����
              IBTableBase.Active:=True;
              IBTableBase.Insert;
              IBTableBase.Fields[0].AsInteger:=ComboBox2.ItemIndex+1; //��� ���
              IBTableBase.Fields[1].AsString:=Edit1.Text; //��� �����
              IBTableBase.Fields[2].AsDateTime:=DateTimePicker2.DateTime;//���� ��������
              IBTableBase.Fields[3].AsFloat:=StrToFloat(Edit3.Text); //����� ��������� ������ � ��������� � �����
              IBTableBase.Fields[4].AsDateTime:=Date; //���� ��������� � ��������� � �����
              IBTableBase.Fields[5].AsString:=Edit5.Text;//���������
              IBTableBase.Fields[6].AsString:=Edit6.Text; //�������������
              IBTableBase.Fields[7].AsString:=Edit7.Text;; //���������
              IBTableBase.Fields[8].AsDateTime:=DateTimePicker1.DateTime; //���� ������������� ����� �� ������� �������� ��������
              IBTableBase.Fields[9].AsDateTime:=DateTimePicker2.DateTime; //���� ����������� ����� �� ������� �������� ��������
              IBTableBase.Post;
              ShowMessage('���������� �������� �����:'+Edit1.Text+' ������ �������!');
             end
              else
               ShowMessage('��������� ���� "���������"');
           end
           else
            ShowMessage('��������� ���� "�������������"');
         end
         else
          ShowMessage('��������� ���� "���������"');
       end

       else
        ShowMessage('��������� ���� "���� ��������"');
     end
    else
      ShowMessage('��������� ���� "���"');
  end
   else
     ShowMessage('��������� ���� "���"');
end;

//����� ������
procedure TForm3.Button7Click(Sender: TObject);
begin
 //������ ��� ����
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

//�������� ������
procedure TForm3.Button8Click(Sender: TObject);
var
 sqlString:String;
begin
 //������� ������
 if ComboBox2.ItemIndex>-1 then
  begin
   if Edit1.Text<>'' then
    begin
     IBQuery1.Close;
     IBQuery1.SQL.Clear;
     sqlString:='DELETE FROM BASE WHERE KOD='+IntToStr(ComboBox2.ItemIndex+1)+
                ' AND FIO ='''+Edit1.Text+'''';
     IBQuery1.SQL.Add(sqlString);
     IBQuery1.ExecSQL; //������� ����������
    end
     else
      ShowMessage('��������� ���� "���"');
  end
   else
    ShowMessage('��������� ���� "���"');
end;

//�� ������� ������ ������ �������
procedure TForm3.Button9Click(Sender: TObject);
 var
 ExpTable1,ExpTable2,ExpTable3:boolean;
begin
 //���� �4 (������� ���������)
 ExpTable1:=CheckBox1.Checked;
 ExpTable2:=CheckBox2.Checked;
 ExpTable3:=CheckBox3.Checked;
 ExportData(ExpTable1,ExpTable2,ExpTable3, Edit2.Text);    //���� �5 (������� ���������)
end;

//�������� �����
procedure TForm3.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if IBDatabase1.Connected then
   IBDatabase1.Connected:=false
end;

//�������� �����
procedure TForm3.FormCreate(Sender: TObject);
var
 i:integer;
begin
 //����������� � ���� ������
 IBDatabase1.Params.Clear();
 IBDatabase1.DatabaseName:=ExtractFilePath(Application.ExeName)+'BASE.FDB';
 IBDatabase1.Params.Add('User_name=SYSDBA');
 IBDatabase1.Params.Add('Password=masterkey');
 IBDatabase1.Params.Add('lc_ctype=WIN1251');
 IBDatabase1.Connected:=true;
 //���������� ������ ���
 IBTableLPU.Active:=True;
 IBTableLPU.Last;
 IBTableLPU.First;
 for i := 1 to IBTableLPU.RecordCount do
  begin
   //���� �� ������� �� ������
   ComboBox1.Items.Add(IBTableLPU.Fields[1].AsString+
                      '(���:'+IBTableLPU.Fields[2].AsString+')');
   //���� �� ������� �� �������� �����
   ComboBox2.Items.Add(IBTableLPU.Fields[1].AsString+
                      '(���:'+IBTableLPU.Fields[2].AsString+')');
   IBTableLPU.Next;
  end;
 IBTableLPU.Active:=False;

end;

//�� �������� ����� �� ������, ������� �������� �����
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
   IBQuery1.Open; //������� ������ ���������� ���� �� ����
   if IBQuery1.RecordCount>0 then
    begin
      ComboBox2.ItemIndex:=IBQuery1.Fields[0].AsInteger-1; //��� ���
      Edit1.Text:=IBQuery1.Fields[1].AsString; //��� �����
      DateTimePicker3.DateTime:=IBQuery1.Fields[2].AsDateTime;//���� ��������
      Edit3.Text:=FloatToStr(IBQuery1.Fields[3].AsFloat); //����� ��������� ������ � ��������� � �����
      DateTimePicker4.DateTime:=IBQuery1.Fields[4].AsDateTime; //���� ��������� � ��������� � �����
      Edit5.Text:=IBQuery1.Fields[5].AsString;//���������
      Edit6.Text:=IBQuery1.Fields[6].AsString; //�������������
      Edit7.Text:=IBQuery1.Fields[7].AsString; //���������
      DateTimePicker1.DateTime:=IBQuery1.Fields[8].AsDateTime; //���� ������������� ����� �� ������� �������� ��������
      DateTimePicker2.DateTime:=IBQuery1.Fields[9].AsDateTime; //���� ����������� ����� �� ������� �������� ��������
    end;
  end;
end;

//������ ������
procedure TForm3.Import(var fileName:String); //������ ������
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
 if fileName<>'' then   //���� �2  (��������� Import)
  begin
   //��������� Excel ����
   XL:=CreateOLEObject('Excel.Application');
   XL.WorkBooks.Open(fileName);
   Sheet:=XL.ActiveWorkBook.Sheets[1];
   //������ ���
   //���� �3  (��������� Import)
   LPUCODE:=-1;
   //���� �4  (��������� Import)
   IBQuery1.Close;
   IBQuery1.SQL.Clear;
   sqlString:='SELECT LPU_KEY FROM LPU WHERE LPU_NAME ='''+Sheet.Cells[4,3].Value+'''';
   IBQuery1.SQL.Add(sqlString);
   IBQuery1.Open; //����������� ����������� ���������
   IBQuery1.Last;
   IBQuery1.First;
   if IBQuery1.RecordCount>0 then
    LPUCODE:=  IBQuery1.Fields[0].AsInteger;
   //��������� ��������
   //���� �5   (��������� Import)
   i:=7;
   EndImport:=false;
   if LPUCODE>-1 then      //���� �6  (��������� Import)
    begin
     while not EndImport do  //���� �7  (��������� Import)
      begin
       tmpStr:=Sheet.Cells[i,1].Value;    //���� �8  (��������� Import)
       if tmpStr<>'' then          //���� �9  (��������� Import)
        begin
         //���� �10   (��������� Import)
         KOD:= LPUCODE;
         FIO:=Sheet.Cells[i,2].Value;
         DB:=StrToDate(Sheet.Cells[i,3].Value);
         NAME_DL:=Sheet.Cells[i,4].Value;
         NAME_SP:=Sheet.Cells[i,5].Value;
         KAT:=Sheet.Cells[i,6].Value;
         DBEG:=StrToDate(Sheet.Cells[i,7].Value);
         DEND:=StrToDate(Sheet.Cells[i,8].Value);

         //������ �� ����� ���������
         IBQuery1.Close;
         IBQuery1.SQL.Clear;
         sqlString:='SELECT NIZ FROM BASE WHERE KOD='+IntToStr(LPUCODE)+
                    ' AND FIO ='''+FIO+'''';
         IBQuery1.SQL.Add(sqlString);
         IBQuery1.Open; //����������� ����������� ���������
         if IBQuery1.RecordCount>0 then
           NIZ:=IBQuery1.Fields[0].AsFloat+1 //��������� ������� � ����������
         else
          NIZ:=1;

         //������� ������
         IBQuery1.Close;
         IBQuery1.SQL.Clear;
         sqlString:='DELETE FROM BASE WHERE KOD='+IntToStr(LPUCODE)+
                    ' AND FIO ='''+FIO+'''';
         IBQuery1.SQL.Add(sqlString);
         IBQuery1.ExecSQL; //������� ������ ���������� ���� �� ����

          //���� �11   (��������� Import)
         //��������� �������� �� �����
         IBTableBase.Active:=True;
         IBTableBase.Insert;
         IBTableBase.Fields[0].AsInteger:=KOD; //��� ���
         IBTableBase.Fields[1].AsString:=FIO; //��� �����
         IBTableBase.Fields[2].AsDateTime:=DB;//���� ��������
         IBTableBase.Fields[3].AsFloat:=NIZ; //����� ��������� ������ � ��������� � �����
         IBTableBase.Fields[4].AsDateTime:=Date; //���� ��������� � ��������� � �����
         IBTableBase.Fields[5].AsString:=NAME_DL;//���������
         IBTableBase.Fields[6].AsString:=NAME_SP; //�������������
         IBTableBase.Fields[7].AsString:=KAT; //���������
         IBTableBase.Fields[8].AsDateTime:=DBEG; //���� ������������� ����� �� ������� �������� ��������
         IBTableBase.Fields[9].AsDateTime:=DEND; //���� ����������� ����� �� ������� �������� ��������
         IBTableBase.Post;
         i:=i+1;      //���� �12   (��������� Import)
        end
        else
         begin
          EndImport:=True;   //���� �13  (��������� Import)
          ShowMessage('���������� ��������� ������ �� ���:'+Sheet.Cells[4,3].Value+' ������ �������!');
          //������ ��� ����� �������� ����� �������
          IBQuery1.Close;
          IBQuery1.SQL.Clear;
          sqlString:='UPDATE BASE SET DEND='''+DateToStr(Date)+''''+
                     ' WHERE DIZ<>'''+DateToStr(Date)+''''+
                     ' AND KOD='+IntToStr(LPUCODE);
          IBQuery1.SQL.Add(sqlString);
          IBQuery1.ExecSQL; //������� ������ ���������� ���� �� ����
          IBTransaction1.Commit;
         end;
      end;
    end
     else
      ShowMessage('������������ �������� ��� � �����: '+fileName);

   //��������� Excel ����
   Sheet:= UnAssigned;
   XL:= UnAssigned;
  end
 else
  ShowMessage('�� ��������� ��� ����� ��� �������!');
end;

//������� ������
procedure TForm3.ExportData(var ExpTable1,ExpTable2,ExpTable3:boolean; dbfFileName:String);
var
 sqlString:String;
 XL,Sheet:Variant;
 i,j:Integer;
 RowIndex:Integer;
 LPUName:String;
 TDBFTable:TTable;
begin
   //������� � Excel ������� 1 (���������� ������ ������� ����� ������� �������� ������������� ���������� �� )
  if ExpTable1 then    //���� �2  (��������� ExportData)
   begin
    //��������� Excel ����
    XL:=CreateOLEObject('Excel.Application');
    XL.WorkBooks.Add;
    XL.Visible:=True;
    Sheet:=XL.ActiveWorkBook.Sheets[1];
    Sheet.Cells[1,1]:='���������� ������ ������� ����� ������� �������� ������������� ���������� �� '+DateToStr(Date);
    Sheet.Cells[2,1]:='������������ ��� ';
    Sheet.Cells[2,2]:='���������� ������';
    IBQuery1.Close;
    IBQuery1.SQL.Clear;
    sqlString:='SELECT KOD FROM BASE GROUP BY KOD';
    IBQuery1.SQL.Add(sqlString);
    IBQuery1.Open; //����������� ����������� ���������
    IBQuery1.Last;
    IBQuery1.First;
    for i := 1 to IBQuery1.RecordCount do //���� �3 (��������� ExportData)
      begin
       //������ ���
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT LPU_NAME FROM LPU WHERE LPU_KEY ='+IntToStr(IBQuery1.Fields[0].AsInteger);
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        if IBQuery2.RecordCount>0 then    //���� �4  (��������� ExportData)
         Sheet.Cells[i+2,1]:= IBQuery2.Fields[0].AsString; //���� �5   (��������� ExportData)

        //���������� ������ ������ ����� �������
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT * FROM BASE WHERE KOD ='+IntToStr(IBQuery1.Fields[0].AsInteger)+
                   ' AND DEND>'''+DateToStr(Date)+'''';
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        IBQuery2.Last;
        IbQuery2.First;
        if IBQuery2.RecordCount>0 then        //���� �6 (��������� ExportData)
         Sheet.Cells[i+2,2]:= IntToStr(IBQuery2.RecordCount);    //���� �7  (��������� ExportData)
      IBQuery1.Next; //���� �8 (��������� ExportData)
      end;
    //��������� Excel ����
   Sheet:= UnAssigned;
   XL:= UnAssigned;
   end;

  //������� � Excel ������� 2 (��������� ��������� ��������������� ������� �������� ������������� ����������)
  if ExpTable2 then //���� �9  (��������� ExportData)
   begin
    //��������� Excel ����
    XL:=CreateOLEObject('Excel.Application');
    XL.WorkBooks.Add;
    XL.Visible:=True;
    Sheet:=XL.ActiveWorkBook.Sheets[1];
    Sheet.Cells[1,1]:='��������� ��������� ��������������� '+
                      '������� �������� ������������� ���������� � '+DateToStr(DateTimePicker5.DateTime)+
                      ' �� '+DateToStr(DateTimePicker6.DateTime);
    Sheet.Cells[2,1]:='������������ ��� ';
    Sheet.Cells[2,2]:='���������� ����� ������ �� ������';
    Sheet.Cells[2,3]:='���������� ������ ���������� ����� �� ������� �������� ������������� ����������';
    Sheet.Cells[2,4]:='������� ���������� ������, �������������� ������� �������� ������������� ����������';
    IBQuery1.Close;
    IBQuery1.SQL.Clear;
    sqlString:='SELECT KOD FROM BASE GROUP BY KOD';
    IBQuery1.SQL.Add(sqlString);
    IBQuery1.Open; //����������� ����������� ���������
    IBQuery1.Last;
    IBQuery1.First;
    for i := 1 to IBQuery1.RecordCount do   //���� �10 (��������� ExportData)
       begin
        //������ ���
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT LPU_NAME FROM LPU WHERE LPU_KEY ='+IntToStr(IBQuery1.Fields[0].AsInteger);
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        if IBQuery2.RecordCount>0 then   //���� �11  (��������� ExportData)
         Sheet.Cells[i+2,1]:= IBQuery2.Fields[0].AsString; //���� �12 (��������� ExportData)

        //���������� ����� ������
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT * FROM BASE WHERE KOD ='+IntToStr(IBQuery1.Fields[0].AsInteger)+
                   ' AND NIZ=1';
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        IBQuery2.Last;
        IbQuery2.First;
        if IBQuery2.RecordCount>0 then    //���� �13 (��������� ExportData)
         Sheet.Cells[i+2,2]:= IntToStr(IBQuery2.RecordCount);    //���� �14 (��������� ExportData)


         //���������� ������ ���������� ����� �������
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT * FROM BASE WHERE KOD ='+IntToStr(IBQuery1.Fields[0].AsInteger)+
                   ' AND (DEND<'''+DateToStr(DateTimePicker5.DateTime)+''''+
                   ' OR DEND='''+DateToStr(DateTimePicker5.DateTime)+''''+')';
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        IBQuery2.Last;
        IbQuery2.First;
        if IBQuery2.RecordCount>0 then    //���� �15  (��������� ExportData)
         Sheet.Cells[i+2,3]:= IntToStr(IBQuery2.RecordCount); //���� �16   (��������� ExportData)

        //���������� ������ ������ ����� �������
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT * FROM BASE WHERE KOD ='+IntToStr(IBQuery1.Fields[0].AsInteger)+
                   ' AND DEND>'''+DateToStr(DateTimePicker6.DateTime)+'''';
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        IBQuery2.Last;
        IbQuery2.First;
        if IBQuery2.RecordCount>0 then        //���� �17 (��������� ExportData)
         Sheet.Cells[i+2,4]:= IntToStr(IBQuery2.RecordCount); //���� �18 (��������� ExportData)
      IBQuery1.Next;     //���� �19 (��������� ExportData)
     end;
    //��������� Excel ����
   Sheet:= UnAssigned;
   XL:= UnAssigned;
   end;

  //������� � Excel ������� 3 (������ ������ �������������� ������� �������� ������������� ����������)
  if ExpTable3 then      //���� �20   (��������� ExportData)
   begin
   //��������� Excel ����
    XL:=CreateOLEObject('Excel.Application');
    XL.WorkBooks.Add;
    XL.Visible:=True;
    Sheet:=XL.ActiveWorkBook.Sheets[1];
    Sheet.Cells[1,1]:='������ ������ �������������� ������� �������� ������������� ���������� �� '+DateToStr(Date);
    Sheet.Cells[2,1]:='������������ ��� ';
    Sheet.Cells[2,2]:='��� �����';
    IBQuery1.Close;
    IBQuery1.SQL.Clear;
    sqlString:='SELECT KOD FROM BASE GROUP BY KOD';
    IBQuery1.SQL.Add(sqlString);
    IBQuery1.Open; //����������� ����������� ���������
    IBQuery1.Last;
    IBQuery1.First;
    RowIndex:=3;      //���� �21 (��������� ExportData)
    for i := 1 to IBQuery1.RecordCount do //���� �22    (��������� ExportData)
     begin
        //������ ���
        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT LPU_NAME FROM LPU WHERE LPU_KEY ='+IntToStr(IBQuery1.Fields[0].AsInteger);
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        if IBQuery2.RecordCount>0 then        //���� �23  (��������� ExportData)
         LPUName:= IBQuery2.Fields[0].AsString;  //���� �24 (��������� ExportData)

        IBQuery2.Close;
        IBQuery2.SQL.Clear;
        sqlString:='SELECT FIO FROM BASE WHERE KOD ='+IntToStr(IBQuery1.Fields[0].AsInteger)+
                   ' AND DEND>'''+DateToStr(Date)+'''';
        IBQuery2.SQL.Add(sqlString);
        IBQuery2.Open;
        IBQuery2.Last;
        IbQuery2.First;
        for j := 1 to IbQuery2.RecordCount do     //���� �25 (��������� ExportData)
         begin
          //���� �26   (��������� ExportData)
          Sheet.Cells[RowIndex,1]:=LPUName;
          Sheet.Cells[RowIndex,2]:=IBQuery2.Fields[0].AsString;
          RowIndex:=RowIndex+1; //���� �27  (��������� ExportData)
          IbQuery2.Next; //���� �28  (��������� ExportData)
         end;
      IBQuery1.Next;  //���� �29  (��������� ExportData)
     end;
    //��������� Excel ����
   Sheet:= UnAssigned;
   XL:= UnAssigned;
   end;

   //�������� � dbf ����
   if dbfFileName<>'' then    //���� �30 (��������� ExportData)
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
     IBQuery1.Open; //����������� ����������� ���������
     IBQuery1.Last;
     IBQuery1.First;
     TDBFTable.Active:=true;
     for i := 1 to IBQuery1.RecordCount do     //���� �31  (��������� ExportData)
      begin
       TDBFTable.Insert;
       //���� �32   (��������� ExportData)
       TDBFTable.Fields[0].AsFloat:=IBQuery1.Fields[0].AsFloat;
       TDBFTable.Fields[1].AsString:=IBQuery1.Fields[1].AsString;
       TDBFTable.Fields[2].AsString:=FormatDateTime('yyyy.mm.dd',IBQuery1.Fields[2].AsDateTime);
       TDBFTable.Fields[3].AsFloat:=IBQuery1.Fields[3].AsFloat;
       TDBFTable.Fields[4].AsString:=DateToStr(IBQuery1.Fields[4].AsDateTime);
       //���� �33   (��������� ExportData)
       TDBFTable.Fields[5].AsString:=IBQuery1.Fields[5].AsString;
       TDBFTable.Fields[6].AsString:=IBQuery1.Fields[6].AsString;
       TDBFTable.Fields[7].AsString:=IBQuery1.Fields[7].AsString;
       TDBFTable.Fields[8].AsString:=DateToStr(IBQuery1.Fields[8].AsDateTime);
       TDBFTable.Fields[9].AsString:=DateToStr(IBQuery1.Fields[9].AsDateTime);
       TDBFTable.Post;
       IBQuery1.Next; //���� �34   (��������� ExportData)
      end;
     TDBFTable.Active:=false;
     ShowMessage('�������� ������ � ����: '+dbfFileName+' ��������� �������!');
    end;
end;
end.
