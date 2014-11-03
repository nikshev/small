unit p1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, Buttons, DB, DBTables, ADODB, xmldom, XMLIntf,
  msxmldom, XMLDoc, ZLib, ZLibConst, ComObj;

type
  TForm1 = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    ComboBox1: TComboBox;
    Label2: TLabel;
    ComboBox2: TComboBox;
    Label3: TLabel;
    ComboBox3: TComboBox;
    GroupBox2: TGroupBox;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    ComboBox4: TComboBox;
    ComboBox5: TComboBox;
    ComboBox6: TComboBox;
    StringGrid1: TStringGrid;
    BitBtn1: TBitBtn;
    ADOConnection1: TADOConnection;
    ADOTableLPU: TADOTable;
    ADOTableObesp: TADOTable;
    ADOTableRepGroup: TADOTable;
    ADOTableReport: TADOTable;
    ADOQuery1: TADOQuery;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    XMLDocument1: TXMLDocument;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ComoboRefresh;
    procedure ReadData(RepName:String; LPU,Obesp:Integer);   //������ ������ �� ����
    procedure SaveData;   //������������ ������ � ����
    procedure PrintData;  //����� �������� IA,RA � �����
    procedure PrintDataW;  //����� �������� IAW,RAW � �����
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure StringGrid1KeyPress(Sender: TObject; var Key: Char); //����� �������� � �����
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  RA: array[1..2,1..9] of Real;
  RAW: array[1..2,1..9] of Real;
  IA: array[1..8,1..9] of Integer;
  IAW: array[1..8,1..9] of Integer;
  E:Integer;

implementation

{$R *.dfm}

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 if ADOConnection1.Connected then
  ADOConnection1.Connected:=False;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
 ConnectionString:String; //������ ����������� � ����
 i,j:Integer;
begin
 //��������� ������������� �������
 StringGrid1.Cells[0,1]:='����������� ��������� ����������������� (�������)';
 StringGrid1.Cells[0,2]:='�������� �������� (��.)';
 StringGrid1.Cells[0,3]:='��������� �������� (��.) ';
 StringGrid1.Cells[0,4]:='�������� �� �� ����� (���. ���.)';
 StringGrid1.Cells[0,5]:='� �.�. ��������� �������� �� ���� ������� ������������� ������� (��.)';
 StringGrid1.Cells[0,6]:='� �.�. ��������� �������� �� ���� ������� ������������� ������� �� ����� (���. ���.) ';
 StringGrid1.Cells[0,7]:='���-�� �������� �� ����������� ����������� (��.) ';
 StringGrid1.Cells[0,8]:='������������� �������� �� ������������ �� ������(��.)';
 StringGrid1.Cells[0,9]:='������ �������� �� ������������� ���������(��.)';
 StringGrid1.Cells[0,10]:='���-�� ��������, ���� �������� ������� ����� � ������ ���������� �� ����������� ����������� (��.)';

 //��������� ������������� ������
 StringGrid1.Cells[1,0]:='�����, � �.�:';
 StringGrid1.Cells[2,0]:='������� ���������';
 StringGrid1.Cells[3,0]:='���� �� 3-� ���';
 StringGrid1.Cells[4,0]:='�������� ������ ��������������� �������� ';
 StringGrid1.Cells[5,0]:='�������� � ��������� ��� ';
 StringGrid1.Cells[6,0]:='�������� (� �.�. ������ ��������������� �������� � � ������������ ������������), ������������� ����������� ������� �������������� � ������ �������� �������� ��';
 StringGrid1.Cells[7,0]:='�������� ������';
 StringGrid1.Cells[8,0]:='������������ ����� ';
 StringGrid1.Cells[9,0]:='�������������� �����������';

 //�������� ������ �����������
  ConnectionString:='Provider=MSDASQL.1;Persist Security Info=False;'+
                    'Extended Properties="DBQ='+ExtractFilePath(Application.ExeName)+
                    'base.accdb;DefaultDir='+ExtractFilePath(Application.ExeName)+
                    ';Driver={Microsoft Access Driver (*.mdb, *.accdb)};DriverId=25;'+
                    'FIL=MS Access;MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;'+
                    'SafeTransactions=0;Threads=3;UID=admin;PWD=1111;UserCommitSync=Yes;"';
 ADOConnection1.ConnectionString:=ConnectionString;
 //����������� � ����
 ADOConnection1.Connected:=True;

 //��������� ComboBox
 ComoboRefresh;

 //�������� ������� ������� � ������� �������� � �����
 for i := 1 to 2 do
  for j := 1 to 9 do
    begin
     RA[i][j]:=0;
     RAW[i][j]:=0;
    end;


 for i := 1 to 7 do
  for j := 1 to 9 do
    begin
     IA[i][j]:=0;
     IAW[i][j]:=0;
    end;

  PrintData; //��������� ������ �������� � �����
end;

//������� ���������
procedure TForm1.BitBtn1Click(Sender: TObject);
var
 FirstRepDate, SecondRepDate:String;
 Obespech:Integer;
 indexR,indexI:Integer;
 i,j:Integer;
begin
   //���� �2
   FirstRepDate:=ComboBox1.Text;
   SecondRepDate:=ComboBox4.Text;
   Obespech:= ComboBox2.ItemIndex+1;
   E:=0;
   if FirstRepDate<>'' then  //���� �3
    begin
     ReadData(ComboBox1.Text,ComboBox3.ItemIndex+2,ComboBox2.ItemIndex+1); //���� �4
    if (E=0) then //���� �5
     begin
      //���� �6
      indexR:=0;
      indexI:=0;
     //���� �7
     for i := 1 to 10 do
      begin
       if (i=4) or (i=6) then   //���� �8
        indexR:=indexR+1        //���� �9
       else
        indexI:=indexI+1;       //���� �10

      for j := 1 to 9 do        //���� �11
       begin
        if (i=4) or (i=6) then   //���� �12
         RAW[indexR][j]:=RA[indexR][j] //���� �13
        else
         IAW[indexI][j]:=IA[indexI][j]; //���� �14
       end;
      end;

    if SecondRepDate<>'' then  //���� �16
     begin
      ReadData(ComboBox4.Text,ComboBox6.ItemIndex+2,ComboBox5.ItemIndex+1); //���� �17
     if (E=0) then //���� �18
     begin
      //���� �19
      indexR:=0;
      indexI:=0;
      for i := 1 to 10 do    //���� �20
       begin
        if (i=4) or (i=6) then   //���� �21
         indexR:=indexR+1        //���� �22
        else
         indexI:=indexI+1;       //���� �23

      for j := 1 to 9 do        //���� �24
       begin
        if (i=4) or (i=6) then   //���� �25
         RAW[indexR][j]:=ABS(RAW[indexR][j]-RA[indexR][j]) //���� �26
        else
         IAW[indexI][j]:=ABS(IAW[indexI][j]-IA[indexI][j]);//���� �27
       end;
      end;
      PrintDataW;
      ShowMessage('��������� ������ ������ �������!');
     end
     else
      ShowMessage('��������� ������ �� �������� ��� ��� � ������� ������������ ������!');
    end;
   end
   else
    ShowMessage('��������� ������ �� �������� ��� ��� � ������� ������������ ������!');
   end
   else
    begin
      //���� �15 �������� �������� � ������
      SaveData; //���� �27
    end;
end;

//��������� �������� xml � ������ ��� �������
procedure TForm1.BitBtn2Click(Sender: TObject);
var
 i,j:Integer;
 indexR,indexI:integer;
 RootNode:IXMLNode;
 LPUNode:IXMLNode;
 RepNode:IXMLNode;
 ObespNode:IXMLNode;
 RowNode:IXMLNode;
 ColNode:IXMLNode;
 CleanString:String;
 CompressionStream:TCompressionStream;
 destStream:TFileStream;
 xmlStream:TFileStream;
begin

 if (ComboBox3.ItemIndex>-1) and (ComboBox2.ItemIndex>-1) and (ComboBox1.ItemIndex>-1) then
  begin
   XMLDocument1.Create(nil);//�������� xml ���������
   XMLDocument1.Active:=True;
   RootNode:=XMLDocument1.AddChild('Root');
   LPUNode:=RootNode.AddChild('LPU-'+ComboBox3.Text);
   RepNode:=LPUNode.AddChild('Rep-'+ComboBox1.Text);
   ObespNode:=RepNode.AddChild(ComboBox2.Text);
   for i := 1 to 10 do
    begin
     CleanString:=StringReplace(StringGrid1.Cells[0,i],' ','_',[rfReplaceAll, rfIgnoreCase]);
     CleanString:=StringReplace(CleanString,'(','_',[rfReplaceAll, rfIgnoreCase]);
     CleanString:=StringReplace(CleanString,')','_',[rfReplaceAll, rfIgnoreCase]);
     CleanString:=StringReplace(CleanString,',','_',[rfReplaceAll, rfIgnoreCase]);
     CleanString:=StringReplace(CleanString,'.','_',[rfReplaceAll, rfIgnoreCase]);
     CleanString:=StringReplace(CleanString,':','_',[rfReplaceAll, rfIgnoreCase]);
     RowNode:=ObespNode.AddChild(CleanString);
    if (i=4) or (i=6) then
     indexR:=indexR+1
    else
     indexI:=indexI+1;

    for j := 1 to 9 do
     begin
      CleanString:=StringReplace(StringGrid1.Cells[j,0],' ','_',[rfReplaceAll, rfIgnoreCase]);
      CleanString:=StringReplace(CleanString,')','_',[rfReplaceAll, rfIgnoreCase]);
      CleanString:=StringReplace(CleanString,'(','_',[rfReplaceAll, rfIgnoreCase]);
      CleanString:=StringReplace(CleanString,',','_',[rfReplaceAll, rfIgnoreCase]);
      CleanString:=StringReplace(CleanString,'.','_',[rfReplaceAll, rfIgnoreCase]);
      CleanString:=StringReplace(CleanString,':','_',[rfReplaceAll, rfIgnoreCase]);
      ColNode:=RowNode.AddChild(CleanString);
      if (i=4) or (i=6) then
       ColNode.ChildValues['value']:= FloatToStrF(RA[indexR][j],ffFixed,8,2)
      else
       ColNode.ChildValues['value']:= IntToStr(IA[indexI][j]);
     end;
    end;
    XMLDocument1.SaveToFile(ExtractFilePath(Application.ExeName)+'temp.xml');
    xmlStream:=TFileStream.Create(ExtractFilePath(Application.ExeName)+'temp.xml',fmOpenRead);
    try
     destStream:=TFileStream.Create(ExtractFilePath(Application.ExeName)+'1.zip',fmOpenReadWrite);
    except
     destStream:=TFileStream.Create(ExtractFilePath(Application.ExeName)+'1.zip',fmCreate);
    end;
    CompressionStream:=TCompressionStream.Create(clMax,destStream);
    CompressionStream.CopyFrom(xmlStream,xmlStream.Size);
    CompressionStream.CompressionRate;
    CompressionStream.Free;
    destStream.Free;
    xmlStream.Free;
    XMLDocument1.Active:=False;
    //XMLDocument1.Free;
  end
 else
  ShowMessage('�� ��������� ���, �����, �����������');


end;

//��������� �������� � Excel
procedure TForm1.BitBtn3Click(Sender: TObject);
var
XL,Sheet:Variant;
conStr:String;
i,j:Integer;
indexI,indexR:Integer;
begin
 if (ComboBox3.ItemIndex>-1) and (ComboBox2.ItemIndex>-1) and (ComboBox1.ItemIndex>-1) then
  begin
   XL:=CreateOLEObject('Excel.Application');
   XL.WorkBooks.Add;
   XL.Visible:=True;
   Sheet:=XL.ActiveWorkBook.Sheets[1];
   Sheet.Cells[3,1]:= '���';

   for i := 1 to 10 do
    begin
    if (i=4) or (i=6) then
     indexR:=indexR+1
    else
     indexI:=indexI+1;
    Sheet.Cells[i+1,1]:=StringGrid1.Cells[0,i];

    for j := 1 to 9 do
     begin
      Sheet.Cells[1,j+1]:=StringGrid1.Cells[j,0];
      if (i=4) or (i=6) then
       Sheet.Cells[i+1,j+1]:= FloatToStrF(RA[indexR][j],ffFixed,8,2)
      else
       Sheet.Cells[i+1,j+1]:= IntToStr(IA[indexI][j]);
     end;
    end;
    Sheet:= UnAssigned;
    XL:= UnAssigned;
  end
 else
  ShowMessage('�� ��������� ���, �����, �����������');

end;

procedure TForm1.ComoboRefresh;
var
 i:Integer;
begin
if ADOConnection1.Connected then
 begin
  //����������� � ������� LPU � ���������� ComboBox3 (���)
  ADOTableLPU.Active:=True;
  ComboBox3.Items.Clear;
  for i:=1 to ADOTableLPU.RecordCount do
   begin
    ComboBox3.Items.Add(ADOTableLPU.FieldByName('LPUCode').AsString);
    ComboBox6.Items.Add(ADOTableLPU.FieldByName('LPUCode').AsString);
    ADOTableLPU.Next
   end;
  ADOTableLPU.Active:=False;

   //����������� � ������� Obesp � ���������� ComboBox3 (���)
  ADOTableObesp.Active:=True;
  ComboBox2.Items.Clear;
  for i:=1 to ADOTableObesp.RecordCount do
   begin
    ComboBox2.Items.Add(ADOTableObesp.FieldByName('ObespName').AsString);
    ComboBox5.Items.Add(ADOTableObesp.FieldByName('ObespName').AsString);
    ADOTableObesp.Next
   end;
  ADOTableObesp.Active:=False;

     //����������� � ������� RepCroup � ���������� ComboBox3 (���)
   ADOTableRepGroup.Active:=True;
   ComboBox1.Items.Clear;
    for i:=1 to ADOTableRepGroup.RecordCount do
     begin
      ComboBox1.Items.Add(ADOTableRepGroup.FieldByName('RepName').AsString);
      ComboBox4.Items.Add(ADOTableRepGroup.FieldByName('RepName').AsString);
      ADOTableRepGroup.Next
     end;
   ADOTableRepGroup.Active:=False;
 end;
end;

//��������� ������ �������� IA,RA � �����
procedure TForm1.PrintData;
var
 i,j:Integer;
 indexR,indexI:Integer;
begin
 indexR:=0;
 indexI:=0;
 for i := 1 to 10 do
  begin

   if (i=4) or (i=6) then
    indexR:=indexR+1
   else
    indexI:=indexI+1;

   for j := 1 to 9 do
    begin
     if (i=4) or (i=6) then
      StringGrid1.Cells[j,i]:=FloatToStrF(RA[indexR][j],ffFixed,8,2)
     else
      StringGrid1.Cells[j,i]:=IntToStr(IA[indexI][j]);
    end;
   end;
end;

//��������� ������ �������� IAW,RAW � �����
procedure TForm1.PrintDataW;
var
 i,j:Integer;
 indexR,indexI:Integer;
begin
 indexR:=0;
 indexI:=0;
 for i := 1 to 10 do
  begin

   if (i=4) or (i=6) then
    indexR:=indexR+1
   else
    indexI:=indexI+1;

   for j := 1 to 9 do
    begin
     if (i=4) or (i=6) then
      StringGrid1.Cells[j,i]:=FloatToStrF(RAW[indexR][j],ffFixed,8,2)
     else
      StringGrid1.Cells[j,i]:=IntToStr(IAW[indexI][j]);
    end;
   end;
end;

//��������� ������ �� ����
procedure TForm1.ReadData(RepName:String; LPU,Obesp:Integer);
var
 i,j:Integer;
 indexR,indexI:Integer;
 SQLString:String;
 tmpInt:Real;
begin
  //���� ��������� ��� � �����������
 if (ADOConnection1.Connected) and (LPU>0)  and (Obesp>0)then
   begin
     //��������� �� ���� ������
     ADOQuery1.Close;
     ADOQuery1.SQL.Clear;
     SQLString:='SELECT * FROM Report'+
                ' WHERE Report.RepName='''+RepName+''''+
                ' and LPU='+IntToStr(LPU)+
                ' and Obespech='+IntToStr(Obesp);
     ADOQuery1.SQL.Add(SQLString);
     ADOQuery1.Open;

     //���� � 2
     indexR:=0;
     indexI:=0;
     //���� � 3
     for i := 1 to 10 do
      begin
       //���� � 4
       if (i=4) or (i=6) then
        indexR:=indexR+1          //���� � 5
       else
        indexI:=indexI+1;         //���� � 6
      //���� � 7
       for j := 1 to 9 do
        begin
         if (i=4) or (i=6) then  //���� � 8
          RA[indexR][j]:=ADOQuery1.Fields[j+3].AsFloat  //���� � 9
         else
          begin
           tmpInt:=ADOQuery1.Fields[j+3].AsFloat;        //���� � 10
           if Round(tmpInt+0.4)=Trunc(tmpInt) then       //���� � 11
            IA[indexI][j]:=Trunc(tmpInt)       //���� � 12
           else
            IA[indexI][j]:=0        //���� � 13
          end;
        end;
         ADOQuery1.Next;
       end;
       ADOQuery1.Close;

     //���� � 14
     indexR:=0;
     indexI:=0;
     E:=0;
     //���� � 15
     for i := 1 to 10 do
      begin
       //���� � 16
       if (i=4) or (i=6) then
        begin
         indexR:=indexR+1;          //���� � 17
         if (RA[IndexR][1]>=RA[IndexR][3]+ RA[IndexR][4]) then //���� � 18
          if (RA[IndexR][1]>= RA[IndexR][7]+RA[IndexR][8]+ RA[IndexR][9]) then //���� � 20
           if (RA[IndexR][1]>= RA[IndexR][2]) then //���� � 22
           else
            E:=E+1 //���� � 23
           else
            E:=E+1 //���� � 21
         else
           E:=E+1 //���� � 19
        end
       else
        begin
         indexI:=indexI+1;         //���� � 24
         if (IA[IndexI][1]>=IA[IndexI][3]+ IA[IndexI][4]) then //���� � 25
          if (IA[IndexI][1]>= IA[IndexI][7]+IA[IndexI][8]+ IA[IndexI][9]) then //���� � 27
           if (IA[IndexI][1]>= IA[IndexI][2]) then //���� � 29
            else
            E:=E+1 //���� � 30
          else
           E:=E+1 //���� � 28
         else
          E:=E+1; //���� � 26
        end;
     end;
     PrintData; //���� � 31
   end;
end;

//��������� ���������� � ����
procedure TForm1.SaveData;
var
 i,j:Integer;
 indexR,indexI:Integer;
 ReportName:String;
 SQLString:String;
begin
 indexR:=0;
 indexI:=0;
 //������� ��� ������ ���� ���� �� ������� ��������� �� ����� ��� ��������
 //���� �� ��������� �� ������� �� ����
 if ComboBox1.ItemIndex>-1 then
   ReportName:=ComboBox1.Text
 else
   ReportName:=DateToStr(Now);

 //���� ��������� ��� � �����������
 if (ADOConnection1.Connected) and (ComboBox3.ItemIndex>-1) and  (ComboBox2.ItemIndex>-1)then
   begin
     //������� �������� ������
     ADOQuery1.SQL.Clear;
     SQLString:='DELETE Report.RepName, Report.LPU, Report.Obespech FROM Report'+
                ' WHERE Report.RepName='''+ReportName+''''+
                ' and LPU='+IntToStr(ComboBox3.ItemIndex+2)+
                ' and Obespech='+IntToStr(ComboBox2.ItemIndex+1);
     ADOQuery1.SQL.Add(SQLString);
     ADOQuery1.ExecSQL;

     indexR:=0;
     indexI:=0;
     for i := 1 to 10 do
      begin
       if (i=4) or (i=6) then
        indexR:=indexR+1
       else
        indexI:=indexI+1;

       for j := 1 to 9 do
        begin
         if (i=4) or (i=6) then
          RA[indexR][j]:=StrToFloat(StringGrid1.Cells[j,i])
         else
          IA[indexI][j]:=StrToInt(StringGrid1.Cells[j,i]);
       end;
   end;

     //��������� � ���� ������ �� ��������
     ADOTableReport.Active:=true;
     indexR:=0;
     indexI:=0;
     for i := 1 to 10 do
      begin
       if (i=4) or (i=6) then
        begin
         indexR:=indexR+1;
         ADOTableReport.Insert;
         ADOTableReport.Fields[1].AsString:=ReportName;
         ADOTableReport.Fields[2].AsInteger:=ComboBox3.ItemIndex+2;
         ADOTableReport.Fields[3].AsInteger:=ComboBox2.ItemIndex+1;
         ADOTableReport.Fields[4].AsFloat:=RA[indexR][1];
         ADOTableReport.Fields[5].AsFloat:=RA[indexR][2];
         ADOTableReport.Fields[6].AsFloat:=RA[indexR][3];
         ADOTableReport.Fields[7].AsFloat:=RA[indexR][4];
         ADOTableReport.Fields[8].AsFloat:=RA[indexR][5];
         ADOTableReport.Fields[9].AsFloat:=RA[indexR][6];
         ADOTableReport.Fields[10].AsFloat:=RA[indexR][7];
         ADOTableReport.Fields[11].AsFloat:=RA[indexR][8];
         ADOTableReport.Fields[12].AsFloat:=RA[indexR][9];
         ADOTableReport.Post;
        end
       else
        begin
         indexI:=indexI+1;
         ADOTableReport.Insert;
         ADOTableReport.Fields[1].AsString:=ReportName;
         ADOTableReport.Fields[2].AsInteger:=ComboBox3.ItemIndex+2;
         ADOTableReport.Fields[3].AsInteger:=ComboBox2.ItemIndex+1;
         ADOTableReport.Fields[4].AsFloat:=IA[indexI][1];
         ADOTableReport.Fields[5].AsFloat:=IA[indexI][2];
         ADOTableReport.Fields[6].AsFloat:=IA[indexI][3];
         ADOTableReport.Fields[7].AsFloat:=IA[indexI][4];
         ADOTableReport.Fields[8].AsFloat:=IA[indexI][5];
         ADOTableReport.Fields[9].AsFloat:=IA[indexI][6];
         ADOTableReport.Fields[10].AsFloat:=IA[indexI][7];
         ADOTableReport.Fields[11].AsFloat:=IA[indexI][8];
         ADOTableReport.Fields[12].AsFloat:=IA[indexI][9];
         ADOTableReport.Post;
        end;
     end;
     ShowMessage('���������� ������ ������� ���������!');
   end
   else
    ShowMessage('��������� ��� � �����������. ���������� �� ��������!');
 ComoboRefresh; //��������� ������
end;

//�������� ����� ���� � ��������� ����������� ���������� �����
procedure TForm1.StringGrid1KeyPress(Sender: TObject; var Key: Char);
begin
if not (key in['0'..'9',decimalseparator,#8,#13,#9,#127]) then
 begin
  key:=#0;
  beep;
 end;
end;

end.
