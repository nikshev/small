unit Unit3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, DB, ADODB, ComObj;

type
  TForm3 = class(TForm)
    TreeView1: TTreeView;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    Button3: TButton;
    Label1: TLabel;
    Edit1: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit7: TEdit;
    Edit8: TEdit;
    Button4: TButton;
    ADOConnection1: TADOConnection;
    ADOTableNameSredstvaGroup: TADOTable;
    ADOTableSredstva: TADOTable;
    ADOQueryGroup: TADOQuery;
    ADOQueryInvNum: TADOQuery;
    ADOQueryParameters: TADOQuery;
    Button2: TButton;
    Button1: TButton;
    ADOQueryInvCard: TADOQuery;
    GroupBox3: TGroupBox;
    Label9: TLabel;
    Edit9: TEdit;
    Button5: TButton;
    Label10: TLabel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Button6: TButton;
    Label11: TLabel;
    Edit10: TEdit;
    Edit11: TEdit;
    Button7: TButton;
    ListBox1: TListBox;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CreateTree; //������� ���������� ������
    procedure ReadInv(InvNumS:String); //������ �������� ������������ ������
    procedure TreeView1GetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure Button3Click(Sender: TObject);   //����� ������
    procedure Button1Click(Sender: TObject);   //���������� ������
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure ListBox1DblClick(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);   //�������� ������
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;
  SelIndex:Integer;
  DobInv:Integer;
  DellInv:Integer;
  RMS,RPR, RIR:Integer;


implementation

{$R *.dfm}

procedure TForm3.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 if ADOConnection1.Connected then   //���� ���������� ���������� � �����
  ADOConnection1.Connected:=False;  //��������� ���������� � �����
end;

procedure TForm3.FormCreate(Sender: TObject);
var
 ConnectionString:String;
begin
 //������� ������ �����������
 ConnectionString:='Provider=MSDASQL.1;Persist Security Info=False;'+
                    'Extended Properties="DBQ='+ExtractFilePath(Application.ExeName)+
                    'base.accdb;DefaultDir='+ExtractFilePath(Application.ExeName)+
                    ';Driver={Microsoft Access Driver (*.mdb, *.accdb)};DriverId=25;'+
                    'FIL=MS Access;MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;'+
                    'SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"';
 ADOConnection1.ConnectionString:=ConnectionString;
 //����������� � ����
 ADOConnection1.Connected:=True;
 CreateTree; //�������� ��������� �������� ������
end;

//�� �������� ������ ����������� �������� ������������ ������
procedure TForm3.ListBox1DblClick(Sender: TObject);
begin
 if (ListBox1.ItemIndex>-1) then
 ReadInv(ListBox1.Items[ListBox1.ItemIndex]);
end;

procedure TForm3.TreeView1GetSelectedIndex(Sender: TObject; Node: TTreeNode);
var
 InvNumI:Integer;
 InvNumS:String;
begin
 try
  InvNumI:=StrToInt(Node.Text);
  InvNumS:=Node.Text;
  ReadInv(InvNumS); //������� �������� ������������
 except
 end;
end;

procedure TForm3.ReadInv(InvNumS:String); //������ �������� ������������ ������
var
 i:Integer;
begin
 if ADOConnection1.Connected then
  begin
   ADOQueryInvCard.Close;
   ADOQueryInvCard.SQL.Clear;
   ADOQueryInvCard.SQL.Add('Select * from Sredstva where InvNum='''+InvNumS+'''');
   ADOQueryInvCard.Open;
   if ADOQueryInvCard.RecordCount>0 then
    begin
     Edit2.Text:= ADOQueryInvCard.Fields[1].AsString;
     Edit3.Text:= ADOQueryInvCard.Fields[2].AsString;
     Edit4.Text:= ADOQueryInvCard.Fields[3].AsString;
     Edit5.Text:= ADOQueryInvCard.Fields[4].AsString;
     Edit6.Text:= ADOQueryInvCard.Fields[5].AsString;
     Edit7.Text:= ADOQueryInvCard.Fields[6].AsString;
     Edit8.Text:= ADOQueryInvCard.Fields[7].AsString;
     if  ADOQueryInvCard.Fields[8].AsInteger=1 then
      begin
       CheckBox1.Checked:=True;
       CheckBox2.Checked:=False;
      end
     else
       begin
       CheckBox1.Checked:=False;
       CheckBox2.Checked:=True;
      end;
    end;
   ADOQueryInvCard.Close;
  end;
end;

//���������� �������� ������������ ������
//���� �4
procedure TForm3.Button1Click(Sender: TObject);
begin
 DobInv:=1; //���� �2 (�������� ���������)
 if DobInv=1 then  //���� �3 (�������� ���������)
  begin
   //���� �4 (�������� ���������)
   if ADOConnection1.Connected then
    begin
     if (Edit3.Text<>'') then
      begin
       ADOQueryInvNum.Close;
       ADOQueryInvNum.SQL.Clear;
       ADOQueryInvNum.SQL.Add('DELETE Sredstva.InvNum '+
                             ' FROM Sredstva where Sredstva.InvNum='''+ Edit3.Text+'''');
       ADOQueryInvNum.ExecSQL;
       ADOQueryInvNum.Close;
       ADOTableSredstva.TableName:='Sredstva';
       ADOTableSredstva.Active:=True;
       ADOTableSredstva.Insert;
       ADOTableSredstva.Fields[1].AsString := Edit2.Text;
       ADOTableSredstva.Fields[2].AsString := Edit3.Text;
       ADOTableSredstva.Fields[3].AsString := Edit4.Text;
       ADOTableSredstva.Fields[4].AsFloat  := StrToFloat(Edit5.Text);
       ADOTableSredstva.Fields[5].AsInteger:= StrToInt(Edit6.Text);
       ADOTableSredstva.Fields[6].AsString := Edit7.Text;
       ADOTableSredstva.Fields[7].AsFloat  := StrToFloat(Edit8.Text);
       if CheckBox1.Checked then
         ADOTableSredstva.Fields[8].AsInteger:=1
       else
         ADOTableSredstva.Fields[8].AsInteger:=0;
       ADOTableSredstva.Post;
       ADOTableSredstva.Active:=False;
       ShowMessage('���������� ������������ ������ '+Edit3.Text+' ������ �������!');
       CreateTree;  //���� �7 (�������� ���������)
      end
     else
      ShowMessage('��������� ����������� ����� ��� ����������!');
   end;
  end;
end;

//�������� ������
procedure TForm3.Button2Click(Sender: TObject);
begin
 DellInv:=1; //���� �2  (�������� ���������)
 if DellInv=1 then  //���� �5 (�������� ���������)
  begin
   //���� �6 (�������� ���������)
   if ADOConnection1.Connected then
    begin
     if (Edit2.Text<>'') then
      begin
       ADOQueryInvNum.Close;
       ADOQueryInvNum.SQL.Clear;
       ADOQueryInvNum.SQL.Add('DELETE Sredstva.InvNum '+
                             ' FROM Sredstva where Sredstva.InvNum='''+ Edit3.Text+'''');
       ADOQueryInvNum.ExecSQL;
       ADOQueryInvNum.Close;
       ShowMessage('�������� ������������ ������ '+Edit3.Text+' ������ �������!');
       CreateTree;   //���� �7 (�������� ���������)
      end
     else
      ShowMessage('��������� ����������� ����� ��� ��������!');
    end;
  end;
end;

//����� ������
procedure TForm3.Button3Click(Sender: TObject);
begin
   Edit2.Text:= '';
   Edit3.Text:= '';
   Edit4.Text:= '';
   Edit5.Text:= '';
   Edit6.Text:= '';
   Edit7.Text:= '';
   Edit8.Text:= '';
   CheckBox1.Checked:=True;
end;

//����� ������������ �� ����� ������
procedure TForm3.Button4Click(Sender: TObject);
var
 i:Integer;
 RecCountI:Integer;
 SQLString:String;
begin
if ADOConnection1.Connected then
  begin
   if (Edit1.Text<>'') then
    begin
     ADOQueryInvNum.Close;
     ADOQueryInvNum.SQL.Clear;
     SQLString:='SELECT Sredstva.InvNum'+
                            ' FROM Sredstva where Sredstva.InvNum Like('''+'%'+ Edit1.Text+'%'+''''+')';
     ADOQueryInvNum.SQL.Add(SQLString);
     ADOQueryInvNum.Open;
     RecCountI:=ADOQueryInvNum.RecordCount;
     ListBox1.Items.Clear;
     for i := 1 to RecCountI do
      begin
       ListBox1.Items.Add(ADOQueryInvNum.Fields[0].AsString);
       ADOQueryInvNum.Next;
      end;
    end
    else
     ShowMessage('��������� ����������� ����� ��� ������!');
  end;
end;

//������� ������������ ������� ������������ ������� �� ����������� � ������� � Excel;
procedure TForm3.Button5Click(Sender: TObject);
var
 i,j:Integer;
 RecCountI:Integer;
 SQLString:String;
 XL,Sheet:Variant;
begin
 if Edit9.Text<>'' then RMS:=1; //���� �2 (�������� ���������)
 if ADOConnection1.Connected then
  begin
   if (RMS=1) then   //���� �9 (�������� ���������)
    begin
     //���� �10 (�������� ���������)
     ADOQueryInvNum.Close;
     ADOQueryInvNum.SQL.Clear;
     SQLString:='SELECT * FROM Sredstva where Sredstva.FIO ='''+Edit9.Text+'''';
     ADOQueryInvNum.SQL.Add(SQLString);
     ADOQueryInvNum.Open;
     RecCountI:=ADOQueryInvNum.RecordCount;
     if RecCountI>0 then
      begin
       XL:=CreateOLEObject('Excel.Application');
       XL.WorkBooks.Add;
       XL.Visible:=True;
       Sheet:=XL.ActiveWorkBook.Sheets[1];
       for i := 1 to RecCountI do
        begin
         if ADOQueryInvNum.Fields[8].AsInteger=1 then
          Sheet.Cells[i,1]:='������������ ������������'
         else
          Sheet.Cells[i,1]:='���������';
         for j := 1 to 7 do
          if j=2 then
           Sheet.Cells[i,j+1]:=''''+ADOQueryInvNum.Fields[j].AsString
          else
           Sheet.Cells[i,j+1]:=ADOQueryInvNum.Fields[j].AsString;
         ADOQueryInvNum.Next
        end;
       Sheet:= UnAssigned;
       XL:= UnAssigned;
      end
      else
       ShowMessage('��� ������ ��� �����������!');
    end;
  end;
end;

//������� ������������ ������� ������������� �������� �� �������� ������;
procedure TForm3.Button6Click(Sender: TObject);
var
 i,j:Integer;
 RecCountI:Integer;
 SQLString:String;
 XL,Sheet:Variant;
 DayB,DayE:Word;
 MonthB,MonthE:Word;
 YearB,YearE:Word;
begin
 RPR:=1;  //���� �2 (�������� ���������)
  if ADOConnection1.Connected then
  begin
   if (RPR=1) then   //���� �11 (�������� ���������)
    begin
     //���� �12 (�������� ���������)
     ADOQueryInvNum.Close;
     ADOQueryInvNum.SQL.Clear;
     DecodeDate(DateTimePicker1.DateTime, YearB,MonthB,DayB);
     DecodeDate(DateTimePicker2.DateTime, YearE,MonthE,DayE);
     SQLString:='SELECT * FROM Sredstva where Sredstva.DateB>#'+
                 IntToStr(MonthB)+'/'+IntToStr(DayB)+'/'+IntToStr(YearB)+
                 '# and Sredstva.DateB<#'+
                 IntToStr(MonthE)+'/'+IntToStr(DayE)+'/'+IntToStr(YearE)+'#';
     ADOQueryInvNum.SQL.Add(SQLString);
     ADOQueryInvNum.Open;
     RecCountI:=ADOQueryInvNum.RecordCount;
     if RecCountI>0 then
      begin
       XL:=CreateOLEObject('Excel.Application');
       XL.WorkBooks.Add;
       XL.Visible:=True;
       Sheet:=XL.ActiveWorkBook.Sheets[1];
       for i := 1 to RecCountI do
        begin
        if ADOQueryInvNum.Fields[8].AsInteger=1 then
          Sheet.Cells[i,1]:='������������ ������������'
         else
          Sheet.Cells[i,1]:='���������';
         for j := 1 to 7 do
          if j=2 then
           Sheet.Cells[i,j+1]:=''''+ADOQueryInvNum.Fields[j].AsString
          else
           Sheet.Cells[i,j+1]:=ADOQueryInvNum.Fields[j].AsString;
         ADOQueryInvNum.Next
        end;
       Sheet:= UnAssigned;
       XL:= UnAssigned;
      end
      else
       ShowMessage('��� ������ ��� �����������!');
    end;
  end;
end;

//������� ������������ ������� �� ������/������� ������������ ��� �������� ����������� �������;
procedure TForm3.Button7Click(Sender: TObject);
var
 i,j:Integer;
 RecCountI:Integer;
 SQLString:String;
 XL,Sheet:Variant;
begin
  if (Edit10.Text<>'') and (Edit11.Text<>'') then
    RIR:=1;  //���� �2 (�������� ���������)
  if ADOConnection1.Connected then
  begin
   if (RIR=1) then   //���� �13 (�������� ���������)
    begin
     //���� �14 (�������� ���������)
     ADOQueryInvNum.Close;
     ADOQueryInvNum.SQL.Clear;
     SQLString:='SELECT * FROM Sredstva where Sredstva.Iznos='+ Edit10.Text+
                ' and Rashod=' +Edit11.Text;
     ADOQueryInvNum.SQL.Add(SQLString);
     ADOQueryInvNum.Open;
     RecCountI:=ADOQueryInvNum.RecordCount;
     if RecCountI>0 then
      begin
       XL:=CreateOLEObject('Excel.Application');
       XL.WorkBooks.Add;
       XL.Visible:=True;
       Sheet:=XL.ActiveWorkBook.Sheets[1];
       for i := 1 to RecCountI do
        begin
         if ADOQueryInvNum.Fields[8].AsInteger=1 then
          Sheet.Cells[i,1]:='������������ ������������'
         else
          Sheet.Cells[i,1]:='���������';
         for j := 1 to 7 do
          if j=2 then
           Sheet.Cells[i,j]:=''''+ADOQueryInvNum.Fields[j].AsString
          else
           Sheet.Cells[i,j]:=ADOQueryInvNum.Fields[j].AsString;
         ADOQueryInvNum.Next
        end;
       Sheet:= UnAssigned;
       XL:= UnAssigned;
      end
      else
       ShowMessage('��� ������ ��� �����������!');
    end;
  end;
end;

//������ ���� checkBox ����� ���� �������
procedure TForm3.CheckBox1Click(Sender: TObject);
begin
 if CheckBox1.Checked then
  CheckBox2.Checked:=False;
end;

//������ ���� checkBox ����� ���� �������
procedure TForm3.CheckBox2Click(Sender: TObject);
begin
 if CheckBox2.Checked then
  CheckBox1.Checked:=False;
end;

//������� ���������� ������
procedure TForm3.CreateTree;
var
 RecCountG,RecCountI,RecCountP:Integer;
 GroupName:String;
 InvNum:String;
 Sotr:String;
 Iznos:Real;
 Rashod:Integer;
 DateB:String;
 Price:Real;
 i,k,j:Integer;
begin
  TreeView1.Items.Clear;   //���� �2 (������� ���������� ������)
  TreeView1.Items.Add(nil,'�������� ��������');
  TreeView1.Items.AddChild(TreeView1.Items[0],'������������ ������������');
  if ADOConnection1.Connected then
   begin
    ADOQueryGroup.Close;
    ADOQueryGroup.SQL.Clear;
    ADOQueryGroup.SQL.Add('SELECT Sredstva.NameSredstva '+
                          ' FROM Sredstva where Uchet=1 GROUP BY Sredstva.NameSredstva;');
    ADOQueryGroup.Open;
    RecCountG:=ADOQueryGroup.RecordCount;
    for i := 1 to RecCountG do   //���� �3 (������� ���������� ������)
     begin
      GroupName:=ADOQueryGroup.Fields[0].AsString;  //���� �4  (������� ���������� ������)
      ADOQueryGroup.Next;
      TreeView1.Items.AddChild(TreeView1.Items[0].Item[0],GroupName); //���� �5  (������� ���������� ������)
      ADOQueryInvNum.Close;
      ADOQueryInvNum.SQL.Clear;
      ADOQueryInvNum.SQL.Add('SELECT Sredstva.InvNum '+
                            ' FROM Sredstva where NameSredstva='''+ GroupName+''''+
                            ' and Uchet=1 GROUP BY Sredstva.InvNum;');
      ADOQueryInvNum.Open;
      RecCountI:=ADOQueryInvNum.RecordCount;
      for k := 1 to RecCountI do               //���� �6  (������� ���������� ������)
       begin
        InvNum:=ADOQueryInvNum.Fields[0].AsString;    //���� �7  (������� ���������� ������)
        ADOQueryInvNum.Next;
        TreeView1.Items.AddChild(TreeView1.Items[0].Item[0].Item[i-1],InvNum);   //���� �8  (������� ���������� ������)
        ADOQueryParameters.Close;
        ADOQueryParameters.SQL.Clear;
        ADOQueryParameters.SQL.Add('SELECT Sredstva.Fio,Sredstva.Iznos,'+
                                   ' Sredstva.Rashod, Sredstva.DateB, Sredstva.Price '+
                                   ' FROM Sredstva where Sredstva.InvNum='''+ InvNum+''''+
                                   ' and Uchet=1');
        ADOQueryParameters.Open;
        RecCountP:=ADOQueryParameters.RecordCount;
        for j := 1 to RecCountP do    //���� �9  (������� ���������� ������)
         begin
          //���� �10  (������� ���������� ������)
           Sotr:=ADOQueryParameters.Fields[0].AsString;
           Iznos:=ADOQueryParameters.Fields[1].AsFloat;
           Rashod:=ADOQueryParameters.Fields[2].AsInteger;
           DateB:=ADOQueryParameters.Fields[3].AsString;
           Price:=ADOQueryParameters.Fields[4].AsFloat;
           TreeView1.Items.AddChild(TreeView1.Items[0].Item[0].Item[i-1].Item[k-1],'���������: '+Sotr); //���� �11  (������� ���������� ������)
           TreeView1.Items.AddChild(TreeView1.Items[0].Item[0].Item[i-1].Item[k-1],'�����,%: '+FloatToStrF(Iznos,ffFixed,8,2));//���� �12  (������� ���������� ������)
           TreeView1.Items.AddChild(TreeView1.Items[0].Item[0].Item[i-1].Item[k-1],'������: '+IntToStr(Rashod));//���� �13  (������� ���������� ������)
           TreeView1.Items.AddChild(TreeView1.Items[0].Item[0].Item[i-1].Item[k-1],'���� ������.: '+DateB); //���� �14  (������� ���������� ������)
           TreeView1.Items.AddChild(TreeView1.Items[0].Item[0].Item[i-1].Item[k-1],'����: '+FloatToStrF(Price,ffFixed,8,2)); //���� �15  (������� ���������� ������)
           ADOQueryParameters.Next;
         end;
         ADOQueryParameters.Close;
       end;
       ADOQueryInvNum.Close;
     end;
    ADOQueryGroup.Close;
   end
  else
   ShowMessage('��� ����������� � ����');

  TreeView1.Items.AddChild(TreeView1.Items[0],'���������');
  if ADOConnection1.Connected then
   begin
    ADOQueryGroup.Close;
    ADOQueryGroup.SQL.Clear;
    ADOQueryGroup.SQL.Add('SELECT Sredstva.NameSredstva '+
                          ' FROM Sredstva where Uchet=0 GROUP BY Sredstva.NameSredstva;');
    ADOQueryGroup.Open;
    RecCountG:=ADOQueryGroup.RecordCount;
    for i := 1 to RecCountG do   //���� �16 (������� ���������� ������)
     begin
      GroupName:=ADOQueryGroup.Fields[0].AsString;  //���� �17  (������� ���������� ������)
      ADOQueryGroup.Next;
      TreeView1.Items.AddChild(TreeView1.Items[0].Item[1],GroupName); //���� �18  (������� ���������� ������)
      ADOQueryInvNum.Close;
      ADOQueryInvNum.SQL.Clear;
      ADOQueryInvNum.SQL.Add('SELECT Sredstva.InvNum '+
                            ' FROM Sredstva where NameSredstva='''+ GroupName+''''+
                            ' and Uchet=0 GROUP BY Sredstva.InvNum;');
      ADOQueryInvNum.Open;
      RecCountI:=ADOQueryInvNum.RecordCount;
      for k := 1 to RecCountI do               //���� �19  (������� ���������� ������)
       begin
        InvNum:=ADOQueryInvNum.Fields[0].AsString;    //���� �20  (������� ���������� ������)
        ADOQueryInvNum.Next;
        TreeView1.Items.AddChild(TreeView1.Items[0].Item[1].Item[i-1],InvNum);   //���� �21  (������� ���������� ������)
        ADOQueryParameters.Close;
        ADOQueryParameters.SQL.Clear;
        ADOQueryParameters.SQL.Add('SELECT Sredstva.Fio,Sredstva.Iznos,'+
                                   ' Sredstva.Rashod, Sredstva.DateB, Sredstva.Price '+
                                   ' FROM Sredstva where Sredstva.InvNum='''+ InvNum+''''+
                                   ' and Uchet=0');
        ADOQueryParameters.Open;
        RecCountP:=ADOQueryParameters.RecordCount;
        for j := 1 to RecCountP do    //���� �22  (������� ���������� ������)
         begin
          //���� �23  (������� ���������� ������)
           Sotr:=ADOQueryParameters.Fields[0].AsString;
           Iznos:=ADOQueryParameters.Fields[1].AsFloat;
           Rashod:=ADOQueryParameters.Fields[2].AsInteger;
           DateB:=ADOQueryParameters.Fields[3].AsString;
           Price:=ADOQueryParameters.Fields[4].AsFloat;
           TreeView1.Items.AddChild(TreeView1.Items[0].Item[1].Item[i-1].Item[k-1],'���������: '+Sotr); //���� �24 (������� ���������� ������)
           TreeView1.Items.AddChild(TreeView1.Items[0].Item[1].Item[i-1].Item[k-1],'�����,%: '+FloatToStrF(Iznos,ffFixed,8,2));//���� �25  (������� ���������� ������)
           TreeView1.Items.AddChild(TreeView1.Items[0].Item[1].Item[i-1].Item[k-1],'������: '+IntToStr(Rashod));//���� �26  (������� ���������� ������)
           TreeView1.Items.AddChild(TreeView1.Items[0].Item[1].Item[i-1].Item[k-1],'���� ������.: '+DateB); //���� �27  (������� ���������� ������)
           TreeView1.Items.AddChild(TreeView1.Items[0].Item[1].Item[i-1].Item[k-1],'����: '+FloatToStrF(Price,ffFixed,8,2)); //���� �28  (������� ���������� ������)
           ADOQueryParameters.Next;
         end;
         ADOQueryParameters.Close;
       end;
       ADOQueryInvNum.Close;
     end;
    ADOQueryGroup.Close;
   end
  else
   ShowMessage('��� ����������� � ����');
end;

end.
