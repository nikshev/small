unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, StdCtrls, ComCtrls, ExtCtrls, TeeProcs, TeEngine, Chart,
  Series, ValEdit, Buttons, ExtDlgs, Menus;

type
  TForm1 = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    Chart1: TChart;
    UpDown1: TUpDown;
    LabeledEdit1: TLabeledEdit;
    StringGrid1: TStringGrid;
    Series1: TLineSeries;
    Series2: TLineSeries;
    LabeledEdit2: TLabeledEdit;
    BitBtn1: TBitBtn;
    StringGrid2: TStringGrid;
    Memo1: TMemo;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    SavePictureDialog1: TSavePictureDialog;
    N6: TMenuItem;
    procedure UpDown1Click(Sender: TObject; Button: TUDBtnType);
    procedure FormActivate(Sender: TObject);
    procedure Chart1Click(Sender: TObject);
    procedure GridCopy;
    procedure TabSheet2Show(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure LabeledEdit1Change(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.GridCopy;
{��������� ��������� ����������� ������ �� 1� ������� �� ������.}
var i:integer;
begin
StringGrid2.RowCount:=StringGrid1.RowCount+1;//��������� ������� ������ �������.
for i:=1 to (StringGrid1.RowCount-1) do
StringGrid2.Cells[0,i]:=StringGrid1.Cells[0,i]; //����������� ���.
end;

procedure TForm1.UpDown1Click(Sender: TObject; Button: TUDBtnType);
{���������� ������� ������ �����/���� ����� ���� �����.
����������� ��� ��������� ���������� ����� � ������� ����� ������.}
begin
StringGrid1.RowCount:=strtoint(LabeledEdit1.Text)+1; //��������� �������.
StringGrid1.Cells[0,StringGrid1.RowCount-1]:='0'; //�������������.
StringGrid1.Cells[1,StringGrid1.RowCount-1]:='0';
StringGrid1.Cells[2,StringGrid1.RowCount-1]:='0';
end;

procedure TForm1.FormActivate(Sender: TObject);
{���������� ��������� �����, �.�. ����������� ��� ������� ����������.}
begin
StringGrid1.Cells[0,0]:='�����';
StringGrid1.Cells[1,0]:='�����������';
StringGrid1.Cells[2,0]:='������������';
StringGrid1.Cells[0,1]:='0';
StringGrid1.Cells[1,1]:='0';
StringGrid1.Cells[2,1]:='0';

StringGrid2.Cells[0,0]:='�����';
StringGrid2.Cells[1,0]:='U [ i ]';
StringGrid2.Cells[2,0]:='U [ i ] ^ 2';
StringGrid2.Cells[3,0]:='S ( t [ i ] )';
StringGrid2.Cells[4,0]:='t [ i ]';
StringGrid2.Cells[5,0]:='S ( t [ i ] ) * t [ i ]';
StringGrid2.Cells[6,0]:='Ksi ( t [ i ] )';
StringGrid2.Cells[7,0]:='U [ s ( t [ i ] ) ]';
StringGrid2.Cells[8,0]:='( U [ s ( t [ i ] ) ] ^ 2';

Form1.GridCopy;
end;

procedure TForm1.Chart1Click(Sender: TObject);
{���������� ������ �� ������� - ��������� ����������.}
var i:integer;
begin
Series1.Clear; //������� ����� ������.
Series2.Clear;
for i:=1 to (StringGrid1.RowCount-1) do
begin
Series1.AddXY(strtofloat(StringGrid1.Cells[0,i]),strtofloat(StringGrid1.Cells[1,i])); //���������� �����.
Series2.AddXY(strtofloat(StringGrid1.Cells[0,i]),strtofloat(StringGrid1.Cells[2,i]));
end;
end;

procedure TForm1.TabSheet2Show(Sender: TObject);
{����������� ������ �� ������� � ������� � ����� ������.}
begin
Form1.GridCopy;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
{����� ������� ��������� - ���������� ���� �������� ����������� ���� ������ ���.
����������� � ���������� �� ���� ���������.}
var c,i,j,PC,N:integer;
Yp,Yt,U,U2,St,T,Ksi,Us,Us2:array[1..1000] of double;
Su2,Sum,Sum2,SST,A11,A12,A22,B1,B2,D,D1,D2,A0,A1,Ksum,Kus2,sumti2,Bst,Stst,Mbst:double;
begin
c:=StringGrid1.RowCount-1;//���������� ����� � �������. ����� - ���������� �����.
StringGrid2.Cells[0,StringGrid2.RowCount-1]:='              -';//������� ������ ���� �������.
StringGrid2.Cells[1,StringGrid2.RowCount-1]:='              -';
Su2:=0; //����������� - �������� �� ����� � 3 �������.
SSt:=0; //����������� - �������� �� ����� � 4 �������.
N:=strtoint(LabeledEdit2.Text);//�������� ����������.
if (N mod 2)=0 then begin //�������� �� ����������.
ShowMessage('�������� ���������� ������ ���� ��������!');
Exit;
end;
PC:=strtoint(LabeledEdit1.Text)-N+1; //���������� ����������� � �������.
A11:=0; //�������� ������� ��� ���������� �0 � �1, ����������� ��� ������� ���.
A12:=0;
A22:=0;
B1:=0;
B2:=0;
sumti2:=0; //����������� - ��� ���������� ����� ��������� T[i];
for i:=1 to c do
  begin
    Yp[i]:=strtofloat(StringGrid1.Cells[1,i]);//������ ������ �� ������ �������.
    Yt[i]:=strtofloat(StringGrid1.Cells[2,i]);
    U[i]:=Yt[i]-Yp[i];//���������� ��������� 1 ������.
    U2[i]:=U[i]*U[i];//���������� ��������� 2 �������.
    Su2:=Su2+U2[i]; //���������� ������������.
    StringGrid2.Cells[1,i]:=floattostr(U[i]); //������� � �������.
    StringGrid2.Cells[2,i]:=floattostr(U2[i]);
  end;

Sum2:=0; //����������� - ����� ��� ���������� ����� 6 �������
for i:=1 to c do
  begin
    if ((i<(N div 2)+1) or (i>c-(N div 2))) then //������� - ����������, ��������� �� ������ ��� i-��� ������ �������. ���� ��� - �������� "-".
    begin
      StringGrid2.Cells[3,i]:='              -'; //����������� ���� � �������� �������.
      StringGrid2.Cells[4,i]:='              -';
      StringGrid2.Cells[5,i]:='              -';
      St[i]:=0;
      T[i]:=0;
    end
    else
      begin
      Sum:=0; //����������� - ����� ��� ���������� ��������� S(t)
      for j:=(i-(N div 2)) to (i+(N div 2)) do Sum:=Sum+U2[j]; //���������� ��������� S(t)
        St[i]:=sqrt(Sum/N); //���������� ���������� �������� S(t)
        if (c mod 2)=0 then T[i]:=i-(c div 2)-0.5 else T[i]:=i-(c div 2)-1; //�������� ���������� 5 �������, � ����� ��������, ����� ����� ��������� ������ ���� �������.
        StringGrid2.Cells[3,i]:=floattostr(St[i]); //������� ������ � �������.
        StringGrid2.Cells[4,i]:=floattostr(T[i]);
        StringGrid2.Cells[5,i]:=floattostr(St[i]*T[i]);
        A11:=A11+1;    //���������� ��������� ������� ��� ���������� ���.
        A12:=A12+A11;
        A22:=A22+A11*A11;
        B1:=B1+St[i]; //���������� ��������� ���������� ����� ��� ���������� ���.
        B2:=B2+St[i]*A11;
      end;
    SSt:=SSt+St[i];  //���������� �������������.
    sum2:=sum2+St[i]*T[i];
    sumti2:=sumti2+T[i]*T[i];
end;

D:=A11*A22-A12*A12; //���������� ������������� �� ������ �������.
D1:=B1*A22-A12*B2;
D2:=A11*B2-A12*B1;
A0:=D1/D; //���������� ������������� �0 � �1, ����������� ��� ���������� ���.
A1:=D2/D;

Ksum:=0; //����������� - ����� ��� ������������ ���.
Kus2:=0; //����������� - ����� ��� ������������ 9 �������.
for i:=1 to c do
  begin
    if ((i<(N div 2)+1) or (i>c-(N div 2))) then  //��������� ����� ���������� ����������.
    begin
      StringGrid2.Cells[6,i]:='              -'; //���������� ������
      StringGrid2.Cells[7,i]:='              -';
      StringGrid2.Cells[8,i]:='              -';
      Ksi[i]:=0;    //��������� ������.
      Us[i]:=0;
      Us2[i]:=0;
    end
    else
      begin
      for j:=(i-(N div 2)) to (i+(N div 2)) do Sum:=Sum+U2[j];
        Ksi[i]:=A0-A1*T[i];  //���������� ��������� �������
        Us[i]:=St[i]-Ksi[i];
        Us2[i]:=Us[i]*Us[i];
        StringGrid2.Cells[6,i]:=floattostr(Ksi[i]);//������� � �������.
        StringGrid2.Cells[7,i]:=floattostr(Us[i]);
        StringGrid2.Cells[8,i]:=floattostr(Us2[i]);
      end;
Ksum:=Ksum+Ksi[i]; //���������� �������������.
Kus2:=Kus2+Us2[i];
end;

StringGrid2.Cells[2,StringGrid2.RowCount-1]:=floattostr(Su2); //������� ������ ������ �������.
StringGrid2.Cells[3,StringGrid2.RowCount-1]:=floattostr(SSt);
StringGrid2.Cells[4,StringGrid2.RowCount-1]:='0';
StringGrid2.Cells[5,StringGrid2.RowCount-1]:=floattostr(Sum2);
StringGrid2.Cells[6,StringGrid2.RowCount-1]:=floattostr(Ksum);
StringGrid2.Cells[7,StringGrid2.RowCount-1]:='              -';
StringGrid2.Cells[8,StringGrid2.RowCount-1]:=floattostr(Kus2);

Memo1.Clear; //������� �������.
Memo1.Lines.Add('������������ ��� ������� Ksi (t[i]):'); //������� ����������� � ��������� �������.
Memo1.Lines.Add('A0 = '+floattostr(A0));    //��� ������������...
Memo1.Lines.Add('A1 = '+floattostr(A1));
Memo1.Lines.Add('');
Memo1.Lines.Add('������� S(t) = '+floattostr(SSt/A11));
Memo1.Lines.Add('');
Bst:=sum2/sumti2;
Memo1.Lines.Add('������� ������ ������������������ b{s(t)} = '+floattostr(Bst));
Memo1.Lines.Add('');
Stst:=sqrt(Kus2/(A11-1));
Memo1.Lines.Add('�������� S(t){s(t)} = '+floattostr(Stst));
Memo1.Lines.Add('');
Mbst:=Stst/sqrt(sumti2);
Memo1.Lines.Add('�������� M b{s(t)} = '+floattostr(Mbst));
Memo1.Lines.Add('');
Memo1.Lines.Add('�������� ��������� ='+floattostr(abs(Bst)/Mbst));

end;

procedure TForm1.N3Click(Sender: TObject);
{���������� ������� ���� "��������� ������"}
var
f:textfile;
i:integer;
begin
  if SaveDialog1.Execute then //����� ������� ����������
    begin
    AssignFile(f,SaveDialog1.FileName); //������
    Rewrite(f); //��������
    Writeln(f,LabeledEdit1.Text); //������ ������ ����������.
    for i:=1 to (StringGrid1.RowCount-1) do //� ����� - ������ ������ ������, � ������� ����� �������, ������ ����.
      begin
      Writeln(f,StringGrid1.Cells[0,i]);
      Writeln(f,StringGrid1.Cells[1,i]);
      Writeln(f,StringGrid1.Cells[2,i]);
      end;
    CloseFile(f); //��������.
    end;
end;

procedure TForm1.N2Click(Sender: TObject);
{���������� ������� ���� "��������� ������"}
var
f:textfile;
i:integer;
s,s0,s1,s2:string; //��������� ������.
begin
  if OpenDialog1.Execute then //����� ������� ��������
    begin
    AssignFile(f,OpenDialog1.FileName);
    Reset(f);
    Readln(f,s); //������ ���������� �������
    LabeledEdit1.Text:=s;
    StringGrid1.RowCount:=strtoint(LabeledEdit1.Text)+1; //��������� �������� �������
    for i:=1 to (StringGrid1.RowCount-1) do  //� ����� - ������ �������.
      begin
      Readln(f,s0); StringGrid1.Cells[0,i]:=s0;
      Readln(f,s1); StringGrid1.Cells[1,i]:=s1;
      Readln(f,s2); StringGrid1.Cells[2,i]:=s2;
      end;
    CloseFile(f); //��������.
    end;
end;

procedure TForm1.N4Click(Sender: TObject);
{���������� ������� ���� "��������� ������"}
begin
    if SavePictureDialog1.Execute then
        Chart1.SaveToBitmapFile(SavePictureDialog1.FileName); //����������.
end;

procedure TForm1.N5Click(Sender: TObject);
{���������� ������� ���� "��������� ����������."}
var
f:textfile;
i,j:integer;
begin
  if SaveDialog1.Execute then
    begin
    AssignFile(f,SaveDialog1.FileName);
    Rewrite(f);
    writeln(f,'������� ��������:');
    writeln(f,'');
    for i:=0 to StringGrid2.ColCount do
    begin
      for j:=0 to StringGrid2.RowCount do write(f,StringGrid2.Cells[i,j]+'      '); //� ������ ��� ������� ����������� �����������.
      writeln(f,'');
    end;
    writeln(f,'');
    writeln(f,Memo1.Text); //���������� ���������� �����������.
    CloseFile(f);
    end
end;

procedure TForm1.N6Click(Sender: TObject);
begin
ShowMessage('1 ��� - �������� ��� ������� ���������� ��������� (�����).');
ShowMessage('2 ��� - ������� ��������������� ������ � ������� ����.');
ShowMessage('3 ��� - �������� �� �������, ����� ��������� ���.');
ShowMessage('4 ��� - ��������� �� ������� "������ ��������", ������� �������� ����������, ������� ��.');
ShowMessage('� ���� "������ � �������" �� ������ ��������� � ��������� ������, � ����� ��������� ���������� � ������.');
end;

procedure TForm1.LabeledEdit1Change(Sender: TObject);
{������������ ������� ������ - ��������� ������� �����}
begin
StringGrid1.RowCount:=strtoint(LabeledEdit1.Text)+1; //��������� �������.
StringGrid1.Cells[0,StringGrid1.RowCount-1]:='0'; //�������������.
StringGrid1.Cells[1,StringGrid1.RowCount-1]:='0';
StringGrid1.Cells[2,StringGrid1.RowCount-1]:='0';
end;

end.
