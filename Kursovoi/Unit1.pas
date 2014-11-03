unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, StdCtrls;

type
  TForm1 = class(TForm)
    grp1: TGroupBox;
    lbl1: TLabel;
    edt1: TEdit;
    lbl2: TLabel;
    edt2: TEdit;
    lbl3: TLabel;
    lbl4: TLabel;
    edt3: TEdit;
    lbl5: TLabel;
    edt4: TEdit;
    lbl6: TLabel;
    edt5: TEdit;
    lbl7: TLabel;
    edt6: TEdit;
    edt7: TEdit;
    edt8: TEdit;
    edt9: TEdit;
    edt10: TEdit;
    edt11: TEdit;
    edt12: TEdit;
    lbl8: TLabel;
    lbl9: TLabel;
    lbl10: TLabel;
    lbl11: TLabel;
    lbl12: TLabel;
    lbl13: TLabel;
    grp2: TGroupBox;
    strngrd1: TStringGrid;
    lbl14: TLabel;
    edt13: TEdit;
    lbl15: TLabel;
    edt14: TEdit;
    lbl16: TLabel;
    edt15: TEdit;
    btn1: TButton;
    procedure btn1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.btn1Click(Sender: TObject);
var
 y:array[1..12] of Integer; //������ ��� �������� ������ � �����
 yy:array[1..12] of Real; //������ �����������
 a:Real; //a
 b1:Real;//b1
 b2:real;//b2
 t:Real;
 tmpSumm1:Real; //��������� �������� ��� ������� a
 tmpSumm2:Real; //��������� �������� ��� ������� b1
 tmpSumm3:Real; //��������� �������� ��� ������� b2
 i:Integer;

begin
 //���������� ������� ������ � ������ 1-������, 12 -�������
 //���� ����� �� ���� ����������� ��������� (������ ��� ������������) �� ����� �������� ������
 y[1]:=StrToInt(Edt1.Text);
 y[2]:=StrToInt(Edt2.Text);
 y[3]:=StrToInt(Edt3.Text);
 y[4]:=StrToInt(Edt4.Text);
 y[5]:=StrToInt(Edt5.Text);
 y[6]:=StrToInt(Edt6.Text);
 y[7]:=StrToInt(Edt7.Text);
 y[8]:=StrToInt(Edt8.Text);
 y[9]:=StrToInt(Edt9.Text);
 y[10]:=StrToInt(Edt10.Text);
 y[11]:=StrToInt(Edt11.Text);
 y[12]:=StrToInt(Edt12.Text);

 //��������� �������� �,b1,b2
  tmpSumm1:=0;
  tmpSumm2:=0;
  tmpSumm3:=0;
  for i:=1 to 12 do
   begin
    t:=(((i-1)*30)*pi)/180; //������� � �������
    tmpSumm1:=tmpSumm1+y[i]; //��������� �������� �������
    tmpSumm2:=tmpSumm2+(y[i]*cos(t)); //������������ ����� ��� b1
    tmpSumm3:=tmpSumm3+(y[i]*sin(t)); //������������ ����� ��� b2
   end;
  a:=tmpSumm1/12; //������� a
  b1:=tmpSumm2/6; //������� b1
  b2:=tmpSumm3/6; //������� b2

  //������� �,b1,b2 �� �����
  edt13.Text:=FloatToStrF(a,ffFixed,3,3);
  edt14.Text:=FloatToStrF(b1,ffFixed,3,3);
  edt15.Text:=FloatToStrF(b2,ffFixed,3,3);

  //������� yyi � ��������� ����� �����������
  for i:=1 to 12 do
   begin
    t:=(((i-1)*30)*pi)/180;  //������� � �������
    yy[i]:=a+b1*cos(t)+b2*sin(t);
    strngrd1.Cells[1,i]:=IntToStr(y[i]);      //������� �� ������� ������ �����
    strngrd1.Cells[2,i]:=IntToStr((i-1)*30);  //������� �������
    strngrd1.Cells[3,i]:=FloatToStrF(cos(t),ffFixed,3,3); //������� cos(ti)
    strngrd1.Cells[4,i]:=FloatToStrF(sin(t),ffFixed,3,3); //������� sin(ti)
    strngrd1.Cells[5,i]:=FloatToStrF(y[i]*cos(t),ffFixed,3,3); //������� yi*cos(ti)
    strngrd1.Cells[6,i]:=FloatToStrF(y[i]*sin(t),ffFixed,3,3); //������� yi*sin(ti)
    strngrd1.Cells[7,i]:=FloatToStrF(yy[i],ffFixed,3,0);   //������� yyi
    strngrd1.Cells[8,i]:=FloatToStrF(y[i]-yy[i],ffFixed,3,0); //������� yi-yyi
   end;

end;

procedure TForm1.FormCreate(Sender: TObject);
begin
 //��������� �������� �����
 strngrd1.Cells[0,0]:='�����, i';
 strngrd1.Cells[0,1]:='������';
 strngrd1.Cells[0,2]:='�������';
 strngrd1.Cells[0,3]:='����';
 strngrd1.Cells[0,4]:='������';
 strngrd1.Cells[0,5]:='���';
 strngrd1.Cells[0,6]:='����';
 strngrd1.Cells[0,7]:='����';
 strngrd1.Cells[0,8]:='������';
 strngrd1.Cells[0,9]:='��������';
 strngrd1.Cells[0,10]:='�������';
 strngrd1.Cells[0,11]:='������';
 strngrd1.Cells[0,12]:='�������';
 strngrd1.Cells[1,0]:='�����';
 strngrd1.Cells[2,0]:='ti';
 strngrd1.Cells[3,0]:='cos(ti)';
 strngrd1.Cells[4,0]:='sin(ti)';
 strngrd1.Cells[5,0]:='yi*cos(ti)';
 strngrd1.Cells[6,0]:='yi*sin(ti)';
 strngrd1.Cells[7,0]:='yyi';
 strngrd1.Cells[8,0]:='yi-yyi';
end;

end.
