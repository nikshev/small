unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Grids;

type
  TForm1 = class(TForm)
    strngrd1: TStringGrid;
    btn1: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure btn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;



implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
var
  i:Integer;
begin
 //��������� ������������� �������
 for i:=1987 to 1997 do
  strngrd1.Cells[0,i-1986]:=IntToStr(i);
 //��������� ������������� ������
  strngrd1.Cells[0,0]:='���';
  strngrd1.Cells[1,0]:='���� �����.';
  strngrd1.Cells[2,0]:='�����������';
  strngrd1.Cells[3,0]:='�����';
  strngrd1.Cells[4,0]:='����������';
  strngrd1.Cells[5,0]:='deltaX';
  strngrd1.Cells[6,0]:='deltaX*(U(YI+1))^2';
  strngrd1.Cells[7,0]:='deltaX^2';
  strngrd1.Cells[8,0]:='deltaY';
  strngrd1.Cells[9,0]:='��������� ��������';
  //� �������� �������� ������� ������� ������ ����������
  strngrd1.Cells[1,1]:='45';
  strngrd1.Cells[2,1]:='13';
  strngrd1.Cells[1,2]:='36';
  strngrd1.Cells[2,2]:='18';
  strngrd1.Cells[1,3]:='47';
  strngrd1.Cells[2,3]:='16,5';
  strngrd1.Cells[1,4]:='33';
  strngrd1.Cells[2,4]:='15,1';
  strngrd1.Cells[1,5]:='42';
  strngrd1.Cells[2,5]:='14,5';
  strngrd1.Cells[1,6]:='51';
  strngrd1.Cells[2,6]:='18';
  strngrd1.Cells[1,7]:='36';
  strngrd1.Cells[2,7]:='22,4';
  strngrd1.Cells[1,8]:='42';
  strngrd1.Cells[2,8]:='18,7';
  strngrd1.Cells[1,9]:='35';
  strngrd1.Cells[2,9]:='23';
  strngrd1.Cells[1,10]:='40';
  strngrd1.Cells[2,10]:='19,8';
  strngrd1.Cells[1,11]:='44';
  strngrd1.Cells[2,11]:='19';

end;

//��������� �������
procedure TForm1.btn1Click(Sender: TObject);
 var
  ti:array[1..11] of Integer;
  xi:array[1..11] of Real;      //���� ���������
  yi:array[1..11] of Real;      //�����������
  yyi:array[1..11] of Real;     //�����
  deltaX:array[1..11] of Real;  //������ x
  Uyi:array[1..11] of Real;     //����������
  xx:Real;                      //������� ���� ���������
  Uyi_deltaX:Real;
  sum_Uyi_deltaX:Real;
  sum_Uyi_Uyi:Real;
  sum_deltaX_deltaX:Real;
  r:Real; //����������� ���������� � ������ ���� � 1 ���
  bx:Real; //����������� ������c��
  i:integer;
 begin
 xx:=0;    //���� �2
 sum_Uyi_Uyi:=0;
 for i:=1 to 11 do //���� �3
  begin
   ti[i]:=-6+i; //���� �4
   //���� �5
   xi[i]:= StrToFloat(strngrd1.Cells[1,i]);
   yi[i]:= StrToFloat(strngrd1.Cells[2,i]);
   //���� �6
   xx:=xx+xi[i];
   yyi[i]:=18+0.6*ti[i];
   Uyi[i]:=yi[i]-yyi[i];
   sum_Uyi_Uyi:=sum_Uyi_Uyi+Uyi[i]*Uyi[i];
   //���� �7
   strngrd1.Cells[3,i]:=FloatToStrF(yyi[i],ffFixed,8,2);
   strngrd1.Cells[4,i]:=FloatToStrF(Uyi[i],ffFixed,8,2);
   strngrd1.Cells[8,i]:=FloatToStrF(Uyi[i]*Uyi[i],ffFixed,8,2);
  end;
   xx:=xx/11; //���� �8
   sum_Uyi_deltaX:=0;
   sum_deltaX_deltaX:=0;

  for i:=1 to 11 do //���� �9
   begin
     deltaX[i]:=xi[i]-xx;   //���� �10
     sum_deltaX_deltaX:=sum_deltaX_deltaX+deltaX[i]*deltaX[i];
     strngrd1.Cells[5,i]:=FloatToStrF(deltaX[i],ffFixed,8,2); //���� �11
     if (i<11) then   //���� �12
      begin
       Uyi_deltaX:=deltaX[i]* Uyi[i+1];  //���� �13
       sum_Uyi_deltaX:=sum_Uyi_deltaX+Uyi_deltaX;
       strngrd1.Cells[6,i]:=FloatToStrF(Uyi_deltaX,ffFixed,8,2); //���� �14
      end;
     strngrd1.Cells[7,i]:=FloatToStrF(deltaX[i]*deltaX[i],ffFixed,8,2); //���� �15
   end;
    //���� �16
    bx:=sum_Uyi_deltaX/sum_deltaX_deltaX;
    r:= sum_Uyi_deltaX/sqrt(sum_deltaX_deltaX*sum_Uyi_Uyi);
  for i:=1 to 11 do  //���� �17
   begin
    //������������ ��������� �������
    if (i>1) then   //���� �18
     strngrd1.Cells[9,i]:=FloatToStrF(yyi[i]+deltaX[i-1]*bx,ffFixed,8,1)   //���� �19
    else
     strngrd1.Cells[9,i]:=FloatToStrF(yyi[i],ffFixed,8,1);    //���� �20
   end;
   ShowMessage('r='+FloatToStrF(r,ffFixed,8,2)); //���� �21
 end;

end.
