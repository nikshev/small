unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, StdCtrls, Buttons;

type
  TForm1 = class(TForm)
    grp1: TGroupBox;
    btn1: TBitBtn;
    lbl5: TLabel;
    edt5: TEdit;
    grp2: TGroupBox;
    lbl1: TLabel;
    edt1: TEdit;
    lbl2: TLabel;
    edt2: TEdit;
    grp3: TGroupBox;
    edt3: TEdit;
    edt4: TEdit;
    lbl4: TLabel;
    lbl3: TLabel;
    grp4: TGroupBox;
    grp5: TGroupBox;
    strngrd1: TStringGrid;
    grp6: TGroupBox;
    lbl6: TLabel;
    edt6: TEdit;
    lbl7: TLabel;
    edt7: TEdit;
    lbl8: TLabel;
    edt8: TEdit;
    lbl9: TLabel;
    edt9: TEdit;
    lbl10: TLabel;
    edt10: TEdit;
    grp7: TGroupBox;
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

procedure TForm1.btn1Click(Sender: TObject);
var
x:array[1..4] of Integer; //Массив X 1..2 для функции x+2y<=12, а 3..4 для функции x+y<=9
y:array[1..4] of Real; //Массив Y 1..2 для функции x+2y<=12, а 3..4 для функции x+y<=9
Z,Zmax,Zmin:Real;
a,b:Real;
D:array[1..2] of Real;
i:Integer;

begin
  //Считываем данные
  x[1]:=StrToInt(edt1.text); //Если не корректно введено число то выскочет ошибка
  x[2]:=StrToInt(edt2.text); //Если не корректно введено число то выскочет ошибка
  x[3]:=StrToInt(edt3.text); //Если не корректно введено число то выскочет ошибка
  x[4]:=StrToInt(edt4.text); //Если не корректно введено число то выскочет ошибка
  Z:=StrToFloat(edt5.Text); //Если не корректно введено число то выскочет ошибка
   if (x[1]>=0) and (x[2]>=0) and (x[3]>=0) and (x[4]>=0) then //По условию x>=0
   if (x[1]<13) and (x[2]<13) and (x[3]<9) and (x[4]<9) then//По условию y>=0
   begin
     //Рассчитываем точки для функции x+2y<=12
     y[1]:=(12-x[1])/2;
     y[2]:=(12-x[2])/2;
     //Рассчитываем точки для функции x+y<=9
     y[3]:=9-x[3];
     y[4]:=9-x[4];
     //Выводим X и Y
     strngrd1.Cells[0,0]:='X';
     strngrd1.Cells[1,0]:='Y';
     for i:=1 to 4 do
      begin
       strngrd1.Cells[0,i]:=IntToStr(x[i]);
       strngrd1.Cells[1,i]:=FloatToStrF(y[i],ffFixed,3,3);
      end;
     //Расчитываем a,b, Zmax, D,Zmin
     a:=Sqrt(Z/2);
     b:=Sqrt(Z);
     Zmax:=2*25+49;
     D[1]:=38/9; //X
     D[2]:=35/9; //Y
     Zmin:=2*sqr((D[1]-5))+sqr((D[2]-7));
     //Выводим a,b Zmax
     edt6.Text:=FloatToStrF(a,ffFixed,3,3);
     edt7.Text:=FloatToStrF(b,ffFixed,3,3);
     edt8.Text:=FloatToStrF(Zmax,ffFixed,3,3);
     edt9.Text:=FloatToStrF(Zmin,ffFixed,3,3);
     edt10.Text:=FloatToStrF(D[1],ffFixed,3,3)+';'+FloatToStrF(D[2],ffFixed,3,3);
   end
   else
    MessageDlg('Не правильно введены исходные данные',mtConfirmation,[mbOK],MB_OK);


end;

end.
