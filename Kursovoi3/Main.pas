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
{Служебная процедура копирования данных из 1й таблицы во вторую.}
var i:integer;
begin
StringGrid2.RowCount:=StringGrid1.RowCount+1;//Установка размера второй таблицы.
for i:=1 to (StringGrid1.RowCount-1) do
StringGrid2.Cells[0,i]:=StringGrid1.Cells[0,i]; //Копирование дат.
end;

procedure TForm1.UpDown1Click(Sender: TObject; Button: TUDBtnType);
{Обработчик нажатия кнопки вверх/вниз возле поля ввода.
Увеличивает или уменьшает количество строк в таблице ввода данных.}
begin
StringGrid1.RowCount:=strtoint(LabeledEdit1.Text)+1; //Изменение размера.
StringGrid1.Cells[0,StringGrid1.RowCount-1]:='0'; //Инициализация.
StringGrid1.Cells[1,StringGrid1.RowCount-1]:='0';
StringGrid1.Cells[2,StringGrid1.RowCount-1]:='0';
end;

procedure TForm1.FormActivate(Sender: TObject);
{Обработчик активации формы, т.е. срабатывает при запуске приложения.}
begin
StringGrid1.Cells[0,0]:='Время';
StringGrid1.Cells[1,0]:='Практически';
StringGrid1.Cells[2,0]:='Теоретически';
StringGrid1.Cells[0,1]:='0';
StringGrid1.Cells[1,1]:='0';
StringGrid1.Cells[2,1]:='0';

StringGrid2.Cells[0,0]:='Время';
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
{Обработчик щелчка по графику - процедура прорисовки.}
var i:integer;
begin
Series1.Clear; //Очистка полей данных.
Series2.Clear;
for i:=1 to (StringGrid1.RowCount-1) do
begin
Series1.AddXY(strtofloat(StringGrid1.Cells[0,i]),strtofloat(StringGrid1.Cells[1,i])); //Добавление точек.
Series2.AddXY(strtofloat(StringGrid1.Cells[0,i]),strtofloat(StringGrid1.Cells[2,i]));
end;
end;

procedure TForm1.TabSheet2Show(Sender: TObject);
{Копирование данных из таблицы в таблицу в самом начале.}
begin
Form1.GridCopy;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
{Самая главная процедура - вычисление всех основных показателей идет именно тут.
Комментарии к переменным по ходу программы.}
var c,i,j,PC,N:integer;
Yp,Yt,U,U2,St,T,Ksi,Us,Us2:array[1..1000] of double;
Su2,Sum,Sum2,SST,A11,A12,A22,B1,B2,D,D1,D2,A0,A1,Ksum,Kus2,sumti2,Bst,Stst,Mbst:double;
begin
c:=StringGrid1.RowCount-1;//Количество строк с данными. Иначе - количество точек.
StringGrid2.Cells[0,StringGrid2.RowCount-1]:='              -';//Забивка лишних граф таблицы.
StringGrid2.Cells[1,StringGrid2.RowCount-1]:='              -';
Su2:=0; //Аккумулятор - отвечает за сумму в 3 столбце.
SSt:=0; //Аккумулятор - отвечает за сумму в 4 столбце.
N:=strtoint(LabeledEdit2.Text);//Величина подпериода.
if (N mod 2)=0 then begin //Проверка на нечетность.
ShowMessage('Величина подпериода должна быть нечетной!');
Exit;
end;
PC:=strtoint(LabeledEdit1.Text)-N+1; //Количество подпериодов в отрезке.
A11:=0; //Элементы матрицы для нахождения А0 и А1, необходимых для расчета Кси.
A12:=0;
A22:=0;
B1:=0;
B2:=0;
sumti2:=0; //Аккумулятор - для нахождения суммы квадратов T[i];
for i:=1 to c do
  begin
    Yp[i]:=strtofloat(StringGrid1.Cells[1,i]);//Выемка данных из первой таблицы.
    Yt[i]:=strtofloat(StringGrid1.Cells[2,i]);
    U[i]:=Yt[i]-Yp[i];//Вычисление элементов 1 стобца.
    U2[i]:=U[i]*U[i];//Вычисление элементов 2 столбца.
    Su2:=Su2+U2[i]; //Увеличение аккумулятора.
    StringGrid2.Cells[1,i]:=floattostr(U[i]); //Выписка в таблицу.
    StringGrid2.Cells[2,i]:=floattostr(U2[i]);
  end;

Sum2:=0; //Аккумулятор - нужен для вычисления суммы 6 столбца
for i:=1 to c do
  begin
    if ((i<(N div 2)+1) or (i>c-(N div 2))) then //Условие - определяет, необходим ли расчет для i-той строки таблицы. Если нет - ставится "-".
    begin
      StringGrid2.Cells[3,i]:='              -'; //Простановка тире в ненужных клетках.
      StringGrid2.Cells[4,i]:='              -';
      StringGrid2.Cells[5,i]:='              -';
      St[i]:=0;
      T[i]:=0;
    end
    else
      begin
      Sum:=0; //Аккумулятор - нужен для вычисления числителя S(t)
      for j:=(i-(N div 2)) to (i+(N div 2)) do Sum:=Sum+U2[j]; //Вычисление числителя S(t)
        St[i]:=sqrt(Sum/N); //вычисление очередного значения S(t)
        if (c mod 2)=0 then T[i]:=i-(c div 2)-0.5 else T[i]:=i-(c div 2)-1; //Условное заполнение 5 столбца, с таким расчетом, чтобы сумма элементов стобца была нулевой.
        StringGrid2.Cells[3,i]:=floattostr(St[i]); //Выписка данных в таблицу.
        StringGrid2.Cells[4,i]:=floattostr(T[i]);
        StringGrid2.Cells[5,i]:=floattostr(St[i]*T[i]);
        A11:=A11+1;    //вычисление элементов матрицы для нахождения Кси.
        A12:=A12+A11;
        A22:=A22+A11*A11;
        B1:=B1+St[i]; //вычисление элементов свободного члена для нахождения Кси.
        B2:=B2+St[i]*A11;
      end;
    SSt:=SSt+St[i];  //Увеличение аккумуляторов.
    sum2:=sum2+St[i]*T[i];
    sumti2:=sumti2+T[i]*T[i];
end;

D:=A11*A22-A12*A12; //Вычисление определителей по методу Крамера.
D1:=B1*A22-A12*B2;
D2:=A11*B2-A12*B1;
A0:=D1/D; //Нахождение коэффициентов А0 и А1, необходимых для вычисления Кси.
A1:=D2/D;

Ksum:=0; //Аккумулятор - нужен для суммирования Кси.
Kus2:=0; //Аккумулятор - нужен для суммирования 9 столбца.
for i:=1 to c do
  begin
    if ((i<(N div 2)+1) or (i>c-(N div 2))) then  //Структура цикла аналогична предыдущей.
    begin
      StringGrid2.Cells[6,i]:='              -'; //Заполнение лишних
      StringGrid2.Cells[7,i]:='              -';
      StringGrid2.Cells[8,i]:='              -';
      Ksi[i]:=0;    //Обнуление лишних.
      Us[i]:=0;
      Us2[i]:=0;
    end
    else
      begin
      for j:=(i-(N div 2)) to (i+(N div 2)) do Sum:=Sum+U2[j];
        Ksi[i]:=A0-A1*T[i];  //Вычисление элементов таблицы
        Us[i]:=St[i]-Ksi[i];
        Us2[i]:=Us[i]*Us[i];
        StringGrid2.Cells[6,i]:=floattostr(Ksi[i]);//выписка в таблицу.
        StringGrid2.Cells[7,i]:=floattostr(Us[i]);
        StringGrid2.Cells[8,i]:=floattostr(Us2[i]);
      end;
Ksum:=Ksum+Ksi[i]; //Увеличение аккумуляторов.
Kus2:=Kus2+Us2[i];
end;

StringGrid2.Cells[2,StringGrid2.RowCount-1]:=floattostr(Su2); //Выписка нижней строки таблицы.
StringGrid2.Cells[3,StringGrid2.RowCount-1]:=floattostr(SSt);
StringGrid2.Cells[4,StringGrid2.RowCount-1]:='0';
StringGrid2.Cells[5,StringGrid2.RowCount-1]:=floattostr(Sum2);
StringGrid2.Cells[6,StringGrid2.RowCount-1]:=floattostr(Ksum);
StringGrid2.Cells[7,StringGrid2.RowCount-1]:='              -';
StringGrid2.Cells[8,StringGrid2.RowCount-1]:=floattostr(Kus2);

Memo1.Clear; //Очистка области.
Memo1.Lines.Add('Коэффициенты для расчета Ksi (t[i]):'); //Выписка результатов в текстовую область.
Memo1.Lines.Add('A0 = '+floattostr(A0));    //Без комментариев...
Memo1.Lines.Add('A1 = '+floattostr(A1));
Memo1.Lines.Add('');
Memo1.Lines.Add('Среднее S(t) = '+floattostr(SSt/A11));
Memo1.Lines.Add('');
Bst:=sum2/sumti2;
Memo1.Lines.Add('Средняя ошибка репрезентативности b{s(t)} = '+floattostr(Bst));
Memo1.Lines.Add('');
Stst:=sqrt(Kus2/(A11-1));
Memo1.Lines.Add('Величина S(t){s(t)} = '+floattostr(Stst));
Memo1.Lines.Add('');
Mbst:=Stst/sqrt(sumti2);
Memo1.Lines.Add('Величина M b{s(t)} = '+floattostr(Mbst));
Memo1.Lines.Add('');
Memo1.Lines.Add('Критерий Стьюдента ='+floattostr(abs(Bst)/Mbst));

end;

procedure TForm1.N3Click(Sender: TObject);
{Обработчик нажатия меню "Сохранить данные"}
var
f:textfile;
i:integer;
begin
  if SaveDialog1.Execute then //Вызов диалога сохранения
    begin
    AssignFile(f,SaveDialog1.FileName); //Связка
    Rewrite(f); //открытие
    Writeln(f,LabeledEdit1.Text); //Запись общего количества.
    for i:=1 to (StringGrid1.RowCount-1) do //в цикле - запись каждой клетки, в порядке слева направо, сверху вниз.
      begin
      Writeln(f,StringGrid1.Cells[0,i]);
      Writeln(f,StringGrid1.Cells[1,i]);
      Writeln(f,StringGrid1.Cells[2,i]);
      end;
    CloseFile(f); //Закрытие.
    end;
end;

procedure TForm1.N2Click(Sender: TObject);
{Обработчик нажатия меню "Загрузить данные"}
var
f:textfile;
i:integer;
s,s0,s1,s2:string; //Служебные строки.
begin
  if OpenDialog1.Execute then //Вызов диалога открытия
    begin
    AssignFile(f,OpenDialog1.FileName);
    Reset(f);
    Readln(f,s); //Чтение количества записей
    LabeledEdit1.Text:=s;
    StringGrid1.RowCount:=strtoint(LabeledEdit1.Text)+1; //Установка размеров таблицы
    for i:=1 to (StringGrid1.RowCount-1) do  //В цикле - чтение записей.
      begin
      Readln(f,s0); StringGrid1.Cells[0,i]:=s0;
      Readln(f,s1); StringGrid1.Cells[1,i]:=s1;
      Readln(f,s2); StringGrid1.Cells[2,i]:=s2;
      end;
    CloseFile(f); //Закрытие.
    end;
end;

procedure TForm1.N4Click(Sender: TObject);
{Обработчик нажатия меню "Сохранить график"}
begin
    if SavePictureDialog1.Execute then
        Chart1.SaveToBitmapFile(SavePictureDialog1.FileName); //Сохранение.
end;

procedure TForm1.N5Click(Sender: TObject);
{Обработчик нажатия меню "Сохранить результаты."}
var
f:textfile;
i,j:integer;
begin
  if SaveDialog1.Execute then
    begin
    AssignFile(f,SaveDialog1.FileName);
    Rewrite(f);
    writeln(f,'Таблица значений:');
    writeln(f,'');
    for i:=0 to StringGrid2.ColCount do
    begin
      for j:=0 to StringGrid2.RowCount do write(f,StringGrid2.Cells[i,j]+'      '); //В циклах вся таблица сохраняется поэлементно.
      writeln(f,'');
    end;
    writeln(f,'');
    writeln(f,Memo1.Text); //Сохранение текстового содержимого.
    CloseFile(f);
    end
end;

procedure TForm1.N6Click(Sender: TObject);
begin
ShowMessage('1 шаг - выберите или введите количество измерений (точек).');
ShowMessage('2 шаг - введите соответствующие данные в таблицу ниже.');
ShowMessage('3 шаг - кликните по графику, чтобы построить его.');
ShowMessage('4 шаг - перейдите на вкладку "Расчет значений", введите величину подпериода, нажмите ОК.');
ShowMessage('В меню "Работа с данными" вы можете сохранять и загружать данные, а также сохранять результаты и график.');
end;

procedure TForm1.LabeledEdit1Change(Sender: TObject);
{Дублирование нажатия кнопок - обработка прямого ввода}
begin
StringGrid1.RowCount:=strtoint(LabeledEdit1.Text)+1; //Изменение размера.
StringGrid1.Cells[0,StringGrid1.RowCount-1]:='0'; //Инициализация.
StringGrid1.Cells[1,StringGrid1.RowCount-1]:='0';
StringGrid1.Cells[2,StringGrid1.RowCount-1]:='0';
end;

end.
