unit IntuitUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Panel1: TPanel;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    Button9: TButton;
    Button10: TButton;
    ProgressBar1: TProgressBar;
    ProgressBar2: TProgressBar;
    Label1: TLabel;
    Label2: TLabel;
    Button11: TButton;
    procedure FormCreate(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure AddNum(Num:String);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure UpdateApplication;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  source:String;
  dest  :String;
  index :Integer;


implementation

{$R *.dfm}

procedure TForm1.Button10Click(Sender: TObject);
begin
 AddNum('0');
end;

procedure TForm1.Button11Click(Sender: TObject);
var
 i:Integer;
begin
 ProgressBar1.Position:=0;
 ProgressBar2.Position:=0;

 //Формируем строку
 source:='';
 dest:='';
 randomize;
 for i := 1 to 36 do
  source:=source+IntToStr(Random(10));
 index:=1;
 Label1.Caption:='Угадано:';
 Label2.Caption:='Пройдено:';
 Panel1.Caption:='';
 UpdateApplication;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
 AddNum('1');
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  AddNum('2');
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
 AddNum('3');
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
 AddNum('4');
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
 AddNum('5');
end;

procedure TForm1.Button6Click(Sender: TObject);
begin
 AddNum('6');
end;

procedure TForm1.Button7Click(Sender: TObject);
begin
 AddNum('7');
end;

procedure TForm1.Button8Click(Sender: TObject);
begin
 AddNum('8');
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
 AddNum('9');
end;

procedure TForm1.FormCreate(Sender: TObject);
var
 i:Integer;
begin
 Panel1.Caption:='';
 ProgressBar1.Position:=0;
 ProgressBar2.Position:=0;
 //Формируем строку
 source:='';
 dest:='';
 randomize;
 for i := 1 to 36 do
  source:=source+IntToStr(Random(10));
 index:=1;
 Label1.Caption:='Угадано:';
 Label2.Caption:='Пройдено:';

end;

procedure TForm1.AddNum(Num:String);
var
 i:Integer;
 U:Integer;
 P:Integer;
begin
 if length(dest)<>36 then
  begin
    dest:=dest+Num;
    Panel1.Caption:=source[index];
    Inc(index);
    i:=1;
    U:=0; //Угадано
    P:=0; //Пройдено в ноль
    while i<index do
     begin
      if dest[i]=source[i] then
       Inc(U);
      Inc(P);
      Inc(i);
     end;
     Label1.Caption:='Угадано '+IntToStr(U)+' из 36:';
     ProgressBar1.Position:=Trunc((U/36)*100);
     Label2.Caption:='Пройдено '+IntToStr(P)+' цифр:';
     ProgressBar2.Position:=Trunc((P/36)*100);
  end
  else
   ShowMessage('Пройдены все 36 цифр!');
end;

procedure TForm1.UpdateApplication;
// Updates (repaints where nesc) all windows of the app
  function UpdateWindow(hWnd: HWND; LParam: longint): bool; stdcall;
  begin
    Result := True;
    Windows.UpdateWindow(hWnd);
  end;
begin
  EnumWindows(@UpdateWindow, 0);
end;
end.
