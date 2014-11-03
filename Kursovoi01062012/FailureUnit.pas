unit FailureUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ExtCtrls, DBCtrls, Grids, DBGrids;

type
  TFailureForm = class(TForm)
    dbgrd1: TDBGrid;
    dbnvgr1: TDBNavigator;
    tbl1: TADOTable;
    ds1: TDataSource;
    atncfldtbl1Key: TAutoIncField;
    wdstrngfldtbl1Description: TWideStringField;
    wdstrngfldtbl1Code: TWideStringField;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FailureForm: TFailureForm;

implementation

{$R *.dfm}

procedure TFailureForm.FormCreate(Sender: TObject);
var
 ConnectionString:string;
begin
  try
   ConnectionString:='Provider=MSDASQL.1;Persist Security Info=False;'+
                     'Extended Properties="DBQ='+ExtractFilePath(Application.ExeName)+
                     'base.accdb;DefaultDir='+ExtractFilePath(Application.ExeName)+
                     ';Driver={Microsoft Access Driver (*.mdb, *.accdb)};DriverId=25;'+
                     'FIL=MS Access;MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;'+
                     'SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"';
   tbl1.ConnectionString:=ConnectionString;
   tbl1.TableName:='Failure';
   tbl1.Active:=True;
  except
   on E : Exception do
      ShowMessage(E.ClassName+' вызвана ошибка, с сообщением : '+E.Message);
  end;
end;

procedure TFailureForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  try
   if tbl1.Active then
    tbl1.Active:=False;
  except
   on E : Exception do
      ShowMessage(E.ClassName+' вызвана ошибка, с сообщением : '+E.Message);
  end;
end;

end.
