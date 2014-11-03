unit EnterprisesUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ExtCtrls, DBCtrls, Grids, DBGrids;

type
  TEnterprisesForm = class(TForm)
    dbgrd1: TDBGrid;
    dbnvgr1: TDBNavigator;
    ds1: TDataSource;
    tbl1: TADOTable;
    atncfldtbl1Key: TAutoIncField;
    wdstrngfldtbl1enterprise: TWideStringField;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  EnterprisesForm: TEnterprisesForm;

implementation

{$R *.dfm}

procedure TEnterprisesForm.FormCreate(Sender: TObject);
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
   tbl1.TableName:='Enterprises';
   tbl1.Active:=True;
  except
   on E : Exception do
      ShowMessage(E.ClassName+' вызвана ошибка, с сообщением : '+E.Message);
  end;
end;

procedure TEnterprisesForm.FormClose(Sender: TObject;
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
