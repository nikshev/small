unit NeTrud1Unit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ExtCtrls, DBCtrls, Grids, DBGrids;

type
  TNeTrud1Form = class(TForm)
    dbgrd1: TDBGrid;
    dbnvgr1: TDBNavigator;
    tbl1: TADOTable;
    ds1: TDataSource;
    tbl1Key: TAutoIncField;
    tbl1Description: TWideStringField;
    tbl1Code: TWideStringField;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  NeTrud1Form: TNeTrud1Form;

implementation

{$R *.dfm}

procedure TNeTrud1Form.FormCreate(Sender: TObject);
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
   tbl1.TableName:='NeTrud1';
   tbl1.Active:=True;
  except
   on E : Exception do
      ShowMessage(E.ClassName+' вызвана ошибка, с сообщением : '+E.Message);
  end;
end;

procedure TNeTrud1Form.FormClose(Sender: TObject;
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
