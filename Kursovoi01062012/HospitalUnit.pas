unit HospitalUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, ExtCtrls, DBCtrls, Grids, DBGrids;

type
  THospitalsForm = class(TForm)
    tbl1: TADOTable;
    dbgrd1: TDBGrid;
    ds1: TDataSource;
    dbnvgr1: TDBNavigator;
    tbl1Key: TAutoIncField;
    tbl1Hospital_name: TWideStringField;
    tbl1Address: TWideStringField;
    tbl1OGRN: TWideStringField;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  HospitalsForm: THospitalsForm;

implementation

{$R *.dfm}

procedure THospitalsForm.FormCreate(Sender: TObject);
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
   tbl1.TableName:='Hospitals';
   tbl1.Active:=True;
  except
   on E : Exception do
      ShowMessage(E.ClassName+' вызвана ошибка, с сообщением : '+E.Message);
  end;
end;

procedure THospitalsForm.FormClose(Sender: TObject;
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
