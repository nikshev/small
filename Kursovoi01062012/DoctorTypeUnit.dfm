object DoctorTypeForm: TDoctorTypeForm
  Left = 214
  Top = 113
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1044#1086#1083#1078#1085#1086#1089#1090#1080
  ClientHeight = 628
  ClientWidth = 341
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object dbgrd1: TDBGrid
    Left = 8
    Top = 0
    Width = 320
    Height = 593
    DataSource = ds1
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'Key'
        Title.Caption = #1050#1086#1076
        Width = 50
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DoctorType'
        Title.Caption = #1044#1086#1083#1078#1085#1086#1089#1090#1100
        Width = 250
        Visible = True
      end>
  end
  object dbnvgr1: TDBNavigator
    Left = 8
    Top = 600
    Width = 320
    Height = 25
    DataSource = ds1
    TabOrder = 1
  end
  object tbl1: TADOTable
    ConnectionString = 
      'Provider=MSDASQL.1;Persist Security Info=False;Extended Properti' +
      'es="DBQ=D:\Eugene\NewWork\Razrabotka\'#1059#1095#1077#1073#1072'\Kursovoi01062012\'#1053#1086#1074#1072 +
      #1103' '#1074#1077#1088#1089#1080#1103'\'#1055#1088#1086#1075#1088#1072#1084#1084#1072'\base.accdb;DefaultDir=D:\Eugene\NewWork\Razra' +
      'botka\'#1059#1095#1077#1073#1072'\Kursovoi01062012\'#1053#1086#1074#1072#1103' '#1074#1077#1088#1089#1080#1103'\'#1055#1088#1086#1075#1088#1072#1084#1084#1072'\;Driver={Mic' +
      'rosoft Access Driver (*.mdb, *.accdb)};DriverId=25;FIL=MS Access' +
      ';MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions' +
      '=0;Threads=3;UID=admin;UserCommitSync=Yes;"'
    CursorType = ctStatic
    TableName = 'DoctorType'
    Left = 128
    Top = 128
    object atncfldtbl1Key: TAutoIncField
      FieldName = 'Key'
      ReadOnly = True
    end
    object wdstrngfldtbl1DoctorType: TWideStringField
      FieldName = 'DoctorType'
      Size = 255
    end
  end
  object ds1: TDataSource
    DataSet = tbl1
    Left = 120
    Top = 256
  end
end
