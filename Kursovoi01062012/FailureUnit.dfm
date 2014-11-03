object FailureForm: TFailureForm
  Left = 209
  Top = 109
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1053#1072#1088#1091#1096#1077#1085#1080#1077' '#1088#1077#1078#1080#1084#1072
  ClientHeight = 628
  ClientWidth = 338
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
        FieldName = 'Code'
        Title.Caption = #1050#1086#1076
        Width = 50
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Description'
        Title.Caption = #1053#1072#1088#1091#1096#1077#1085#1080#1077' '#1088#1077#1078#1080#1084#1072
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
    CursorType = ctStatic
    TableName = 'Failure'
    Left = 56
    Top = 56
    object atncfldtbl1Key: TAutoIncField
      FieldName = 'Key'
      ReadOnly = True
    end
    object wdstrngfldtbl1Description: TWideStringField
      FieldName = 'Description'
      Size = 255
    end
    object wdstrngfldtbl1Code: TWideStringField
      FieldName = 'Code'
      Size = 255
    end
  end
  object ds1: TDataSource
    DataSet = tbl1
    Left = 56
    Top = 112
  end
end
