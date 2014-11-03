object OtherForm: TOtherForm
  Left = 325
  Top = 109
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1048#1085#1086#1077
  ClientHeight = 630
  ClientWidth = 336
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
        Title.Caption = #1048#1085#1086#1077
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
    TableName = 'Other'
    Left = 80
    Top = 136
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
    Left = 96
    Top = 240
  end
end
