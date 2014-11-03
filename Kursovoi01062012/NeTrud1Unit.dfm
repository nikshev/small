object NeTrud1Form: TNeTrud1Form
  Left = 585
  Top = 139
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1055#1088#1080#1095#1080#1085#1072' '#1085#1077#1090#1088#1091#1076#1086#1089#1087#1086#1089#1086#1073#1085#1086#1089#1090#1080
  ClientHeight = 638
  ClientWidth = 339
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
    Top = 8
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
        Title.Caption = #1054#1087#1080#1089#1072#1085#1080#1077
        Width = 250
        Visible = True
      end>
  end
  object dbnvgr1: TDBNavigator
    Left = 8
    Top = 608
    Width = 320
    Height = 25
    DataSource = ds1
    TabOrder = 1
  end
  object tbl1: TADOTable
    CursorType = ctStatic
    TableName = 'NeTrud1'
    Left = 40
    Top = 32
    object tbl1Key: TAutoIncField
      FieldName = 'Key'
      ReadOnly = True
    end
    object tbl1Description: TWideStringField
      FieldName = 'Description'
      Size = 255
    end
    object tbl1Code: TWideStringField
      FieldName = 'Code'
      Size = 255
    end
  end
  object ds1: TDataSource
    DataSet = tbl1
    Left = 80
    Top = 80
  end
end
