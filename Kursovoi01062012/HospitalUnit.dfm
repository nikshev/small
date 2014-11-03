object HospitalsForm: THospitalsForm
  Left = 490
  Top = 212
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #1051#1077#1095#1077#1073#1085#1099#1077' '#1091#1095#1077#1088#1077#1078#1076#1077#1085#1080#1103
  ClientHeight = 625
  ClientWidth = 417
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
    Width = 401
    Height = 577
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
        FieldName = 'Hospital_name'
        Title.Caption = #1051#1077#1095#1077#1073#1085#1086#1077' '#1091#1095#1077#1088#1077#1078#1076#1077#1085#1080#1077
        Width = 132
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Address'
        Title.Caption = #1040#1076#1088#1077#1089
        Width = 124
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'OGRN'
        Title.Caption = #1054#1043#1056#1053'/'#1054#1043#1056#1053#1048
        Width = 126
        Visible = True
      end>
  end
  object dbnvgr1: TDBNavigator
    Left = 8
    Top = 592
    Width = 400
    Height = 25
    DataSource = ds1
    TabOrder = 1
  end
  object tbl1: TADOTable
    CursorType = ctStatic
    TableName = 'Hospitals'
    Left = 48
    Top = 64
    object tbl1Key: TAutoIncField
      FieldName = 'Key'
      ReadOnly = True
    end
    object tbl1Hospital_name: TWideStringField
      FieldName = 'Hospital_name'
      Size = 255
    end
    object tbl1Address: TWideStringField
      FieldName = 'Address'
      Size = 255
    end
    object tbl1OGRN: TWideStringField
      FieldName = 'OGRN'
      Size = 255
    end
  end
  object ds1: TDataSource
    DataSet = tbl1
    Left = 48
    Top = 120
  end
end
