object Form3: TForm3
  Left = 0
  Top = 0
  Caption = #1052#1086#1085#1080#1090#1086#1088#1080#1085#1075' '#1076#1080#1089#1087#1072#1085#1089#1077#1088#1080#1079#1072#1094#1080#1080' '#1076#1077#1090#1077#1081' '#1074' '#1074#1086#1079#1088#1072#1089#1090#1077' 14 '#1083#1077#1090
  ClientHeight = 589
  ClientWidth = 868
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 439
    Top = 8
    Width = 425
    Height = 569
    Caption = #1056#1077#1079#1091#1083#1100#1090#1072#1090
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    object StringGrid1: TStringGrid
      Left = 16
      Top = 24
      Width = 406
      Height = 529
      ColCount = 4
      RowCount = 27
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSizing, goColSizing]
      TabOrder = 0
      ColWidths = (
        64
        102
        110
        83)
      RowHeights = (
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24
        24)
    end
  end
  object GroupBox2: TGroupBox
    Left = 8
    Top = 8
    Width = 425
    Height = 113
    Caption = #1055#1086#1080#1089#1082
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    object Label1: TLabel
      Left = 18
      Top = 24
      Width = 29
      Height = 16
      Caption = #1051#1055#1059
    end
    object Label2: TLabel
      Left = 18
      Top = 46
      Width = 39
      Height = 16
      Caption = #1054#1090#1095#1077#1090
    end
    object ComboBox1: TComboBox
      Left = 139
      Top = 21
      Width = 262
      Height = 24
      ItemHeight = 16
      TabOrder = 0
    end
    object ComboBox2: TComboBox
      Left = 139
      Top = 51
      Width = 262
      Height = 24
      ItemHeight = 16
      TabOrder = 1
    end
    object Button1: TButton
      Left = 139
      Top = 81
      Width = 166
      Height = 25
      Caption = #1055#1086#1080#1089#1082
      TabOrder = 2
      OnClick = Button1Click
    end
  end
  object GroupBox3: TGroupBox
    Left = 8
    Top = 127
    Width = 425
    Height = 90
    Caption = #1048#1084#1087#1086#1088#1090' '#1076#1072#1085#1085#1099#1093
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    object Label3: TLabel
      Left = 18
      Top = 24
      Width = 73
      Height = 16
      Caption = #1048#1084#1103' '#1092#1072#1081#1083#1072
    end
    object Edit1: TEdit
      Left = 139
      Top = 24
      Width = 215
      Height = 24
      ReadOnly = True
      TabOrder = 0
    end
    object Button2: TButton
      Left = 360
      Top = 22
      Width = 41
      Height = 25
      Caption = '>>'
      TabOrder = 1
      OnClick = Button2Click
    end
    object Button3: TButton
      Left = 139
      Top = 54
      Width = 166
      Height = 25
      Caption = #1048#1084#1087#1086#1088#1090
      TabOrder = 2
      OnClick = Button3Click
    end
  end
  object GroupBox4: TGroupBox
    Left = 8
    Top = 223
    Width = 425
    Height = 146
    Caption = #1057#1074#1086#1076#1085#1099#1081' '#1086#1090#1095#1077#1090
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 3
    object Label4: TLabel
      Left = 18
      Top = 72
      Width = 73
      Height = 16
      Caption = #1048#1084#1103' '#1092#1072#1081#1083#1072
    end
    object Label7: TLabel
      Left = 10
      Top = 24
      Width = 50
      Height = 16
      Caption = #1055#1077#1088#1080#1086#1076
    end
    object Label8: TLabel
      Left = 18
      Top = 46
      Width = 29
      Height = 16
      Caption = #1051#1055#1059
    end
    object Edit2: TEdit
      Left = 139
      Top = 78
      Width = 215
      Height = 24
      ReadOnly = True
      TabOrder = 0
    end
    object Button4: TButton
      Left = 360
      Top = 76
      Width = 41
      Height = 25
      Caption = '>>'
      TabOrder = 1
      OnClick = Button4Click
    end
    object Button5: TButton
      Left = 139
      Top = 108
      Width = 166
      Height = 25
      Caption = #1069#1082#1089#1087#1086#1088#1090
      TabOrder = 2
      OnClick = Button5Click
    end
    object DateTimePicker1: TDateTimePicker
      Left = 139
      Top = 18
      Width = 102
      Height = 24
      Date = 40848.891848414350000000
      Time = 40848.891848414350000000
      TabOrder = 3
    end
    object DateTimePicker2: TDateTimePicker
      Left = 247
      Top = 18
      Width = 107
      Height = 24
      Date = 40848.892638159720000000
      Time = 40848.892638159720000000
      TabOrder = 4
    end
    object ComboBox5: TComboBox
      Left = 139
      Top = 48
      Width = 262
      Height = 24
      ItemHeight = 16
      TabOrder = 5
    end
  end
  object GroupBox5: TGroupBox
    Left = 8
    Top = 375
    Width = 425
    Height = 113
    Caption = #1057#1088#1072#1074#1085#1077#1085#1080#1077
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 4
    object Label5: TLabel
      Left = 18
      Top = 29
      Width = 41
      Height = 16
      Caption = #1051#1055#1059' 2'
    end
    object Label6: TLabel
      Left = 18
      Top = 51
      Width = 51
      Height = 16
      Caption = #1054#1090#1095#1077#1090' 2'
    end
    object ComboBox3: TComboBox
      Left = 139
      Top = 21
      Width = 262
      Height = 24
      ItemHeight = 16
      TabOrder = 0
    end
    object ComboBox4: TComboBox
      Left = 139
      Top = 51
      Width = 262
      Height = 24
      ItemHeight = 16
      TabOrder = 1
    end
    object Button6: TButton
      Left = 139
      Top = 81
      Width = 166
      Height = 25
      Caption = #1057#1088#1072#1074#1085#1080#1090#1100
      TabOrder = 2
      OnClick = Button6Click
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = '*.doc|*.doc|*.*|*.*'
    Left = 64
    Top = 512
  end
  object IBDatabase1: TIBDatabase
    LoginPrompt = False
    Left = 104
    Top = 512
  end
  object IBTable1: TIBTable
    Database = IBDatabase1
    Transaction = IBTransaction1
    Left = 152
    Top = 512
  end
  object IBTransaction1: TIBTransaction
    DefaultDatabase = IBDatabase1
    Left = 24
    Top = 552
  end
  object OpenDialog2: TOpenDialog
    Filter = '*.doc|*.doc'
    Left = 24
    Top = 512
  end
  object IBQuery1: TIBQuery
    Database = IBDatabase1
    Transaction = IBTransaction1
    Left = 200
    Top = 512
  end
  object IBQuery2: TIBQuery
    Database = IBDatabase1
    Transaction = IBTransaction1
    Left = 248
    Top = 512
  end
  object IBQuery3: TIBQuery
    Database = IBDatabase1
    Transaction = IBTransaction1
    Left = 288
    Top = 512
  end
  object IBTable2: TIBTable
    Database = IBDatabase1
    Transaction = IBTransaction1
    Left = 104
    Top = 552
  end
end
