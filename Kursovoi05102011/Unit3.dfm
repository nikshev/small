object Form3: TForm3
  Left = 0
  Top = 0
  Caption = #1054#1089#1085#1086#1074#1085#1099#1077' '#1089#1088#1077#1076#1089#1090#1074
  ClientHeight = 448
  ClientWidth = 1001
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
  object TreeView1: TTreeView
    Left = 16
    Top = 16
    Width = 257
    Height = 385
    Indent = 19
    TabOrder = 0
    OnGetSelectedIndex = TreeView1GetSelectedIndex
  end
  object GroupBox1: TGroupBox
    Left = 591
    Top = 8
    Width = 410
    Height = 263
    Caption = #1055#1086#1080#1089#1082' '#1087#1086#1080#1089#1082' '#1087#1086' '#1080#1085#1074#1077#1085#1090#1072#1088#1085#1086#1084#1091
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    object Label1: TLabel
      Left = 16
      Top = 24
      Width = 40
      Height = 16
      Caption = #1055#1086#1080#1089#1082
    end
    object Edit1: TEdit
      Left = 72
      Top = 21
      Width = 242
      Height = 24
      TabOrder = 0
    end
    object Button4: TButton
      Left = 320
      Top = 21
      Width = 75
      Height = 25
      Caption = #1055#1086#1080#1089#1082
      TabOrder = 1
      OnClick = Button4Click
    end
    object ListBox1: TListBox
      Left = 16
      Top = 51
      Width = 377
      Height = 198
      ItemHeight = 16
      TabOrder = 2
      OnDblClick = ListBox1DblClick
    end
  end
  object GroupBox2: TGroupBox
    Left = 279
    Top = 8
    Width = 306
    Height = 393
    Caption = #1050#1072#1088#1090#1086#1095#1082#1072' '#1086#1089#1085#1086#1074#1085#1086#1075#1086' '#1089#1088#1077#1076#1089#1090#1074#1072
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    object Label2: TLabel
      Left = 25
      Top = 32
      Width = 47
      Height = 16
      Caption = #1043#1088#1091#1087#1087#1072
    end
    object Label3: TLabel
      Left = 25
      Top = 62
      Width = 70
      Height = 16
      Caption = #1048#1085#1074'.'#1085#1086#1084#1077#1088
    end
    object Label4: TLabel
      Left = 25
      Top = 87
      Width = 28
      Height = 16
      Caption = #1060#1048#1054
    end
    object Label5: TLabel
      Left = 25
      Top = 117
      Width = 38
      Height = 16
      Caption = #1048#1079#1085#1086#1089
    end
    object Label6: TLabel
      Left = 25
      Top = 147
      Width = 47
      Height = 16
      Caption = #1056#1072#1089#1093#1086#1076
    end
    object Label7: TLabel
      Left = 25
      Top = 177
      Width = 92
      Height = 16
      Caption = #1044#1072#1090#1072' '#1087#1086#1082#1091#1087#1082#1080
    end
    object Label8: TLabel
      Left = 25
      Top = 207
      Width = 34
      Height = 16
      Caption = #1062#1077#1085#1072
    end
    object Button3: TButton
      Left = 17
      Top = 353
      Width = 105
      Height = 25
      Caption = #1053#1086#1074#1072#1103' '#1079#1072#1087#1080#1089#1100
      TabOrder = 0
      OnClick = Button3Click
    end
    object Edit2: TEdit
      Left = 136
      Top = 29
      Width = 153
      Height = 24
      TabOrder = 1
    end
    object Edit3: TEdit
      Left = 136
      Top = 59
      Width = 153
      Height = 24
      TabOrder = 2
    end
    object Edit4: TEdit
      Left = 136
      Top = 89
      Width = 153
      Height = 24
      TabOrder = 3
    end
    object Edit5: TEdit
      Left = 136
      Top = 117
      Width = 153
      Height = 24
      TabOrder = 4
    end
    object Edit6: TEdit
      Left = 136
      Top = 147
      Width = 153
      Height = 24
      TabOrder = 5
    end
    object Edit7: TEdit
      Left = 136
      Top = 177
      Width = 153
      Height = 24
      TabOrder = 6
    end
    object Edit8: TEdit
      Left = 136
      Top = 207
      Width = 153
      Height = 24
      TabOrder = 7
    end
    object Button2: TButton
      Left = 216
      Top = 353
      Width = 75
      Height = 25
      Caption = #1059#1076#1072#1083#1080#1090#1100
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 8
      OnClick = Button2Click
    end
    object Button1: TButton
      Left = 128
      Top = 353
      Width = 82
      Height = 25
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 9
      OnClick = Button1Click
    end
    object CheckBox1: TCheckBox
      Left = 136
      Top = 237
      Width = 153
      Height = 17
      Caption = #1055#1077#1088#1077#1092#1077#1088#1080#1081#1085#1086#1077
      TabOrder = 10
      OnClick = CheckBox1Click
    end
    object CheckBox2: TCheckBox
      Left = 136
      Top = 260
      Width = 97
      Height = 17
      Caption = #1050#1072#1088#1090#1088#1080#1076#1078#1080
      TabOrder = 11
      OnClick = CheckBox2Click
    end
  end
  object GroupBox3: TGroupBox
    Left = 591
    Top = 277
    Width = 410
    Height = 124
    Caption = #1069#1082#1089#1087#1086#1088#1090' '#1074' Excel'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 3
    object Label9: TLabel
      Left = 16
      Top = 24
      Width = 71
      Height = 16
      Caption = #1057#1086#1090#1088#1091#1076#1085#1080#1082
    end
    object Label10: TLabel
      Left = 16
      Top = 55
      Width = 50
      Height = 16
      Caption = #1055#1077#1088#1080#1086#1076
    end
    object Label11: TLabel
      Left = 16
      Top = 88
      Width = 76
      Height = 16
      Caption = #1048#1079#1085#1086#1089'/'#1056#1072#1089#1093
    end
    object Edit9: TEdit
      Left = 96
      Top = 24
      Width = 216
      Height = 24
      TabOrder = 0
    end
    object Button5: TButton
      Left = 318
      Top = 24
      Width = 75
      Height = 25
      Caption = #1069#1082#1089#1087#1086#1088#1090
      TabOrder = 1
      OnClick = Button5Click
    end
    object DateTimePicker1: TDateTimePicker
      Left = 96
      Top = 55
      Width = 105
      Height = 24
      Date = 40828.654520879630000000
      Time = 40828.654520879630000000
      TabOrder = 2
    end
    object DateTimePicker2: TDateTimePicker
      Left = 207
      Top = 55
      Width = 105
      Height = 24
      Date = 40828.654520879630000000
      Time = 40828.654520879630000000
      TabOrder = 3
    end
    object Button6: TButton
      Left = 318
      Top = 55
      Width = 75
      Height = 25
      Caption = #1069#1082#1089#1087#1086#1088#1090
      TabOrder = 4
      OnClick = Button6Click
    end
    object Edit10: TEdit
      Left = 96
      Top = 85
      Width = 105
      Height = 24
      TabOrder = 5
    end
    object Edit11: TEdit
      Left = 207
      Top = 85
      Width = 105
      Height = 24
      TabOrder = 6
    end
    object Button7: TButton
      Left = 318
      Top = 86
      Width = 75
      Height = 25
      Caption = #1069#1082#1089#1087#1086#1088#1090
      TabOrder = 7
      OnClick = Button7Click
    end
  end
  object ADOConnection1: TADOConnection
    Left = 16
    Top = 408
  end
  object ADOTableNameSredstvaGroup: TADOTable
    Connection = ADOConnection1
    Left = 72
    Top = 408
  end
  object ADOTableSredstva: TADOTable
    Connection = ADOConnection1
    Left = 184
    Top = 408
  end
  object ADOQueryGroup: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 272
    Top = 408
  end
  object ADOQueryInvNum: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 368
    Top = 408
  end
  object ADOQueryParameters: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 456
    Top = 408
  end
  object ADOQueryInvCard: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 552
    Top = 408
  end
end
