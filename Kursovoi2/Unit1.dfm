object Form1: TForm1
  Left = 202
  Top = 156
  Width = 870
  Height = 387
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object grp1: TGroupBox
    Left = 8
    Top = 8
    Width = 217
    Height = 337
    Caption = #1048#1089#1093#1086#1076#1085#1099#1077' '#1076#1072#1085#1085#1099#1077
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    object lbl5: TLabel
      Left = 16
      Top = 216
      Width = 7
      Height = 16
      Caption = 'Z'
    end
    object btn1: TBitBtn
      Left = 8
      Top = 296
      Width = 193
      Height = 25
      Caption = #1056#1072#1089#1095#1077#1090' '#1080' '#1087#1086#1089#1090#1088#1086#1077#1085#1080#1077
      TabOrder = 0
      OnClick = btn1Click
    end
    object edt5: TEdit
      Left = 40
      Top = 216
      Width = 121
      Height = 24
      TabOrder = 1
    end
    object grp2: TGroupBox
      Left = 8
      Top = 24
      Width = 193
      Height = 89
      Caption = 'x+2*y<=12'
      TabOrder = 2
      object lbl1: TLabel
        Left = 8
        Top = 24
        Width = 15
        Height = 16
        Caption = 'x1'
      end
      object lbl2: TLabel
        Left = 8
        Top = 48
        Width = 15
        Height = 16
        Caption = 'x2'
      end
      object edt1: TEdit
        Left = 32
        Top = 24
        Width = 121
        Height = 24
        TabOrder = 0
      end
      object edt2: TEdit
        Left = 32
        Top = 48
        Width = 121
        Height = 24
        TabOrder = 1
      end
    end
    object grp3: TGroupBox
      Left = 8
      Top = 120
      Width = 193
      Height = 89
      Caption = 'x+y<=9'
      TabOrder = 3
      object lbl4: TLabel
        Left = 8
        Top = 25
        Width = 15
        Height = 16
        Caption = 'x3'
      end
      object lbl3: TLabel
        Left = 8
        Top = 49
        Width = 15
        Height = 16
        Caption = 'x4'
      end
      object edt3: TEdit
        Left = 32
        Top = 25
        Width = 121
        Height = 24
        TabOrder = 0
      end
      object edt4: TEdit
        Left = 32
        Top = 49
        Width = 121
        Height = 24
        TabOrder = 1
      end
    end
  end
  object grp4: TGroupBox
    Left = 232
    Top = 16
    Width = 617
    Height = 329
    Caption = #1056#1077#1079#1091#1083#1100#1090#1072#1090#1099
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    object grp5: TGroupBox
      Left = 8
      Top = 16
      Width = 153
      Height = 153
      TabOrder = 0
      object strngrd1: TStringGrid
        Left = 8
        Top = 16
        Width = 137
        Height = 129
        ColCount = 2
        FixedCols = 0
        TabOrder = 0
      end
    end
    object grp6: TGroupBox
      Left = 8
      Top = 168
      Width = 153
      Height = 145
      TabOrder = 1
      object lbl6: TLabel
        Left = 8
        Top = 16
        Width = 8
        Height = 16
        Caption = #1072
      end
      object lbl7: TLabel
        Left = 8
        Top = 40
        Width = 8
        Height = 16
        Caption = 'b'
      end
      object lbl8: TLabel
        Left = 8
        Top = 64
        Width = 33
        Height = 16
        Caption = 'Zmax'
      end
      object lbl9: TLabel
        Left = 8
        Top = 88
        Width = 29
        Height = 16
        Caption = 'Zmin'
      end
      object lbl10: TLabel
        Left = 8
        Top = 112
        Width = 9
        Height = 16
        Caption = 'D'
      end
      object edt6: TEdit
        Left = 48
        Top = 16
        Width = 97
        Height = 24
        ReadOnly = True
        TabOrder = 0
      end
      object edt7: TEdit
        Left = 48
        Top = 40
        Width = 97
        Height = 24
        ReadOnly = True
        TabOrder = 1
      end
      object edt8: TEdit
        Left = 48
        Top = 64
        Width = 97
        Height = 24
        ReadOnly = True
        TabOrder = 2
      end
      object edt9: TEdit
        Left = 48
        Top = 88
        Width = 97
        Height = 24
        ReadOnly = True
        TabOrder = 3
      end
      object edt10: TEdit
        Left = 48
        Top = 112
        Width = 97
        Height = 24
        ReadOnly = True
        TabOrder = 4
      end
    end
    object grp7: TGroupBox
      Left = 168
      Top = 16
      Width = 433
      Height = 297
      TabOrder = 2
    end
  end
end
