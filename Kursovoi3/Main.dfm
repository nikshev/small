object Form1: TForm1
  Left = 306
  Top = 187
  Width = 908
  Height = 596
  Caption = #1042#1099#1095#1080#1089#1083#1077#1085#1080#1077' '#1089#1082#1086#1083#1100#1079#1103#1097#1080#1093' '#1087#1086#1082#1072#1079#1072#1090#1077#1083#1077#1081
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 13
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 900
    Height = 550
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 0
    object TabSheet1: TTabSheet
      Caption = #1047#1072#1076#1072#1085#1080#1077' '#1091#1089#1083#1086#1074#1080#1081
      DesignSize = (
        892
        522)
      object Chart1: TChart
        Left = 280
        Top = 10
        Width = 642
        Height = 502
        Hint = #1065#1077#1083#1082#1085#1080' '#1076#1083#1103' '#1087#1088#1086#1088#1080#1089#1086#1074#1082#1080'!'
        AnimatedZoom = True
        BackWall.Brush.Color = clWhite
        BackWall.Brush.Style = bsClear
        MarginBottom = 5
        MarginLeft = 5
        MarginRight = 10
        MarginTop = 0
        Title.Font.Charset = DEFAULT_CHARSET
        Title.Font.Color = clBlack
        Title.Font.Height = -16
        Title.Font.Name = 'Arial'
        Title.Font.Style = [fsBold]
        Title.Text.Strings = (
          #1044#1080#1085#1072#1084#1080#1082#1072' '#1079#1072#1076#1072#1085#1085#1086#1075#1086' '#1087#1086#1082#1072#1079#1072#1090#1077#1083#1103)
        BottomAxis.Axis.Mode = pmBlack
        BottomAxis.AxisValuesFormat = '#,##0'
        BottomAxis.LabelsFont.Charset = DEFAULT_CHARSET
        BottomAxis.LabelsFont.Color = clBlack
        BottomAxis.LabelsFont.Height = -8
        BottomAxis.LabelsFont.Name = 'Arial'
        BottomAxis.LabelsFont.Style = []
        BottomAxis.Title.Caption = #1043#1086#1076
        BottomAxis.Title.Font.Charset = DEFAULT_CHARSET
        BottomAxis.Title.Font.Color = clBlack
        BottomAxis.Title.Font.Height = -13
        BottomAxis.Title.Font.Name = 'Arial'
        BottomAxis.Title.Font.Style = [fsBold]
        BottomAxis.TitleSize = 1
        DepthAxis.Title.Font.Charset = DEFAULT_CHARSET
        DepthAxis.Title.Font.Color = clBlack
        DepthAxis.Title.Font.Height = -13
        DepthAxis.Title.Font.Name = 'Arial'
        DepthAxis.Title.Font.Style = [fsBold]
        LeftAxis.Title.Caption = #1055#1086#1082#1072#1079#1072#1090#1077#1083#1100
        LeftAxis.Title.Font.Charset = DEFAULT_CHARSET
        LeftAxis.Title.Font.Color = clBlack
        LeftAxis.Title.Font.Height = -13
        LeftAxis.Title.Font.Name = 'Arial'
        LeftAxis.Title.Font.Style = [fsBold]
        LeftAxis.TitleSize = 1
        Legend.Alignment = laTop
        Legend.Font.Charset = DEFAULT_CHARSET
        Legend.Font.Color = clBlack
        Legend.Font.Height = -13
        Legend.Font.Name = 'Arial'
        Legend.Font.Style = [fsBold]
        Legend.LegendStyle = lsSeries
        Legend.ShadowSize = 0
        Legend.VertMargin = 3
        TopAxis.Axis.Style = psClear
        View3D = False
        BevelInner = bvLowered
        TabOrder = 0
        Anchors = [akLeft, akTop, akRight, akBottom]
        OnClick = Chart1Click
        object Series2: TLineSeries
          Marks.ArrowLength = 8
          Marks.Visible = False
          SeriesColor = clGreen
          Title = #1060#1072#1082#1090#1080#1095#1077#1089#1082#1080
          Pointer.InflateMargins = True
          Pointer.Style = psRectangle
          Pointer.Visible = False
          XValues.DateTime = False
          XValues.Name = 'X'
          XValues.Multiplier = 1.000000000000000000
          XValues.Order = loAscending
          YValues.DateTime = False
          YValues.Name = 'Y'
          YValues.Multiplier = 1.000000000000000000
          YValues.Order = loNone
        end
        object Series1: TLineSeries
          Marks.ArrowLength = 8
          Marks.Style = smsXValue
          Marks.Visible = False
          SeriesColor = clRed
          Title = #1058#1088#1077#1085#1076
          Pointer.InflateMargins = True
          Pointer.Style = psRectangle
          Pointer.Visible = False
          XValues.DateTime = False
          XValues.Name = 'X'
          XValues.Multiplier = 1.000000000000000000
          XValues.Order = loAscending
          YValues.DateTime = False
          YValues.Name = 'Y'
          YValues.Multiplier = 1.000000000000000000
          YValues.Order = loNone
        end
      end
      object UpDown1: TUpDown
        Left = 260
        Top = 20
        Width = 12
        Height = 25
        Associate = LabeledEdit1
        Min = 1
        Max = 1000
        Position = 1
        TabOrder = 1
        OnClick = UpDown1Click
      end
      object LabeledEdit1: TLabeledEdit
        Left = 10
        Top = 20
        Width = 250
        Height = 25
        AutoSize = False
        EditLabel.Width = 121
        EditLabel.Height = 13
        EditLabel.Caption = #1050#1086#1083#1080#1095#1077#1089#1090#1074#1086' '#1080#1079#1084#1077#1088#1077#1085#1080#1081':'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
        Text = '1'
        OnChange = LabeledEdit1Change
      end
      object StringGrid1: TStringGrid
        Left = 10
        Top = 50
        Width = 262
        Height = 462
        Hint = #1042#1074#1086#1076#1080#1090#1077' '#1076#1072#1085#1085#1099#1077
        Anchors = [akLeft, akTop, akBottom]
        ColCount = 3
        DefaultColWidth = 84
        DefaultRowHeight = 20
        FixedCols = 0
        RowCount = 2
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing, goAlwaysShowEditor]
        TabOrder = 3
      end
    end
    object TabSheet2: TTabSheet
      Caption = #1056#1072#1089#1095#1077#1090' '#1079#1085#1072#1095#1077#1085#1080#1081
      ImageIndex = 1
      OnShow = TabSheet2Show
      DesignSize = (
        892
        522)
      object LabeledEdit2: TLabeledEdit
        Left = 10
        Top = 20
        Width = 40
        Height = 25
        Hint = #1044#1086#1083#1078#1085#1072' '#1073#1099#1090#1100' '#1085#1077#1095#1077#1090#1085#1086#1081'!'
        AutoSize = False
        EditLabel.Width = 114
        EditLabel.Height = 13
        EditLabel.Caption = #1042#1077#1083#1080#1095#1080#1085#1072' '#1087#1086#1076#1087#1077#1088#1080#1086#1076#1072':'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        MaxLength = 3
        ParentFont = False
        TabOrder = 0
        Text = '1'
      end
      object BitBtn1: TBitBtn
        Left = 60
        Top = 20
        Width = 75
        Height = 25
        TabOrder = 1
        OnClick = BitBtn1Click
        Kind = bkOK
      end
      object StringGrid2: TStringGrid
        Left = 10
        Top = 50
        Width = 872
        Height = 462
        Anchors = [akLeft, akTop, akRight, akBottom]
        Color = clCaptionText
        ColCount = 9
        DefaultColWidth = 95
        DefaultRowHeight = 20
        FixedCols = 0
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
      end
    end
    object TabSheet3: TTabSheet
      Caption = #1056#1077#1079#1091#1083#1100#1090#1072#1090#1099
      ImageIndex = 2
      DesignSize = (
        892
        522)
      object Memo1: TMemo
        Left = 10
        Top = 10
        Width = 872
        Height = 502
        Anchors = [akLeft, akTop, akRight, akBottom]
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Pitch = fpFixed
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
    end
  end
  object MainMenu1: TMainMenu
    Left = 936
    object N1: TMenuItem
      Caption = #1056#1072#1073#1086#1090#1072' '#1089' '#1076#1072#1085#1085#1099#1084#1080'...'
      object N2: TMenuItem
        Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1076#1072#1085#1085#1099#1077'...'
        OnClick = N2Click
      end
      object N3: TMenuItem
        Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1076#1072#1085#1085#1099#1077'...'
        OnClick = N3Click
      end
      object N4: TMenuItem
        Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1075#1088#1072#1092#1080#1082'...'
        OnClick = N4Click
      end
      object N5: TMenuItem
        Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1088#1077#1079#1091#1083#1100#1090#1072#1090#1099'...'
        OnClick = N5Click
      end
    end
    object N6: TMenuItem
      Caption = #1057#1087#1088#1072#1074#1082#1072
      OnClick = N6Click
    end
  end
  object OpenDialog1: TOpenDialog
    DefaultExt = '*.txt'
    FileName = 'MyFile'
    Filter = #1058#1077#1082#1089#1090#1086#1074#1099#1081' '#1092#1072#1081#1083'|*.txt'
    Title = #1042#1099#1073#1077#1088#1080#1090#1077' '#1092#1072#1081#1083' '#1076#1083#1103' '#1079#1072#1075#1088#1091#1079#1082#1080'...'
    Left = 820
    Top = 42
  end
  object SaveDialog1: TSaveDialog
    DefaultExt = '*.txt'
    FileName = 'MyFile'
    Filter = #1058#1077#1082#1089#1090#1086#1074#1099#1081' '#1092#1072#1081#1083'|*.txt'
    Title = #1042#1099#1073#1077#1088#1080#1090#1077' '#1092#1072#1081#1083' '#1076#1083#1103' '#1089#1086#1093#1088#1072#1085#1077#1085#1080#1103'.'
    Left = 784
    Top = 40
  end
  object SavePictureDialog1: TSavePictureDialog
    DefaultExt = '*.bmp'
    FileName = 'MyChart'
    Filter = 'Bitmaps (*.bmp)|*.bmp'
    Title = #1042#1099#1073#1077#1088#1080#1090#1077' '#1092#1072#1081#1083' '#1076#1083#1103' '#1089#1086#1093#1088#1072#1085#1077#1085#1080#1103' '#1075#1088#1072#1092#1080#1082#1072'...'
    Left = 752
    Top = 40
  end
end
