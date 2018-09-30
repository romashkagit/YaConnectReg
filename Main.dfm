object YaConnectReg: TYaConnectReg
  Left = 0
  Top = 0
  Caption = 'YaConnectReg'
  ClientHeight = 556
  ClientWidth = 486
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCanResize = FormCanResize
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object pcMain: TPageControl
    Left = 0
    Top = 0
    Width = 486
    Height = 556
    ActivePage = tshComand
    Align = alClient
    TabOrder = 0
    object tshParam: TTabSheet
      Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099
      object Panel1: TPanel
        Left = 0
        Top = 0
        Width = 478
        Height = 528
        Align = alClient
        TabOrder = 0
        object pnlTop: TPanel
          Left = 1
          Top = 1
          Width = 476
          Height = 73
          Align = alTop
          TabOrder = 0
          object lblUrl: TLabel
            Left = 21
            Top = 6
            Width = 181
            Height = 13
            Caption = 'URL '#1076#1083#1103' '#1087#1086#1083#1091#1095#1077#1085#1080#1103' '#1090#1086#1082#1077#1085#1072' '#1076#1086#1084#1077#1085#1072':'
          end
          object Edit1: TEdit
            Left = 15
            Top = 25
            Width = 357
            Height = 21
            BorderStyle = bsNone
            Color = clCream
            TabOrder = 0
            Text = 'https://oauth.yandex.ru/authorize?response_type=token&client_id='
          end
        end
        object pnl2: TPanel
          Left = 1
          Top = 146
          Width = 476
          Height = 72
          Align = alTop
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          TabOrder = 1
          DesignSize = (
            476
            72)
          object lblDomen2: TLabel
            Left = 17
            Top = 15
            Width = 41
            Height = 13
            Caption = #1044#1086#1084#1077#1085' 2'
          end
          object lblToken2: TLabel
            Left = 19
            Top = 46
            Width = 39
            Height = 13
            Caption = #1058#1086#1082#1077#1085' 2'
          end
          object edDom2: TEdit
            Left = 64
            Top = 14
            Width = 390
            Height = 21
            Anchors = [akLeft, akTop, akRight, akBottom]
            TabOrder = 0
            OnChange = edDomChange
          end
          object edToken2: TEdit
            Left = 64
            Top = 41
            Width = 390
            Height = 21
            Anchors = [akLeft, akTop, akRight, akBottom]
            TabOrder = 1
            OnChange = edTokenChange
          end
        end
        object pnl3: TPanel
          Left = 1
          Top = 218
          Width = 476
          Height = 72
          Align = alTop
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          DesignSize = (
            476
            72)
          object lblDom3: TLabel
            Left = 16
            Top = 15
            Width = 41
            Height = 13
            Caption = #1044#1086#1084#1077#1085' 3'
          end
          object lblToken3: TLabel
            Left = 18
            Top = 46
            Width = 39
            Height = 13
            Caption = #1058#1086#1082#1077#1085' 3'
          end
          object edDom3: TEdit
            Left = 63
            Top = 14
            Width = 391
            Height = 21
            Anchors = [akLeft, akTop, akRight, akBottom]
            TabOrder = 0
            OnChange = edDomChange
          end
          object edToken3: TEdit
            Left = 63
            Top = 41
            Width = 391
            Height = 21
            Anchors = [akLeft, akTop, akRight, akBottom]
            TabOrder = 1
            OnChange = edTokenChange
          end
        end
        object pnl4: TPanel
          Left = 1
          Top = 290
          Width = 476
          Height = 72
          Align = alTop
          Alignment = taLeftJustify
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          TabOrder = 3
          DesignSize = (
            476
            72)
          object lblDom4: TLabel
            Left = 17
            Top = 15
            Width = 41
            Height = 13
            Caption = #1044#1086#1084#1077#1085' 4'
          end
          object lblToken4: TLabel
            Left = 19
            Top = 46
            Width = 39
            Height = 13
            Caption = #1058#1086#1082#1077#1085' 4'
          end
          object edToken4: TEdit
            Left = 64
            Top = 40
            Width = 390
            Height = 21
            Anchors = [akLeft, akTop, akRight, akBottom]
            TabOrder = 0
            OnChange = edTokenChange
          end
          object edDom4: TEdit
            Left = 64
            Top = 13
            Width = 390
            Height = 21
            Margins.Top = 10
            Anchors = [akLeft, akTop, akRight, akBottom]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 1
            OnChange = edDomChange
          end
        end
        object pnl5: TPanel
          Left = 1
          Top = 362
          Width = 476
          Height = 165
          Align = alClient
          TabOrder = 4
          object lb1: TLabel
            Left = 15
            Top = 6
            Width = 219
            Height = 13
            Caption = #1050#1086#1085#1089#1090#1088#1091#1082#1090#1086#1088' '#1087#1072#1088#1086#1083#1103' '#1076#1083#1103' '#1087#1086#1095#1090#1086#1074#1086#1075#1086' '#1103#1097#1080#1082#1072
          end
          object edPass: TEdit
            Left = 19
            Top = 39
            Width = 435
            Height = 21
            Hint = 
              #1042#1074#1077#1076#1080#1090#1077' '#1084#1072#1089#1082#1091' '#1076#1083#1103' '#1089#1086#1079#1076#1072#1085#1080#1103' '#1087#1072#1088#1086#1083#1103' '#1076#1083#1103' '#1087#1086#1095#1090#1086#1074#1099#1093' '#1103#1097#1080#1082#1086#1074'.'#13#10' '#1052#1086#1078#1085#1086' '#1080 +
              #1089#1087#1086#1083#1100#1079#1086#1074#1072#1090#1100' '#1083#1072#1090#1080#1085#1089#1082#1080#1077' '#1073#1091#1082#1074#1099', '#1094#1080#1092#1088#1099' '#1080' '#1089#1080#1084#1074#1086#1083#1099' '#1080#1079' '#1089#1087#1080#1089#1082#1072':'#13#10'` ! @ #' +
              ' $ % ^ & * ( ) _ = + - [ ] { } ; : "  , . < > \ / ?'#13#10#1052#1086#1078#1085#1086' '#1080#1089#1087#1086#1083 +
              #1100#1079#1086#1074#1072#1090#1100' '#1083#1086#1075#1080#1085', '#1076#1083#1103' '#1101#1090#1086#1075#1086' '#1074#1085#1077#1089#1080#1090#1077' '#1087#1077#1088#1077#1084#1077#1085#1085#1091#1102' $LOGIN.'#13#10#1053#1072#1087#1088#1080#1084#1077#1088': $' +
              'LOGIN_Cc'
            ParentCustomHint = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clGrayText
            Font.Height = -11
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            TabOrder = 0
            OnChange = edPassChange
            OnClick = edPassClick
          end
        end
        object pnl1: TPanel
          Left = 1
          Top = 74
          Width = 476
          Height = 72
          Align = alTop
          Alignment = taLeftJustify
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          TabOrder = 5
          DesignSize = (
            476
            72)
          object lblDomen1: TLabel
            Left = 17
            Top = 15
            Width = 41
            Height = 13
            Caption = #1044#1086#1084#1077#1085' 1'
          end
          object lblToken1: TLabel
            Left = 19
            Top = 46
            Width = 39
            Height = 13
            Caption = #1058#1086#1082#1077#1085' 1'
          end
          object edToken1: TEdit
            Left = 64
            Top = 40
            Width = 390
            Height = 21
            Anchors = [akLeft, akTop, akRight, akBottom]
            TabOrder = 0
            OnChange = edTokenChange
          end
          object edDom1: TEdit
            Left = 64
            Top = 13
            Width = 390
            Height = 21
            Margins.Top = 10
            Anchors = [akLeft, akTop, akRight, akBottom]
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 1
            OnChange = edDomChange
          end
        end
      end
    end
    object tshComand: TTabSheet
      Caption = #1050#1086#1084#1072#1085#1076#1099
      ImageIndex = 1
      object V: TPanel
        Left = 0
        Top = 0
        Width = 478
        Height = 196
        Align = alTop
        AutoSize = True
        TabOrder = 0
        object lblListDom: TLabel
          Left = 21
          Top = 9
          Width = 82
          Height = 13
          Caption = #1057#1087#1080#1089#1086#1082' '#1076#1086#1084#1077#1085#1086#1074
        end
        object lblListLogin: TLabel
          Left = 20
          Top = 37
          Width = 80
          Height = 13
          Caption = #1057#1087#1080#1089#1086#1082' '#1083#1086#1075#1080#1085#1086#1074
        end
        object btnRefresh: TSpeedButton
          Left = 363
          Top = 2
          Width = 94
          Height = 21
          Caption = 'Refresh'
          Glyph.Data = {
            36030000424D3603000000000000360000002800000010000000100000000100
            1800000000000003000000000000000000000000000000000000FF00FFFF00FF
            A37875A37875A37875A37875A37875A37875A37875A37875A37875A37875A378
            75A3787590615EFF00FFFF00FFFF00FFA67C76F2E2D3F2E2D3FFE8D1EFDFBBFF
            E3C5FFDEBDFFDDBAFFD8B2FFD6AEFFD2A5FFD2A3936460FF00FFFF00FFFF00FF
            AB8078F3E7DAF3E7DA019901AFD8A071C57041AA3081BB5EEFD4A6FFD6AEFFD2
            A3FFD2A3966763FF00FFFF00FFFF00FFB0837AF4E9DDF4E9DD01990101990101
            990101990101990141AA2FFFD8B2FFD4A9FFD4A99A6A65FF00FFFF00FFFF00FF
            B6897DF5EDE4F5EDE4019901019901119E0ECFD6A3FFE4C821A21AFFD8B2FFD7
            B0FFD7B09E6D67FF00FFFF00FFFF00FFBC8E7FF7EFE8F7EFE801990101990101
            9901019901FFE4C8EFDEBAFFD8B2FFD7B0FFD9B4A27069FF00FFFF00FFFF00FF
            C39581F8F3EFF8F3EFF8F3EFFFF4E8FFF4E8FFF4E8EFE3C4EFE3C4FFE4C8FFDE
            BDFFDDBBA5746BFF00FFFF00FFFF00FFCA9B84F9F5F2FBFBFBFFF4E8FFF4E8FF
            F4E8019901019901019901FFE8D1FFE3C5FFE1C2A8776DFF00FFFF00FFFF00FF
            D2A187F9F9F9FBFBFB119F10AFD8A0FFF4E8AFD8A0019901019901FFE8D1FFE4
            C8FFE3C6AC7A6FFF00FFFF00FFFF00FFD9A88AFBFBFBFFFFFF71C57001990101
            9901019901019901019901FFE8D1FFE8D1FFE6CEAE7C72FF00FFFF00FFFF00FF
            DFAE8CFCFCFCFFFFFFFFFFFF71C570019901019901AFD8A0019901FFE8D1FFC8
            C2FFB0B0B07E73FF00FFFF00FFFF00FFE5B38FFDFDFDFDFDFDFFFFFFFFFFFFFF
            FFFEFFFAF6FFF9F3FFF5EAF4DECEB28074B28074B28074FF00FFFF00FFFF00FF
            EAB891FEFEFEFEFEFEFFFFFFFFFFFFFFFFFFFFFFFEFFFAF6FFF9F3F5E1D2B280
            74EDA755CBA390FF00FFFF00FFFF00FFEFBC92FFFFFFFFFFFFFCFCFCFAFAFAF7
            F7F7F5F5F5F2F1F1F0EDEAE9DAD0B28074D2AA93C9CED0FF00FFFF00FFFF00FF
            F2BF94DCA987DCA987DCA987DCA987DCA987DCA987DCA987DCA987DCA987B280
            74CACED0FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
            00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
          Layout = blGlyphRight
          OnClick = btnRefreshClick
        end
        object cbLog: TComboBox
          Left = 120
          Top = 29
          Width = 225
          Height = 21
          AutoDropDown = True
          Style = csDropDownList
          DropDownCount = 10
          Sorted = True
          TabOrder = 0
          Items.Strings = (
            '')
        end
        object btnEnter2box: TButton
          Left = 363
          Top = 29
          Width = 94
          Height = 21
          Caption = 'Enter'
          TabOrder = 1
          OnClick = btnEnter2boxClick
        end
        object btnStart1: TButton
          Left = 363
          Top = 78
          Width = 94
          Height = 21
          Caption = 'Start'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnClick = btnStart1Click
        end
        object btnStop1: TButton
          Left = 363
          Top = 105
          Width = 94
          Height = 21
          Caption = 'Stop'
          TabOrder = 3
          OnClick = btnStop1Click
        end
        object rgAct: TRadioGroup
          Left = 21
          Top = 73
          Width = 324
          Height = 53
          Columns = 3
          ItemIndex = 2
          Items.Strings = (
            'Save'
            'Restore'
            'Sync')
          TabOrder = 4
          OnClick = rgActClick
        end
        object cbAll: TCheckBox
          Left = 19
          Top = 132
          Width = 249
          Height = 17
          Caption = #1055#1088#1086#1074#1077#1088#1080#1090#1100' '#1076#1086#1088#1077#1075#1080#1089#1090#1088#1072#1094#1080#1102' '#1074#1089#1077#1093' '#1103#1097#1080#1082#1086#1074
          TabOrder = 5
        end
        object cbDel: TCheckBox
          Left = 19
          Top = 155
          Width = 346
          Height = 17
          Caption = #1059#1076#1072#1083#1080#1090#1100' '#1103#1097#1080#1082#1080', '#1076#1083#1103' '#1082#1086#1090#1086#1088#1099#1093' '#1085#1077#1090' '#1091#1095#1077#1090#1085#1099#1093' '#1079#1072#1087#1080#1089#1077#1081' '#1085#1072' '#1089#1077#1088#1074#1077#1088#1077
          TabOrder = 6
          OnClick = cbDelClick
        end
        object cbPass: TCheckBox
          Left = 20
          Top = 178
          Width = 346
          Height = 17
          Caption = #1047#1072#1084#1077#1085#1080#1090#1100' '#1087#1072#1088#1086#1083#1080' '#1087#1086' '#1096#1072#1073#1083#1086#1085#1091
          TabOrder = 7
          OnClick = cbPassClick
        end
        object cbDomen: TComboBox
          Left = 120
          Top = 1
          Width = 225
          Height = 21
          AutoDropDown = True
          DropDownCount = 4
          TabOrder = 8
          OnChange = cbDomenChange
          OnDropDown = cbDomenDropDown
        end
      end
      object meOut: TRichEdit
        Left = 0
        Top = 196
        Width = 478
        Height = 275
        Align = alClient
        Color = clBlack
        Font.Charset = RUSSIAN_CHARSET
        Font.Color = clYellow
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        ScrollBars = ssVertical
        TabOrder = 1
        Zoom = 100
      end
      object pnlProg: TPanel
        Left = 0
        Top = 471
        Width = 478
        Height = 57
        Align = alBottom
        TabOrder = 2
        object lblProg: TLabel
          Left = 32
          Top = 35
          Width = 155
          Height = 13
          Caption = #1054#1073#1085#1086#1074#1083#1077#1085#1080#1077' '#1089#1087#1080#1089#1082#1072' '#1083#1086#1075#1080#1085#1086#1074'...'
        end
        object pbRefresh: TProgressBar
          Left = 32
          Top = 13
          Width = 369
          Height = 16
          TabOrder = 0
        end
      end
    end
    object tsExcel: TTabSheet
      Caption = #1047#1072#1075#1088#1091#1079#1082#1072' '#1080#1079' Excel'
      ImageIndex = 2
      object pnlload: TPanel
        Left = 0
        Top = 0
        Width = 478
        Height = 105
        Align = alTop
        TabOrder = 0
        object btnloadfrxl: TButton
          Left = 40
          Top = 16
          Width = 163
          Height = 25
          Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1080#1079' Excel'
          TabOrder = 0
        end
        object btnLoadYA: TButton
          Left = 40
          Top = 47
          Width = 163
          Height = 25
          Caption = #1048#1079#1084#1077#1085#1080#1090#1100' '#1074' '#1087#1086#1095#1090#1086#1074#1086#1084' '#1089#1077#1088#1074#1077#1088#1077
          Enabled = False
          TabOrder = 1
          OnClick = btnLoadYAClick
        end
      end
      object PageControl1: TPageControl
        Left = 0
        Top = 105
        Width = 478
        Height = 423
        ActivePage = TabSheet1
        Align = alClient
        TabOrder = 1
        object TabSheet1: TTabSheet
          Caption = #1044#1086#1084#1077#1085#1099
          object sgDomens: TStringGrid
            Left = 0
            Top = 0
            Width = 470
            Height = 395
            Align = alClient
            FixedCols = 0
            Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSizing, goColSizing]
            TabOrder = 0
            ColWidths = (
              64
              64
              64
              64
              64)
            RowHeights = (
              24
              24
              24
              24
              24)
          end
        end
        object TabSheet2: TTabSheet
          Caption = #1056#1072#1089#1089#1099#1083#1082#1080
          ImageIndex = 1
          object sgMailList: TStringGrid
            Left = 0
            Top = 0
            Width = 470
            Height = 395
            Align = alClient
            FixedCols = 0
            Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSizing, goColSizing]
            TabOrder = 0
            ColWidths = (
              64
              64
              64
              64
              64)
            RowHeights = (
              24
              24
              24
              24
              24)
          end
        end
        object TabSheet3: TTabSheet
          Caption = #1055#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1080
          ImageIndex = 2
          object sgUsers: TStringGrid
            Left = 0
            Top = 0
            Width = 470
            Height = 395
            Align = alClient
            FixedCols = 0
            Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goRowSizing, goColSizing]
            TabOrder = 0
            ColWidths = (
              64
              64
              64
              64
              64)
            RowHeights = (
              24
              24
              24
              24
              24)
          end
        end
      end
    end
  end
end
