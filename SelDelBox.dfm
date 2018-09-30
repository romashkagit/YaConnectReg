object Form2: TForm2
  Left = 0
  Top = 0
  ClientHeight = 176
  ClientWidth = 287
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object clbDel: TCheckListBox
    Left = 0
    Top = 41
    Width = 287
    Height = 94
    Align = alClient
    AutoComplete = False
    ItemHeight = 13
    PopupMenu = pmSel
    Sorted = True
    TabOrder = 0
  end
  object pnlSelTop: TPanel
    Left = 0
    Top = 0
    Width = 287
    Height = 41
    Align = alTop
    TabOrder = 1
    object lblSel: TLabel
      Left = 8
      Top = 0
      Width = 272
      Height = 26
      Caption = 
        #1057#1087#1080#1089#1086#1082' '#1083#1086#1075#1080#1085#1086#1074' '#1087#1086#1095#1090#1086#1074#1099#1093' '#1103#1097#1080#1082#1086#1074','#1076#1083#1103'  '#1082#1086#1090#1086#1088#1099#1093' '#1085#1077#1090' '#1089#1086#1086#1090#1074#1077#1090#1074#1091#1102#1097#1080#1093' '#1091#1095 +
        #1077#1090#1085#1099#1093' '#1079#1072#1087#1080#1089#1077#1081' '#1085#1072' '#1089#1077#1088#1074#1077#1088#1077
      WordWrap = True
    end
  end
  object pnlSelBot: TPanel
    Left = 0
    Top = 135
    Width = 287
    Height = 41
    Align = alBottom
    TabOrder = 2
    object btnOK: TButton
      Left = 80
      Top = 6
      Width = 114
      Height = 25
      Caption = #1059#1076#1072#1083#1080#1090#1100' '#1103#1097#1080#1082#1080
      Default = True
      ModalResult = 1
      TabOrder = 0
      OnClick = btnOKClick
    end
    object btnCancel: TButton
      Left = 200
      Top = 6
      Width = 75
      Height = 25
      Cancel = True
      Caption = #1054#1090#1084#1077#1085#1072
      ModalResult = 2
      TabOrder = 1
    end
  end
  object pmSel: TPopupMenu
    Left = 200
    Top = 56
    object miSelect: TMenuItem
      Caption = #1042#1099#1076#1077#1083#1080#1090#1100' '#1074#1089#1077
      OnClick = miSelectClick
    end
    object miUnselect: TMenuItem
      Caption = #1057#1085#1103#1090#1100' '#1074#1089#1077' '#1074#1099#1076#1077#1083#1077#1085#1080#1103
      OnClick = miUnselectClick
    end
  end
end
