object FDocuments: TFDocuments
  Left = 342
  Top = 222
  Width = 927
  Height = 508
  Caption = #1057#1086#1079#1076#1072#1085#1080#1077' '#1086#1090#1095#1077#1090#1086#1074
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 13
  object Pn1: TPanel
    Left = 0
    Top = 49
    Width = 433
    Height = 421
    Align = alLeft
    TabOrder = 0
    object LbAllPeople: TListBox
      Left = 16
      Top = 40
      Width = 169
      Height = 369
      ItemHeight = 13
      Sorted = True
      TabOrder = 0
    end
    object MeSelectPeople: TMemo
      Left = 256
      Top = 40
      Width = 161
      Height = 369
      ScrollBars = ssBoth
      TabOrder = 1
    end
    object BtAddSelectPeople: TButton
      Left = 192
      Top = 208
      Width = 57
      Height = 25
      Caption = '>>>'
      TabOrder = 2
      OnClick = BtAddSelectPeopleClick
    end
    object BtAddSelectPeopleAll: TButton
      Left = 192
      Top = 40
      Width = 57
      Height = 25
      Caption = 'All'
      TabOrder = 3
      OnClick = BtAddSelectPeopleAllClick
    end
  end
  object Pn2: TPanel
    Left = 0
    Top = 0
    Width = 911
    Height = 49
    Align = alTop
    TabOrder = 1
    object BtLOadDocDocument: TButton
      Left = 24
      Top = 8
      Width = 193
      Height = 25
      Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1096#1072#1073#1083#1086#1085' '#1076#1086#1082#1091#1084#1077#1085#1090#1072' .doc'
      TabOrder = 0
      OnClick = BtLOadDocDocumentClick
    end
    object BtGoFile: TButton
      Left = 432
      Top = 8
      Width = 153
      Height = 25
      Caption = #1057#1086#1079#1076#1072#1090#1100' '#1092#1072#1081#1083#1099
      TabOrder = 1
      OnClick = BtGoFileClick
    end
  end
  object PnP: TPanel
    Left = 433
    Top = 49
    Width = 478
    Height = 421
    Align = alClient
    Caption = 'PnP'
    TabOrder = 2
    object SgSetting: TStringGrid
      Left = 1
      Top = 1
      Width = 476
      Height = 419
      Align = alClient
      ColCount = 2
      DefaultColWidth = 220
      RowCount = 23
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
      TabOrder = 0
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
        24)
    end
  end
  object Od1: TOpenDialog
    Left = 224
    Top = 8
  end
end
