object Frm_Principal: TFrm_Principal
  Left = 0
  Top = 0
  Caption = 'Windows e Acessibilidade'
  ClientHeight = 451
  ClientWidth = 585
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Lbl_Acessibilidade: TLabel
    Left = 6
    Top = 14
    Width = 320
    Height = 13
    Caption = 
      'Interface IAccessible ( Nome, Valor, Role, Classe, Texto Windows' +
      ')'
  end
  object Lbl_TextoLeitor: TLabel
    Left = 344
    Top = 14
    Width = 221
    Height = 13
    Caption = 'Texto Principal para um possivel Leitor de Tela'
  end
  object MemoInfo: TMemo
    Left = 6
    Top = 32
    Width = 306
    Height = 402
    ScrollBars = ssBoth
    TabOrder = 0
    WordWrap = False
  end
  object Mmo_TxtLeitor: TMemo
    Left = 329
    Top = 32
    Width = 250
    Height = 402
    ScrollBars = ssVertical
    TabOrder = 1
  end
  object Tmr_Info: TTimer
    Interval = 100
    OnTimer = Tmr_InfoTimer
    Left = 89
    Top = 176
  end
end
