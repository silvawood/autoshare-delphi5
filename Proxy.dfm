object frmProxy: TfrmProxy
  Left = 434
  Top = 310
  Width = 259
  Height = 279
  HelpContext = 800
  Caption = 'Proxy Server Settings'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Visible = True
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object lblAddress: TLabel
    Left = 8
    Top = 52
    Width = 41
    Height = 13
    Caption = 'Address:'
  end
  object lblPort: TLabel
    Left = 27
    Top = 84
    Width = 22
    Height = 13
    Caption = 'Port:'
  end
  object edtAddress: TEdit
    Left = 56
    Top = 48
    Width = 187
    Height = 21
    TabOrder = 0
  end
  object edtPort: TEdit
    Left = 56
    Top = 80
    Width = 41
    Height = 21
    TabOrder = 1
  end
  object chkAuthorization: TCheckBox
    Left = 136
    Top = 16
    Width = 105
    Height = 17
    Caption = 'With authorization'
    TabOrder = 2
    OnClick = chkAuthorizationClick
  end
  object btnOK: TButton
    Left = 8
    Top = 216
    Width = 75
    Height = 25
    Caption = '&OK'
    TabOrder = 3
    OnClick = btnOKClick
  end
  object btnCancel: TButton
    Left = 88
    Top = 216
    Width = 75
    Height = 25
    Caption = '&Cancel'
    TabOrder = 4
    OnClick = btnCancelClick
  end
  object grpAuthorization: TGroupBox
    Left = 8
    Top = 112
    Width = 235
    Height = 89
    Caption = 'Authorization'
    TabOrder = 5
    object lblUsername: TLabel
      Left = 14
      Top = 24
      Width = 51
      Height = 13
      Caption = 'Username:'
    end
    object lblPassword: TLabel
      Left = 16
      Top = 56
      Width = 49
      Height = 13
      Caption = 'Password:'
    end
    object edtUsername: TEdit
      Left = 72
      Top = 20
      Width = 137
      Height = 21
      TabOrder = 0
    end
    object edtPassword: TEdit
      Left = 72
      Top = 52
      Width = 137
      Height = 21
      PasswordChar = '*'
      TabOrder = 1
    end
  end
  object chkUseProxy: TCheckBox
    Left = 8
    Top = 16
    Width = 105
    Height = 17
    Caption = 'Use proxy server'
    TabOrder = 6
    OnClick = chkUseProxyClick
  end
  object btnHelp: TButton
    Left = 168
    Top = 216
    Width = 75
    Height = 25
    Caption = '&Help'
    TabOrder = 7
    OnClick = btnHelpClick
  end
end
