VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL16 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   8910
   ClientLeft      =   2115
   ClientTop       =   360
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8910
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox ChkSSL 
      Caption         =   "SSL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   48
      Top             =   6600
      Value           =   1  'Aktiviert
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   19
      Left            =   7920
      MaxLength       =   8
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Index           =   18
      Left            =   7920
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   6120
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   17
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   5640
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   16
      Left            =   7920
      MaxLength       =   50
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   5160
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   15
      Left            =   7920
      MaxLength       =   30
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   4080
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   14
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   4080
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   13
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   6480
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   12
      Left            =   1920
      MaxLength       =   34
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   5160
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   11
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   9
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   10
      Left            =   4200
      MaxLength       =   13
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   9
      Left            =   1920
      MaxLength       =   13
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   8
      Left            =   7920
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3600
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3600
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3120
      Width           =   5415
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   8760
      TabIndex        =   11
      Top             =   8160
      Width           =   2895
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   8760
      TabIndex        =   10
      Top             =   7560
      Width           =   2895
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Speichern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   7920
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   4680
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   1920
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   1920
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   720
      Width           =   5415
   End
   Begin sevCommand3.Command Command1 
      Height          =   255
      Index           =   2
      Left            =   9720
      TabIndex        =   49
      Top             =   6600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Test - Email"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Standard = 587"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   25
      Left            =   9000
      TabIndex        =   47
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   24
      Left            =   6720
      TabIndex        =   46
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "PW:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   23
      Left            =   6720
      TabIndex        =   44
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "User:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   22
      Left            =   6720
      TabIndex        =   42
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Einstellungen für den Email-Versand"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   21
      Left            =   7200
      TabIndex        =   40
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "SMTP:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   20
      Left            =   6720
      TabIndex        =   39
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "Ort:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   19
      Left            =   3360
      TabIndex        =   37
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "IBAN:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   18
      Left            =   6120
      TabIndex        =   34
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "BIC:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   17
      Left            =   360
      TabIndex        =   33
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   16
      Left            =   120
      TabIndex        =   32
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "Steuernr.:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   15
      Left            =   120
      TabIndex        =   31
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "1 - 9  = Filial-Nummer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   14
      Left            =   3240
      TabIndex        =   29
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "0 = keine Filialen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   13
      Left            =   3240
      TabIndex        =   28
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "Filial-Nr.:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   27
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Rechnungsempfänger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   11
      Left            =   4200
      TabIndex        =   26
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Warenempfänger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   10
      Left            =   1920
      TabIndex        =   25
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "ILN-Nr.:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   24
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "Konto:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   8
      Left            =   6360
      TabIndex        =   23
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "BLZ:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "Bank:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "Fax:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   5
      Left            =   6120
      TabIndex        =   17
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "Telefon:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "PLZ:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "Straße:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Unternehmensdaten"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmWKL16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ChkSSL_Click()
On Error GoTo LOKAL_ERROR
    
    If ChkSSL.value = vbChecked Then
        Label1(25).Caption = "Standard = 587"
    Else
        Label1(25).Caption = "Standard = 25"
    End If
    
    Label1(25).Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ChkSSL_Click"
    Fehler.gsFehlertext = "Im Programmteil Unternehmensdaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    LogtoEnd Me
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LadeUnternehmensDatenWKL16()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFilnr As String
    Dim iFileNr As Integer
    
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    Text1(5).Text = ""
    Text1(6).Text = ""
    Text1(7).Text = ""
    Text1(8).Text = ""
    Text1(9).Text = ""
    Text1(10).Text = ""
    Text1(11).Text = ""
    Text1(12).Text = ""
    Text1(13).Text = ""
    Text1(14).Text = ""
    Text1(15).Text = ""
    
    Text1(16).Text = ""
    Text1(17).Text = ""
    Text1(18).Text = ""
    Text1(19).Text = ""
    
    cSQL = "Select * from FIRMA"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!name) Then
            Text1(0).Text = rsrs!name
        End If
        If Not IsNull(rsrs!strasse) Then
            Text1(1).Text = rsrs!strasse
        End If
        If Not IsNull(rsrs!Plz) Then
            Text1(2).Text = rsrs!Plz
        End If
        If Not IsNull(rsrs!Ort) Then
            Text1(3).Text = rsrs!Ort
        End If
        If Not IsNull(rsrs!Tel) Then
            Text1(4).Text = rsrs!Tel
        End If
        If Not IsNull(rsrs!Fax) Then
            Text1(5).Text = rsrs!Fax
        End If
        If Not IsNull(rsrs!BankName) Then
            Text1(6).Text = rsrs!BankName
        End If
        If Not IsNull(rsrs!BLZ) Then
            Text1(7).Text = rsrs!BLZ
        End If
        If Not IsNull(rsrs!Konto) Then
            Text1(8).Text = rsrs!Konto
        End If
        If Not IsNull(rsrs!Steuernr) Then
            Text1(12).Text = rsrs!Steuernr
        End If
        If Not IsNull(rsrs!Email) Then
            Text1(13).Text = rsrs!Email
        End If
        If Not IsNull(rsrs!ILN_1) Then
            Text1(9).Text = rsrs!ILN_1
        End If
        If Not IsNull(rsrs!ILN_2) Then
            Text1(10).Text = rsrs!ILN_2
        End If
        If Not IsNull(rsrs!BIC) Then
            Text1(14).Text = rsrs!BIC
        End If
        If Not IsNull(rsrs!IBAN) Then
            Text1(15).Text = rsrs!IBAN
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Text1(11).Text = gcFilNr
    
    
    
    

    
    Text1(16).Text = gcSMTP_SERVER
    Text1(17).Text = gcSMTP_USER
    Text1(18).Text = gcSMTP_PW
    Text1(19).Text = gcSMTP_PORT
    
    If gbSMTP_SSL Then
        ChkSSL.value = vbChecked
    Else
        ChkSSL.value = vbUnchecked
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LadeUnternehmensDatenWKL16"
    Fehler.gsFehlertext = "Im Programmteil Unternehmensdaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SchreibeDatenFirmaWKL16()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFilnr As String
    Dim iFileNr As Integer
    
    cSQL = "Select * from FIRMA"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    
    rsrs!name = Text1(0).Text
    rsrs!strasse = Trim$(Text1(1).Text)
    rsrs!Plz = Trim$(Text1(2).Text)
    rsrs!Ort = Trim$(Text1(3).Text)
    rsrs!Tel = Trim$(Text1(4).Text)
    rsrs!Fax = Trim$(Text1(5).Text)
    rsrs!BankName = Trim$(Text1(6).Text)
    rsrs!BLZ = Trim$(Text1(7).Text)
    rsrs!Konto = Trim$(Text1(8).Text)
    rsrs!ILN_1 = Trim$(Text1(9).Text)
    rsrs!ILN_2 = Trim$(Text1(10).Text)
    rsrs!Steuernr = Trim(Text1(12).Text)
    rsrs!Email = Trim(Text1(13).Text)
    rsrs!BIC = Trim(Text1(14).Text)
    rsrs!IBAN = Trim(Text1(15).Text)
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    
    
    
    
    'SMTP
    
    gcSMTP_SERVER = Text1(16).Text
    gcSMTP_USER = Text1(17).Text
    gcSMTP_PW = Text1(18).Text
    gcSMTP_PORT = Text1(19).Text
        
    If (Text1(16).Text = "smtp.strato.de") And (Text1(17).Text = "bestsend@kisswws.de") And (Text1(18).Text = "geheim") Then

        gcSMTP_PW = "Ki55!Ww52020"
    
    End If
    
    
    
    If ChkSSL.value = vbChecked Then
        gbSMTP_SSL = True
        
    Else
        gbSMTP_SSL = False
    End If
    
    
    If gcSMTP_SERVER <> "" And gcSMTP_USER <> "" And gcSMTP_PW <> "" And gcSMTP_PORT <> "" Then
    
        'Update Kassein
        
        
        cSQL = "Update KASSEIN Set SMTP_SERVER = '" & gcSMTP_SERVER & "'"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update KASSEIN Set SMTP_USER = '" & gcSMTP_USER & "'"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update KASSEIN Set SMTP_PW = '" & gcSMTP_PW & "'"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update KASSEIN Set SMTP_PORT = '" & gcSMTP_PORT & "'"
        gdBase.Execute cSQL, dbFailOnError
        
        If gbSMTP_SSL = True Then
            cSQL = "Update KASSEIN Set SMTP_SSL = True"
            gdBase.Execute cSQL, dbFailOnError
        Else
            cSQL = "Update KASSEIN Set SMTP_SSL = False"
            gdBase.Execute cSQL, dbFailOnError
        End If
        
        
        
        
        
        
        
    Else
    
        cSQL = "Update KASSEIN Set SMTP_SERVER = ''"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update KASSEIN Set SMTP_USER = ''"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update KASSEIN Set SMTP_PW = ''"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update KASSEIN Set SMTP_PORT = ''"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update KASSEIN Set SMTP_SSL = False "
        gdBase.Execute cSQL, dbFailOnError
        
        gcSMTP_SERVER = ""
        gcSMTP_USER = ""
        gcSMTP_PW = ""
        gcSMTP_PORT = ""
        gbSMTP_SSL = False
    
    End If
    
    
    
    
    'Ende SMTP
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

    gcFilNr = Trim(Text1(11).Text)
    
    If gcFilNr = "" Then
        gcFilNr = "1"
    End If
    
    '****************Fila
    
    cSQL = "Delete from FILA "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into FILA  (fil)"
    cSQL = cSQL & "values ("
    cSQL = cSQL & gcFilNr
    cSQL = cSQL & " ) "
    gdBase.Execute cSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("Fila", dbOpenTable)
    If Not rsrs.EOF Then
        gcFilNr = rsrs!fil
    End If
    rsrs.Close: Set rsrs = Nothing
    
    gbFilNr = True
    '****************Fila Ende
    
    frmWKL00.Label1(13).Caption = "F " & gcFilNr
    
    LeseFirmenDaten

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenFirmaWKL16"
    Fehler.gsFehlertext = "Im Programmteil Unternehmensdaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    Select Case index
        Case Is = 0     'Speichern
            SchreibeDatenFirmaWKL16
            Unload frmWKL16
        Case Is = 1     'Beenden
            Unload frmWKL16
        Case Is = 2     'Testmail
            SchreibeDatenFirmaWKL16
            SendeTestMail
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Unternehmensdaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SendeTestMail()
    On Error GoTo LOKAL_ERROR
    
    Dim cAbsenderEmail As String
    cAbsenderEmail = ermFirmenMail
    
    If cAbsenderEmail <> "" Then
        Dim sAttachment As String
        sAttachment = ""
        
        
        
        Dim sMess As String
    
        sMess = "Die Test-Email wurde erfolgreich zugestellt."
    
        schickeMailimHintergrundSSL ermFirmenBez, cAbsenderEmail, cAbsenderEmail, cAbsenderEmail _
        , cAbsenderEmail, gcSMTP_SERVER, gcSMTP_PORT, gcSMTP_USER, gcSMTP_PW, "Test-Email", sMess, sAttachment
                    
    End If
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SendeTestMail"
    Fehler.gsFehlertext = "Im Programmteil Unternehmensdaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    LadeUnternehmensDatenWKL16
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Unternehmensdaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(index).BackColor = glSelBack1
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Unternehmensdaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    
    Select Case index
        Case Is = 11
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Unternehmensdaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(index).BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Unternehmensdaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


