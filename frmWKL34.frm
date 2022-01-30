VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL34 
   BackColor       =   &H00C0C000&
   Caption         =   "Bonus auf Bon"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWKL34.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   13
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   53
      Top             =   6720
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   12
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   52
      Top             =   6240
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   11
      Left            =   8760
      MaxLength       =   2
      TabIndex        =   49
      Text            =   "10"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   10
      Left            =   3960
      MaxLength       =   32
      TabIndex        =   48
      Top             =   5400
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   15
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   46
      Top             =   5760
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   9
      Left            =   5160
      MaxLength       =   3
      TabIndex        =   43
      Top             =   4800
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Caption         =   "mit Barcode"
      Height          =   255
      Left            =   6360
      TabIndex        =   42
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Index           =   8
      Left            =   6360
      MaxLength       =   6
      TabIndex        =   40
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "nur mit Kundenbindung"
      Height          =   255
      Left            =   6360
      TabIndex        =   39
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   11160
      TabIndex        =   36
      Text            =   "Euro"
      Top             =   3460
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   6
      Left            =   8640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   35
      Top             =   3720
      Width           =   3015
   End
   Begin sevCommand3.Command Command5 
      Height          =   405
      Index           =   3
      Left            =   120
      TabIndex        =   34
      Top             =   3600
      Width           =   1095
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
      Caption         =   "Test"
      Enabled         =   0   'False
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bonus"
      Height          =   975
      Left            =   1560
      TabIndex        =   26
      Top             =   3600
      Width           =   4695
      Begin VB.OptionButton Option2 
         Caption         =   "% vom Warenwert"
         Height          =   210
         Index           =   1
         Left            =   2640
         TabIndex        =   33
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "fester € - Wert"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   28
         Text            =   "10"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Text            =   "5"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "oder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   1560
         TabIndex        =   31
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "%"
         Height          =   255
         Index           =   15
         Left            =   3240
         TabIndex        =   30
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "€"
         Enabled         =   0   'False
         Height          =   255
         Index           =   14
         Left            =   720
         TabIndex        =   29
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   7560
      TabIndex        =   24
      Text            =   "50"
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   8640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   23
      Top             =   2880
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "keine Variante wählen"
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
      Index           =   3
      Left            =   6480
      TabIndex        =   22
      Top             =   7560
      Width           =   3015
   End
   Begin sevCommand3.Command Command5 
      Height          =   405
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   1680
      Width           =   1095
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
      Caption         =   "Test"
      Enabled         =   0   'False
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "Variante 3 wählen"
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
      Index           =   2
      Left            =   9240
      TabIndex        =   20
      Top             =   6600
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "Variante 2 wählen"
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
      Index           =   1
      Left            =   9240
      TabIndex        =   19
      Top             =   4320
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "Variante 1 wählen"
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
      Index           =   0
      Left            =   9240
      TabIndex        =   18
      Top             =   2400
      Width           =   2415
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   1
      Left            =   9720
      TabIndex        =   11
      Top             =   7440
      Width           =   1935
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
      Height          =   375
      Index           =   1
      Left            =   7560
      TabIndex        =   8
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   7
      Text            =   "Sie haben jetzt einen Bonusstand von"
      Top             =   1560
      Width           =   4815
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   0
      Left            =   9720
      TabIndex        =   1
      Top             =   7920
      Width           =   1935
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
   Begin sevCommand3.Command Command5 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   0
      Top             =   360
      Width           =   345
      _ExtentX        =   0
      _ExtentY        =   0
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
         Size            =   6
         Charset         =   0
         Weight          =   400
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
      Caption         =   "?"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   405
      Index           =   4
      Left            =   120
      TabIndex        =   45
      Top             =   6240
      Width           =   1095
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
      Caption         =   "Test"
      Enabled         =   0   'False
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   54
      Top             =   3960
      Width           =   255
      _ExtentX        =   450
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
      Caption         =   "S"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Barcode für Artikelrabatt"
      Height          =   255
      Index           =   13
      Left            =   8760
      TabIndex        =   51
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "%"
      Height          =   255
      Index           =   7
      Left            =   9360
      TabIndex        =   50
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label lbl6 
      Caption         =   "Bontext: spezieller Bon, der bei Eingabe (Artikel 9999) im Kassenprogrammteil ausgedruckt wird. "
      Height          =   1095
      Index           =   18
      Left            =   1560
      TabIndex        =   47
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Gültigkeit der Bonuseinlösung (Tage)"
      Height          =   255
      Index           =   20
      Left            =   1560
      TabIndex        =   44
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Artnr für Bonusabzug"
      Height          =   255
      Index           =   19
      Left            =   6360
      TabIndex        =   41
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "5,24"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   18
      Left            =   9840
      TabIndex        =   38
      Top             =   3500
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "€"
      Enabled         =   0   'False
      Height          =   255
      Index           =   17
      Left            =   8280
      TabIndex        =   37
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "ab Warenwert:"
      Height          =   255
      Index           =   12
      Left            =   6120
      TabIndex        =   25
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "gewählt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   11
      Left            =   120
      TabIndex        =   17
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "gewählt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "gewählt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Variante 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   14
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   3
      X1              =   120
      X2              =   11640
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   2
      X1              =   120
      X2              =   11640
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "z.B.: Ab einem Warenwert von 52,39 € bekommt der Kunde für den nächsten Einkauf 5 € Bonus."
      Height          =   615
      Index           =   6
      Left            =   1560
      TabIndex        =   13
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Variante 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   11640
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Vorraussetzung: ein Verkauf mit Kundenbindung"
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   10
      Top             =   2040
      Width           =   7455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00C0C000&
      Caption         =   "145"
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   9
      Top             =   1620
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Beispiel bei 145,67 Euro Bonus:"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   6
      Top             =   1200
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Sie drucken am Ende des Bons den derzeitigen Bonuswert in Punkten. "
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Variante 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Bonus auf Bon"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   9255
   End
End
Attribute VB_Name = "frmWKL34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Check2_Click()
On Error GoTo LOKAL_ERROR

    If Check2.Value = vbChecked Then
        Label2(19).Visible = True
        Text1(8).Visible = True
        Command5(5).Visible = True
    Else
        Label2(19).Visible = False
        Text1(8).Visible = False
        Command5(5).Visible = False
    End If
    
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check2_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonus auf Bon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim sPfad As String
    Dim sSQL As String
    Dim sdbPfad As String
   
    Select Case Index
        Case 0
            Command5_Click 1
            Unload frmWKL34
        Case 1
                'Speichern
            If Option1(0).Value = True Then
                speicherBonus 0
            ElseIf Option1(1).Value = True Then
                speicherBonus 1
            ElseIf Option1(2).Value = True Then
                speicherBonus 2
            ElseIf Option1(3).Value = True Then
                sSQL = "Delete from Bonusart"
                gdBase.Execute sSQL, dbFailOnError
                
                loeschNEW "BONUSBONTEXTE", gdBase
            End If
            
            leseBonusArt
            
        Case 2 'Test Variante 1
            Command5_Click 1
        
            speicherBonus 0
            leseBonusBonTexte
            frmWKLar.Show 1
        Case 3 'Test Variante 2
            Command5_Click 1
        
            speicherBonus 1
            leseBWWBonTexte
            frmWKLar.Show 1
            
        Case 4 'Test Variante 3
            Command5_Click 1
        
            speicherBonus 2
            
            frmWKLar.Show 1

        Case 5
            gcSuch = ""
            gsARTNR = ""
            frmWKL70.Show 1
            Me.Refresh
            If gsARTNR <> "" Then
                Text1(8).Text = gsARTNR
                gsARTNR = ""
            End If
        Case 11
            gsHelpstring = "Bonus auf Bon"
            frmWKL110.Show 1
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonus auf Bon ist ein Fehler aufgetreten."
    
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
Private Sub speicherBonus(iBonusart As Integer)
On Error GoTo LOKAL_ERROR

Dim sSQL As String
    Select Case iBonusart
    
        Case 0
            sSQL = "Delete from Bonusart"
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Insert into Bonusart (Nr,Beschreib) values (" & iBonusart & " , 'Variante 1')"
            gdBase.Execute sSQL, dbFailOnError
            
            speicherBonusBonTexte
            
        Case 1
            sSQL = "Delete from Bonusart"
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Insert into Bonusart (Nr,Beschreib) values (" & iBonusart & " , 'Variante 2')"
            gdBase.Execute sSQL, dbFailOnError
            
            speicherBWWBonTexte
            
        Case 2
            sSQL = "Delete from Bonusart"
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Insert into Bonusart (Nr,Beschreib) values (" & iBonusart & " , 'Variante 2')"
            gdBase.Execute sSQL, dbFailOnError
            
            SpeicherBenutzerBonus
    End Select
    
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherBonus"
    Fehler.gsFehlertext = "Im Programmteil Bonus auf Bon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherBenutzerBonus()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
   
    gsSpezBontext = ""
    If Text1(15).Text <> "" Then

        sSQL = "Update KASSEIN Set SPEZBONTEXT = '" & Trim(Text1(15).Text) & "'"
        gdBase.Execute sSQL, dbFailOnError

        gsSpezBontext = Trim(Text1(15).Text)
    End If
    
    gsSpezBontext2 = ""
    If Text1(12).Text <> "" Then

        sSQL = "Update KASSEIN Set SPEZBONTEXT2 = '" & Trim(Text1(12).Text) & "'"
        gdBase.Execute sSQL, dbFailOnError

        gsSpezBontext2 = Trim(Text1(12).Text)
    End If
    
    gsSpezBontext3 = ""
    If Text1(13).Text <> "" Then

        sSQL = "Update KASSEIN Set SPEZBONTEXT3 = '" & Trim(Text1(13).Text) & "'"
        gdBase.Execute sSQL, dbFailOnError

        gsSpezBontext3 = Trim(Text1(13).Text)
    End If
    
    gsSpezBontextU = ""
    If Text1(10).Text <> "" Then

        sSQL = "Update KASSEIN Set SPEZBONTEXTU = '" & Trim(Text1(10).Text) & "'"
        gdBase.Execute sSQL, dbFailOnError

        gsSpezBontextU = Trim(Text1(10).Text)
    End If
    
    gsSpezBonArtRab = ""
    If Text1(11).Text <> "" Then

        sSQL = "Update KASSEIN Set SPEZBONARTRAB = '" & Trim(Text1(11).Text) & "'"
        gdBase.Execute sSQL, dbFailOnError

        gsSpezBonArtRab = Trim(Text1(11).Text)
    End If
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherBenutzerBonus"
    Fehler.gsFehlertext = "Im Programmteil Einstellungen an der Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherBonusBonTexte()
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim sTextvor As String
Dim sTextnach As String

sTextvor = Text1(0).Text
sTextnach = Text1(1).Text

loeschNEW "BONUSBONTEXTE", gdBase
CreateTableT2 "BONUSBONTEXTE", gdBase

sSQL = "Insert into BonusBonTexte (Textvor,Textnach) values ('" & sTextvor & "' , '" & sTextnach & "')"
gdBase.Execute sSQL, dbFailOnError
            
anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherBonusBonTexte"
    Fehler.gsFehlertext = "Im Programmteil Bonus auf Bon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherBWWBonTexte()
On Error GoTo LOKAL_ERROR

Dim sSQL            As String
Dim sTextvor        As String
Dim sTextnach       As String
Dim sZeichen        As String
Dim sArt            As String
Dim sWert           As String
Dim sSchwellenWert  As String
Dim sBonusArtnr     As String
Dim sGDAUER         As String
Dim bo1             As Integer

sTextvor = Text1(2).Text
sZeichen = Text1(7).Text
sTextnach = Text1(6).Text
sSchwellenWert = Val(Text1(3).Text)
sBonusArtnr = Val(Text1(8).Text)
sGDAUER = Val(Text1(9).Text)

If Check1.Value = vbChecked Then
    bo1 = 0
Else
    bo1 = -1
End If

If Option2(0).Value = True Then
    sArt = "Euro"
    sWert = Text1(4).Text
Else
    sArt = "Prozent"
    sWert = Text1(5).Text
End If

loeschNEW "BWWBONTEXTE", gdBase
CreateTableT2 "BWWBONTEXTE", gdBase

sSQL = "Insert into BWWBONTEXTE (Textvor,Zeichen,Textnach,Art,Wert,Schwellenwert,Kundbi,BonusArtnr,GDAUER) "
sSQL = sSQL & " values ('" & sTextvor & "','" & sZeichen & "','" & sTextnach & "','" & sArt & "','" & sWert & "','" & sSchwellenWert & "'," & bo1 & "," & sBonusArtnr & "," & sGDAUER & ")"
gdBase.Execute sSQL, dbFailOnError
            
anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherBWWBonTexte"
    Fehler.gsFehlertext = "Im Programmteil Bonus auf Bon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    leseBonusArt
    If giBonusNr <> -1 Then
        Option1(giBonusNr).Value = True
    Else
        Option1(3).Value = True
    End If
    
    Standardfuellung 0
    Standardfuellung 1
    Standardfuellung 2
    
    If giBonusNr = 0 Then
        leseBonusBonTexte
        
        If gsTextVor <> "" Then
            Text1(0).Text = gsTextVor
        Else
            Text1(0).Text = "Sie haben jetzt einen Bonusstand von"
        End If
        
        If gsTextNach <> "" Then
            Text1(1).Text = gsTextNach
        Else
            Text1(1).Text = "Punkten erreicht."
        End If
    ElseIf giBonusNr = 1 Then
        leseBWWBonTexte
        
        If gsTextVor <> "" Then
            Text1(2).Text = gsTextVor
        Else
            Text1(2).Text = ""
        End If
        
        If gsWWZeichen <> "" Then
            Text1(7).Text = gsWWZeichen
        Else
            Text1(7).Text = ""
        End If
        
        If gsTextNach <> "" Then
            Text1(6).Text = gsTextNach
        Else
            Text1(6).Text = ""
        End If
        
        If gsWWArt <> "" Then
            If gsWWArt = "Prozent" Then
                Option2(1).Value = vbChecked
                Text1(5).Text = gsWWwert
            Else
                Option2(0).Value = vbChecked
                Text1(4).Text = gsWWwert
            End If
        Else
            Option2(1).Value = vbChecked
            Text1(5).Text = "3"
        End If
        
        If gsWWSchwellenwert <> "" Then
            Text1(3).Text = gsWWSchwellenwert
        Else
            Text1(3).Text = ""
        End If
        
        If gbWWKundBi = True Then
            Check1.Value = vbChecked
        Else
            Check1.Value = vbUnchecked
        End If
        
        If gsWWBonusArtnr <> "0" Then
            Check2.Value = vbChecked
            Label2(19).Visible = True
            Text1(8).Visible = True
            Text1(8).Text = gsWWBonusArtnr
            Command5(5).Visible = True
        Else
            Check2.Value = vbUnchecked
            Label2(19).Visible = False
            Text1(8).Visible = False
            Text1(8).Text = "0"
            Command5(5).Visible = False
        End If
        
        Text1(9).Text = gsWWBonusGDAUER
    End If
    
    anzeige "normal", "", Label1(4)
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Bonus auf Bon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
    
        Case 0
            Label2(9).Visible = True
            Label2(9).ForeColor = vbRed
            Command5(2).Enabled = True
            
            Label2(11).Visible = False
            Label2(10).Visible = False
            
            Command5(3).Enabled = False
            Command5(4).Enabled = False
            
            Standardfuellung 0
            
        Case 1
            Label2(10).Visible = True
            Label2(10).ForeColor = vbRed
            Command5(3).Enabled = True
            
            Label2(9).Visible = False
            Label2(11).Visible = False
            
            Command5(2).Enabled = False
            Command5(4).Enabled = False
            
            Standardfuellung 1
            
            leseBWWBonTexte
        
            If gsTextVor <> "" Then
                Text1(2).Text = gsTextVor
            Else
                Text1(2).Text = ""
            End If
            
            If gsWWZeichen <> "" Then
                Text1(7).Text = gsWWZeichen
            Else
                Text1(7).Text = ""
            End If
            
            If gsTextNach <> "" Then
                Text1(6).Text = gsTextNach
            Else
                Text1(6).Text = ""
            End If
            
            If gsWWArt <> "" Then
                If gsWWArt = "Prozent" Then
                    Option2(1).Value = vbChecked
                    Text1(5).Text = gsWWwert
                Else
                    Option2(0).Value = vbChecked
                    Text1(4).Text = gsWWwert
                End If
            Else
                Option2(1).Value = vbChecked
                Text1(5).Text = "3"
            End If
            
            If gsWWSchwellenwert <> "" Then
                Text1(3).Text = gsWWSchwellenwert
            Else
                Text1(3).Text = ""
            End If
            
            If gbWWKundBi = True Then
                Check1.Value = vbChecked
            Else
                Check1.Value = vbUnchecked
            End If
            
            If gsWWBonusArtnr <> "0" Then
                Check2.Value = vbChecked
                Label2(19).Visible = True
                Text1(8).Visible = True
                Text1(8).Text = gsWWBonusArtnr
                Command5(5).Visible = True
            Else
                Check2.Value = vbUnchecked
                Label2(19).Visible = False
                Text1(8).Visible = False
                Text1(8).Text = "0"
                Command5(5).Visible = False
            End If
            
            Text1(9).Text = gsWWBonusGDAUER
            
        Case 2
            Label2(11).Visible = True
            Label2(11).ForeColor = vbRed
            
            Label2(9).Visible = False
            Label2(10).Visible = False
            
            Command5(4).Enabled = True
            
            Command5(2).Enabled = False
            Command5(3).Enabled = False
            
            Standardfuellung 2
            
            
        Case 3
            Label2(11).Visible = False
            Label2(9).Visible = False
            Label2(10).Visible = False
            
            Command5(2).Enabled = False
            Command5(3).Enabled = False
    End Select
    
'    Command5_Click 1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonus auf Bon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Standardfuellung(iVar As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case iVar
        Case 0
            Text1(0).Text = "Sie haben jetzt einen Bonusstand von"
            Text1(1).Text = "Punkten erreicht."
            
        Case 1
            'Standardfüllung
            Text1(2).Text = "Sommerspezial:               10% Rabatt bis zum 31.08.2018. Lösen Sie"
            Text1(7).Text = "Euro"
            Text1(6).Text = "gegen Vorlage des Kassenbons beim nächsten Besuch in unserem Geschäft ein."
            Text1(3).Text = "50,00"
            
            Text1(5).Text = "10"
            Option2(1).Value = True
        Case 2
            Text1(15).Text = gsSpezBontext
            Text1(12).Text = gsSpezBontext2
            Text1(13).Text = gsSpezBontext3
            Text1(10).Text = gsSpezBontextU
            Text1(11).Text = gsSpezBonArtRab
    End Select
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonus auf Bon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Option2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case 1
            If Option2(0).Value = vbChecked Then
                Text1(5).Enabled = False
                Label2(15).Enabled = False
                Text1(4).Enabled = True
                Label2(14).Enabled = True
            Else
                Text1(5).Enabled = True
                Label2(15).Enabled = True
                Text1(4).Enabled = False
                Label2(14).Enabled = False
            End If
        Case 0
            If Option2(1).Value = vbChecked Then
                Text1(4).Enabled = False
                Label2(14).Enabled = False
                Text1(5).Enabled = True
                Label2(15).Enabled = True
            Else
                Text1(4).Enabled = True
                Label2(14).Enabled = True
                Text1(5).Enabled = False
                Label2(15).Enabled = False
            End If
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option2_Click"
    Fehler.gsFehlertext = "Im Programmteil Bonus auf Bon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bonus auf Bon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = vbWhite

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Bonus auf Bon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    Select Case Index
        Case 0, 1, 2, 6, 7
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%€,:;_-."
        Case 3, 4, 5, 11
            cValid = "1234567890," & Chr$(8)
        Case 8, 9
            cValid = "1234567890" & Chr$(8)
        Case 15, 10, 12, 13
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46)  '& - .
            cValid = cValid & "+äÄÜüÖöß/:\%()!§$=?"
    End Select

    cZeichen = Chr$(KeyAscii)

    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(cZeichen)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Bonus auf Bon ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
