VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmWKL47 
   Caption         =   "Bestellung zusammenstellen"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL47.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.OptionButton Option1 
      Caption         =   "Bestellmenge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   35
      Top             =   1440
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "VPE-Faktor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4560
      TabIndex        =   34
      Top             =   1200
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   8760
      TabIndex        =   33
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picprogress 
      Height          =   135
      Left            =   9840
      ScaleHeight     =   75
      ScaleWidth      =   1275
      TabIndex        =   32
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   -120
      TabIndex        =   2
      Top             =   4680
      Width           =   10575
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   120
         MultiSelect     =   1  '1 -Einfach
         TabIndex        =   4
         Top             =   2400
         Visible         =   0   'False
         Width           =   11445
      End
      Begin sevCommand3.Command cmdAnfuegen 
         Height          =   495
         Index           =   1
         Left            =   9480
         TabIndex        =   11
         Top             =   6840
         Width           =   2055
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdAnfuegen 
         Height          =   495
         Index           =   0
         Left            =   9480
         TabIndex        =   10
         ToolTipText     =   "Ohne Eingabe erhalten Sie alle  Artikel des Lieferanten."
         Top             =   360
         Width           =   2055
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
         Caption         =   "Suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   2640
         TabIndex        =   9
         Top             =   1320
         Width           =   3732
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   2640
         MaxLength       =   13
         TabIndex        =   8
         Top             =   960
         Width           =   2652
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   7
         Top             =   600
         Width           =   1572
      End
      Begin sevCommand3.Command cmdAnfuegen 
         Height          =   495
         Index           =   2
         Left            =   9480
         TabIndex        =   6
         Top             =   960
         Width           =   2055
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
         Caption         =   "Auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   11445
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   2640
         TabIndex        =   3
         Top             =   1680
         Width           =   3732
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   6960
         Width           =   7095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel Anfügen:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   120
         Width           =   9855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Artikelbezeichnung :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "EAN :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   14
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Artikelnummer :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferantenbestellnr :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   2415
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   1
      Top             =   7800
      Width           =   2055
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   5055
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   23
      FixedCols       =   0
      BackColor       =   16777215
      BackColorSel    =   16711680
      ForeColorSel    =   65535
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   17
      Top             =   7280
      Width           =   2055
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
      Caption         =   "Anfügen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   7520
      TabIndex        =   19
      Top             =   7280
      Width           =   2055
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
      Caption         =   "Drucken"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   255
      Index           =   5
      Left            =   11280
      TabIndex        =   24
      Top             =   960
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "F2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   23
      Top             =   1320
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   255
      Index           =   3
      Left            =   7560
      TabIndex        =   22
      Top             =   960
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Bestellwerte leeren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      TabIndex        =   27
      Top             =   1600
      Width           =   615
   End
   Begin sevCommand3.Command Command5 
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   29
      Top             =   1560
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "Bestellwert"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9840
      MaxLength       =   6
      TabIndex        =   30
      Top             =   1320
      Width           =   1815
   End
   Begin sevCommand3.Command Command5 
      Height          =   310
      Index           =   7
      Left            =   9840
      TabIndex        =   31
      Top             =   360
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
      Caption         =   "MDE"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   360
      Left            =   11280
      TabIndex        =   40
      Top             =   360
      Width           =   405
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      ToolTip         =   "Spaltenanordung der Tabelle bestimmen"
      ToolTipTitle    =   "Spaltenanordung"
      ButtonStyle     =   2
      Caption         =   ""
      Filename        =   "D:\Thomas\VB6\Winkiss\Zubehör\tab24.gif"
      Picture         =   "frmWKL47.frx":0442
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   8
      Left            =   7520
      TabIndex        =   37
      Top             =   7800
      Width           =   2055
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
      Caption         =   "Senden"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      MaxLength       =   6
      TabIndex        =   38
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label lbl6 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      Caption         =   "AuftragsNr.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   6120
      TabIndex        =   39
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mengenangabe als"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   36
      Top             =   960
      Width           =   2175
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
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Einkaufswert"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   28
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Standardwert Mengenangabe:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   26
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   7920
      Width           =   9375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ihre Zusammenstellung:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Bestellung zusammenstellen"
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
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Lieferant:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Index           =   1
      Left            =   9840
      TabIndex        =   25
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmWKL47"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerBEstellen  As Integer
Dim SpaltennummerArtnr      As Integer
Dim SpaltennummerBEZEICH    As Integer
Dim SpaltennummerAWM        As Integer

Dim mdeErr As Boolean

Dim sFormat As String
Dim sKUNDNR As String
Dim sLieferant As String

Private Sub cmdAnfuegen_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            ArtikelSuchenWKL47
        Case 1
            Frame1.Visible = False
            anzeigeNew "Normal", "", Label9
        Case 2
            ArtikelAnfuegenWKL47
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdAnfuegen_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
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
Private Sub ArtikelAnfuegenWKL47()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount      As Long
    Dim lrow        As Long
    Dim lcol        As Long
    Dim ctmp        As String
    Dim cSQL        As String
    Dim cJahr       As String
    Dim cMonat      As String
    Dim cJahrV      As String
    Dim cMonatV     As String
    Dim iTmp        As Integer
    Dim rsrs        As Recordset
    Dim rsArt       As Recordset
    Dim rsLIEF      As Recordset
    Dim rsRS2       As Recordset
    Dim rsRsF       As Recordset
    Dim bFound      As Boolean
    Dim cArtANFU    As String
    Dim lswMenge    As Long
    Dim sMin_Linr   As String
    
    Screen.MousePointer = 11
    
    lswMenge = CLng(Val(Combo1.Text))
    bFound = False
    
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            bFound = True
            Exit For
        End If
    Next lcount
    
    If Not bFound Then
        anzeigeNew "rot", "Bitte markieren Sie einen Artikel", Label9
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    anzeigeNew "Normal", "Artikel werden angefügt...", Label9
    
    cArtANFU = ""
    
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            cArtANFU = List2.list(lcount)
            cArtANFU = Trim(Left(cArtANFU, 6))
            
            Set rsArt = gdBase.OpenRecordset("select * from Artikel where artnr = " & cArtANFU)
            If Not rsArt.EOF Then
            
                sMin_Linr = ermLiefmitkleinstenLEKPR(cArtANFU)
                
                Set rsRS2 = gdApp.OpenRecordset("select * from MANUB where artnr = " & cArtANFU)
                If Not rsRS2.EOF Then
                    rsRS2.Edit
                    
                    If Option1(0).Value = True Then
                        'Das ist Faktor * VPE
                        rsRS2!ZUBEST = rsRS2!ZUBEST + (Val(lswMenge) * Val(rsRS2!MINMEN))
                    ElseIf Option1(1).Value = True Then
                        'Das ist Menge
                        rsRS2!ZUBEST = rsRS2!ZUBEST + Val(lswMenge)
                    End If
                    
                    rsRS2.Update
                Else
                    rsRS2.AddNew
                    rsRS2!artnr = rsArt!artnr
                    rsRS2!BEZEICH = rsArt!BEZEICH
                    rsRS2!BESTAND = rsArt!BESTAND
                    rsRS2!AGN = rsArt!AGN
                    rsRS2!vkpr = rsArt!vkpr
                    rsRS2!KVKPR1 = rsArt!KVKPR1
                    rsRS2!EAN = rsArt!EAN
'                    rsRS2!RKZ = rsArt!RKZ
                    rsRS2!LPZ = rsArt!LPZ
                    rsRS2!NOTIZEN = rsArt!NOTIZEN
                    rsRS2!MINBEST = rsArt!MINBEST
                    rsRS2!AWM = "99"
                    
                    rsRS2!linr = sMin_Linr
    
                    Set rsLIEF = gdBase.OpenRecordset("select * from Artlief where artnr = " & cArtANFU & " and linr = " & sMin_Linr)
                    If Not rsLIEF.EOF Then
                        If Not IsNull(rsLIEF!LIBESNR) Then
                            rsRS2!LIBESNR = rsLIEF!LIBESNR
                        Else
                            rsRS2!LIBESNR = ""
                        End If
    
                        If Not IsNull(rsLIEF!lekpr) Then
                            rsRS2!lekpr = rsLIEF!lekpr
                        Else
                            rsRS2!lekpr = 0
                        End If
                        
                        If Not IsNull(rsLIEF!MINMEN) Then
                            rsRS2!MINMEN = rsLIEF!MINMEN
                        Else
                            rsRS2!MINMEN = 0
                        End If
                        
                        If Not IsNull(rsLIEF!RKZ) Then
                            rsRS2!RKZ = rsLIEF!RKZ
                        Else
                            rsRS2!RKZ = "N"
                        End If
                    Else
                        rsRS2!LIBESNR = ""
                        rsRS2!lekpr = 0
                        rsRS2!MINMEN = 0
                        rsRS2!RKZ = "N"
                    End If
                    rsLIEF.Close: Set rsLIEF = Nothing
                    
                    
                    
                    If Option1(0).Value = True Then
                        'Das ist Faktor * VPE
                        rsRS2!ZUBEST = Val(lswMenge) * Val(rsRS2!MINMEN)
                    ElseIf Option1(1).Value = True Then
                        'Das ist Menge
                        rsRS2!ZUBEST = Val(lswMenge)
                    End If
                    
                    
                    
                    
                    rsRS2.Update
                End If
                  
                rsRS2.Close: Set rsRS2 = Nothing
            End If
            rsArt.Close: Set rsArt = Nothing
        End If
    Next lcount
    
    GridFuellen
    
    anzeigeNew "Normal", "angefügte Artikel sind in blauer Schriftfarbe hervorgehoben", Label9
    
    Frame1.Visible = False
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ArtikelAnfuegenWKL47"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub ArtikelSuchenWKL47()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp        As String
    Dim cSQL        As String
    Dim cPfad       As String
    Dim cPfad1      As String
    Dim cArtNr      As String
    Dim cBez        As String
    Dim cEAN        As String
    Dim cLiBesNr    As String
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim i           As Long
    Dim rsrs        As Recordset
    Dim cMinArt     As String
    Dim cdatei      As String
    
    
    Screen.MousePointer = 11
    cPfad = gcPfad  'Anwendungspfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad1 = gcDBPfad   'Datenbankpfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
   
    cArtNr = Trim$(Text6.Text)
    cEAN = Trim$(Text7.Text)
    cBez = Trim$(Text8.Text)
    cLiBesNr = Trim$(Text2.Text)
    
    anzeigeNew "Normal", "Artikel werden gesucht...", Label1
    
    List1.Clear
    List2.Clear
    ctmp = "ArtNr" & Space$(2) & "Artikelbezeichnung" & Space$(19) & "EAN" & Space$(11) & "BestellNr." & Space$(4) & "Bestand" & Space$(1) & "RKZ"
    List1.AddItem ctmp
    
    loeschNEW "Winanfu1", gdBase
    
    
    cSQL = "Select min(artnr) as minart from artikel "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.RecordCount = 0 Then
        Screen.MousePointer = 0
        rsrs.Close: Set rsrs = Nothing
        Exit Sub
    Else
        cMinArt = rsrs!minart
    End If
    rsrs.Close: Set rsrs = Nothing

    cSQL = "Select * into WINANFU1 from Artikel where ARTNR = " & cMinArt & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from WINANFU1"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into WINANFU1 SELECT "
    cSQL = cSQL & " B.ARTNR "
    cSQL = cSQL & ", A.BEZEICH"
    cSQL = cSQL & ", A.AGN"
    cSQL = cSQL & ", B.LEKPR"
    cSQL = cSQL & ", A.VKPR"
    cSQL = cSQL & ", A.MWST"
    cSQL = cSQL & ", B.LINR"
    cSQL = cSQL & ", B.LIBESNR"
    cSQL = cSQL & ", A.EAN"
    cSQL = cSQL & ", A.EAN2"
    
    cSQL = cSQL & ", A.EAN3"
    cSQL = cSQL & ", A.ETIMERK "
    cSQL = cSQL & ", A.MOPREIS"
    cSQL = cSQL & ", B.RKZ"
    cSQL = cSQL & ", A.LPZ"
    cSQL = cSQL & ", A.NOTIZEN"
    cSQL = cSQL & ", A.BESTAND"
    cSQL = cSQL & ", A.VKMENGE"
    cSQL = cSQL & ", A.VKDATUM"
    cSQL = cSQL & ", B.MINMEN"
    
    cSQL = cSQL & ", A.INHALT"
    cSQL = cSQL & ", A.INHALTBEZ"
    cSQL = cSQL & ", A.GRUNDPREIS"
    cSQL = cSQL & ", A.MINBEST"
    cSQL = cSQL & ", A.RABATT_OK"
    cSQL = cSQL & ", A.GEFUEHRT"
    cSQL = cSQL & ", A.KVKPR1"
    cSQL = cSQL & ", A.EKPR"
    cSQL = cSQL & ", A.PREISSCHU"
    cSQL = cSQL & ", A.BONUS_OK"
    
    cSQL = cSQL & ", A.UMS_OK"
    cSQL = cSQL & ", A.AWM"
    cSQL = cSQL & ", A.LASTDATE"
    cSQL = cSQL & ", A.LASTTIME"
    cSQL = cSQL & ", A.AUFDAT"
    cSQL = cSQL & ", A.EXDAT"
    cSQL = cSQL & ", A.FARBNR"
    cSQL = cSQL & ", A.GROESSE"
    cSQL = cSQL & ", A.SPANNE"
    
    cSQL = cSQL & " from ARTLIEF B inner join ARTIKEL A on A.ARTNR = B.ARTNR where B.LINR <> 0 "

    If cArtNr <> "" Then cSQL = cSQL & " and B.ARTNR = " & cArtNr & " "
    If cBez <> "" Then cSQL = cSQL & " and A.BEZEICH like '" & cBez & "*' "
    If cEAN <> "" Then cSQL = cSQL & " and A.EAN like '" & cEAN & "*' "
    If cLiBesNr <> "" Then cSQL = cSQL & " and B.LIBESNR like '" & cLiBesNr & "*' "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschapp "WINANFU"

    cdatei = cPfad & "kissapp.mdb"
    cSQL = "Select winanfu1.* INTO winanfu IN '" & cdatei & "' from winanfu1"
    gdBase.Execute cSQL, dbFailOnError
        

    cSQL = "Delete from WINANFU where ARTNR in (Select ARTNR from Manub where winanfu.artnr = manub.artnr)"
    gdApp.Execute cSQL, dbFailOnError
    
    

    
    
    
    cSQL = "Select ARTNR,BEZEICH,EAN,LIBESNR,Bestand,RKZ from WINANFU group by artnr,BEZEICH,EAN,LIBESNR,Bestand,RKZ order by BEZEICH"
    Set rsrs = gdApp.OpenRecordset(cSQL)

    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(6 - Len(cFeld))
            cLBSatz = cFeld
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = Space$(1) & cFeld & Space$(36 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld
            
            If Not IsNull(rsrs!EAN) Then
                cFeld = rsrs!EAN
            Else
                cFeld = ""
            End If
            cFeld = Space$(14 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld
            
            If Not IsNull(rsrs!LIBESNR) Then
                cFeld = Trim(rsrs!LIBESNR)
            Else
                cFeld = ""
            End If
            
            cFeld = Space$(1) & cFeld & Space$(15 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld
            
            If Not IsNull(rsrs!BESTAND) Then
                cFeld = rsrs!BESTAND
            Else
                cFeld = ""
            End If
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld
            
            If Not IsNull(rsrs!RKZ) Then
                cFeld = rsrs!RKZ
            Else
                cFeld = ""
            End If
            cFeld = Space$(4 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld
            
            List2.AddItem cLBSatz
            rsrs.MoveNext
        Loop
        List1.Visible = True
        List2.Visible = True
    Else
        
        anzeigeNew "rot", "Der gesuchte Artikel wird entweder bereits angezeigt oder existiert nicht!", Label1
    End If
    rsrs.Close: Set rsrs = Nothing
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 58 Or err.Number = 3167 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ArtikelSuchenWKL47"
        Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub


'Private Sub Combo2_GotFocus()
'    On Error GoTo LOKAL_ERROR
'
'    Combo2.BackColor = glSelBack1
'    Combo2.SelStart = Len(Combo2.Text)
'Exit Sub
'LOKAL_ERROR:
'
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Combo2_GotFocus"
'    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
'Private Sub Combo2_LostFocus()
'    On Error GoTo LOKAL_ERROR
'
'    Combo2.BackColor = vbWhite
'
'Exit Sub
'LOKAL_ERROR:
'
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Combo2_LostFocus"
'    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub

'Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
'    On Error GoTo LOKAL_ERROR
'
'    AutocompleteCombo KeyCode, Shift, Combo2
'
'    If KeyCode = vbKeyF2 Then
'        gF2Prompt.cFeld = ""
'        gF2Prompt.cWert = ""
'        gF2Prompt.cWert2 = ""
'        gF2Prompt.cWahl = ""
'        gF2Prompt.bMultiple = False
'
'
'        gF2Prompt.cFeld = "LINR"
'
'        If gF2Prompt.cFeld <> "" Then
'            frmWK00a.Show 1
'        End If
'
'        If gF2Prompt.cWahl <> "" Then
'            Combo2.Text = gF2Prompt.cWahl
'
'        End If
'        Combo2.SetFocus
'
'    End If
'
'    If KeyCode = vbKeyEscape Then
'        Unload Me
'    End If
'
'
'Exit Sub
'LOKAL_ERROR:
'
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Combo2_KeyUp"
'    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'
'End Sub
Private Sub Text1_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus()
    On Error GoTo LOKAL_ERROR

    Text1.BackColor = glSelBack1
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cValid As String
    cValid = "1234567890" & Chr$(8)
        
    
    If InStr(cValid, Chr$(KeyAscii)) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(Chr$(KeyAscii))
    End If
        

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        
        
        gF2Prompt.bMultiple = False
        gF2Prompt.cFeld = "LINR"
        
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
        End If
        If gF2Prompt.cWahl <> "" Then
            Text1.Text = gF2Prompt.cWahl
        End If
                
        Text1.SetFocus
    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
End Sub

Private Sub Command1_Click()
    
    
    gsZSpalte = "Artnr"
    gstab = "BESTMAN"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim ctmp As String

    Select Case Index
        Case 0
            voreinstellungspeichernE47B
            speicherezubest
            Unload frmWKL47
        Case 1
            Frame1.Visible = True
            anzeigeNew "normal", "", Label9
        Case 2
            speicherezubest
            SchreibeBest True
        Case 3 'Bestellwerte leeren
            BestLeer
            Label6(6).Caption = "0,00 EUR"
            Label6(6).Refresh
        Case 4 'löschen
            ctmp = "Möchten Sie die gesamte Bestellung löschen, dann drücken Sie 'Ja'" & vbCrLf
            ctmp = ctmp & "Oder möchten Sie nur Artikel mit Bestellmenge '0' löschen, dann drücken Sie 'Nein'" & vbCrLf
            iRet = MsgBox(ctmp, vbQuestion + vbYesNoCancel, "Winkiss Frage:")
            If iRet = vbYes Then
                BestdEL
            ElseIf iRet = vbNo Then
                speicherezubest
                BestDelNull
                GridFuellen
            End If
        Case 5    'F2 Lieferant
            Text1_KeyUp vbKeyF2, 0
        Case 6 'Bestwertrechnen
            speicherezubest
            GridFuellen
        Case 7 'MDE auslesen
            MDElesen
        Case 8
            speicherezubest
            SchreibeBest False
            
            If sLieferant <> "" Then
                Verbindungseinstellunglesen sLieferant
                
                If sFormat = "EDIERNST" Then
                    Do Until sKUNDNR <> ""
                        
                        Verbindungseinstellunglesen sLieferant
                             
                        If sKUNDNR = "" Then
                            MsgBox "Es konnte keine Bestellung im Spezialformat geschrieben werden. Bitte vervollständigen Sie Ihre Kundendaten!", vbCritical, "Winkiss Fehler:"
                            gsLinr = sLieferant
                            frmWKL17.Show 1
                            
                        End If
                    Loop
                    Uebertragung sLieferant, BestDateiinOrdner(sLieferant), Text12.Text, sKUNDNR
                End If
                
                If Text12.Text <> "" Then
                    If IsNumeric(Text12.Text) Then
                        schreibeAufnr CLng(Text12.Text), "" ' cZiel
                    End If
                End If
                
                reportbildschirmApp "WKL004ac", "aWKL43a"
                
            End If
            
            
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function BestDateiinOrdner(cLinr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    BestDateiinOrdner = ""
    
    VerzVorhanden cLinr, cPfad & "BESTSIC\"
    
    Select Case UCase(sFormat)
        Case "EDIERNST"
            BestDateiinOrdner = StandardFormat(cLinr)
    End Select
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BestDateiinOrdner"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function StandardFormat(cLinr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim lRows           As Long
    Dim lCols           As Long
    Dim lColArtNr       As Long
    Dim lColBestell     As Long
    Dim lcol            As Long
    Dim lrow            As Long
    Dim lPos            As Long
    Dim lWert           As Long
    Dim cArtNr          As String
    Dim cBestMenge      As String
    Dim cBezeich        As String
    Dim cPfad           As String
    Dim iFileNr         As Integer
    Dim sTime           As String
    Dim ctmp            As String
    Dim cStsatz         As String
    
    StandardFormat = ""
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    sTime = Left(sTime, 5)
    
    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM")
    
    ctmp = ctmp & sTime
    ctmp = SwapStr(ctmp, ".", "")
    ctmp = SwapStr(ctmp, ":", "")
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    VerzVorhanden ctmp, cPfad & "BESTSIC\" & cLinr & "\"
    
    cPfad = cPfad & "BESTSIC\" & cLinr & "\" & ctmp & "\"
    
    StandardFormat = ctmp
    
    lRows = MSFlexGrid2.Rows
    lCols = MSFlexGrid2.Cols
    
    lColArtNr = SpaltennummerArtnr
    
    'Detaildaten
    lColBestell = SpaltennummerBEstellen
    
    iFileNr = FreeFile
    Open cPfad & "bestell.txt" For Binary As #iFileNr
    cStsatz = ""
    For lrow = 1 To lRows - 1
        MSFlexGrid2.Row = lrow
        MSFlexGrid2.Col = lColBestell
        cBestMenge = MSFlexGrid2.Text
        cBestMenge = Trim$(Str$(Val(cBestMenge)))
        If Val(cBestMenge) > 0 Then
            MSFlexGrid2.Col = SpaltennummerBEZEICH
            cBezeich = Trim(MSFlexGrid2.Text)
        
            MSFlexGrid2.Col = lColArtNr
            cArtNr = MSFlexGrid2.Text

            cStsatz = cArtNr & vbTab
            cStsatz = cStsatz & cBezeich & vbTab
            cStsatz = cStsatz & cBestMenge & vbTab
            
            cStsatz = cStsatz & vbCrLf
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cStsatz
        End If
    Next lrow
    
    Close iFileNr

Exit Function
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "StandardFormat"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub Uebertragung(cLinr As String, cverz As String, sAuftragsnummer As String, sKund As String)
On Error GoTo LOKAL_ERROR

    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    
    'kopiere erst aus dem BESTSIC ins EDI Verzeichnis
    cQuelle = gcDBPfad
    If Right$(cQuelle, 1) <> "\" Then
        cQuelle = cQuelle & "\"
    End If
    cQuelle = cQuelle & "BESTSIC\" & cLinr & "\" & cverz & "\bestell.txt"
    cQuelle = ShortPath(cQuelle)
    
    cZiel = App.Path
    cZiel = ShortPath(cZiel)
    cZiel = cZiel & "\EDI\" & sKund & "_" & sAuftragsnummer & ".txt"
    
    lRet = CopyFile(cQuelle, cZiel, lfail)

    If lRet <> 0 Then
        giKissFtpMode = 33
        frmWKL38.Show 1
    Else
        MsgBox "Die Bestelldatei konnte nicht kopiert werden.", vbCritical, "Winkiss Hinweis:"
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Uebertragung"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Verbindungseinstellunglesen(ByRef sLifnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "select * from lisrt where linr = " & sLifnr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!Kundnr) Then
            sKUNDNR = rsrs!Kundnr
        Else
            sKUNDNR = ""
        End If
        
        If Not IsNull(rsrs!Format) Then
            sFormat = rsrs!Format
        Else
            sFormat = ""
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Verbindungseinstellunglesen"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MDElesen()
    On Error GoTo LOKAL_ERROR
    
    If MDEeinlesenOhneLinr(Label9, txtStatus, picprogress, frmWKL47) = False Then
        anzeigeNew "rot", "Es konnten keine Daten aus dem MDE - Gerät ausgelesen werden.", Label9
    Else
        anzeigeNew "normal", "", Label9
        MdeVerarbeitung
        
        GridFuellen
        
        If mdeErr Then
            anzeigeNew "normal", "nicht erkannte Artikel werden angezeigt...", Label9
            reportbildschirm "", "aWKL46e" 'Error artikel mde
        End If

    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MDElesen"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MdeVerarbeitung()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL                As String
    Dim rsMDE               As Recordset
    Dim rsRS2               As Recordset
    Dim rsFilBu             As Recordset
    Dim rsArt               As Recordset
    Dim rsLIEF              As Recordset
    Dim seekEAN             As String
    Dim lMenge              As Long
    Dim lscanfolge          As Long
    Dim sArtnr              As String
    Dim dKleinsterLEKPR     As Double
    Dim lVPE                As Long
    Dim sMin_Linr           As String
    
    Screen.MousePointer = 11
    
    If Not NewTableSuchenDBKombi("ARTT23", gdBase) Then
        CreateTable "ARTT23", gdBase
    End If
    
    loeschNEW "ARTERRIN", gdBase
    CreateTable "ARTERRIN", gdBase
    
    Set rsFilBu = gdBase.OpenRecordset("ARTERRIN")
    
    mdeErr = False
    lscanfolge = 0
    
    anzeigeNew "normal", "Die Daten aus dem MDE - Gerät werden verarbeitet...", Label9
    
    Set rsMDE = gdBase.OpenRecordset("mdeinh")
    If Not rsMDE.EOF Then
        rsMDE.MoveFirst
        Do While Not rsMDE.EOF
        
            lscanfolge = lscanfolge + 1
            If Not IsNull(rsMDE!eancode) Then
            
                seekEAN = Trim(rsMDE!eancode)
                seekEAN = checkean(seekEAN)
                
                
                If Len(seekEAN) = 11 Then
                    seekEAN = "0" & seekEAN
            
                    sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                ElseIf Len(seekEAN) = 8 Then
                    If Left(seekEAN, 1) = "2" Then
                        seekEAN = Mid$(seekEAN, 2, 6)
                        sSQL = "select * from artikel where artnr = " & seekEAN
                    Else
                        sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                    End If
                Else
                    sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                End If

                
                Set rsArt = gdBase.OpenRecordset(sSQL)
                If Not rsArt.EOF Then
                    sArtnr = Trim(rsArt!artnr)
                    
                    sMin_Linr = ermLiefmitkleinstenLEKPR(sArtnr)
                    
                    Set rsRS2 = gdApp.OpenRecordset("select * from MANUB where artnr = " & sArtnr)
                    If Not rsRS2.EOF Then
                    
                        rsRS2.Edit
                        
                        If Option1(0).Value = True Then
                            'Das ist Faktor * VPE
                            rsRS2!ZUBEST = rsRS2!ZUBEST + (Val(rsMDE!Menge) * Val(rsRS2!MINMEN))
                        ElseIf Option1(1).Value = True Then
                            'Das ist Menge
                            rsRS2!ZUBEST = rsRS2!ZUBEST + Val(rsMDE!Menge)
                        End If
                        
                        rsRS2.Update
                        
                    Else
                        rsRS2.AddNew
                        rsRS2!artnr = rsArt!artnr
                        rsRS2!BEZEICH = rsArt!BEZEICH
                        rsRS2!BESTAND = rsArt!BESTAND
                        rsRS2!AGN = rsArt!AGN
                        rsRS2!vkpr = rsArt!vkpr
                        rsRS2!KVKPR1 = rsArt!KVKPR1
                        rsRS2!EAN = rsArt!EAN
                        rsRS2!RKZ = rsArt!RKZ
                        rsRS2!LPZ = rsArt!LPZ
                        rsRS2!NOTIZEN = rsArt!NOTIZEN
                        rsRS2!MINBEST = rsArt!MINBEST
                        rsRS2!AWM = "99"
                        
                        
                        rsRS2!linr = sMin_Linr
    
                        Set rsLIEF = gdBase.OpenRecordset("select * from Artlief where artnr = " & sArtnr & " and linr = " & sMin_Linr)
                        If Not rsLIEF.EOF Then
                            If Not IsNull(rsLIEF!LIBESNR) Then
                                rsRS2!LIBESNR = rsLIEF!LIBESNR
                            Else
                                rsRS2!LIBESNR = ""
                            End If
    
                            If Not IsNull(rsLIEF!lekpr) Then
                                rsRS2!lekpr = rsLIEF!lekpr
                            Else
                                rsRS2!lekpr = 0
                            End If
                            
                            If Not IsNull(rsLIEF!MINMEN) Then
                                rsRS2!MINMEN = rsLIEF!MINMEN
                            Else
                                rsRS2!MINMEN = 0
                            End If
                        Else
                            rsRS2!LIBESNR = ""
                            rsRS2!lekpr = 0
                            rsRS2!MINMEN = 0
                        End If
                        rsLIEF.Close: Set rsLIEF = Nothing
                        
                        
                        If Option1(0).Value = True Then
                            'Das ist Faktor * VPE
                            rsRS2!ZUBEST = Val(rsMDE!Menge) * Val(rsRS2!MINMEN)
                        ElseIf Option1(1).Value = True Then
                            'Das ist Menge
                            rsRS2!ZUBEST = Val(rsMDE!Menge)
                        End If
                        
                        
                        
                        rsRS2.Update
                    End If
                      
                    rsRS2.Close: Set rsRS2 = Nothing
                
                Else 'hier die unbekannten
                
                    mdeErr = True
                    rsFilBu.AddNew
                    rsFilBu!EAN = seekEAN
                    rsFilBu!Menge = rsMDE!Menge
                    rsFilBu!lfnr = lscanfolge
                    rsFilBu.Update
                    
                End If
                rsArt.Close: Set rsArt = Nothing
            End If
            rsMDE.MoveNext
        Loop
    
    End If
    
    rsMDE.Close: Set rsMDE = Nothing
    
    rsFilBu.Close: Set rsFilBu = Nothing
    
    anzeigeNew "normal", "Der Einlesevorgang ist beendet.", Label9
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MdeVerarbeitung"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub SchreibeBest(bmitDruck As Boolean)
    On Error GoTo LOKAL_ERROR

    Dim cSQL         As String
    Dim cLieferant   As String
    Dim cWert        As String
    Dim cLiefFax     As String
    Dim cBLinBez     As String
    Dim cKundnr      As String
    Dim cKundFax     As String
    Dim cMitt        As String
    Dim dWert        As Double
    Dim rsrs         As Recordset
    Dim cPfad        As String
    
    cPfad = gcPfad & "\kissapp.mdb"
    
    If Trim$(Text1.Text) = "" Then
        anzeigeNew "rot", "Bitte geben Sie einen Lieferanten an!", Label9
        Text1.SetFocus
        sLieferant = ""
        Exit Sub
    End If

    sLieferant = Trim(Text1.Text)
    If Trim(sLieferant) = "" Then
        anzeigeNew "rot", "Es konnte keine Lieferantennummer ermittelt werden.", Label9
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    anzeigeNew "normal", "Druckvorschau wird erstellt...", Label9
    
    loeschapp "KDFBEST"
    CreateTable "KDFBEST", gdApp
    
    Text12.Text = SwapStr(Text12.Text, "'", "")
    
    cWert = Label6(6).Caption

    cWert = fnMoveComma2Point$(cWert)
    dWert = Val(cWert)
    cWert = Trim$(Str$(dWert))
    If InStr(cWert, ",") > 0 Then
        cWert = fnMoveComma2Point$(cWert)
    End If

    cSQL = "Insert into KDFBEST "
    cSQL = cSQL & "Select '' as LIEFBEZ, "
    cSQL = cSQL & cWert & " as EK_WERT, "
    cSQL = cSQL & "'" & sLieferant & "' as LINR, "
    cSQL = cSQL & "'" & gFirma.FirmaName & "' as FIRMANAME, "
    cSQL = cSQL & "'" & gFirma.strasse & "' as STRASSE, "
    cSQL = cSQL & "'" & gFirma.Plz & "' as PLZ, "
    cSQL = cSQL & "'" & gFirma.Ort & "' as ORT, "
    cSQL = cSQL & "'" & gFirma.Tel & "' as TEL, "
    cSQL = cSQL & "'" & gFirma.Fax & "' as FAX, "
    cSQL = cSQL & "'" & Text12.Text & "' as ANUMMER, "
    cSQL = cSQL & "'" & gFirma.FirmaMail & "' as FIRMAMAIL "
    gdApp.Execute cSQL, dbFailOnError

    loeschapp "DDFBEST"
    CreateTable "DDFBEST", gdApp
    
    cSQL = "Insert into DDFBEST Select "
    cSQL = cSQL & " artnr "
    cSQL = cSQL & ", bezeich "
    cSQL = cSQL & ", libesnr "
    cSQL = cSQL & ", lpz"
    cSQL = cSQL & ", linr as blinr"
    cSQL = cSQL & ", zubest as bestvor "
    cSQL = cSQL & ", LEKPR "
    cSQL = cSQL & ", KVKPR1 "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & " from manub where zubest > 0 "
    cSQL = cSQL & " order by linr, LPZ,BEZEICH"
    gdApp.Execute cSQL, dbFailOnError

    'lisrt holen
    loeschapp "lisrt"
    
    cSQL = "Select lisrt.* into lisrt in '" & cPfad & "' from lisrt "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update KDFBEST inner join lisrt on KDFBEST.linr = lisrt.linr "
    cSQL = cSQL & " set KDFBEST.LIEF_FAX = lisrt.fax"
    cSQL = cSQL & ", KDFBEST.KUNDNR = lisrt.KUNDNR"
    cSQL = cSQL & ", KDFBEST.KTEXT = lisrt.KTEXT"
    cSQL = cSQL & ", KDFBEST.Liefbez = lisrt.Liefbez"
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Update DDFBEST inner join lisrt on ddfbest.Blinr = lisrt.linr "
    cSQL = cSQL & "set DDFBEST.LINBEZ = lisrt.liefbez "
    gdApp.Execute cSQL, dbFailOnError

    Screen.MousePointer = 0
    
    If bmitDruck Then
        reportbildschirmApp "WKL004ac", "aWKL43a"
    End If
    
    anzeigeNew "normal", "", Label9

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SchreibeBest"
        Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Sub
Private Sub BestLeer()
    On Error GoTo LOKAL_ERROR
    
    Dim j As Integer
    
    Screen.MousePointer = 11

    With MSFlexGrid2
        .Redraw = False
        For j = 2 To .Rows - 1
            .Row = j
            .Col = SpaltennummerBEstellen
            .Text = "0"
        Next j
        .Redraw = True
    End With
    
    Screen.MousePointer = 0
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BestLeer"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub BestDelNull()
    On Error GoTo LOKAL_ERROR

    Dim cSQL   As String

    Screen.MousePointer = 11

    cSQL = "Delete from MANUB where zubest <= 0 "
    gdApp.Execute cSQL, dbFailOnError
       
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BestDelNull"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub BestdEL()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11
    
    loeschNEW "Manub", gdApp
    CreateTable "MANUB", gdApp
    
    cSQL = "Create Index ARTNR on MANUB (ARTNR)"
    gdApp.Execute cSQL, dbFailOnError
    
    zeigemanuB
    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BestdEL"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    
    Screen.MousePointer = 11
    
    Modul6.Farbform Me, lblUeberschrift
    PositionierenWKL47
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    
    If Not NewTableSuchenDBKombi("MANUB", gdApp) Then
        CreateTable "MANUB", gdApp
        
        cSQL = "Create Index ARTNR on MANUB (ARTNR)"
        gdApp.Execute cSQL, dbFailOnError
    End If
    
    cSQL = "Update MANUB set AWM = '0' where AWM = '99' "  'löschen der "Artikel anfügen" Farbe
    gdApp.Execute cSQL, dbFailOnError
    
    zeigemanuB
    
    cboswfuell
    
    If NewTableSuchenDBKombi("E47B", gdApp) = False Then
        
        CreateTable "E47B", gdApp
        
    End If
    
    voreinstellungladenE47B
    
    Text12.Text = ermMaxAufnr
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichernE47B()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim bo0             As Integer
    Dim bo1             As Integer
    Dim bo2             As Integer
    
    loeschNEW "E47B", gdApp
    CreateTable "E47B", gdApp
    
    bo1 = Option1(0).Value
    bo2 = Option1(1).Value

    sSQL = "Insert into E47B ( bo1,bo2) "
    sSQL = sSQL & " values (" & bo1 & "," & bo2 & ")"
    gdApp.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichernE47B"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladenE47B()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    
    Set rs = gdApp.OpenRecordset("E47B")
    If Not rs.EOF Then
    
        Option1(0).Value = rs!bo1
        Option1(1).Value = rs!bo2
        
    End If
    rs.Close: Set rs = Nothing
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladenE47B"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Bestellwert() As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rs As Recordset
    
    Screen.MousePointer = 11
    Bestellwert = ""
    
    cSQL = "Select sum(zubest * LEKPR) as bestwert from Manub "
    Set rs = gdApp.OpenRecordset(cSQL)
    
    If Not rs.EOF Then
        If Not IsNull(rs!Bestwert) Then
            Bestellwert = rs!Bestwert
        Else
            Bestellwert = ""
        End If
    End If
    
    rs.Close: Set rs = Nothing
    
    Screen.MousePointer = 0
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Bestellwert"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub cboswfuell()
    On Error GoTo LOKAL_ERROR
    
    With Combo1
        .Clear
        .Visible = True
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
        .AddItem "9"
        .AddItem "10"
        
        .Text = "1"
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboswfuell"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherezubest()
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim cART As String
    Dim lBest As Long
    Dim i As Integer
    Dim j As Integer
    Dim lrow As Long
    
    Screen.MousePointer = 11
    
    
    With MSFlexGrid2
    
    .Redraw = False
    
    For j = 2 To .Rows - 1
        .Row = j
        .Col = SpaltennummerArtnr
        cART = .Text
        If cART <> "" Then
            .Col = SpaltennummerBEstellen
            lBest = CLng(Val(.Text))
            
            cSQL = "Update MANUB set zubest = " & lBest & " where Artnr  = " & cART
            gdApp.Execute cSQL, dbFailOnError
        End If
    Next j
    .Redraw = True
    
    End With
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherezubest"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigemanuB()
On Error GoTo LOKAL_ERROR

    Dim j As Integer
    
    Screen.MousePointer = 11
    
    Tabcheck "BESTMAN"
    FormatGridOverTablay "BESTMAN"

    With MSFlexGrid2
    
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 25
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
            aBreite(j) = TextWidth(.TextMatrix(0, j)) ' * 1.8
        Next j
    End With
    
    

    'Grid fuellen
    anzeigeNew "normal", "Die Daten werden angezeigt...", Label9
    
    GridFuellen
    
    
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigemanuB"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub GridFuellen()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim counter     As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim sSQL        As String
    
    sSQL = "Select * from MANUB order by Linr,lpz,Bezeich"
    Set rsrs = gdApp.OpenRecordset(sSQL)
    
    With MSFlexGrid2
        .Redraw = False
        
        lrow = 1
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                lrow = lrow + 1
                .Rows = lrow + 1
                .Col = 0
                
                For i = 0 To byAnzahlSpalten - 1
                    .Row = 0
                    .Col = i
                    
                    If sSpaltenname(i) = .Text Then
                        
                        Select Case sSpaltenname(i)
                            Case Is = "Listen - EK", "Listen - VK", "Kassen - VK"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "####0.00")
                                
                            Case Is = "Preisschutz", "Geführt"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "N"
                                End If
                                .Row = lrow
                                .Text = sWert
    
                            Case Is = "Rabatt"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "J"
                                End If
                                .Row = lrow
                                .Text = sWert
                             
                            Case Is = "MinBest"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = sWert
                            
                            Case Else
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = ""
                                End If
                                .Row = lrow
                                .Text = sWert
                        End Select
                        
                        If TextWidth(.TextMatrix(lrow, i)) > aBreite(i) Then
                            aBreite(i) = TextWidth(.TextMatrix(lrow, i))
                        End If
                        
                    End If
                Next i
    
                rsrs.MoveNext
            Loop
        End If
        
        For i = 0 To byAnzahlSpalten - 1
            .Col = i
            .ColWidth(i) = aBreite(i) * 1.6
        Next i
            
        rsrs.Close: Set rsrs = Nothing
    
        If byAnzahlSpalten < 2 Then
        
        Else
            .FixedCols = 1
        End If
        
        .RowHeight(1) = 0
        
        lrow = lrow - 1
        
        If lrow = 0 Then
            anzeigeNew "normal", "Neuer Bestellvorschlag - Artikel anfügen", Label9
        ElseIf lrow = 1 Then
            anzeigeNew "Normal", "Ihr Bestellvorschlag enthält einen Artikel.", Label9
        Else
            anzeigeNew "Normal", "Ihr Bestellvorschlag enthält " & lrow & " Artikel.", Label9
        End If
        
        .Redraw = True
        .Visible = True
    End With
    
    Label6(6).Caption = Format(Bestellwert, "##0.00" & Space(1) & "EUR")
    Label6(6).Refresh
    
    Tabellenbreiteanpassen MSFlexGrid2, 1.25 * gdTabfak
    
    ermittlespalten
    
    FaerbenGrid MSFlexGrid2, CInt(SpaltennummerAWM), CInt(SpaltennummerArtnr)
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Gridfuellen"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub FaerbenGrid(grid As MSFlexGrid, iawmSpalte As Integer, Izufarbspalte As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim j As Integer
    
    Dim cAWM                As String
    
    With grid
        .Redraw = False
    
        For i = 0 To .Rows - 1
            .Row = i
            For j = 0 To .Cols - 1
            .Col = j
                If .Col = iawmSpalte Then
                    cAWM = .TextMatrix(i, j)
                    If cAWM = "" Then cAWM = "0"
                    FaerbenFlex cAWM, grid, Izufarbspalte, i
                End If
                
            Next j
        Next i
        .Redraw = True
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "FaerbenGrid"
    Fehler.gsFehlertext = "Beim Faerben eines Grids ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSFlexGrid, siEigFak As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim siFak       As Single
    Dim bBreit()    As Integer
    Dim i           As Integer
    Dim j           As Integer
    
    With gridx
        ReDim bBreit(.Cols - 1)
        
        For j = 0 To .Rows - 1
            For i = 0 To .Cols - 1
                If TextWidth(.TextMatrix(j, i)) > bBreit(i) Then
                    bBreit(i) = TextWidth(.TextMatrix(j, i))
                End If
            Next i
        Next j
        
        Select Case Screen.Height
            Case Is > 15000
                siFak = 1.5
            Case Is > 12000
                siFak = 1.4
            Case Is > 11000
                siFak = 1.2
            Case Is > 10000
                siFak = 1.1
            Case Is > 8000
                siFak = 1#
        End Select
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = bBreit(i) * siFak * siEigFak
        Next i
    
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenWKL47()
On Error GoTo LOKAL_ERROR
    
    With Frame1
        .Height = 7575
        .Left = 120
        .Top = 960
        .Width = 11655
        .Visible = False
        .BorderStyle = 0
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL47"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case (sSpaltenbez(i))
            Case Is = "Artnr"
                SpaltennummerArtnr = i
            Case Is = "ZuBest"
                SpaltennummerBEstellen = i
            Case Is = "AWM"
                SpaltennummerAWM = i
            Case Is = "Bezeich"
                SpaltennummerBEZEICH = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid2_DblClick()
    On Error GoTo LOKAL_ERROR
    
    sortierenGrid MSFlexGrid2
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim lcol     As Long
    Dim lrow     As Long
    Dim cZeichen As String
    Dim cFeld    As String
    
    lcol = MSFlexGrid2.Col
    lrow = MSFlexGrid2.Row

    If MSFlexGrid2.Col = SpaltennummerBEstellen Then
    
        cZeichen = Chr$(KeyAscii)
        cZeichen = UCase$(cZeichen)
        KeyAscii = Asc(cZeichen)
    
        cFeld = MSFlexGrid2.Text
        
        
        Select Case KeyAscii
            Case Is = 8
                If Len(cFeld) > 0 Then
                    cFeld = Left$(cFeld, Len(cFeld) - 1)
                End If
            Case 48 To 57
                cFeld = cFeld & Chr$(KeyAscii)
            
            Case Else
                cFeld = cFeld
                
        End Select
    
        MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, MSFlexGrid2.Col) = cFeld
        MSFlexGrid2.Refresh
        

        MSFlexGrid2.Col = lcol
        MSFlexGrid2.Row = lrow
        MSFlexGrid2.SetFocus
    End If

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim lcol As Long
    
    
    lrow = MSFlexGrid2.Row
    lcol = MSFlexGrid2.Col
    

    If MSFlexGrid2.Col = SpaltennummerBEstellen Then
        If iKeypress = 0 And KeyCode <> vbKeyBack Then
            MSFlexGrid2.Row = lrow
            MSFlexGrid2.Col = lcol
            MSFlexGrid2.Text = ""
        End If
        iKeypress = iKeypress + 1
    End If
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid2_LeaveCell()
On Error GoTo LOKAL_ERROR
    
    iKeypress = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_LeaveCell"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text2_LostFocus()
On Error GoTo LOKAL_ERROR
    
    Text2.BackColor = vbWhite

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR

    Dim cValid As String
    
    cValid = "1234567890" & Chr$(8)
        
    If InStr(cValid, Chr$(KeyAscii)) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(Chr$(KeyAscii))
    End If
        

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text6_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        cmdAnfuegen_Click 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text6_Keyup"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR

    Dim cValid As String
    
    cValid = "1234567890" & Chr$(8)
        
    If InStr(cValid, Chr$(KeyAscii)) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(Chr$(KeyAscii))
    End If
        

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text7_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        cmdAnfuegen_Click 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text8_Keyup"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        cmdAnfuegen_Click 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text7_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        cmdAnfuegen_Click 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_Keyup"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text6_LostFocus()
On Error GoTo LOKAL_ERROR
    
    Text6.BackColor = vbWhite

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text6_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Text7_LostFocus()
On Error GoTo LOKAL_ERROR
    
    Text7.BackColor = vbWhite

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text7_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Text8_LostFocus()
On Error GoTo LOKAL_ERROR
    
    Text8.BackColor = vbWhite

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text8_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Text6_GotFocus()
On Error GoTo LOKAL_ERROR
    
    Text6.BackColor = glSelBack1
    Text6.SelStart = 0
    Text6.SelLength = Len(Text6.Text)
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text6_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text7_GotFocus()
On Error GoTo LOKAL_ERROR
    
    Text7.BackColor = glSelBack1
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7.Text)
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text7_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text8_GotFocus()
On Error GoTo LOKAL_ERROR
    
    Text8.BackColor = glSelBack1
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.Text)
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text8_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_GotFocus()
On Error GoTo LOKAL_ERROR
    
    Text2.BackColor = glSelBack1
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bestellung manuell ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtStatus_Change()
On Error GoTo LOKAL_ERROR
    
    Dim nProz As Long
  
    nProz = Val(txtStatus.Text)
    ShowProgress picprogress, nProz, 0, 100, True
    picprogress.Refresh

Exit Sub

LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtstatus_Change"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
