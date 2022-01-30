VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWK21f 
   Caption         =   "Zusammenfassung von Tagesabschlüssen"
   ClientHeight    =   8595
   ClientLeft      =   1155
   ClientTop       =   1530
   ClientWidth     =   11880
   Icon            =   "frmWK21f.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   1920
      TabIndex        =   132
      Top             =   1560
      Visible         =   0   'False
      Width           =   5295
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   135
         Top             =   720
         Width           =   4935
      End
      Begin sevCommand3.Command Command5 
         Height          =   345
         Left            =   4680
         TabIndex        =   134
         Top             =   120
         Width           =   345
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         Caption         =   "x"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label0 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Wechselgeld der anderen Kassenschubladen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   240
         TabIndex        =   133
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Caption         =   "Zusammengefaßte Tagesabschlüsse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   6855
      Left            =   120
      TabIndex        =   23
      Top             =   1680
      Width           =   11535
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kassen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   56
         Left            =   3000
         TabIndex        =   145
         ToolTipText     =   "Wechselgeld der nicht abgerechneten Kassen"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "offenes WG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   55
         Left            =   2640
         TabIndex        =   144
         ToolTipText     =   "Wechselgeld der nicht abgerechneten Kassen"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kreditkarte, nicht umsr. Verkäufe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   54
         Left            =   6600
         TabIndex        =   143
         Top             =   4440
         Width           =   3855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   51
         Left            =   9600
         TabIndex        =   142
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "offenes WG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   53
         Left            =   2640
         TabIndex        =   131
         ToolTipText     =   "Wechselgeld in anderen Kassenschubladen"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kreditkarte,insgesamt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   52
         Left            =   6600
         TabIndex        =   127
         Top             =   6720
         Width           =   3855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   52
         Left            =   10080
         TabIndex        =   126
         Top             =   6720
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikeltrabatt Anz.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   51
         Left            =   6720
         TabIndex        =   125
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   53
         Left            =   9480
         TabIndex        =   124
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Dukaten"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   46
         Left            =   360
         TabIndex        =   123
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   50
         Left            =   3840
         TabIndex        =   122
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   49
         Left            =   10080
         TabIndex        =   21
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kassendifferenz, aufgelaufene"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   6600
         TabIndex        =   22
         Top             =   6480
         Width           =   3855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   48
         Left            =   9600
         TabIndex        =   121
         Top             =   6240
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kassendifferenz jetzt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   49
         Left            =   6600
         TabIndex        =   120
         Top             =   6240
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   47
         Left            =   3720
         TabIndex        =   119
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   46
         Left            =   3720
         TabIndex        =   118
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Wechselgeld"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   48
         Left            =   720
         TabIndex        =   117
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "- Abschöpfung"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   47
         Left            =   720
         TabIndex        =   116
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Gutschein über Gutschein"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   115
         Top             =   6840
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   45
         Left            =   3720
         TabIndex        =   114
         Top             =   6840
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   44
         Left            =   3720
         TabIndex        =   113
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "nicht umsatzrelevante Verkäufe:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   240
         TabIndex        =   112
         Top             =   5160
         Width           =   3615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   43
         Left            =   3840
         TabIndex        =   111
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Tilgung in Bar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   44
         Left            =   720
         TabIndex        =   110
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   42
         Left            =   3720
         TabIndex        =   109
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Gutschein über Bar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   43
         Left            =   720
         TabIndex        =   108
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   41
         Left            =   3720
         TabIndex        =   107
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "- Gutscheinauszahlungen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   42
         Left            =   720
         TabIndex        =   106
         Top             =   4200
         Width           =   3135
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   40
         Left            =   9480
         TabIndex        =   105
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   39
         Left            =   9480
         TabIndex        =   104
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   38
         Left            =   9480
         TabIndex        =   103
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   37
         Left            =   9480
         TabIndex        =   102
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   36
         Left            =   9480
         TabIndex        =   101
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Umsatz ohne MWSt:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   41
         Left            =   6720
         TabIndex        =   100
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Betrag erm MWSt:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   40
         Left            =   6720
         TabIndex        =   99
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Umsatz erm. MWSt:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   6720
         TabIndex        =   98
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Betrag volle MWSt:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   38
         Left            =   6720
         TabIndex        =   97
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Umsatz volle MWSt:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   6720
         TabIndex        =   96
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   35
         Left            =   3720
         TabIndex        =   95
         Top             =   6600
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Gutschein über Lastschrift"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   480
         TabIndex        =   94
         Top             =   6600
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   34
         Left            =   3840
         TabIndex        =   93
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Lastschrift"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   360
         TabIndex        =   92
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   33
         Left            =   3720
         TabIndex        =   91
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "daraus Rest-Gutscheine:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   480
         TabIndex        =   90
         Top             =   4920
         Width           =   3135
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   32
         Left            =   3720
         TabIndex        =   89
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "eingereichte Gutscheine:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   240
         TabIndex        =   88
         Top             =   4680
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kredit-Tilgungen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   6600
         TabIndex        =   87
         Top             =   4800
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   27
         Left            =   9480
         TabIndex        =   86
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Tilgung in Bar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   6720
         TabIndex        =   85
         Top             =   5040
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Tilgung per Scheck"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   6720
         TabIndex        =   84
         Top             =   5280
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Tilgung per Gutschein"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   6720
         TabIndex        =   83
         Top             =   5520
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Tilgung per Karte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   6720
         TabIndex        =   82
         Top             =   5760
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   28
         Left            =   9600
         TabIndex        =   81
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   29
         Left            =   9600
         TabIndex        =   80
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   30
         Left            =   9600
         TabIndex        =   79
         Top             =   5520
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   31
         Left            =   9600
         TabIndex        =   78
         Top             =   5760
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   26
         Left            =   3720
         TabIndex        =   77
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Schecks:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   480
         TabIndex        =   76
         Top             =   4440
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   25
         Left            =   3720
         TabIndex        =   75
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Bar-Verkäufe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   720
         TabIndex        =   74
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   24
         Left            =   3720
         TabIndex        =   73
         Top             =   6360
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   23
         Left            =   3720
         TabIndex        =   72
         Top             =   6120
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   22
         Left            =   3720
         TabIndex        =   71
         Top             =   5880
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   21
         Left            =   3720
         TabIndex        =   70
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Gutschein über Karte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   480
         TabIndex        =   69
         Top             =   6360
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Gutschein über Kredite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   480
         TabIndex        =   68
         Top             =   6120
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Gutschein über Scheck"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   480
         TabIndex        =   67
         Top             =   5880
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Gutschein über Bar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   480
         TabIndex        =   66
         Top             =   5640
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   20
         Left            =   3840
         TabIndex        =   65
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Umsatz aus Gutscheinen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   360
         TabIndex        =   64
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   19
         Left            =   9480
         TabIndex        =   63
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   18
         Left            =   9480
         TabIndex        =   62
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   17
         Left            =   9480
         TabIndex        =   61
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   16
         Left            =   9480
         TabIndex        =   60
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   15
         Left            =   9480
         TabIndex        =   59
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   14
         Left            =   9480
         TabIndex        =   58
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   13
         Left            =   9480
         TabIndex        =   57
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   12
         Left            =   9480
         TabIndex        =   56
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   11
         Left            =   9480
         TabIndex        =   55
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   10
         Left            =   9480
         TabIndex        =   54
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   9
         Left            =   3720
         TabIndex        =   53
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   3720
         TabIndex        =   52
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   3720
         TabIndex        =   51
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   3720
         TabIndex        =   50
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   3720
         TabIndex        =   49
         Top             =   5400
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   48
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   47
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   46
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   45
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 DEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   44
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Storno Anz.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   6720
         TabIndex        =   43
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Stornosumme:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   6720
         TabIndex        =   42
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Gesamtrabatt Anz.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   6720
         TabIndex        =   41
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Gesamtrabatt:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   6720
         TabIndex        =   40
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikelrabatt:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   6720
         TabIndex        =   39
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Schublade geöffnet:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   6720
         TabIndex        =   38
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Sonderpreise Anz.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   6720
         TabIndex        =   37
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Sonderpreise:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   6720
         TabIndex        =   36
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kundenschnitt:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   6720
         TabIndex        =   35
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kundenzahl:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   6600
         TabIndex        =   34
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Bargeld:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   33
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kassensoll:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   32
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "- Auszahlungen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   720
         TabIndex        =   31
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Einzahlungen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   720
         TabIndex        =   30
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Gutschein-Verkäufe:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   29
         Top             =   5400
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Kreditkarten"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   28
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Kredite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   27
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Schecks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   26
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "+ Bar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Umsatz gesamt:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Caption         =   "Zusammenführen von Tagesabschlüssen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1575
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   12015
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kasse 8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   6840
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kasse 7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   6840
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kasse 6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   6840
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kasse 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6840
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kasse 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5640
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kasse 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5640
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kasse 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kasse 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'Kein
         Enabled         =   0   'False
         Height          =   975
         Left            =   2760
         TabIndex        =   16
         Top             =   480
         Width           =   2775
         Begin MSComCtl2.DTPicker Text1 
            Height          =   375
            Index           =   0
            Left            =   600
            TabIndex        =   19
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OLEDropMode     =   1
            CalendarTitleBackColor=   12615680
            Format          =   112263169
            UpDown          =   -1  'True
            CurrentDate     =   38425
         End
         Begin MSComCtl2.DTPicker Text1 
            Height          =   375
            Index           =   1
            Left            =   600
            TabIndex        =   20
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarTitleBackColor=   12615680
            Format          =   112263169
            UpDown          =   -1  'True
            CurrentDate     =   38425
         End
         Begin sevCommand3.Command Command0 
            Height          =   360
            Index           =   0
            Left            =   1920
            TabIndex        =   129
            ToolTipText     =   "Kalender"
            Top             =   0
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
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
            ToolTip         =   "Wählen Sie hier das Datum aus."
            ToolTipTitle    =   "Kalender"
            ButtonStyle     =   2
            Caption         =   ""
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   360
            Index           =   1
            Left            =   1920
            TabIndex        =   130
            ToolTipText     =   "Kalender"
            Top             =   480
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
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
            ToolTip         =   "Wählen Sie hier das Datum aus."
            ToolTipTitle    =   "Kalender"
            ButtonStyle     =   2
            Caption         =   ""
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin VB.Label Label0 
            BackColor       =   &H00C0C000&
            Caption         =   " bis"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label0 
            BackColor       =   &H00C0C000&
            Caption         =   "von"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            TabIndex        =   17
            Top             =   120
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         Height          =   855
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2055
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   375
            Index           =   1
            Left            =   1080
            TabIndex        =   1
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label0 
            BackColor       =   &H00C0C000&
            Caption         =   "Jahr"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   15
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label0 
            BackColor       =   &H00C0C000&
            Caption         =   "Monat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Zeitraum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   11
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Monat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
      Begin sevCommand3.Command Command98 
         Height          =   315
         Left            =   4680
         TabIndex        =   128
         Top             =   120
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
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
         Picture         =   "frmWK21f.frx":0442
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand1 
         Height          =   375
         Index           =   0
         Left            =   9720
         TabIndex        =   136
         Top             =   200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
            Size            =   14.25
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
      Begin sevCommand3.Command SSCommand1 
         Height          =   375
         Index           =   2
         Left            =   9720
         TabIndex        =   137
         Top             =   970
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
            Size            =   14.25
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
      Begin sevCommand3.Command SSCommand1 
         Height          =   375
         Index           =   1
         Left            =   9720
         TabIndex        =   138
         Top             =   580
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
            Size            =   14.25
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
         Caption         =   "Z-Bon drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand1 
         Height          =   375
         Index           =   3
         Left            =   8480
         TabIndex        =   139
         Top             =   200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
            Size            =   14.25
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
         Caption         =   "AGN"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand1 
         Height          =   375
         Index           =   4
         Left            =   8480
         TabIndex        =   140
         Top             =   580
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
            Size            =   14.25
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
         Caption         =   "Zeitung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand1 
         Height          =   375
         Index           =   5
         Left            =   8480
         TabIndex        =   141
         Top             =   970
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
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
            Size            =   14.25
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
         Caption         =   "all"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmWK21f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command5_Click()
On Error GoTo LOKAL_ERROR
    
    Frame5.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub

Private Sub Command98_Click()
On Error GoTo LOKAL_ERROR
    
    gsZSpalte = ""
    gstab = "ZBON"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command98_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "Agtemp", gdBase
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
Private Function fnPruefeEingabeDialogWK21f() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld As String
    Dim dWert As Double
    Dim lVon As Long
    Dim lBis As Long
    Dim lcount As Long
    Dim bgefunden As Boolean
    
    If Option1(0).value = True Then
        cFeld = MaskEdBox1(0).Text
        dWert = Val(cFeld)
        If dWert < 1 Or dWert > 12 Then
            fnPruefeEingabeDialogWK21f = 1
            Exit Function
        End If
        cFeld = MaskEdBox1(1).Text
        dWert = Val(cFeld)
        If dWert < 1980 Or dWert > 2100 Then
            fnPruefeEingabeDialogWK21f = 2
            Exit Function
        End If
    End If
    
    If Option1(1).value = True Then
        cFeld = Text1(0).value
        If Not IsDate(cFeld) Then
            fnPruefeEingabeDialogWK21f = 3
            Exit Function
        End If
        lVon = DateValue(cFeld)
        
        cFeld = Text1(1).value
        If Not IsDate(cFeld) Then
            fnPruefeEingabeDialogWK21f = 4
            Exit Function
        End If
        lBis = DateValue(cFeld)
    
        If lVon > lBis Then
            fnPruefeEingabeDialogWK21f = 5
            Exit Function
        End If
    End If
    
    bgefunden = False
    For lcount = 0 To 7
        If Check1(lcount).value = vbChecked Then
            bgefunden = True
            Exit For
        End If
    Next lcount
    
    If Not bgefunden Then
        fnPruefeEingabeDialogWK21f = 6
        Exit Function
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeDialogWK21f"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub HoleKassenAbschlussDatenWK21f()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim iFileNr As Integer
    Dim lWert1 As Long
    Dim lWert2 As Long
    Dim rsrs1 As Recordset
    Dim rsRs2 As Recordset
    Dim lcount As Long
    Dim tdQ As TableDef
    Dim tdZ As TableDef
    Dim tdP As TableDef
    Dim cFeldNameQ As String
    Dim cFeldNameZ As String
    Dim bgefunden As Boolean
    Dim bAus As Boolean
    Dim bOr As Boolean
    Dim sSQL As String
    
    loeschNEW "ZZKopf", gdBase
    CreateTable "ZZKOPF", gdBase
    
    sSQL = " Insert into zzkopf (K0,K1,K2,K3,K4,K5,K6,K7) "
    sSQL = sSQL & " Values (True,True,True,True,True,True,True,True)"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "AFCSTATS", gdBase
    
    'check

    If Not SpalteInTabellegefundenNEW("AFCSTATP", "GUTSCHGUTSCH", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "GUTSCHGUTSCH", "double", gdBase

        sSQL = "Update AFCSTATP set GUTSCHGUTSCH = 0 "
        gdBase.Execute sSQL, dbFailOnError
    End If

    If Not SpalteInTabellegefundenNEW("AFCSTATP", "ABSCHOPF", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "ABSCHOPF", "double", gdBase

        sSQL = "Update AFCSTATP set ABSCHOPF = 0 "
        gdBase.Execute sSQL, dbFailOnError
    End If

    If Not SpalteInTabellegefundenNEW("AFCSTATP", "KDIFF", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "KDIFF", "double", gdBase

        sSQL = "Update AFCSTATP set KDIFF = 0 "
        gdBase.Execute sSQL, dbFailOnError
    End If

    If Not SpalteInTabellegefundenNEW("AFCSTATP", "TDIFF", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "TDIFF", "double", gdBase

        sSQL = "Update AFCSTATP set TDIFF = 0 "
        gdBase.Execute sSQL, dbFailOnError
    End If

    If Not SpalteInTabellegefundenNEW("AFCSTATP", "DUKA", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "DUKA", "double", gdBase

        sSQL = "Update AFCSTATP set DUKA = 0 "
        gdBase.Execute sSQL, dbFailOnError
    End If

    If Not SpalteInTabellegefundenNEW("AFCSTATP", "WECHSEL", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "WECHSEL", "double", gdBase

        sSQL = "Update AFCSTATP set WECHSEL = 0 "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If Not SpalteInTabellegefundenNEW("AFCSTATP", "NUMSKARTE", gdBase) Then
        SpalteAnfuegenNEW "AFCSTATP", "NUMSKARTE", "double", gdBase

        sSQL = "Update AFCSTATP set NUMSKARTE = 0 "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    'check ende
    cSQL = "Select SUM(UMS_BAR) as UMS_BAR1"
    cSQL = cSQL & ", SUM(UMS_KRED) as UMS_KRED1"
    cSQL = cSQL & ", SUM(UMS_SCHECK) as UMS_SCHECK1"
    cSQL = cSQL & ", SUM(UMS_KARTE) as UMS_KARTE1"
    cSQL = cSQL & ", SUM(UMS_LAST) as UMS_LAST1"
    cSQL = cSQL & ", SUM(SPREIS_ANZ) as SPREIS_ANZ1"
    cSQL = cSQL & ", SUM(SPREIS_GES) as SPREIS_GES1"
    cSQL = cSQL & ", SUM(ANZSCHECK) as ANZSCHECK1"
    cSQL = cSQL & ", SUM(KUNDENZAHL) as KUNDENZAHL1"
    cSQL = cSQL & ", SUM(GELDFACH) as GELDFACH1"
    cSQL = cSQL & ", SUM(ARTRAB_ANZ) as ARTRAB_ANZ1"
    cSQL = cSQL & ", SUM(ARTRAB_GES) as ARTRAB_GES1"
    cSQL = cSQL & ", SUM(GESRAB_ANZ) as GESRAB_ANZ1"
    cSQL = cSQL & ", SUM(GESRAB_GES) as GESRAB_GES1"
    cSQL = cSQL & ", SUM(STORNO_ANZ) as STORNO_ANZ1"
    cSQL = cSQL & ", SUM(STORNO_GES) as STORNO_GES1"
    cSQL = cSQL & ", SUM(EINZAHLUNG) as EINZAHLUNG1"
    cSQL = cSQL & ", SUM(AUSZAHLUNG) as AUSZAHLUNG1"
    cSQL = cSQL & ", SUM(GUTSCHEIN) as GUTSCHEIN1"
    cSQL = cSQL & ", SUM(BELEGNR) as BELEGNR1"
    cSQL = cSQL & ", SUM(ZHLGGUTSCH) as ZHLGGUTSCH1"
    cSQL = cSQL & ", SUM(GUTSCHBAR) as GUTSCHBAR1"
    cSQL = cSQL & ", SUM(GUTSCHSCH) as GUTSCHSCH1"
    cSQL = cSQL & ", SUM(GUTSCHKRE) as GUTSCHKRE1"
    cSQL = cSQL & ", SUM(GUTSCHKAR) as GUTSCHKAR1"
    cSQL = cSQL & ", SUM(GUTSCHLAST) as GUTSCHLAST1"
    cSQL = cSQL & ", SUM(BARVERKAUF) as BARVERKAUF1"
    cSQL = cSQL & ", SUM(SCHVERKAUF) as SCHVERKAUF1"
    cSQL = cSQL & ", SUM(TILGBAR) as TILGBAR1"
    cSQL = cSQL & ", SUM(TILGSCH) as TILGSCH1"
    cSQL = cSQL & ", SUM(TILGGUT) as TILGGUT1"
    cSQL = cSQL & ", SUM(TILGKAR) as TILGKAR1"
    cSQL = cSQL & ", SUM(EINRGUTSCH) as EINRGUTSCH1"
    cSQL = cSQL & ", SUM(RESTGUTSCH) as RESTGUTSCH1"
    cSQL = cSQL & ", SUM(GUTSCHGUTSCH) as GUTSCHGUTSCH1"
    cSQL = cSQL & ", SUM(ABSCHOPF) as ABSCHOPF1"
    cSQL = cSQL & ", SUM(KDIFF) as KDIFF1"
    cSQL = cSQL & ", SUM(TDIFF) as TDIFF1"
    cSQL = cSQL & ", SUM(DUKA) as DUKA1"
    cSQL = cSQL & ", SUM(WECHSEL) as WECHSEL1"
    cSQL = cSQL & ", SUM(AUSZGUTSCH) as AUSZGUTSCH1"
    cSQL = cSQL & ", SUM(NUMSKARTE) as NUMSKARTE1"
    cSQL = cSQL & " into AFCSTATS from AFCSTATP where "
    
    If Option1(0).value = True Then
        lWert1 = Val(MaskEdBox1(0).Text)
        lWert2 = Val(MaskEdBox1(1).Text)
        cSQL = cSQL & "YEAR(ADATE) = " & Trim$(Str$(lWert2)) & " "
        cSQL = cSQL & "and MONTH(ADATE) = " & Trim$(Str$(lWert1)) & " "
        
        sSQL = "Update zzKopf set Monat = " & lWert1
        sSQL = sSQL & " , Jahr = " & lWert2
        gdBase.Execute sSQL, dbFailOnError
        
        'Speicher hier Monat/Jahr
    Else
        lWert1 = Text1(0).value '     DateValue(MaskEdBox2(0).Text)
        lWert2 = Text1(1).value 'DateValue(MaskEdBox2(1).Text)
        cSQL = cSQL & "ADATE >= " & Trim$(Str$(lWert1)) & " "
        cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lWert2)) & " "
        
        sSQL = "Update zzKopf set Datvon = " & lWert1
        sSQL = sSQL & " , Datbis = " & lWert2
        gdBase.Execute sSQL, dbFailOnError
        'Speicher Zeitraum
    End If
        
    bAus = False
    For lcount = 0 To 7
        If Check1(lcount).value = vbUnchecked Then
            bAus = True
            Exit For
        End If
    Next lcount
    
    If bAus Then
        cSQL = cSQL & " and ( "
        bOr = False
        For lcount = 0 To 7
            If Check1(lcount).value = vbChecked Then
                If bOr Then
                    cSQL = cSQL & " or "
                End If
                cSQL = cSQL & "KASNUM = " & Trim$(Str$(lcount + 1)) & " "
                bOr = True
            Else
                sSQL = "Update zzKopf set K" & lcount & " = False"
                gdBase.Execute sSQL, dbFailOnError
            End If
        Next lcount
        cSQL = cSQL & " ) "
    End If
    gdBase.Execute cSQL, dbFailOnError
    
    LeseDatenWK21f
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleKassenAbschlussDatenWK21f"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub LeseDatenWK21f()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim dWert As Double
    Dim dSumme As Double
    Dim dUmsatz As Double
    Dim cMWSK As String
    Dim dEinzahlung As Double
    Dim dAuszahlung As Double
    Dim dAuszGutsch As Double
    Dim dBar As Double
    Dim dKunden As Double
    Dim dScheck As Double
    Dim dKredit As Double
    Dim dKarte As Double
    Dim dLast As Double
    Dim dUmsBar As Double
    Dim dUmsScheck As Double
    
    Dim dZhlgGutsch As Double
    
    Dim dKasse As Double
    Dim dKassenBargeld As Double
    Dim dKassenSchecks As Double
    Dim dSchVerkauf As Double
    Dim dBarVerkauf As Double
    
    Dim dGutschein As Double
    Dim dGutschBar As Double
    Dim dGutschSch As Double
    Dim dGutschKre As Double
    Dim dGutschKar As Double
    Dim dGutschLast As Double
    Dim dGutschGUTSCH As Double
    Dim dABSCHOPF As Double
    Dim dKDIFF As Double
    Dim dTDIFF As Double
    Dim dDUKA As Double
    Dim dWECHSEL As Double
    Dim dEinrGutsch As Double
    
    Dim dTilgung As Double
    Dim dTilgBar As Double
    Dim dTilgSch As Double
    Dim dTilgGut As Double
    Dim dTilgKar As Double
    
'    NUMSKARTE1
    Dim dNichtUmsReleKar As Double
    
    Dim sVon As String
    Dim sBis As String
    
    Dim lWert1 As Long
    Dim lWert2 As Long
    
    If Option1(0).value = True Then
        lWert1 = Val(MaskEdBox1(0).Text)
        lWert2 = Val(MaskEdBox1(1).Text)
        ctmp = "and YEAR(ADATE) = " & Trim$(Str$(lWert2)) & " and MONTH(ADATE) = " & Trim$(Str$(lWert1)) & " "
    Else
        lWert1 = Text1(0).value
        lWert2 = Text1(1).value
        ctmp = "and ADATE >= " & Trim$(Str$(lWert1)) & " and ADATE <= " & Trim$(Str$(lWert2)) & " "
    End If
    
    
    
    
    '*****Kassennummern
    Dim cSQLKass As String
    Dim bOr As Boolean
    Dim bAus As Boolean
    Dim lcount As Long
    
    bAus = False
    For lcount = 0 To 7
        If Check1(lcount).value = vbUnchecked Then
            bAus = True
            Exit For
        End If
    Next lcount
    
    If bAus Then
        cSQLKass = " where  ( "
        bOr = False
        For lcount = 0 To 7
            If Check1(lcount).value = vbChecked Then
                If bOr Then
                    cSQLKass = cSQLKass & " or "
                End If
                cSQLKass = cSQLKass & " KASNUM = " & Trim$(Str$(lcount + 1)) & " "
                bOr = True
            End If
        Next lcount
        cSQLKass = cSQLKass & " ) "
    End If
    '******Ende Kassennummern
    
    
'  666666"

    
    
    
    
'    loeschNEW "AfcTempo", gdBase
'    cSQL = "Select * into AFCTempo from Kassjour where UMS_OK = 'N' "
'    cSQL = cSQL & ctmp
'    gdBase.Execute cSQL, dbFailOnError
'
'    dWert = 0
'
'    cSQL = "Select SUM(PREIS) as UMSATZ from AFCTempo " & cSQLKass
'    Set rsrs = gdBase.OpenRecordset(cSQL)
'    If Not rsrs.EOF Then
'        rsrs.MoveFirst
'        If Not IsNull(rsrs!UMSATZ) Then
'            dWert = rsrs!UMSATZ
'        End If
'    End If
'    rsrs.Close: Set rsrs = Nothing
'    loeschNEW "AfcTempo", gdBase
'
'    Label3(44).Caption = Format$(dWert, "######0.00") & " " & gcWaehrung
    
    
    
    'ich suche das offene Wechselgeld aller anderen Kassen
    'Afcstat, Wechsel, Kasnum
    
    
    
    
    
    Dim cVon As String
    Dim cBis As String
    Dim cDatum As String
        
    
    
    
    Dim dOffenenesWechselgeldKasse1 As Double
    Dim dOffenenesWechselgeldKasse2 As Double
    Dim dOffenenesWechselgeldKasse3 As Double
    Dim dOffenenesWechselgeldKasse4 As Double
    Dim dOffenenesWechselgeldKasse5 As Double
    Dim dOffenenesWechselgeldKasse6 As Double
    Dim dOffenenesWechselgeldKasse7 As Double
    Dim dOffenenesWechselgeldKasse8 As Double
    
    dOffenenesWechselgeldKasse1 = 0
    dOffenenesWechselgeldKasse2 = 0
    dOffenenesWechselgeldKasse3 = 0
    dOffenenesWechselgeldKasse4 = 0
    dOffenenesWechselgeldKasse5 = 0
    dOffenenesWechselgeldKasse6 = 0
    dOffenenesWechselgeldKasse7 = 0
    dOffenenesWechselgeldKasse8 = 0
    
    Dim dWechselgeldKasseausKABUCH1 As Double
    Dim dWechselgeldKasseausKABUCH2 As Double
    Dim dWechselgeldKasseausKABUCH3 As Double
    Dim dWechselgeldKasseausKABUCH4 As Double
    Dim dWechselgeldKasseausKABUCH5 As Double
    Dim dWechselgeldKasseausKABUCH6 As Double
    Dim dWechselgeldKasseausKABUCH7 As Double
    Dim dWechselgeldKasseausKABUCH8 As Double
    
    dWechselgeldKasseausKABUCH1 = 0
    dWechselgeldKasseausKABUCH2 = 0
    dWechselgeldKasseausKABUCH3 = 0
    dWechselgeldKasseausKABUCH4 = 0
    dWechselgeldKasseausKABUCH5 = 0
    dWechselgeldKasseausKABUCH6 = 0
    dWechselgeldKasseausKABUCH7 = 0
    dWechselgeldKasseausKABUCH8 = 0
    
    Dim bgleicherTagundHeute As Boolean
    bgleicherTagundHeute = False
    
    Dim bZeitraumleztzerTagundHeuteundalleGew As Boolean
    bZeitraumleztzerTagundHeuteundalleGew = False
    
    Dim dOffenenesWechselgeldDEAKTIV As Double
    Dim dOffenenesWechselgeldAndere As Double
    
    dOffenenesWechselgeldAndere = 0
    dOffenenesWechselgeldDEAKTIV = 0
    
    If Option1(1).value = True Then
    
        lWert1 = CLng(Text1(0).value)
        lWert2 = CLng(Text1(1).value)
        
        If DateValue(Now) = lWert2 And lWert2 = lWert1 Then
            bgleicherTagundHeute = True
        End If
            
    ElseIf Option1(0).value = True Then
    
        'hier kann nur der MonatsEndetag gelten
        lWert1 = Val(MaskEdBox1(0).Text) 'Monat
        lWert2 = Val(MaskEdBox1(1).Text) 'Jahr
        
        Select Case lWert1
            Case 1, 3, 5, 7, 8, 10, 12
                cDatum = "31."
            
            Case 2
                If lWert2 = 2016 Then
                    cDatum = "29."
                ElseIf lWert2 = 2020 Then
                    cDatum = "29."
                ElseIf lWert2 = 2024 Then
                    cDatum = "29."
                ElseIf lWert2 = 2028 Then
                    cDatum = "29."
                Else
                    cDatum = "28."
                End If
            
            Case Else
                cDatum = "30."
        End Select
            
        cDatum = cDatum & lWert1 & "." & lWert2
        lWert2 = CLng(DateValue(cDatum))
        
    End If
    
    'Zeitraum?
    'letzter Tag = Heute?
    'alle gewählt?
    
    If lWert1 <> lWert2 Then
        If lWert2 = DateValue(Now) Then
        
            Dim bAlle As Boolean
            
            bAlle = True
            
            
            For lcount = 0 To 7
                If Check1(lcount).value = vbUnchecked Then
                    bAlle = False
                End If
            Next lcount
            
            If bAlle = True Then
                bZeitraumleztzerTagundHeuteundalleGew = True
            End If
        End If
    End If
    
    For lcount = 0 To 7

        cSQL = "select sum(wechsel) as sWechsel from Afcstat where kasnum = " & Trim$(Str$(lcount + 1)) & " "
        cSQL = cSQL & " and ADATE <= " & Trim$(Str$(lWert2)) & " "
    
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            If Not IsNull(rsrs!SWECHSEL) Then
                Select Case lcount
                    Case Is = 0: dOffenenesWechselgeldKasse1 = CDbl(rsrs!SWECHSEL)
                    Case Is = 1: dOffenenesWechselgeldKasse2 = CDbl(rsrs!SWECHSEL)
                    Case Is = 2: dOffenenesWechselgeldKasse3 = CDbl(rsrs!SWECHSEL)
                    Case Is = 3: dOffenenesWechselgeldKasse4 = CDbl(rsrs!SWECHSEL)
                    Case Is = 4: dOffenenesWechselgeldKasse5 = CDbl(rsrs!SWECHSEL)
                    Case Is = 5: dOffenenesWechselgeldKasse6 = CDbl(rsrs!SWECHSEL)
                    Case Is = 6: dOffenenesWechselgeldKasse7 = CDbl(rsrs!SWECHSEL)
                    Case Is = 7: dOffenenesWechselgeldKasse8 = CDbl(rsrs!SWECHSEL)
                End Select
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    
    Next lcount
    
    
    
    
    
    
    'Ende offenes Wechselgeld für alle
    
    
    
    
    
    'Ende Suche
    
    loeschNEW "temp_KABUCH", gdBase
    
    cSQL = "Select * into temp_KABUCH "
    cSQL = cSQL & " from KABUCH where BEZUMS = 'Wechselgeld' "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "temp_KABUCH_Wechsel", gdBase
    
    cSQL = "Select max(EURBAR) as euronenbar ,Max(autopos) as maxi,datum,kasnum into temp_KABUCH_Wechsel from temp_KABUCH group by datum,kasnum  "
    gdBase.Execute cSQL, dbFailOnError

    If Option1(0).value = True Then
        'zeitraum Monat Jahr
        lWert1 = Val(MaskEdBox1(0).Text)
        lWert2 = Val(MaskEdBox1(1).Text)
        
        cVon = Format("01." & Val(MaskEdBox1(0).Text) & "." & Val(MaskEdBox1(1).Text), "DD.MM.YYYY")
        
        Select Case Val(MaskEdBox1(0).Text)
            Case 1, 3, 5, 7, 8, 10, 12
                cBis = Format("31." & Val(MaskEdBox1(0).Text) & "." & Val(MaskEdBox1(1).Text), "DD.MM.YYYY")
            Case 2
                If Val(MaskEdBox1(1).Text) = 2016 Then
                    cBis = Format("29." & Val(MaskEdBox1(0).Text) & "." & Val(MaskEdBox1(1).Text), "DD.MM.YYYY")
                ElseIf Val(MaskEdBox1(1).Text) = 2012 Then
                    cBis = Format("29." & Val(MaskEdBox1(0).Text) & "." & Val(MaskEdBox1(1).Text), "DD.MM.YYYY")
                Else
                    cBis = Format("28." & Val(MaskEdBox1(0).Text) & "." & Val(MaskEdBox1(1).Text), "DD.MM.YYYY")
                End If
            Case Else
                cBis = Format("30." & Val(MaskEdBox1(0).Text) & "." & Val(MaskEdBox1(1).Text), "DD.MM.YYYY")
        End Select
        
        lWert1 = CLng(DateValue(cVon))
        lWert2 = CLng(DateValue(cBis))
    Else
        lWert1 = CLng(Text1(0).value)
        lWert2 = CLng(Text1(1).value)
        'Speicher Zeitraum TAG TAG
    End If
    
    
    For lcount = 0 To 7
        
        cSQL = "Select SUM(euronenbar) as SWECHSEL"
        cSQL = cSQL & " from temp_KABUCH_Wechsel   "
        cSQL = cSQL & " where DATUM >= " & Trim$(Str$(lWert1)) & " "
        cSQL = cSQL & " and DATUM <= " & Trim$(Str$(lWert2)) & " "
        cSQL = cSQL & " and kasnum = " & Trim$(Str$(lcount + 1)) & " "
            
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            If Not IsNull(rsrs!SWECHSEL) Then
                Select Case lcount
                    Case Is = 0: dWechselgeldKasseausKABUCH1 = CDbl(rsrs!SWECHSEL)
                    Case Is = 1: dWechselgeldKasseausKABUCH2 = CDbl(rsrs!SWECHSEL)
                    Case Is = 2: dWechselgeldKasseausKABUCH3 = CDbl(rsrs!SWECHSEL)
                    Case Is = 3: dWechselgeldKasseausKABUCH4 = CDbl(rsrs!SWECHSEL)
                    Case Is = 4: dWechselgeldKasseausKABUCH5 = CDbl(rsrs!SWECHSEL)
                    Case Is = 5: dWechselgeldKasseausKABUCH6 = CDbl(rsrs!SWECHSEL)
                    Case Is = 6: dWechselgeldKasseausKABUCH7 = CDbl(rsrs!SWECHSEL)
                    Case Is = 7: dWechselgeldKasseausKABUCH8 = CDbl(rsrs!SWECHSEL)
                End Select
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    
    Next lcount
        
    loeschNEW "temp_KABUCH", gdBase
    loeschNEW "temp_KABUCH_Wechsel", gdBase
    
    
    'offenes Wechselgeld anderer Kassen, also nicht gewählter anzeigen
    dOffenenesWechselgeldAndere = 0
    For lcount = 0 To 7
        If Check1(lcount).value = vbUnchecked Then
            Select Case lcount
                Case Is = 0: dOffenenesWechselgeldAndere = dOffenenesWechselgeldAndere + dOffenenesWechselgeldKasse1
                Case Is = 1: dOffenenesWechselgeldAndere = dOffenenesWechselgeldAndere + dOffenenesWechselgeldKasse2
                Case Is = 2: dOffenenesWechselgeldAndere = dOffenenesWechselgeldAndere + dOffenenesWechselgeldKasse3
                Case Is = 3: dOffenenesWechselgeldAndere = dOffenenesWechselgeldAndere + dOffenenesWechselgeldKasse4
                Case Is = 4: dOffenenesWechselgeldAndere = dOffenenesWechselgeldAndere + dOffenenesWechselgeldKasse5
                Case Is = 5: dOffenenesWechselgeldAndere = dOffenenesWechselgeldAndere + dOffenenesWechselgeldKasse6
                Case Is = 6: dOffenenesWechselgeldAndere = dOffenenesWechselgeldAndere + dOffenenesWechselgeldKasse7
                Case Is = 7: dOffenenesWechselgeldAndere = dOffenenesWechselgeldAndere + dOffenenesWechselgeldKasse8
            End Select
        End If
    Next lcount
    
    If dOffenenesWechselgeldAndere <> 0 Then 'offenes Wechselgeld anderer Kassen
        ctmp = Format$(dOffenenesWechselgeldAndere, "###,###,##0.00")
        ctmp = "(" & ctmp & " " & gcWaehrung & ")"
        Label1(53).Caption = ctmp
        Label1(53).Visible = True
    Else
        Label1(53).Caption = "0"
        Label1(53).Visible = False
    End If
    
    
    Dim cKassen As String
    
    cKassen = ""
    
    Label1(56).Caption = ""
    
    Label1(55).Caption = "0"
    Label1(55).Visible = False
    
    'für die Zeitraumsauswertung
    If bZeitraumleztzerTagundHeuteundalleGew = True Then
    
        For lcount = 0 To 7
            Select Case lcount
                Case Is = 0
                
                    If DatendrinSQL("select * from kabuch where kasnum = " & lcount + 1 & " and BEZUMS = 'Wechselgeld' and Datum = " & Trim$(Str$(lWert2)) & "  ", gdBase) = False Then
                        If dOffenenesWechselgeldKasse1 > 0 Then
                            cKassen = cKassen & "1,"
                            dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse1
                        End If
                    End If
                    
                Case Is = 1
                    
                    If DatendrinSQL("select * from kabuch where kasnum = " & lcount + 1 & " and BEZUMS = 'Wechselgeld' and Datum = " & Trim$(Str$(lWert2)) & "  ", gdBase) = False Then
                    
                    If dOffenenesWechselgeldKasse2 > 0 Then
                        cKassen = cKassen & "2,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse2
                    End If
                    End If
                Case Is = 2
                    If DatendrinSQL("select * from kabuch where kasnum = " & lcount + 1 & " and BEZUMS = 'Wechselgeld' and Datum = " & Trim$(Str$(lWert2)) & "  ", gdBase) = False Then
                    
                    If dOffenenesWechselgeldKasse3 > 0 Then
                        cKassen = cKassen & "3,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse3
                    End If
                    End If
                Case Is = 3
                    If DatendrinSQL("select * from kabuch where kasnum = " & lcount + 1 & " and BEZUMS = 'Wechselgeld' and Datum = " & Trim$(Str$(lWert2)) & "  ", gdBase) = False Then
                    
                    If dOffenenesWechselgeldKasse4 > 0 Then
                        cKassen = cKassen & "4,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse4
                    End If
                    End If
                Case Is = 4
                    If DatendrinSQL("select * from kabuch where kasnum = " & lcount + 1 & " and BEZUMS = 'Wechselgeld' and Datum = " & Trim$(Str$(lWert2)) & "  ", gdBase) = False Then
                    
                    If dOffenenesWechselgeldKasse5 > 0 Then
                        cKassen = cKassen & "5,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse5
                    End If
                    End If
                Case Is = 5
                    If DatendrinSQL("select * from kabuch where kasnum = " & lcount + 1 & " and BEZUMS = 'Wechselgeld' and Datum = " & Trim$(Str$(lWert2)) & "  ", gdBase) = False Then
                    
                    If dOffenenesWechselgeldKasse6 > 0 Then
                        cKassen = cKassen & "6,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse6
                    End If
                    End If
                Case Is = 6
                    If DatendrinSQL("select * from kabuch where kasnum = " & lcount + 1 & " and BEZUMS = 'Wechselgeld' and Datum = " & Trim$(Str$(lWert2)) & "  ", gdBase) = False Then
                    
                    If dOffenenesWechselgeldKasse7 > 0 Then
                        cKassen = cKassen & "7,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse7
                    End If
                    End If
                Case Is = 7
                    If DatendrinSQL("select * from kabuch where kasnum = " & lcount + 1 & " and BEZUMS = 'Wechselgeld' and Datum = " & Trim$(Str$(lWert2)) & "  ", gdBase) = False Then
                    
                    If dOffenenesWechselgeldKasse8 > 0 Then
                        cKassen = cKassen & "8,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse8
                    End If
                    End If
            End Select
            
        Next lcount
        

        If Right(cKassen, 1) = "," Then
             cKassen = Left(cKassen, Len(cKassen) - 1)
        End If
        
        Label1(56).Caption = cKassen
    
    
        If dOffenenesWechselgeldDEAKTIV <> 0 Then 'offenes Wechselgeld nicht abgerechnet Kassen
            ctmp = Format$(dOffenenesWechselgeldDEAKTIV, "###,###,##0.00")
            ctmp = "(" & ctmp & " " & gcWaehrung & ")"
            Label1(55).Caption = ctmp
            Label1(55).Visible = True
        
            
        End If
    
    End If
    
    'Ende für die Zeitraumsauswertung
    
    
    
    
    
    
    
    
    
    
    
    
    If bgleicherTagundHeute = True Then
    
        For lcount = 0 To 7
            Select Case lcount
                Case Is = 0
                    If dWechselgeldKasseausKABUCH1 = 0 Then
                        If dOffenenesWechselgeldKasse1 > 0 Then cKassen = cKassen & "1,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse1
                    End If
                Case Is = 1
                    
                    If dWechselgeldKasseausKABUCH2 = 0 Then
                        If dOffenenesWechselgeldKasse2 > 0 Then cKassen = cKassen & "2,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse2
                    End If
                Case Is = 2
                    If dWechselgeldKasseausKABUCH3 = 0 Then
                        If dOffenenesWechselgeldKasse3 > 0 Then cKassen = cKassen & "3,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse3
                    End If
                Case Is = 3
                    If dWechselgeldKasseausKABUCH4 = 0 Then
                        If dOffenenesWechselgeldKasse4 > 0 Then cKassen = cKassen & "4,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse4
                    End If
                Case Is = 4
                    If dWechselgeldKasseausKABUCH5 = 0 Then
                        If dOffenenesWechselgeldKasse5 > 0 Then cKassen = cKassen & "5,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse5
                    End If
                Case Is = 5
                    If dWechselgeldKasseausKABUCH6 = 0 Then
                        If dOffenenesWechselgeldKasse6 > 0 Then cKassen = cKassen & "6,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse6
                    End If
                Case Is = 6
                    If dWechselgeldKasseausKABUCH7 = 0 Then
                        If dOffenenesWechselgeldKasse7 > 0 Then cKassen = cKassen & "7,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse7
                    End If
                Case Is = 7
                    If dWechselgeldKasseausKABUCH8 = 0 Then
                        If dOffenenesWechselgeldKasse8 > 0 Then cKassen = cKassen & "8,"
                        dOffenenesWechselgeldDEAKTIV = dOffenenesWechselgeldDEAKTIV + dOffenenesWechselgeldKasse8
                    End If
            End Select
            
        Next lcount
        

        If Right(cKassen, 1) = "," Then
             cKassen = Left(cKassen, Len(cKassen) - 1)
        End If
        
        Label1(56).Caption = cKassen
    
    
        If dOffenenesWechselgeldDEAKTIV <> 0 Then 'offenes Wechselgeld nicht abgerechnet Kassen
            ctmp = Format$(dOffenenesWechselgeldDEAKTIV, "###,###,##0.00")
            ctmp = "(" & ctmp & " " & gcWaehrung & ")"
            Label1(55).Caption = ctmp
            Label1(55).Visible = True
        
            
        End If
    
    End If
    
    dWECHSEL = 0
    
    For lcount = 0 To 7
        If Check1(lcount).value = vbChecked Then
            Select Case lcount
                Case Is = 0: dWECHSEL = dWECHSEL + dWechselgeldKasseausKABUCH1
                Case Is = 1: dWECHSEL = dWECHSEL + dWechselgeldKasseausKABUCH2
                Case Is = 2: dWECHSEL = dWECHSEL + dWechselgeldKasseausKABUCH3
                Case Is = 3: dWECHSEL = dWECHSEL + dWechselgeldKasseausKABUCH4
                Case Is = 4: dWECHSEL = dWECHSEL + dWechselgeldKasseausKABUCH5
                Case Is = 5: dWECHSEL = dWECHSEL + dWechselgeldKasseausKABUCH6
                Case Is = 6: dWECHSEL = dWECHSEL + dWechselgeldKasseausKABUCH7
                Case Is = 7: dWECHSEL = dWECHSEL + dWechselgeldKasseausKABUCH8
            End Select
        End If
    Next lcount
    
    cSQL = "Select SUM(Geldwert) as SABSCHOPF"
    cSQL = cSQL & " from ABSCHOPF where "
    
    If Option1(0).value = True Then
        lWert1 = Val(MaskEdBox1(0).Text)
        lWert2 = Val(MaskEdBox1(1).Text)
        cSQL = cSQL & "YEAR(ADATE) = " & Trim$(Str$(lWert2)) & " "
        cSQL = cSQL & "and MONTH(ADATE) = " & Trim$(Str$(lWert1)) & " "
        'Speicher hier Monat/Jahr
    Else
        lWert1 = Text1(0).value
        lWert2 = Text1(1).value
        cSQL = cSQL & "ADATE >= " & Trim$(Str$(lWert1)) & " "
        cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lWert2)) & " "
        'Speicher Zeitraum
    End If
        
    bAus = False
    For lcount = 0 To 7
        If Check1(lcount).value = vbUnchecked Then
            bAus = True
            Exit For
        End If
    Next lcount
    
    If bAus Then
        cSQL = cSQL & " and ( "
        bOr = False
        For lcount = 0 To 7
            If Check1(lcount).value = vbChecked Then
                If bOr Then
                    cSQL = cSQL & " or "
                End If
                cSQL = cSQL & "KASNUM = " & Trim$(Str$(lcount + 1)) & " "
                bOr = True
            End If
        Next lcount
        cSQL = cSQL & " ) "
    End If
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!SABSCHOPF) Then
            dABSCHOPF = rsrs!SABSCHOPF
        Else
            dABSCHOPF = 0
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
        
    
    cSQL = "Select "
    cSQL = cSQL & "  SUM(UMS_BAR1) as SUMS_BAR"
    cSQL = cSQL & ", SUM(UMS_KRED1) as SUMS_KRED"
    cSQL = cSQL & ", SUM(UMS_KARTE1) as SUMS_KARTE"
    cSQL = cSQL & ", SUM(UMS_SCHECK1) as SUMS_SCHEC"
    cSQL = cSQL & ", SUM(UMS_LAST1) as SUMS_LAST"
    
    cSQL = cSQL & ", SUM(TILGBAR1) as STILGBAR"
    cSQL = cSQL & ", SUM(TILGSCH1) as STILGSCH"
    cSQL = cSQL & ", SUM(TILGGUT1) as STILGGUT"
    cSQL = cSQL & ", SUM(TILGKAR1) as STILGKAR"
    
    cSQL = cSQL & ", SUM(GUTSCHBAR1) as SGUTSCHBAR"
    cSQL = cSQL & ", SUM(GUTSCHSCH1) as SGUTSCHSCH"
    cSQL = cSQL & ", SUM(GUTSCHKRE1) as SGUTSCHKRE"
    cSQL = cSQL & ", SUM(GUTSCHKAR1) as SGUTSCHKAR"
    cSQL = cSQL & ", SUM(GUTSCHLAST1) as SGUTSCHLAS"
    cSQL = cSQL & ", SUM(GUTSCHGUTSCH1) as SGUTSCHGUTSCH"
'    cSQL = cSQL & ", SUM(ABSCHOPF1) as SABSCHOPF"
    cSQL = cSQL & ", SUM(KDIFF1) as SKDIFF"
    cSQL = cSQL & ", SUM(TDIFF1) as STDIFF"
    cSQL = cSQL & ", SUM(DUKA1) as SDUKA"
'    cSQL = cSQL & ", SUM(WECHSEL1) as SWECHSEL"
    
    cSQL = cSQL & ", SUM(BARVERKAUF1) as SBARVERKAU"
    cSQL = cSQL & ", SUM(SCHVERKAUF1) as SSCHVERKAU"
    
    cSQL = cSQL & ", SUM(AUSZAHLUNG1) as SAUSZAHLUN"
    cSQL = cSQL & ", SUM(EINZAHLUNG1) as SEINZAHLUN"
    cSQL = cSQL & ", SUM(AUSZGUTSCH1) as SAUSZGUTSC"
    
    cSQL = cSQL & ", SUM(SPREIS_GES1) as SSPREIS_GE"
    cSQL = cSQL & ", SUM(SPREIS_ANZ1) as SSPREIS_AN"
    cSQL = cSQL & ", SUM(GESRAB_GES1) as SGESRAB_GE"
    cSQL = cSQL & ", SUM(GESRAB_ANZ1) as SGESRAB_AN"
    cSQL = cSQL & ", SUM(ARTRAB_GES1) as SARTRAB_GE"
    cSQL = cSQL & ", SUM(ARTRAB_ANZ1) as SARTRAB_AN"
    cSQL = cSQL & ", SUM(STORNO_GES1) as SSTORNO_GE"
    cSQL = cSQL & ", SUM(STORNO_ANZ1) as SSTORNO_AN"
    
    cSQL = cSQL & ", SUM(ZHLGGUTSCH1) as SZHLGGUTSC"
    cSQL = cSQL & ", SUM(KUNDENZAHL1) as SKUNDENZAH"
    cSQL = cSQL & ", SUM(GELDFACH1) as SGELDFACH"
    
    cSQL = cSQL & ", SUM(EINRGUTSCH1) as SEINRGUTSC"
    cSQL = cSQL & ", SUM(RESTGUTSCH1) as SRESTGUTSC"
    cSQL = cSQL & ", SUM(GUTSCHEIN1) as SGUTSCH"
    cSQL = cSQL & ", SUM(NUMSKARTE1) as SNUMSKARTE"
    
    cSQL = cSQL & " from AFCSTATS "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        dWert = rsrs.RecordCount
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!SUMS_BAR) Then
            dWert = rsrs!SUMS_BAR
        Else
            dWert = 0
        End If
        dUmsBar = dWert
        
        If Not IsNull(rsrs!SUMS_SCHEC) Then
            dWert = rsrs!SUMS_SCHEC
        Else
            dWert = 0
        End If
        dUmsScheck = dWert
        
        If Not IsNull(rsrs!SUMS_KARTE) Then
            dWert = rsrs!SUMS_KARTE
        Else
            dWert = 0
        End If
        dKarte = dWert
                
        If Not IsNull(rsrs!SUMS_KRED) Then
            dWert = rsrs!SUMS_KRED
        Else
            dWert = 0
        End If
        dKredit = dWert
        
        
        If Not IsNull(rsrs!SUMS_LAST) Then
            dWert = rsrs!SUMS_LAST
        Else
            dWert = 0
        End If
        dLast = dWert
        
        If Not IsNull(rsrs!SDUKA) Then
            dWert = rsrs!SDUKA
        Else
            dWert = 0
        End If
        dDUKA = dWert
        
        
        If Not IsNull(rsrs!STILGBAR) Then
            dWert = rsrs!STILGBAR
        Else
            dWert = 0
        End If
        dTilgBar = dWert
        
        If Not IsNull(rsrs!STILGSCH) Then
            dWert = rsrs!STILGSCH
        Else
            dWert = 0
        End If
        dTilgSch = dWert
        
        If Not IsNull(rsrs!STILGGUT) Then
            dWert = rsrs!STILGGUT
        Else
            dWert = 0
        End If
        dTilgGut = dWert
        
        If Not IsNull(rsrs!STILGKAR) Then
            dWert = rsrs!STILGKAR
        Else
            dWert = 0
        End If
        dTilgKar = dWert
        
        dTilgung = dTilgBar + dTilgSch + dTilgGut + dTilgKar
        
        If Not IsNull(rsrs!SGUTSCHBAR) Then
            dWert = rsrs!SGUTSCHBAR
        Else
            dWert = 0
        End If
        dGutschBar = dWert
        
        If Not IsNull(rsrs!SGUTSCHSCH) Then
            dWert = rsrs!SGUTSCHSCH
        Else
            dWert = 0
        End If
        dGutschSch = dWert
        
        If Not IsNull(rsrs!SGUTSCHKRE) Then
            dWert = rsrs!SGUTSCHKRE
        Else
            dWert = 0
        End If
        dGutschKre = dWert
        
        If Not IsNull(rsrs!SGUTSCHKAR) Then
            dWert = rsrs!SGUTSCHKAR
        Else
            dWert = 0
        End If
        dGutschKar = dWert
        
        If Not IsNull(rsrs!SNUMSKARTE) Then
            dWert = rsrs!SNUMSKARTE
        Else
            dWert = 0
        End If
        dNichtUmsReleKar = dWert
        
        If Not IsNull(rsrs!SGUTSCHLAS) Then
            dWert = rsrs!SGUTSCHLAS
        Else
            dWert = 0
        End If
        dGutschLast = dWert
        
        If Not IsNull(rsrs!SGUTSCHGUTSCH) Then
            dWert = rsrs!SGUTSCHGUTSCH
        Else
            dWert = 0
        End If
        dGutschGUTSCH = dWert
        
        If Not IsNull(rsrs!sGutsch) Then
            dWert = rsrs!sGutsch
        Else
            dWert = 0
        End If
        dGutschein = dWert
        

        
        If Not IsNull(rsrs!SSCHVERKAU) Then
            dWert = rsrs!SSCHVERKAU
        Else
            dWert = 0
        End If
        dSchVerkauf = dWert
        
        dKassenSchecks = dSchVerkauf + dGutschSch + dTilgSch
        
        If Not IsNull(rsrs!SAUSZAHLUN) Then
            dWert = rsrs!SAUSZAHLUN
        Else
            dWert = 0
        End If
        dAuszahlung = dWert
        
        If Not IsNull(rsrs!SEINZAHLUN) Then
            dWert = rsrs!SEINZAHLUN
        Else
            dWert = 0
        End If
        dEinzahlung = dWert
        
        If Not IsNull(rsrs!SAUSZGUTSC) Then
            dWert = rsrs!SAUSZGUTSC
        Else
            dWert = 0
        End If
        dAuszGutsch = dWert
        
        If Not IsNull(rsrs!SBARVERKAU) Then
            dWert = rsrs!SBARVERKAU
        Else
            dWert = 0
        End If
        dBarVerkauf = dWert
        
'        If Not IsNull(rsrs!SABSCHOPF) Then
'            dWert = rsrs!SABSCHOPF
'        Else
'            dWert = 0
'        End If
'        dABSCHOPF = dWert
        
        If Not IsNull(rsrs!SKDIFF) Then
            dWert = rsrs!SKDIFF
        Else
            dWert = 0
        End If
        dKDIFF = dWert
        
        If Not IsNull(rsrs!StDIFF) Then
            dWert = rsrs!StDIFF
        Else
            dWert = 0
        End If
        dTDIFF = 0
        
        If Not IsNull(rsrs!SDUKA) Then
            dWert = rsrs!SDUKA
        Else
            dWert = 0
        End If
        dDUKA = dWert
        

        
        dKassenBargeld = dBarVerkauf + dGutschBar + dTilgBar + dEinzahlung - dAuszahlung - dAuszGutsch - dABSCHOPF + dWECHSEL
        
        dKasse = dKassenBargeld + dKassenSchecks
        
        If Not IsNull(rsrs!SZHLGGUTSC) Then
            dWert = rsrs!SZHLGGUTSC
        Else
            dWert = 0
        End If
        dZhlgGutsch = dWert
                
        dScheck = dKassenSchecks - dGutschSch - dTilgSch
        dBar = dBarVerkauf
            
        dUmsatz = dZhlgGutsch + dKarte + dKredit + dUmsScheck + dUmsBar + dLast + dDUKA
        
           
        ctmp = Format$(dUmsatz, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(0).Caption = ctmp
           
        ctmp = Format$(dUmsBar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(1).Caption = ctmp
        
        ctmp = Format$(dUmsScheck, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(2).Caption = ctmp
        
        ctmp = Format$(dKredit, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(3).Caption = ctmp
        
        ctmp = Format$(dKarte, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(4).Caption = ctmp
        
        ctmp = Format$(dZhlgGutsch, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(20).Caption = ctmp
        
        ctmp = Format$(dKasse, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(8).Caption = ctmp
        
        ctmp = Format$(dKassenBargeld, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(9).Caption = ctmp
        
        ctmp = Format$(dBarVerkauf, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(25).Caption = ctmp
        
        ctmp = Format$(dEinzahlung, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(6).Caption = ctmp
        
        ctmp = Format$(dAuszahlung, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(7).Caption = ctmp
        
        ctmp = Format$(dABSCHOPF, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(47).Caption = ctmp
        
        ctmp = Format$(dTDIFF, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(48).Caption = ctmp
        
        ctmp = Format$(dKDIFF, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(49).Caption = ctmp
        
        ctmp = Format$(dDUKA, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(50).Caption = ctmp
        
        ctmp = Format$(dWECHSEL, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(46).Caption = ctmp
        
        ctmp = Format$(dKassenSchecks, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(26).Caption = ctmp
        
        ctmp = Format$(dGutschein, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(5).Caption = ctmp
        
        ctmp = Format$(dGutschBar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(21).Caption = ctmp
        Label3(42).Caption = ctmp
        
        ctmp = Format$(dGutschSch, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(22).Caption = ctmp
        
        ctmp = Format$(dGutschKre, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(23).Caption = ctmp
        
        ctmp = Format$(dGutschKar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(24).Caption = ctmp
        
        ctmp = Format$(dBarVerkauf, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(25).Caption = ctmp
        
        ctmp = Format$(dKassenSchecks, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(26).Caption = ctmp
        
        ctmp = Format$(dTilgung, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(27).Caption = ctmp
        
        ctmp = Format$(dTilgBar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(28).Caption = ctmp
        Label3(43).Caption = ctmp
        
        ctmp = Format$(dTilgSch, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(29).Caption = ctmp
        
        ctmp = Format$(dTilgGut, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(30).Caption = ctmp
        
        ctmp = Format$(dTilgKar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(31).Caption = ctmp
        
        ctmp = Format$(dLast, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(34).Caption = ctmp
        
        ctmp = Format$(dGutschLast, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(35).Caption = ctmp
        
        ctmp = Format$(dGutschGUTSCH, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(45).Caption = ctmp
        
        If Not IsNull(rsrs!SKUNDENZAH) Then
            dWert = rsrs!SKUNDENZAH
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(10).Caption = ctmp
        
        If dKunden = 0 Then
            dKunden = 1
        End If
        
        dWert = dUmsatz / dKunden
        ctmp = Format$(dWert, "###,###,##0.00")
        Label3(11).Caption = ctmp & " " & gcWaehrung
        
        If Not IsNull(rsrs!SSPREIS_GE) Then
            dWert = rsrs!SSPREIS_GE
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(12).Caption = ctmp
    
        If Not IsNull(rsrs!SSPREIS_AN) Then
            dWert = rsrs!SSPREIS_AN
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(13).Caption = ctmp
    
        If Not IsNull(rsrs!SGELDFACH) Then
            dWert = rsrs!SGELDFACH
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(14).Caption = ctmp
    
        If Not IsNull(rsrs!SARTRAB_GE) Then
            dWert = rsrs!SARTRAB_GE
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(15).Caption = ctmp
    
        If Not IsNull(rsrs!SGESRAB_GE) Then
            dWert = rsrs!SGESRAB_GE
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(16).Caption = ctmp
    
        If Not IsNull(rsrs!SGESRAB_AN) Then
            dWert = rsrs!SGESRAB_AN
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(17).Caption = ctmp
        
        If Not IsNull(rsrs!SARTRAB_AN) Then
            dWert = rsrs!SARTRAB_AN
        Else
            dWert = 0
        End If
        
        ctmp = Format$(dWert, "###,###,##0")
        Label3(53).Caption = ctmp
    
        If Not IsNull(rsrs!SSTORNO_GE) Then
            dWert = rsrs!SSTORNO_GE
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(18).Caption = ctmp
    
        If Not IsNull(rsrs!SSTORNO_AN) Then
            dWert = rsrs!SSTORNO_AN
        Else
            dWert = 0
        End If
        dKunden = dWert
        ctmp = Format$(dWert, "###,###,##0")
        Label3(19).Caption = ctmp
        
        If Not IsNull(rsrs!SEINRGUTSC) Then
            dWert = rsrs!SEINRGUTSC
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00 ") & gcWaehrung
        Label3(32).Caption = ctmp
    
        If Not IsNull(rsrs!SRESTGUTSC) Then
            dWert = rsrs!SRESTGUTSC
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00 ") & gcWaehrung
        Label3(33).Caption = ctmp
    
        If Not IsNull(rsrs!SAUSZGUTSC) Then
            dWert = rsrs!SAUSZGUTSC
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "###,###,##0.00 ") & gcWaehrung
        Label3(41).Caption = ctmp
        
        ctmp = Format$(dNichtUmsReleKar, "###,###,##0.00")
        ctmp = ctmp & " " & gcWaehrung
        Label3(51).Caption = ctmp
        
        If gbGutscheinBeiVKversteuern = True Then
        
            ctmp = Format$(dKarte + dTilgKar + dNichtUmsReleKar, "###,###,##0.00 ") & gcWaehrung
            
        Else
            ctmp = Format$(dKarte + dGutschKar + dTilgKar + dNichtUmsReleKar, "###,###,##0.00 ") & gcWaehrung
            
        End If
       
       
       
        
        Label3(52).Caption = ctmp
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    If Option1(0).value = True Then
        lWert1 = Val(MaskEdBox1(0).Text)
        lWert2 = Val(MaskEdBox1(1).Text)
        ctmp = "and YEAR(ADATE) = " & Trim$(Str$(lWert2)) & " and MONTH(ADATE) = " & Trim$(Str$(lWert1)) & " "
        
    Else
        sVon = Text1(0).value ' MaskEdBox2(0).Text
        sBis = Text1(1).value ' MaskEdBox2(1).Text
        
        lWert1 = DateValue(sVon)
        lWert2 = DateValue(sBis)
        ctmp = "and ADATE >= " & Trim$(Str$(lWert1)) & " and ADATE <= " & Trim$(Str$(lWert2)) & " "
        
    End If
    
    Label3(36).Caption = "0,00 " & gcWaehrung
    Label3(37).Caption = "0,00 " & gcWaehrung
    Label3(38).Caption = "0,00 " & gcWaehrung
    Label3(39).Caption = "0,00 " & gcWaehrung
    Label3(40).Caption = "0,00 " & gcWaehrung
    
    
    
    
    
    
    
    
    
    
    
    
    Dim dUmsatzausKassjour As Double
    Dim dUmsatzausKassjourGesamt As Double
    
    dUmsatzausKassjour = 0
    dUmsatzausKassjourGesamt = 0
    
    Dim dNichtUmsGutschbetrag As Double
    dNichtUmsGutschbetrag = 0
    
    If gbGutscheinBeiVKversteuern = True Then
    
        
        cSQL = "Select SUM(Wert) as UMSATZ from Gemischte_ZP "
        
        
        If Trim(cSQLKass) <> "" Then
            cSQL = cSQL & cSQLKass
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        
        
        cSQL = cSQL & "  Thema = 'nicht ums GUTSCHBETRAG'"
        
        
        cSQL = cSQL & ctmp
        
        
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            If Not IsNull(rsrs!UMSATZ) Then
                dNichtUmsGutschbetrag = rsrs!UMSATZ
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        
        
        
        cSQL = "Select MWST, SUM(PREIS) as UMSATZ from Kassjour "
        If Trim(cSQLKass) <> "" Then
            cSQL = cSQL & cSQLKass
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " UMS_OK <> 'N' "
        cSQL = cSQL & ctmp
        cSQL = cSQL & " group by MWST "
        
        
    Else
    
        cSQL = "Select MWST, SUM(PREIS) as UMSATZ from Kassjour "
        If Trim(cSQLKass) <> "" Then
            cSQL = cSQL & cSQLKass
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & "  ARTNR <> 666666 " 'and UMS_OK <> 'N'
        cSQL = cSQL & ctmp
        cSQL = cSQL & " group by MWST "
    

    
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!MWST) Then
                cMWSK = rsrs!MWST
            Else
                    
            End If
            
            If Not IsNull(rsrs!UMSATZ) Then
                dUmsatzausKassjour = rsrs!UMSATZ
            Else
                dUmsatzausKassjour = 0
            End If
            
            Select Case cMWSK
                Case Is = "V"
                
                
                    dUmsatzausKassjour = dUmsatzausKassjour - dNichtUmsGutschbetrag
                
                    Label3(36).Caption = Format$(dUmsatzausKassjour, "######0.00") & " " & gcWaehrung
                    Label3(37).Caption = Format$((dUmsatzausKassjour / (gdMWStV + 100)) * gdMWStV, "######0.00") & " " & gcWaehrung
                    
                    dUmsatzausKassjourGesamt = dUmsatzausKassjourGesamt + dUmsatzausKassjour
                    
                    
                    
                Case Is = "E"
                    Label3(38).Caption = Format$(dUmsatzausKassjour, "######0.00") & " " & gcWaehrung
                    Label3(39).Caption = Format$((dUmsatzausKassjour / (gdMWStE + 100)) * gdMWStE, "######0.00") & " " & gcWaehrung
                    dUmsatzausKassjourGesamt = dUmsatzausKassjourGesamt + dUmsatzausKassjour
                Case Is = "O"
                    Label3(40).Caption = Format$(dUmsatzausKassjour, "######0.00") & " " & gcWaehrung
                    dUmsatzausKassjourGesamt = dUmsatzausKassjourGesamt + dUmsatzausKassjour
                Case Else
                
            End Select
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Round(dUmsatzausKassjourGesamt, 2) <> Round(dUmsatz, 2) Then
    
        Label3(40).ForeColor = glWarn
        Label3(38).ForeColor = glWarn
        Label3(36).ForeColor = glWarn
    Else
        Label3(40).ForeColor = Label1(37).ForeColor
        Label3(38).ForeColor = Label1(37).ForeColor
        Label3(36).ForeColor = Label1(37).ForeColor
    
    End If
    
    
    
    
    
    loeschNEW "AfcTempo", gdBase
    cSQL = "Select * into AFCTempo from Kassjour where UMS_OK = 'N' "
    cSQL = cSQL & ctmp
    gdBase.Execute cSQL, dbFailOnError
    
    dWert = 0

    cSQL = "Select SUM(PREIS) as UMSATZ from AFCTempo " & cSQLKass
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!UMSATZ) Then
            dWert = rsrs!UMSATZ
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    loeschNEW "AfcTempo", gdBase
    
    dWert = dWert + dNichtUmsGutschbetrag
    
    Label3(44).Caption = Format$(dWert, "######0.00") & " " & gcWaehrung
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseDatenWK21f"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
    
End Sub
Private Sub SucheKassenAbschlussWK21f()
    On Error GoTo LOKAL_ERROR
    
    Dim lRet As Long
    
    lRet = fnPruefeEingabeDialogWK21f()
    Select Case lRet
        Case Is = 0     'alles okay
            HoleKassenAbschlussDatenWK21f
            
        Case Is = 1     'kein oder falscher Monat
            MsgBox "Falsche oder fehlende Monatsangabe!", vbInformation, "Winkiss Hinweis:"
            MaskEdBox1(0).SetFocus
            
        Case Is = 2     'kein oder falsches Jahr
            MsgBox "Falsche oder fehlende Jahresangabe!", vbInformation, "Winkiss Hinweis:"
            MaskEdBox1(1).SetFocus
        Case Is = 3     'kein oder falsches Von-Datum
            MsgBox "Falsches oder fehlendes Von-Datum!", vbInformation, "Winkiss Hinweis:"

        Case Is = 4     'kein oder falsches Bis-Datum
            MsgBox "Falsches oder fehlendes Bis-Datum!", vbInformation, "Winkiss Hinweis:"

            
        Case Is = 5     'Von-Datum > Bis-Datum
            MsgBox "Von-Datum ist größer als Bis-Datum!", vbInformation, "Winkiss Hinweis:"


        Case Is = 6     'Keine Kassennummer angegeben
            MsgBox "Bitte mindestens eine Kasse markieren!", vbInformation, "Winkiss Hinweis:"
            Check1(0).SetFocus

    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheKassenAbschlussWK21f"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Command0_Click(index As Integer)

    On Error GoTo LOKAL_ERROR
    
    Select Case index
    
        Case Is = 0
            Text1(0).value = Format(Datumschreiben11a(3700, 260), "DD.MM.YY")
            Text1(1).value = Text1(0).value
            
        Case Is = 1
            Text1(1).value = Format(Datumschreiben11a(5600, 260), "DD.MM.YY")
            'fertig
        
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum As String
    Dim lAnz As Long
    Dim lcount As Long
    
    Screen.MousePointer = 11
    
    Positionieren21f
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    For lAnz = 0 To 53
        Label3(lAnz).Caption = "0,00 " & gcWaehrung
    Next lAnz
    
    
    For lAnz = 0 To 7
        Check1(lAnz).value = vbChecked
        Check1(lAnz).BackColor = glfarbe(lAnz + 1)
    Next lAnz
    
    Text1(0).value = DateValue(Now)
    Text1(1).value = DateValue(Now)
    
    Text1(0).Enabled = False
    Text1(1).Enabled = False
    Command0(0).Enabled = False
    Command0(1).Enabled = False
    
    Option1_Click 1
    Option1(1).value = True
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub

Private Sub Label1_Click(index As Integer)
On Error GoTo LOKAL_ERROR
     
    If index = 53 Then
        Frame5.Visible = True
        Frame5.BackColor = glH2
        
        Zeige_Wechselgeld_an
    End If

    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Zeige_Wechselgeld_an()
On Error GoTo LOKAL_ERROR
     
    'ich suche das offene Wechselgeld aller anderen Kassen
    'Afcstat, Wechsel, Kasnum
    
    
    anzeige "normal", "Wechselgeld der anderen Kassenschubladen", Label0(4)
    
    Dim sSQL As String
    Dim lcount As Long
    Dim lWert1 As Long
    Dim lWert2 As Long
    Dim cDatum As String
    
    loeschNEW "WECHSEL_ANZEIGE", gdBase
    
    sSQL = "Create Table WECHSEL_ANZEIGE ("
    sSQL = sSQL & " Wechselgeld Double"
    sSQL = sSQL & ", KASNUM long"
    sSQL = sSQL & ") "
    gdBase.Execute sSQL, dbFailOnError
    
    Dim dOffenenesWechselgeld As Double
    dOffenenesWechselgeld = 0
    
    If Option1(1).value = True Then
    
        lWert1 = CLng(Text1(0).value)
        lWert2 = CLng(Text1(1).value)
            

    ElseIf Option1(0).value = True Then
        'hier kann nur der MonatsEndetag gelten
        lWert1 = Val(MaskEdBox1(0).Text) 'Monat
        lWert2 = Val(MaskEdBox1(1).Text) 'Jahr
        
        Select Case lWert1
            Case 1, 3, 5, 7, 8, 10, 12
                cDatum = "31."
            
            Case 2
                If lWert2 = 2016 Then
                    cDatum = "29."
                ElseIf lWert2 = 2020 Then
                    cDatum = "29."
                ElseIf lWert2 = 2024 Then
                    cDatum = "29."
                ElseIf lWert2 = 2028 Then
                    cDatum = "29."
                Else
                    cDatum = "28."
                End If
            
            Case Else
                cDatum = "30."
        End Select
            
        cDatum = cDatum & lWert1 & "." & lWert2
        lWert2 = CLng(DateValue(cDatum))
        
    End If
    
    
    For lcount = 0 To 7
        If Check1(lcount).value = vbUnchecked Then

            sSQL = "Insert into WECHSEL_ANZEIGE "
            sSQL = sSQL & " select sum(wechsel) as Wechselgeld," & Trim$(Str$(lcount + 1)) & " as Kasnum from Afcstat where kasnum = " & Trim$(Str$(lcount + 1)) & " "
            sSQL = sSQL & " and ADATE <= " & Trim$(Str$(lWert2)) & " "
            gdBase.Execute sSQL, dbFailOnError
        End If
    Next lcount
    
     
    Dim rsrs As Recordset
    Dim cLBSatz As String
    Dim sFeld As String
    
    List2.Clear

    sSQL = "Select * from WECHSEL_ANZEIGE where wechselgeld > 0"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
        
            cLBSatz = ""
            sFeld = ""
        
            If Not IsNull(rsrs!kasnum) Then
                sFeld = rsrs!kasnum
            End If
            
            cLBSatz = cLBSatz & Space$(6 - Len(sFeld))
            cLBSatz = cLBSatz & sFeld & Space$(2)
            
            If Not IsNull(rsrs!Wechselgeld) Then
                sFeld = Format(rsrs!Wechselgeld, "######0.00")
            End If
            
            cLBSatz = cLBSatz & Space$(10 - Len(sFeld))
            cLBSatz = cLBSatz & sFeld & Space$(2)
            
            List2.AddItem cLBSatz
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    

    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zeige_Wechselgeld_an"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Zeige_WechselgeldausKABUCH_an()
On Error GoTo LOKAL_ERROR
     
    'ich suche das  Wechselgeld aller  Kassen aus KABUCH
    
    
    anzeige "normal", "Wechselgeld der einzelnen Tage (laut Kassenbuch)", Label0(4)
    
    
    Dim sSQL As String
    Dim cVon As String
    Dim cBis As String
    Dim lcount As Long
    Dim lWert1 As Long
    Dim lWert2 As Long
    Dim lWert3 As Long
    Dim cDatum As String
    
    loeschNEW "WECHSEL_ANZEIGE", gdBase
    
    sSQL = "Create Table WECHSEL_ANZEIGE ("
    sSQL = sSQL & " Wechselgeld Double"
    sSQL = sSQL & ", KASNUM long"
    sSQL = sSQL & ", Datum Datetime"
    sSQL = sSQL & ") "
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "temp_KABUCH", gdBase
    
    sSQL = "Select * into temp_KABUCH "
    sSQL = sSQL & " from KABUCH where BEZUMS = 'Wechselgeld' "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "temp_KABUCH_Wechsel", gdBase
    
    sSQL = "Select max(EURBAR) as euronenbar ,Max(autopos) as maxi,datum,kasnum into temp_KABUCH_Wechsel from temp_KABUCH group by datum,kasnum  "
    gdBase.Execute sSQL, dbFailOnError
    
    If Option1(0).value = True Then
        'zeitraum Monat Jahr
        lWert1 = Val(MaskEdBox1(0).Text)
        lWert2 = Val(MaskEdBox1(1).Text)
        
        cVon = Format("01." & Val(MaskEdBox1(0).Text) & "." & Val(MaskEdBox1(1).Text), "DD.MM.YYYY")
        
        Select Case Val(MaskEdBox1(0).Text)
            Case 1, 3, 5, 7, 8, 10, 12
                cBis = Format("31." & Val(MaskEdBox1(0).Text) & "." & Val(MaskEdBox1(1).Text), "DD.MM.YYYY")
            Case 2
                If Val(MaskEdBox1(1).Text) = 2016 Then
                    cBis = Format("29." & Val(MaskEdBox1(0).Text) & "." & Val(MaskEdBox1(1).Text), "DD.MM.YYYY")
                ElseIf Val(MaskEdBox1(1).Text) = 2012 Then
                    cBis = Format("29." & Val(MaskEdBox1(0).Text) & "." & Val(MaskEdBox1(1).Text), "DD.MM.YYYY")
                Else
                    cBis = Format("28." & Val(MaskEdBox1(0).Text) & "." & Val(MaskEdBox1(1).Text), "DD.MM.YYYY")
                End If
            Case Else
                cBis = Format("30." & Val(MaskEdBox1(0).Text) & "." & Val(MaskEdBox1(1).Text), "DD.MM.YYYY")
        End Select
        
        lWert1 = CLng(DateValue(cVon))
        lWert2 = CLng(DateValue(cBis))
    Else
        lWert1 = CLng(Text1(0).value)
        lWert2 = CLng(Text1(1).value)
        'Speicher Zeitraum TAG TAG
    End If
    
    For lcount = 0 To 7
    
        If Check1(lcount).value = vbChecked Then
        
        
        
            sSQL = "Insert into WECHSEL_ANZEIGE "
            sSQL = sSQL & " select euronenbar as Wechselgeld," & Trim$(Str$(lcount + 1)) & " as Kasnum, datum from temp_KABUCH_Wechsel where kasnum = " & Trim$(Str$(lcount + 1)) & " "
            sSQL = sSQL & " and DATUM >= " & Trim$(Str$(lWert1)) & " "
            sSQL = sSQL & " and DATUM <= " & Trim$(Str$(lWert2)) & " "
            sSQL = sSQL & " "
            gdBase.Execute sSQL, dbFailOnError
        
        End If
    
    Next lcount
    
     
    Dim rsrs As Recordset
    Dim cLBSatz As String
    Dim sFeld As String
    
    List2.Clear

    sSQL = "Select * from WECHSEL_ANZEIGE where wechselgeld > 0 order by Datum, kasnum"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
        
            cLBSatz = ""
            sFeld = ""
            
            If Not IsNull(rsrs!Datum) Then
                sFeld = Format(rsrs!Datum, "DD.MM.YY")
            End If
            
            cLBSatz = sFeld
            
            If Not IsNull(rsrs!kasnum) Then
                sFeld = rsrs!kasnum
            End If
            
            cLBSatz = cLBSatz & Space$(3 - Len(sFeld))
            cLBSatz = cLBSatz & sFeld & Space$(2)
            
            If Not IsNull(rsrs!Wechselgeld) Then
                sFeld = Format(rsrs!Wechselgeld, "######0.00")
            End If
            
            cLBSatz = cLBSatz & Space$(10 - Len(sFeld))
            cLBSatz = cLBSatz & sFeld & Space$(2)
            
            List2.AddItem cLBSatz
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    

    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zeige_WechselgeldausKABUCH_an"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Positionieren21f()
    On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 0
    Frame1.Left = 0
    Frame1.Height = 1575
    Frame1.Width = 12015
    Frame1.Visible = True
    
    Frame4.Top = 1440
    Frame4.Left = 0
    Frame4.Height = 7215
    Frame4.Width = 12015
    Frame4.Visible = True
    
    Frame5.Top = 2520
    Frame5.Left = 600
    Frame5.Height = 2295
    Frame5.Width = 5295
    Frame5.Visible = False


    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Positionieren21f"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub


Private Sub Label3_Click(index As Integer)
On Error GoTo LOKAL_ERROR
     
    If index = 46 Then
        Frame5.Visible = True
        Frame5.BackColor = glH2
        
        Zeige_WechselgeldausKABUCH_an
    End If

    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label3_Click"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub MaskEdBox1_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    MaskEdBox1(index).BackColor = glSelBack1
    MaskEdBox1(index).SelStart = 0
    MaskEdBox1(index).SelLength = Len(MaskEdBox1(index).Text)

    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub


Private Sub MaskEdBox1_LostFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    MaskEdBox1(index).BackColor = vbWhite

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Option1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case index
        Case Is = 0
            Frame2.Enabled = True
            Label0(0).Enabled = True
            Label0(1).Enabled = True
            Frame3.Enabled = False
            Label0(2).Enabled = False
            Label0(3).Enabled = False
            Text1(0).Enabled = False
            Text1(1).Enabled = False
            Command0(0).Enabled = False
            Command0(1).Enabled = False
            

            MaskEdBox1(0).SetFocus
            
        Case Is = 1
            Frame3.Enabled = True
            Label0(2).Enabled = True
            Label0(3).Enabled = True
            Frame2.Enabled = False
            Label0(0).Enabled = False
            Label0(1).Enabled = False
            MaskEdBox1(0).Text = "__"
            MaskEdBox1(1).Text = "____"
            
            Text1(0).Enabled = True
            Text1(1).Enabled = True
            
            Command0(0).Enabled = True
            Command0(1).Enabled = True

    End Select
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SSCommand1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    
    Screen.MousePointer = 11
    
    Select Case index
        Case Is = 0     'Suchen
            SucheKassenAbschlussWK21f
            
        Case Is = 1     'Drucken
            iRet = MsgBox("Wollen Sie die Daten auf dem Bondrucker ausdrucken?", vbQuestion + vbYesNo, "Winkiss Frage:")
            If iRet <> vbYes Then
            
                DruckeTagesabschlussNeuWK21f 1
                reportbildschirm "WKL003b", "aWKL21fa"
                
            Else
                DruckeTagesAbschlussAufBonDruckerWK21f
            End If
        
        Case Is = 2     'Schließen
            Unload frmWK21f
            
        Case Is = 3     'Agenauswertung
            iRet = MsgBox("Wollen Sie die Daten auf dem Bondrucker ausdrucken?", vbQuestion + vbYesNo, "Winkiss Frage:")
            If iRet <> vbYes Then
                DruckeKassenAgnAuswertungaufListendrucker
            Else
                DruckeKassenAgnAuswertungaufBondrucker
            End If
        Case Is = 4
            Screen.MousePointer = 0
            frmWKL08.Show 1 'Zeitung
        Case 5
        
            Dim lWert1 As Long
            Dim lWert2 As Long
            Dim i As Integer
            Dim cTag As String
            
            If Option1(0).value = True Then
            
                lWert1 = Val(MaskEdBox1(0).Text)
                lWert2 = Val(MaskEdBox1(1).Text)
                
                
                Option1(1).value = True
                
                For i = 1 To 31
                
                    cTag = Format(i, "0#")
                    
                    Text1(0).value = cTag & "." & CStr(lWert1) & "." & CStr(lWert2)
                    Text1(1).value = cTag & "." & CStr(lWert1) & "." & CStr(lWert2)
                
                    SucheKassenAbschlussWK21f
                    
'                    MsgBox Label3(0).Caption
                    If Label3(0).Caption <> "0,00 EUR" Then
                    
                        DruckeTagesabschlussNeuWK21f 1
                        
                        reportbildschirmtoPDF "aWKL21fa", "C:\Tag\Tagesprotokoll_" & cTag & Format(CStr(lWert1), "0#") & CStr(lWert2) & ".rtf"
                        Pause (1)
                    
                    End If
                    
                Next i
            End If
        

    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub DruckeKassenAgnAuswertungaufBondrucker()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim cAgn As String
    Dim cMenge As String
    Dim cAGNBEZEICH As String
    Dim dAnteil As Double
    Dim cAnteil As String
    Dim cPreis As String
    
    Dim dWert As Double
    Dim dPreis As Double
    Dim dSummeEin As Double
    Dim dSummeAus As Double
    
    
    Dim lAnz As Long
    Dim lcount As Long
    Dim aDeviceName As String
    Dim cEscapeSequenz As String
    Dim iLenZeile As Integer
    Dim cDaten As String
    Dim lAnzZeile As Long
    ReDim cDruckZeile(1 To 1) As String
    Dim iRet As Integer
    Dim dSum As Double
    
    Dim lWert1 As Long
    Dim lWert2 As Long
    Dim bAus As Boolean
    Dim bChecky As Boolean
    Dim bOr As Boolean
    
    bChecky = False

    Screen.MousePointer = 11
    
    loeschNEW "ABAGN", gdBase
    CreateTable "ABAGN", gdBase
    
    cSQL = "insert into ABAGN Select artnr, preis, menge, agn from Kassjour where  "

    If Option1(0).value = True Then
        lWert1 = Val(MaskEdBox1(0).Text)
        lWert2 = Val(MaskEdBox1(1).Text)
        cSQL = cSQL & "YEAR(ADATE) = " & Trim$(Str$(lWert2)) & " "
        cSQL = cSQL & "and MONTH(ADATE) = " & Trim$(Str$(lWert1)) & " "
        'Speicher hier Monat/Jahr
    Else
        lWert1 = Text1(0).value '     DateValue(MaskEdBox2(0).Text)
        lWert2 = Text1(1).value 'DateValue(MaskEdBox2(1).Text)
        cSQL = cSQL & "ADATE >= " & Trim$(Str$(lWert1)) & " "
        cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lWert2)) & " "
        'Speicher Zeitraum
    End If
        
    bAus = False
    For lcount = 0 To 7
        If Check1(lcount).value = vbUnchecked Then
            bAus = True
            Exit For
        End If
    Next lcount
    
    If bAus Then
        cSQL = cSQL & " and ( "
        bOr = False
        For lcount = 0 To 7
            If Check1(lcount).value = vbChecked Then
                If bOr Then
                    cSQL = cSQL & " or "
                End If
                cSQL = cSQL & " KASNUM = " & Trim$(Str$(lcount + 1)) & " "
                bOr = True
                
            End If
        Next lcount
        cSQL = cSQL & " ) "
    End If
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "ABAGNS", gdBase
    CreateTable "ABAGNS", gdBase
    
    cSQL = "insert into ABAGNS Select agn,sum(preis) as apreis,sum(menge) as amenge from ABAGN group by agn "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update ABAGNS inner join AGNDBF on ABAGNS.AGN = AGNDBF.AGN set ABAGNS.AGNBEZEICH = AGNDBF.AGTEXT"
    gdBase.Execute cSQL, dbFailOnError
    
    dSum = 0
    cSQL = "Select sum(apreis) as maxi from ABAGNS  "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dSum = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Select * from ABAGNS order by AGN "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        '********************************************
        '*** 1.Schritt: Umschalten auf BonDrucker ***
        '********************************************
        setzedrucker gcBonDrucker
        
    
        '********************************************************
        '*** 2.Schritt: Drucker an, Display aus, Init Drucker ***
        '********************************************************
    
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = Chr$(27) + Chr$(61) + Chr$(1) + Chr$(27) + Chr$(64)
        OpenDrawer aDeviceName, cEscapeSequenz
    
    
        iLenZeile = 32
        'Drucker ist bereits auf BonDrucker geschaltet
        aDeviceName = gcBonDrucker
    
        '******************************************************************
        If gbDEMO Then
            'HIER FÜR DEMO FESTTEXT
            cDaten = "K.I.S.S. Warenwirtschaft"
        Else
            cDaten = gcBonText(0)
        End If
        If Trim$(cDaten) <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To 1) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        '******************************************************************
        If gbDEMO Then
            'HIER FÜR DEMO FESTTEXT
            cDaten = "Carsten Schröder"
        Else
            cDaten = gcBonText(1)
        End If
        If Trim$(cDaten) <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        '******************************************************************
        If gbDEMO Then
            'HIER FÜR DEMO FESTTEXT
            cDaten = "DEMO-VERSION!"
        Else
            cDaten = gcBonText(4)
        End If
        If Trim$(cDaten) <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        '***********************************************
        'Kopfdaten 4.Zeile an Drucker senden
        '***********************************************
    
        If gbDEMO Then
            'HIER FÜR DEMO FESTTEXT
            cDaten = "DEMO - VERSION"
        Else
            cDaten = gcBonText(12)
        End If
        
        If Trim$(cDaten) <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        '******************************************************************
        
        cDaten = String$(iLenZeile, "-")
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        '******************************************************************
        
        Dim cVon As String
        Dim cBis As String
        
        Dim cMon As String
        Dim cJahr As String
        
        If Option1(0).value = True Then
            cMon = Val(MaskEdBox1(0).Text)
            cJahr = Val(MaskEdBox1(1).Text)
            cDaten = "ABSCHLUSS   " & cMon & "/" & cJahr
            
            'Speicher hier Monat/Jahr
        Else
            cVon = Text1(0).value
            cBis = Text1(1).value
    
            
            cDaten = "ABSCHLUSS " & Format$(cVon, "DD.MM.YY") & " - " & Format$(cBis, "DD.MM.YY")
            'Speicher Zeitraum
            
        End If
        
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '******************************************************************
        cDaten = "Artikelgruppenauswertung"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        '******************************************************************
    
        cDaten = String$(iLenZeile, "-")
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '*****Kasse?
        bAus = False
        For lcount = 0 To 7
            If Check1(lcount).value = vbUnchecked Then
                bAus = True
                Exit For
            End If
        Next lcount
        
        If bAus Then
            cDaten = "Kasse: "
            For lcount = 0 To 7
                If Check1(lcount).value = vbChecked Then
                    If bChecky = False Then
                        cDaten = cDaten & lcount + 1
                    Else
                        cDaten = cDaten & ", " & lcount + 1
                    End If
                    
                    bChecky = True
                End If
            Next lcount
        Else
            cDaten = "Kasse: alle Kassen"
        End If
        '*******Kasse? Ende
    
'        cDaten = "Kasse: " & gcKasNum
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '********************************************************
        '*** 3.Schritt: Daten drucken                         ***
        '********************************************************

        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!AGN) Then
                cAgn = rsrs!AGN
            Else
                cAgn = "0"
            End If
            
            If Not IsNull(rsrs!aMenge) Then
                cMenge = rsrs!aMenge
            Else
                cMenge = "0"
            End If
            
            cMenge = cMenge & " Stück"
            
            If Not IsNull(rsrs!AGNBEZEICH) Then
                cAGNBEZEICH = rsrs!AGNBEZEICH
            Else
                cAGNBEZEICH = ""
            End If
            
            


            dPreis = 0
        
            If Not IsNull(rsrs!APREIS) Then
                cPreis = rsrs!APREIS
                dPreis = rsrs!APREIS
            Else
                cPreis = "0"
            End If
            
            cPreis = Format$(cPreis, "#####0.00")
            
            dAnteil = Format((dPreis * 100) / dSum, "##0.00")
            cAnteil = CStr(dAnteil)
            cAnteil = cAnteil & " %"
            '***** erste Zeile *****
            
            cDaten = cAgn & Space$(6 - Len(cAgn)) & cAnteil & Space$(10 - Len(cAnteil)) & Space$(16 - Len(cMenge)) & cMenge
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            '***** zweite Zeile *****
            
            cDaten = Left(cAGNBEZEICH, 22) & Space$(22 - Len(Left(cAGNBEZEICH, 22))) & Space$(10 - Len(cPreis)) & cPreis
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            '***** dritte Zeile *****
            cDaten = String$(32, "-")
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            rsrs.MoveNext
        Loop
        
        rsrs.Close: Set rsrs = Nothing
        
        '//Leerzeilen
        For lcount = 1 To 9
            If lcount = 9 Then
                cEscapeSequenz = "." & vbCrLf
            Else
                cEscapeSequenz = " " & vbCrLf
            End If
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        Next lcount
        '//end
        
        If gbAPI = True Then
            OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
        Else
            OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
        End If
        
        'Bon-Daten sichern
        SichernBonDaten cDruckZeile(), lAnzZeile, "", "", False
        
        '//potong
        '//Aenderung
        'Kassenbon abschneiden
        If gbAPI = True Then
            aDeviceName = Printer.DeviceName
            cEscapeSequenz = gcSchneiden
            OpenDrawer aDeviceName, cEscapeSequenz
        End If
        '//End Aenderung
        
        
        '*******************************************************
        '*** Letzter Schritt: Umschalten auf ListenDrucker   ***
        '*** und ggf. Löschen der gedruckten Daten           ***
        '*******************************************************
        setzedrucker gcListenDrucker
        
    Else
    
        '******************************
        '*** Keine Daten vorhanden
        '******************************
        
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeKassenAgnAuswertungaufBondrucker"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeKassenAgnAuswertungaufListendrucker()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim cJahr As String
    Dim cMon As String
    Dim cVon As String
    Dim cBis As String
    
    Dim lAnz As Long
    Dim lcount As Long
    Dim aDeviceName As String
    Dim cEscapeSequenz As String
    Dim iLenZeile As Integer
    Dim cDaten As String
    Dim lAnzZeile As Long
    ReDim cDruckZeile(1 To 1) As String
    Dim iRet As Integer
    Dim dSum As Double
    
    Dim lWert1 As Long
    Dim lWert2 As Long
    Dim j As Integer
    Dim i As Integer


    '///////      bringe MWSt. Satz aus der Tabelle MWSTSATZ von Datenbank Kissdata.mdb
    
    Dim TempMwstV As Integer
    Dim TempMwstE As Integer
    Dim TempMwstO As Integer
    
    cSQL = "SELECT VOLL,ERM,OHNE FROM MWSTSATZ WHERE CDate('" & Date & "')> vonD AND bisD is NULL"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
        If Not rsrs.EOF Then
        
           rsrs.MoveFirst
           
               If Not IsNull(rsrs!VOLL) Then
                 TempMwstV = rsrs!VOLL
               End If
           
               If Not IsNull(rsrs!ERM) Then
                TempMwstE = rsrs!ERM
               End If
           
               If Not IsNull(rsrs!OHNE) Then
                TempMwstO = rsrs!OHNE
               End If
                 
        End If
        
    cSQL = ""
    rsrs.Close: Set rsrs = Nothing
        
    '/////// Ende bringe MWSt. Satz aus der Tabelle MWSTSATZ von Datenbank Kissdata.mdb
    
    
    Screen.MousePointer = 11
    
    loeschNEW "ABAGNC", gdBase
    CreateTable "ABAGNC", gdBase
    
    
    loeschNEW "ABAGN", gdBase
    CreateTable "ABAGN", gdBase
    
    
    
    
    cSQL = "insert into ABAGNC (agn,ust) values (100,'" & TempMwstE & "')"
    gdBase.Execute cSQL, dbFailOnError
    
'    cSQL = "insert into ABAGNC (agn,ust) values (2100,'" & TempMwstE & "')"
'    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (3000,'" & TempMwstE & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (2901,'" & TempMwstE & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "insert into ABAGNC (agn,ust) values (111,'" & TempMwstV & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (2900,'" & TempMwstV & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (8510,'" & TempMwstV & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (8610,'" & TempMwstV & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (9001,'" & TempMwstV & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (3017,'" & TempMwstV & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (8002,'" & TempMwstV & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (4000,'" & TempMwstV & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (6000,'" & TempMwstV & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (9500,'" & TempMwstV & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (9998,'" & TempMwstV & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (9999,'" & TempMwstV & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (8007,'" & TempMwstO & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (2000,'" & TempMwstO & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (2001,'" & TempMwstO & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    
    'neu
    cSQL = "insert into ABAGNC (agn,ust) values (2020,'" & TempMwstO & "')"
    gdBase.Execute cSQL, dbFailOnError
    cSQL = "insert into ABAGNC (agn,ust) values (2019,'" & TempMwstO & "')"
    gdBase.Execute cSQL, dbFailOnError
    'neu ende
    
    
    
    
    cSQL = "insert into ABAGNC (agn,ust) values (5101,'" & TempMwstO & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (8000,'" & TempMwstO & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (8001,'" & TempMwstO & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (8006,'" & TempMwstO & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (7000,'" & TempMwstO & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "insert into ABAGNC (agn,ust) values (2100,'" & TempMwstO & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update ABAGNC  Set apreis3 = 0"
    cSQL = cSQL & ", apreis2 = 0 "
    cSQL = cSQL & ", apreis1 = 0 "
    cSQL = cSQL & ", abs1 = 0 "
    cSQL = cSQL & ", abs2 = 0 "
    cSQL = cSQL & ", abs3 = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    If Option1(0).value = True Then
        cMon = Val(MaskEdBox1(0).Text)
        cJahr = Val(MaskEdBox1(1).Text)
        cDaten = "ABSCHLUSS   " & cMon & "/" & cJahr
        
        'Speicher hier Monat/Jahr
    Else
        cVon = Text1(0).value
        cBis = Text1(1).value

        cDaten = "ABSCHLUSS " & Format$(cVon, "DD.MM.YY") & " - " & Format$(cBis, "DD.MM.YY")
        'Speicher Zeitraum
        
    End If
    cSQL = "Update ABAGNC  Set DatText = '" & cDaten & "'"
    gdBase.Execute cSQL, dbFailOnError
    

    
    For i = 1 To 3
        loeschNEW "ABAGN", gdBase
        CreateTable "ABAGN", gdBase
        
        cSQL = "insert into ABAGN Select artnr, preis, menge, agn from Kassjour where  "
        If Option1(0).value = True Then
            lWert1 = Val(MaskEdBox1(0).Text)
            lWert2 = Val(MaskEdBox1(1).Text)
            cSQL = cSQL & "YEAR(ADATE) = " & Trim$(Str$(lWert2)) & " "
            cSQL = cSQL & "and MONTH(ADATE) = " & Trim$(Str$(lWert1)) & " "
            'Speicher hier Monat/Jahr
        Else
            lWert1 = Text1(0).value '     DateValue(MaskEdBox2(0).Text)
            lWert2 = Text1(1).value 'DateValue(MaskEdBox2(1).Text)
            cSQL = cSQL & "ADATE >= " & Trim$(Str$(lWert1)) & " "
            cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lWert2)) & " "
            'Speicher Zeitraum
        End If
        cSQL = cSQL & " and KASNUM =  " & i
        gdBase.Execute cSQL, dbFailOnError
        
        loeschNEW "Agtemp", gdBase
        
        Select Case i
            Case 1
                cSQL = "Select agn,sum(preis) as apreis1 into Agtemp from ABAGN group by agn "
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Update ABAGNC inner join Agtemp on ABAGNC.agn = Agtemp.agn Set ABAGNC.apreis1 = Agtemp.apreis1  "
                gdBase.Execute cSQL, dbFailOnError
            Case 2
                cSQL = "Select agn,sum(preis) as apreis2 into Agtemp from ABAGN group by agn "
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Update ABAGNC inner join Agtemp on ABAGNC.agn = Agtemp.agn Set ABAGNC.apreis2 = Agtemp.apreis2  "
                gdBase.Execute cSQL, dbFailOnError
            Case 3
                cSQL = "Select agn,sum(preis) as apreis3 into Agtemp from ABAGN group by agn "
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Update ABAGNC inner join Agtemp on ABAGNC.agn = Agtemp.agn Set ABAGNC.apreis3 = Agtemp.apreis3  "
                gdBase.Execute cSQL, dbFailOnError
        End Select
    Next i
    
    cSQL = "Update ABAGNC inner join AGNDBF on ABAGNC.AGN = AGNDBF.AGN set ABAGNC.AGNBEZEICH = AGNDBF.AGTEXT"
    gdBase.Execute cSQL, dbFailOnError
    
    'Kasse
    For i = 1 To 3
    cSQL = "select sum(geldwert) as maxi from ABSCHOPF where  "
    If Option1(0).value = True Then
        lWert1 = Val(MaskEdBox1(0).Text)
        lWert2 = Val(MaskEdBox1(1).Text)
        cSQL = cSQL & "YEAR(ADATE) = " & Trim$(Str$(lWert2)) & " "
        cSQL = cSQL & "and MONTH(ADATE) = " & Trim$(Str$(lWert1)) & " "
        'Speicher hier Monat/Jahr
    Else
        lWert1 = Text1(0).value '     DateValue(MaskEdBox2(0).Text)
        lWert2 = Text1(1).value 'DateValue(MaskEdBox2(1).Text)
        cSQL = cSQL & "ADATE >= " & Trim$(Str$(lWert1)) & " "
        cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lWert2)) & " "
        'Speicher Zeitraum
    End If
    
    cSQL = cSQL & " and kasnum = " & i
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!maxi) Then
            Select Case i
                Case 1
                    cSQL = "Update ABAGNC set ABAGNC.abs1 = '" & rsrs!maxi & "'"
                Case 2
                    cSQL = "Update ABAGNC set ABAGNC.abs2 = '" & rsrs!maxi & "'"
                Case 3
                    cSQL = "Update ABAGNC set ABAGNC.abs3 = '" & rsrs!maxi & "'"
            End Select
            
           
            gdBase.Execute cSQL, dbFailOnError
        End If
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Next i
    
    cSQL = "Update ABAGNC  Set apreisg = apreis3 + apreis2 + apreis1 "
    gdBase.Execute cSQL, dbFailOnError
    
    reportbildschirm "", "aWKL21m"
        
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeKassenAgnAuswertungaufListendrucker"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub DruckeTagesAbschlussAufBonDruckerWK21f()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim lAnz As Long
    Dim lAktSatz As Long
    Dim lAnzSatz As Long
    Dim bReturn As Boolean
    Dim cDrucker As String

    Dim aDeviceName As String
    Dim cEscapeSequenz As String
    Dim cDaten As String
    Dim dSumme As Double
    
    Dim iLenZeile As Integer
    Dim lAnzZeile As Long
    
    Dim cSQL As String
    Dim rsrs As Recordset
    ReDim cFeldName(0 To 0) As String
    ReDim iFeldPos(0 To 0) As Integer
    ReDim cFeldDruck(0 To 0) As String
    
    ReDim cDruckZeile(1 To 1) As String
    
    '***********************************************************
    '*** Lese Konfiguration des Tagesprotokolls ***
    '***********************************************************
    Tabcheck "ZBON"
    
    cSQL = "Select * from ZBONLay where ANZEIGE = 'J' and Tabname = 'ZBON'"
    cSQL = cSQL & " order by Reihenf "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        
        lAnzSatz = rsrs.RecordCount
        
        ReDim cFeldName(1 To lAnzSatz) As String
        ReDim iFeldPos(1 To lAnzSatz) As Integer
        ReDim cFeldDruck(1 To lAnzSatz) As String
        
        rsrs.MoveFirst
        lAktSatz = 0
        Do While Not rsrs.EOF
            lAktSatz = lAktSatz + 1
            If Not IsNull(rsrs!Spaltenbez) Then
                cFeldName(lAktSatz) = rsrs!Spaltenbez
            Else
                cFeldName(lAktSatz) = ""
            End If
            If Not IsNull(rsrs!Reihenf) Then
                iFeldPos(lAktSatz) = rsrs!Reihenf
            Else
                iFeldPos(lAktSatz) = -1
            End If
            If Not IsNull(rsrs!anzeige) Then
                cFeldDruck(lAktSatz) = rsrs!anzeige
            Else
                cFeldDruck(lAktSatz) = "N"
            End If
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    '********************************************
    '*** 1.Schritt: Umschalten auf BonDrucker ***
    '********************************************
    
    setzedrucker gcBonDrucker


    '********************************************************
    '*** 2.Schritt: Drucker an, Display aus, Init Drucker ***
    '********************************************************

    aDeviceName = Printer.DeviceName
    cEscapeSequenz = Chr$(27) + Chr$(61) + Chr$(1) + Chr$(27) + Chr$(64)
    OpenDrawer aDeviceName, cEscapeSequenz


    iLenZeile = 32
    'Drucker ist bereits auf BonDrucker geschaltet
    aDeviceName = gcBonDrucker

    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "K.I.S.S. Warenwirtschaft"
    Else
        cDaten = gcBonText(0)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To 1) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "Carsten Schröder"
    Else
        cDaten = gcBonText(1)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO-VERSION!"
    Else
        cDaten = gcBonText(4)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Kopfdaten 4.Zeile an Drucker senden
    '***********************************************

    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION"
    Else
        cDaten = gcBonText(12)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '******************************************************************
    
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    
    Dim cVon As String
    Dim cBis As String
    
    Dim cMon As String
    Dim cJahr As String
    
    If Option1(0).value = True Then
        cMon = Val(MaskEdBox1(0).Text)
        cJahr = Val(MaskEdBox1(1).Text)
        cDaten = "ABSCHLUSS   " & cMon & "/" & cJahr
        
        'Speicher hier Monat/Jahr
    Else
        cVon = Text1(0).value
        cBis = Text1(1).value
'        cVon = DateValue(MaskEdBox2(0).Text)
'        cBis = DateValue(MaskEdBox2(1).Text)
        
        cDaten = "ABSCHLUSS " & Format$(cVon, "DD.MM.YY") & " - " & Format$(cBis, "DD.MM.YY")
        'Speicher Zeitraum
        
    End If
    
    
    
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    '*****Kasse?
    Dim bAus As Boolean
    Dim bChecky As Boolean
    
    bAus = False
    For lcount = 0 To 7
        If Check1(lcount).value = vbUnchecked Then
            bAus = True
            Exit For
        End If
    Next lcount
    
    If bAus Then
        cDaten = "Kasse: "
        For lcount = 0 To 7
            If Check1(lcount).value = vbChecked Then
                If bChecky = False Then
                    cDaten = cDaten & lcount + 1
                Else
                    cDaten = cDaten & ", " & lcount + 1
                End If
                
                bChecky = True
            End If
        Next lcount
    Else
        cDaten = "Kasse: alle Kassen"
    End If
    '*******Kasse? Ende

'        cDaten = "Kasse: " & gcKasNum
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = " " & vbCrLf
        
    '******************************************************************************
    '***** An Bondrucker in gewünschter Reihenfolge und ohne ausgeblendete Zeilen *
    '******************************************************************************
    
    For lAktSatz = 1 To lAnzSatz
        If cFeldDruck(lAktSatz) = "J" Then
            Select Case cFeldName(lAktSatz)
                Case Is = "Umsatz gesamt:"
                    '******************************************************************
                    'Umsatz gesamt
                    '******************************************************************
                    cDaten = Trim$(Label3(0).Caption)
                    cDaten = Space$(15 - Len(cDaten)) & cDaten
                    cDaten = "Umsatz gesamt:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Bar"
                    '******************************************************************
                    '+ Bar
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(1).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Bar:             " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Schecks"
                    '******************************************************************
                    '+ Schecks
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(2).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Schecks:         " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz

                Case Is = "+ Kredite"
                    '******************************************************************
                    '+ Kredite
                    
                    cDaten = Trim$(Label3(3).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Kredite:         " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Kreditkarte"
                    '******************************************************************
                    '+ Kreditkarten
                    
                    cDaten = Trim$(Label3(4).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Kreditkarten:    " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "+ Umsatz aus Gutscheinen:"
                    '******************************************************************
                    '+ Gutscheine
                    
                    cDaten = Trim$(Label3(20).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Umsatz Gutsch:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Lastschrift"
                    '******************************************************************
                    '+ Lastschrift
                    
                    cDaten = Trim$(Label3(34).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Lastschrift:     " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "+ Dukaten"
                    '******************************************************************
                    '+ Dukaten
                    
                    cDaten = Trim$(Label3(50).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Dukaten:         " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Kassensoll:"
                    '******************************************************************
                    'Kassensoll
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(8).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Kassensoll:        " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Bargeld:"
                    '******************************************************************
                    '+ Bargeld
                    '******************************************************************
                    cDaten = Trim$(Label3(9).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Bargeld:         " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Barverkäufe"
                    '******************************************************************
                    '+ Barverkäufe
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(25).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  + Barverkäufe:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                
                Case Is = "+ Einzahlungen:"
                    '******************************************************************
                    '+ Einzahlungen
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(6).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  + Einzahlungen:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ GutschVk Bar"
                    '******************************************************************
                    '+ Gutscheinverkäufe gegen Bar
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(42).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  + GutschVk Bar:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "+ Tilgung Bar"
                    '******************************************************************
                    '+ Tilgung gegen Bar
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(43).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  + Tilgung Bar:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "+ Wechselgeld:"
                    '******************************************************************
                    '+ Tilgung gegen Bar
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(46).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  + Wechselgeld:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Differenz summiert"
                    '******************************************************************
                    'Differenz summiert
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(49).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    
                    cDaten = "Differenz summiert " & cDaten
                    
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Differenz jetzt:"
                    '******************************************************************
                    'Differenz jetzt:
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(48).Caption)
                    
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    
                    cDaten = "Differenz jetzt:   " & cDaten
                
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                    
                Case Is = "- Abschöpfung:"
                    '******************************************************************
                    '- Auszahlungen
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(47).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  - Abschöpfung:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "- Auszahlungen:"
                    '******************************************************************
                    '- Auszahlungen
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(7).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  - Auszahlungen:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                
                Case Is = "- Gutscheinauszahlungen:"
                    '******************************************************************
                    '- Gutscheinauszahlungen
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(41).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "  - GutschAuszahl: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                
                Case Is = "+ Schecks:"
                    '******************************************************************
                    '+ Schecks
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(26).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ Schecks:         " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Kredit-Tilgungen"
                    '******************************************************************
                    'Tilgung
                    '******************************************************************
                    cDaten = Trim$(Label3(27).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Kredittilgung:     " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Tilgung in Bar"
                    '******************************************************************
                    '+ Tilgung in Bar
                    '******************************************************************
                    cDaten = Trim$(Label3(28).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ in Bar:          " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "+ Tilgung per Scheck"
                    '******************************************************************
                    '+ Tilgung per Scheck
                    '******************************************************************
                    cDaten = Trim$(Label3(29).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ per Scheck:      " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Tilgung per Gutschein"
                    '******************************************************************
                    '+ Tilgung per Gutschein
                    '******************************************************************
                    cDaten = Trim$(Label3(30).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ per Gutschein:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Tilgung per Karte"
                    '******************************************************************
                    '+ Tilgung per Kreditkarte
                    '******************************************************************
                    cDaten = Trim$(Label3(31).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "+ per Kreditkarte: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "eingereichte Gutscheine:"
                    '******************************************************************
                    'eingelöste Gutsch
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(32).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "eingelöste Gutsch. " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz

                Case Is = "daraus Rest-Gutscheine:"
                    '******************************************************************
                    'Restgutscheine
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(33).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "daraus RestGutsch. " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Kundenzahl:"
                    '******************************************************************
                    'Kundenzahl
                    '******************************************************************
                    cDaten = Trim$(Label3(10).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Kundenzahl:        " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Kundenschnitt:"
                    '******************************************************************
                    'Kundenschitt
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(11).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Kundenschnitt:     " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Sonderpreise:"
                    '******************************************************************
                    'Sonderpreise
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(12).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Sonderpreise:      " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Sonderpreise Anz.:"
                    '******************************************************************
                    'Sonderpreise Anz
                    '******************************************************************
                    cDaten = Trim$(Label3(13).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Sonderpreise Anz.: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Schublade geöffnet:"
                    '******************************************************************
                    'Schublade offen
                    '******************************************************************
                    cDaten = Trim$(Label3(14).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Schublade offen:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Artikelrabatt:"
                    '******************************************************************
                    'Artikelrabatt
                    '******************************************************************
                    cDaten = Trim$(Label3(15).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Artikelrabatt:     " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Artikelrabatt Anz.:"
                    '******************************************************************
                    'Gesamtrabatt Anz
                    '******************************************************************
                    cDaten = Trim$(Label3(53).Caption)
'                    If Len(cDaten) > 13 Then
'                        cMeld = "Zu großer Wert bei 'Artikelrabatt Anz.' (mehr als 13 Stellen)!" & vbCrLf
'                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
'                        MsgBox cMeld, vbCritical, "STOP!"
'                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Artikelrabatt Anz.:" & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Gesamtrabatt:"
                    '******************************************************************
                    'Gesamtrabatt
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(16).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gesamtrabatt:      " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Gesamtrabatt Anz.:"
                    '******************************************************************
                    'Gesamtrabatt Anz
                    '******************************************************************
                    cDaten = Trim$(Label3(17).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gesamtrabatt Anz.: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Stornosumme:"
                    '******************************************************************
                    'Stornosumme
                    '******************************************************************
                    cDaten = Trim$(Label3(18).Caption)
                    cDaten = Space$(14 - Len(cDaten)) & cDaten
                    cDaten = "Stornosumme:      " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Storno Anz.:"
                    '******************************************************************
                    'Storno Anz
                    '******************************************************************
                    cDaten = Trim$(Label3(19).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Storno Anz.:       " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                
                
                
                Case Is = "Gutschein-Verkäufe:"
                    '******************************************************************
                    'Gutschein-Verkauf Gesamt
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(5).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein-Verkauf: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Gutscheine über Bar"
                    '******************************************************************
                    'Gutschein-Verkauf BAR
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(21).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein Bar:     " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Gutscheine über Scheck"
                    '******************************************************************
                    'Gutschein-Verkauf SCHECK
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(22).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein Scheck:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Gutscheine über Kredite"
                    '******************************************************************
                    'Gutschein-Verkauf KREDIT
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(23).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein Kredit:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Gutscheine über Karte"
                    '******************************************************************
                    'Gutschein-Verkauf KARTE
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(24).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein Karte:   " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "+ Gutscheine über Lastschrift"
                    '******************************************************************
                    'Gutschein-Verkauf LASTSCHRIFT
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(35).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein Lastschr:" & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "+ Gutscheine über Gutscheine"
                    '******************************************************************
                    'Gutschein-Verkauf Gutscheine
                    '******************************************************************
                    
                    cDaten = Trim$(Label3(45).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Gutschein Gutschei:" & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Umsatz volle MWSt:"
                    '******************************************************************
                    'Umsatz volle MWSt
                    '******************************************************************
                    cDaten = Trim$(Label3(36).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Umsatz volle MWSt: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Betrag volle MWSt:"
                    '******************************************************************
                    'Betrag volle MWSt
                    '******************************************************************
                    cDaten = Trim$(Label3(37).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Betrag volle MWSt: " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Umsatz erm. MWSt:"
                    '******************************************************************
                    'Umsatz erm. MWSt
                    '******************************************************************
                    cDaten = Trim$(Label3(38).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Umsatz erm. MWSt:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Betrag erm. MWSt:"
                    '******************************************************************
                    'Betrag erm. MWSt
                    '******************************************************************
                    cDaten = Trim$(Label3(39).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Betrag erm. MWSt:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Is = "Umsatz ohne MWSt:"
                    '******************************************************************
                    'Umsatz ohne MWSt
                    '******************************************************************
                    cDaten = Trim$(Label3(40).Caption)
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Umsatz ohne MWSt:  " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Gesamtumsatz:"
                    
                    '******************************************************************
                    'Gesamtumsatz(E,V,O)
                    '******************************************************************
                    
                    dSumme = CDbl(Left(Label3(40).Caption, Len(Label3(40).Caption) - Len(gcWaehrung) - 1))
                    dSumme = dSumme + CDbl(Left(Label3(38).Caption, Len(Label3(38).Caption) - Len(gcWaehrung) - 1))
                    dSumme = dSumme + CDbl(Left(Label3(36).Caption, Len(Label3(36).Caption) - Len(gcWaehrung) - 1))
                    
                    cDaten = Format$(dSumme, "####,##0.00") & " " & gcWaehrung
                    cDaten = Space$(19 - Len(cDaten)) & cDaten
                    
                    cDaten = "Gesamtumsatz:" & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                Case Is = "Kreditkarte gesamt"
                    '******************************************************************
                    'Kreditkarte gesamt
                    '******************************************************************
                    cDaten = Trim$(Label3(52).Caption)
'                    If Len(cDaten) > 13 Then
'                        cMeld = "Zu großer Wert bei 'Kreditkarte gesamt' (mehr als 13 Stellen)!" & vbCrLf
'                        cMeld = cMeld & "Bon kann nicht gedruckt werden!"
'                        MsgBox cMeld, vbCritical, "STOP!"
'                    End If
                    cDaten = Space$(13 - Len(cDaten)) & cDaten
                    cDaten = "Kreditkarte gesamt " & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                Case Else
'                    MsgBox cFeldName(lAktSatz)
            End Select
        End If
    Next lAktSatz
        
    '******************************************************************
    'nicht umsatzrelevante Verkäufe
    '******************************************************************
    cDaten = Trim$(Label3(44).Caption)
    cDaten = Space$(13 - Len(cDaten)) & cDaten
    cDaten = "nicht umsatzrelv.: " & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    
    '******************************************************************
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = " " & vbCrLf
        
    
    '******************************************************************
    
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    cDaten = Format$(Now, "DD.MM.YYYY") & "                 " & Format$(Now, "HH:MM")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "KEIN GÜLTIGER KASSENBON!"
    Else
        cDaten = gcBonText(2)
    End If
    
    cDaten = Trim$(cDaten)
    If cDaten <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(3)
    End If
    
    
    cDaten = Trim$(cDaten)
    If cDaten <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = ""
    Else
        cDaten = gcBonText(5)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If

    
    '******************************************************************
    
    For lcount = 1 To 9
        cEscapeSequenz = vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
'        OpenDrawer aDeviceName, cEscapeSequenz
    Next lcount
    
    
    If gbAPI = True Then
        OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
    Else
        OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
    End If
    
    Erase cDruckZeile
    
BON_SCHNEIDEN:

    'Kassenbon abschneiden
    If gbAPI Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = Chr$(27) + Chr$(105)
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
ENDE:

    '...und tschüß!

    '*******************************************************
    '*** Letzter Schritt: Umschalten auf ListenDrucker   ***
    '*******************************************************
    
    setzedrucker gcListenDrucker

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeTagesAbschlussAufBonDruckerWK21f"
    Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
    
End Sub

Private Sub DruckeTagesabschlussNeuWK21f(iAuswahl As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim lcount As Long
    Dim rsrs As Recordset
    Dim ctmp As String
    Dim cTmp2 As String
    Dim cDaten As String
    Dim dKey As Double
    Dim cPfad As String
    Dim cWaeCode As String
    
    
    If gcWaehrung = gcWaehrung Then
        cWaeCode = "alle Preise kummuliert in " & gcWaehrung
    Else
        cWaeCode = "alle Preise kummuliert in EURO"
    End If
    
    '****************************************
    '* dieser Teil wird immer durchlaufen!  *
    '****************************************
    
    loeschNEW "Tagkopf", gdBase
    
    cSQL = "Create Table TAGKOPF "
    cSQL = cSQL & "(SCHLUESSEL double"
    cSQL = cSQL & ", WAE_CODE TEXT(30)"
    cSQL = cSQL & ", DATEN1 TEXT(50)"
    cSQL = cSQL & ", DATEN2 TEXT(50)"
    cSQL = cSQL & ", DATEN3 TEXT(50)"
    cSQL = cSQL & ", DATEN4 TEXT(50)"
    cSQL = cSQL & ", DATEN5 TEXT(50)"
    cSQL = cSQL & ", DATEN6 TEXT(50)"
    cSQL = cSQL & ", DATEN7 TEXT(50)"
    cSQL = cSQL & ", DATEN8 TEXT(50)"
    cSQL = cSQL & ", DATEN9 TEXT(50)"
    cSQL = cSQL & ", DATEN10 TEXT(50)"
    cSQL = cSQL & ", DATEN11 TEXT(50)"
    cSQL = cSQL & ", DATEN12 TEXT(50)"
    cSQL = cSQL & ", DATEN13 TEXT(50)"
    cSQL = cSQL & ", DATEN14 TEXT(50)"
    cSQL = cSQL & ", DATEN15 TEXT(50)"
    cSQL = cSQL & ", DATEN16 TEXT(50)"
    cSQL = cSQL & ", DATEN17 TEXT(50)"
    cSQL = cSQL & ", DATEN18 TEXT(50)"
    cSQL = cSQL & ", DATEN19 TEXT(50)"
    cSQL = cSQL & ", DATEN20 TEXT(50)"
    cSQL = cSQL & ", DATEN21 TEXT(50)"
    cSQL = cSQL & ", DATEN22 TEXT(50)"
    cSQL = cSQL & ", DATEN23 TEXT(50)"
    cSQL = cSQL & ", DATEN24 TEXT(50)"
    cSQL = cSQL & ", DATEN25 TEXT(50)"
    cSQL = cSQL & ", DATEN26 TEXT(50)"
    cSQL = cSQL & ", DATEN27 TEXT(50)"
    cSQL = cSQL & ", DATEN28 TEXT(50)"
    cSQL = cSQL & ", DATEN29 TEXT(50)"
    cSQL = cSQL & ", DATEN30 TEXT(50)"
    cSQL = cSQL & ", DATEN31 TEXT(50)"
    cSQL = cSQL & ", DATEN32 TEXT(50)"
    cSQL = cSQL & ", DATEN33 TEXT(50)"
    cSQL = cSQL & ", DATEN34 TEXT(50)"
    cSQL = cSQL & ", DATEN35 TEXT(50)"
    cSQL = cSQL & ", DATEN36 TEXT(50)"
    cSQL = cSQL & ", DATEN37 TEXT(50)"
    cSQL = cSQL & ", DATEN38 TEXT(50)"
    cSQL = cSQL & ", DATEN39 TEXT(50)"
    cSQL = cSQL & ", DATEN40 TEXT(50)"
    cSQL = cSQL & ", DATEN41 TEXT(50)"
    cSQL = cSQL & ", DATEN42 TEXT(50)"
    cSQL = cSQL & ", DATEN43 TEXT(50)"
    cSQL = cSQL & ", DATEN44 TEXT(50)"
    cSQL = cSQL & ", DATEN45 TEXT(50)"
    cSQL = cSQL & ", DATEN47 TEXT(50)"
    cSQL = cSQL & ", DATEN48 TEXT(50)"
    cSQL = cSQL & ", DATEN49 TEXT(50)"
    cSQL = cSQL & ", DATEN50 TEXT(50)"
    cSQL = cSQL & ", DATEN51 TEXT(50)"
    cSQL = cSQL & ", DATEN52 TEXT(50)"
    cSQL = cSQL & ", DATEN53 TEXT(50)"
    cSQL = cSQL & ", DATEN54 TEXT(50)"
    cSQL = cSQL & ", DATEN TEXT(150)"
    cSQL = cSQL & ", ALTERABSCH TEXT(50)"
    cSQL = cSQL & ", NEUERABSCH TEXT(50)"
    cSQL = cSQL & ", DATEN55 TEXT(50)"
    cSQL = cSQL & ", DATEN56 TEXT(50)"
    
    cSQL = cSQL & ", DATEN57 TEXT(50)"
    cSQL = cSQL & ", DATEN58 TEXT(50)"
    cSQL = cSQL & ", DATEN59 TEXT(50)"
    
    
    
    cSQL = cSQL & ")"
    gdBase.Execute cSQL, dbFailOnError
   
    cSQL = "Drop Index SCHLUESSEL on TAGKOPF"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index SCHLUESSEL on TAGKOPF (SCHLUESSEL)"
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    Dim cGewählteKassen As String
    Dim bAlle As Boolean
    
    bAlle = True
    cGewählteKassen = "Alle"
    
    For lcount = 0 To 7
        If Check1(lcount).value = vbUnchecked Then
            bAlle = False
        End If
    Next lcount
    
    If bAlle = False Then
        cGewählteKassen = ""
        For lcount = 0 To 7
            If Check1(lcount).value = vbChecked Then
                cGewählteKassen = cGewählteKassen & CStr(lcount + 1) & ","
            End If
        Next lcount
    End If
    
    If Right(cGewählteKassen, 1) = "," Then
        cGewählteKassen = Left(cGewählteKassen, Len(cGewählteKassen) - 1)
    End If
    
    
    

    dKey = Now
    dKey = Fix(dKey * 1000)
    
    '****************************************
    '* dieser Teil produziert Kopfdaten!    *
    '****************************************
    
    If iAuswahl = 1 Or iAuswahl = 3 Then
    
        cSQL = "Select * from TAGKOPF"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If rsrs.EOF Then
            rsrs.AddNew
            rsrs!schluessel = dKey
            
            rsrs!WAE_CODE = cWaeCode

            '**************************************
            ' linke Seite der Kopfdaten
            '**************************************

            'Umsatz gesamt
            cTmp2 = Label3(0).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN1 = cTmp2
            
            '+ Bar
            cTmp2 = Label3(1).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN2 = cTmp2
            
            '+ Schecks
            cTmp2 = Label3(2).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN3 = cTmp2
            
            '+ Kredite
            cTmp2 = Label3(3).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN4 = cTmp2
            
            '+ Kreditkarten
            cTmp2 = Label3(4).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN5 = cTmp2
            
            '+ Gutscheine
            cTmp2 = Label3(20).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN6 = cTmp2
            
            '+ Lastschrift
            cTmp2 = Label3(34).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN7 = cTmp2
            
            '+ Dukaten
            cTmp2 = Label3(50).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN52 = cTmp2
            
            'Kassensoll
            cTmp2 = Label3(8).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN8 = cTmp2
            
            '+ Bargeld
            cTmp2 = Label3(9).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN9 = cTmp2
            
            '+ Barverkäufe
            cTmp2 = Label3(25).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN10 = cTmp2
            
            '+ Einzahlungen
            cTmp2 = Label3(6).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN11 = cTmp2
            
            '+ Wechselgeld
            cTmp2 = Label3(46).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN49 = cTmp2
            
            '- Abschöpfung
            cTmp2 = Label3(47).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN48 = cTmp2
            
            '- Auszahlungen
            cTmp2 = Label3(7).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN12 = cTmp2
            
            '- Gutschein Auszahlungen
            cTmp2 = Label3(41).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN13 = cTmp2
            
            '+ Schecks
            cTmp2 = Label3(26).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN14 = cTmp2
            
            'Kredittilgungen
            cTmp2 = Label3(27).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN15 = cTmp2
            
            '+ Tilgung Bar
            cTmp2 = Label3(28).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN16 = cTmp2
            
            '+ Tilgung Scheck
            cTmp2 = Label3(29).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN17 = cTmp2
            
            '+ Tilgung Gutschein
            cTmp2 = Label3(30).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN18 = cTmp2
            
            '+ Tilgung Karte
            cTmp2 = Label3(31).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN19 = cTmp2
            
            'Eingelöste Gutscheine
            cTmp2 = Label3(32).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN20 = cTmp2
            
            'Daraus Restgutscheine
            cTmp2 = Label3(33).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN21 = cTmp2
            
            'nicht umsatzrelevante Verkäufe
            cTmp2 = Label3(44).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN45 = cTmp2
            
            '**************************************
            ' rechte Seite der Kopfdaten
            '**************************************
            
            'Kundenzahl
            cTmp2 = Label3(10).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN22 = cTmp2
            
            'Kundenschnitt
            cTmp2 = Label3(11).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN23 = cTmp2
            
            'Sonderpreise
            cTmp2 = Label3(12).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN24 = cTmp2
            
            'Sonderpreise Anzahl
            cTmp2 = Label3(13).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN25 = cTmp2
            
            'Schublade geöffnet
            cTmp2 = Label3(14).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN26 = cTmp2
            
            'Artikelrabatt
            
            cTmp2 = Label3(15).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN27 = cTmp2
            
            'artikelrabatt Anz
            cTmp2 = Label3(53).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN54 = cTmp2
            
            'Gesamtrabatt
            cTmp2 = Label3(16).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN28 = cTmp2
            
            'Gesamtrabatt Anz
            cTmp2 = Label3(17).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN29 = cTmp2
            
            'Stornosumme
            cTmp2 = Label3(18).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN30 = cTmp2
            
            'Stornosumme Anz
            cTmp2 = Label3(19).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN31 = cTmp2
            
            'Gutschein Verkauf
            cTmp2 = Label3(5).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN32 = cTmp2
            
            'Gutschein in Bar
            
            cTmp2 = Label3(21).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN33 = cTmp2
            
            'Gutschein per Scheck
            cTmp2 = Label3(22).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN34 = cTmp2
            
            'Gutschein per Kredit
            cTmp2 = Label3(23).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN35 = cTmp2
            
            'Gutschein per Karte
            cTmp2 = Label3(24).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN36 = cTmp2
            
            'Gutschein per Lastschrift
            cTmp2 = Label3(35).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN37 = cTmp2
            
            'Gutschein per Gutschein
            cTmp2 = Label3(45).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN47 = cTmp2
            
            'Umsatz volle MWSt
            cTmp2 = Label3(36).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN38 = cTmp2
            'Betrag volle MWSt
            
            cTmp2 = Label3(37).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN39 = cTmp2
            
            'Umsatz erm. MWSt
            cTmp2 = Label3(38).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN40 = cTmp2
            
            'Betrag erm. MWSt
            cTmp2 = Label3(39).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN41 = cTmp2
            
            'Umsatz ohne MWSt
            cTmp2 = Label3(40).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN42 = cTmp2
            
            '+ GutschVk Bar
            cTmp2 = Label3(42).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN43 = cTmp2
            
            '+ Tilgung Bar
            cTmp2 = Label3(43).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN44 = cTmp2
            
            'differenz jetzt
            cTmp2 = Label3(48).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN50 = cTmp2
            
            'differenz kumu
            cTmp2 = Label3(49).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN51 = cTmp2
            
            'karte insgesamt
            cTmp2 = Label3(52).Caption
            cTmp2 = Trim$(cTmp2)
            rsrs!DATEN53 = cTmp2
            
            'Wechselgeld in anderen Kassenschubladen
            cTmp2 = Label1(53).Caption
            cTmp2 = Trim$(cTmp2)
            cTmp2 = SwapStr(cTmp2, "(", "")
            cTmp2 = SwapStr(cTmp2, ")", "")
            rsrs!DATEN55 = cTmp2
            
            'Kreditkarten, nichtumsrele verkäufe
            cTmp2 = Label3(51).Caption
            cTmp2 = Trim$(cTmp2)
            cTmp2 = SwapStr(cTmp2, "(", "")
            cTmp2 = SwapStr(cTmp2, ")", "")
            rsrs!DATEN56 = cTmp2
            
            'Wechselgeld in nicht abgerechneten Kassenschubladen
            cTmp2 = Label1(55).Caption
            cTmp2 = Trim$(cTmp2)
            cTmp2 = SwapStr(cTmp2, "(", "")
            cTmp2 = SwapStr(cTmp2, ")", "")
            rsrs!DATEN57 = cTmp2
            
            'Kassennummern nicht abgerechneter Kassenschubladen
            rsrs!DATEN58 = Label1(56).Caption
            
            'gewählte Kassen
            rsrs!DATEN59 = cGewählteKassen
            
            rsrs.Update
        End If
    
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
        
Exit Sub
LOKAL_ERROR:
    If err.Number = 3372 Or err.Number = 3051 Or err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DruckeTagesabschlussNeuWK21f"
        Fehler.gsFehlertext = "Im Programmteil Zusammenfassung der Tagesabschlüsse ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        

    End If
End Sub

