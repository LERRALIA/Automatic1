VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWKLam 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Winkiss Artikelinfo"
   ClientHeight    =   8625
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   11910
   Icon            =   "frmWKLam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'Kein
      Height          =   3975
      Left            =   0
      TabIndex        =   60
      Top             =   4440
      Width           =   6615
      Begin sevCommand3.Command Command1 
         Height          =   345
         Index           =   1
         Left            =   2400
         TabIndex        =   1
         Top             =   2790
         Width           =   375
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
         Caption         =   "-"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   345
         Index           =   0
         Left            =   3600
         TabIndex        =   2
         Top             =   2790
         Width           =   375
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
         Caption         =   "+"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   4440
         TabIndex        =   106
         Top             =   0
         Width           =   435
      End
      Begin VB.Label Label14 
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
         Left            =   4920
         TabIndex        =   105
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Summe Umsatz:"
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
         Left            =   3360
         TabIndex        =   104
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         Left            =   1680
         TabIndex        =   103
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Summe Anzahl:"
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
         Left            =   120
         TabIndex        =   102
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   5520
         TabIndex        =   101
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   5040
         TabIndex        =   100
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   4560
         TabIndex        =   99
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   4080
         TabIndex        =   98
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   97
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   96
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   95
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   94
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   93
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   2
         Left            =   1200
         TabIndex        =   92
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   1
         Left            =   720
         TabIndex        =   91
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Index           =   0
         Left            =   240
         TabIndex        =   90
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   24
         Left            =   5640
         TabIndex        =   89
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   5145
         TabIndex        =   88
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   4650
         TabIndex        =   87
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   4185
         TabIndex        =   86
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   3705
         TabIndex        =   85
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   3255
         TabIndex        =   84
         Top             =   2400
         Width           =   165
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   2775
         TabIndex        =   83
         Top             =   2400
         Width           =   165
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   2250
         TabIndex        =   82
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   1785
         TabIndex        =   81
         Top             =   2400
         Width           =   225
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1290
         TabIndex        =   80
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   840
         TabIndex        =   79
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   375
         TabIndex        =   78
         Top             =   2400
         Width           =   165
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   11
         Left            =   5520
         Top             =   2400
         Width           =   435
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   10
         Left            =   5040
         Top             =   2400
         Width           =   435
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   9
         Left            =   4560
         Top             =   2400
         Width           =   435
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   8
         Left            =   4080
         Top             =   2400
         Width           =   435
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   7
         Left            =   3600
         Top             =   2400
         Width           =   435
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   6
         Left            =   3120
         Top             =   2400
         Width           =   435
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   5
         Left            =   2640
         Top             =   2400
         Width           =   435
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   4
         Left            =   2160
         Top             =   2400
         Width           =   435
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   3
         Left            =   1680
         Top             =   2400
         Width           =   435
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   2
         Left            =   1200
         Top             =   2400
         Width           =   435
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   1
         Left            =   720
         Top             =   2400
         Width           =   435
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   0
         Left            =   240
         Top             =   2400
         Width           =   435
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   5040
         TabIndex        =   77
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   4560
         TabIndex        =   76
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   4080
         TabIndex        =   75
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   3600
         TabIndex        =   74
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   3120
         TabIndex        =   73
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   2640
         TabIndex        =   72
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1680
         TabIndex        =   71
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1200
         TabIndex        =   70
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   720
         TabIndex        =   69
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   240
         TabIndex        =   68
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   2160
         TabIndex        =   67
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   5520
         TabIndex        =   66
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   49
         Left            =   2880
         TabIndex        =   65
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
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
         Index           =   17
         Left            =   4920
         TabIndex        =   64
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
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
         Index           =   21
         Left            =   5880
         TabIndex        =   63
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
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
         Index           =   23
         Left            =   6360
         TabIndex        =   62
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Verkaufszahlen (Menge, Umsatz)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   6960
      TabIndex        =   53
      Top             =   4560
      Width           =   4815
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2775
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   11
         Cols            =   13
      End
      Begin sevCommand3.Command Command1 
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   123
         Top             =   240
         Width           =   375
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
         Caption         =   "-"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   345
         Index           =   3
         Left            =   1560
         TabIndex        =   124
         Top             =   240
         Width           =   375
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
         Caption         =   "+"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Filial-Verkäufe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   42
         Left            =   120
         TabIndex        =   57
         Top             =   0
         Width           =   4575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "JAHR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   600
         TabIndex        =   54
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6960
      TabIndex        =   49
      Top             =   3000
      Width           =   3375
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferant"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   14
         Left            =   720
         TabIndex        =   126
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bei:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   125
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   13
         Left            =   2640
         TabIndex        =   117
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "letzte Bestellung:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   45
         Left            =   -240
         TabIndex        =   108
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   12
         Left            =   1800
         TabIndex        =   107
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Außenstände"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   43
         Left            =   120
         TabIndex        =   58
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Stück(e) in Bestellung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   39
         Left            =   1200
         TabIndex        =   52
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   51
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "von diesem Artikel sind z.Zt."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   38
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   6960
      TabIndex        =   47
      Top             =   120
      Width           =   4815
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   27
         Left            =   120
         TabIndex        =   122
         Top             =   2280
         Width           =   4575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel zu beziehen bei:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   44
         Left            =   120
         TabIndex        =   59
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   6735
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   26
         Left            =   5400
         TabIndex        =   121
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "PGN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   23
         Left            =   4680
         TabIndex        =   120
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Merkmal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   22
         Left            =   2400
         TabIndex        =   119
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   25
         Left            =   3600
         TabIndex        =   118
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Farbbeschreibung:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   21
         Left            =   4560
         TabIndex        =   116
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   24
         Left            =   4560
         TabIndex        =   115
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "geräumt am:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   20
         Left            =   4200
         TabIndex        =   114
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   23
         Left            =   5400
         TabIndex        =   113
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Basisinformation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   41
         Left            =   0
         TabIndex        =   56
         Top             =   120
         Width           =   6375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   22
         Left            =   1560
         TabIndex        =   46
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   21
         Left            =   5400
         TabIndex        =   45
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   20
         Left            =   3600
         TabIndex        =   44
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   19
         Left            =   1560
         TabIndex        =   43
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Preisschutz:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   19
         Left            =   360
         TabIndex        =   42
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus OK:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   18
         Left            =   4320
         TabIndex        =   41
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Rabatt OK:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   17
         Left            =   2400
         TabIndex        =   40
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Geführt:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   39
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404000&
         Index           =   1
         X1              =   120
         X2              =   6600
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   18
         Left            =   3600
         TabIndex        =   38
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Räumung:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   15
         Left            =   2400
         TabIndex        =   37
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   17
         Left            =   1560
         TabIndex        =   36
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Bestand:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   35
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   16
         Left            =   3600
         TabIndex        =   34
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   15
         Left            =   1560
         TabIndex        =   33
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   14
         Left            =   3600
         TabIndex        =   32
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   13
         Left            =   1560
         TabIndex        =   31
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   12
         Left            =   3600
         TabIndex        =   30
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   11
         Left            =   1560
         TabIndex        =   29
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Einheit:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   13
         Left            =   2400
         TabIndex        =   28
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Inhalt:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   27
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   10
         Left            =   5400
         TabIndex        =   26
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   9
         Left            =   3600
         TabIndex        =   25
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   24
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "EANs:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   23
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   7
         Left            =   5400
         TabIndex        =   22
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "MWSt.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   10
         Left            =   4560
         TabIndex        =   21
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   20
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Lief.-Best.-Nr.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   19
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   18
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Haupt-Lieferant:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404000&
         Index           =   0
         X1              =   120
         X2              =   6600
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Linie:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   16
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Kassen-VK:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   15
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Listen-VK:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   14
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Schnitt-EK:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   12
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Listen-EK:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Art.-Gruppe:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   840
         Width           =   5535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Bezeichnung:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Art.Nr.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   495
      Left            =   7080
      TabIndex        =   0
      Top             =   8160
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Zurück"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   255
      Index           =   1
      Left            =   10440
      TabIndex        =   109
      Top             =   3360
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "lt. Verkauf"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   255
      Index           =   2
      Left            =   10440
      TabIndex        =   110
      Top             =   3720
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Lagerumschlag"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   255
      Index           =   0
      Left            =   10440
      TabIndex        =   111
      Top             =   3000
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "lt. Einkauf"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   255
      Index           =   3
      Left            =   10440
      TabIndex        =   112
      Top             =   4080
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "lt. Einkaufspreis"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404000&
      Index           =   6
      X1              =   240
      X2              =   6720
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404000&
      Index           =   3
      X1              =   7080
      X2              =   11640
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404000&
      Index           =   4
      X1              =   6960
      X2              =   11760
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404000&
      Index           =   5
      X1              =   6840
      X2              =   6840
      Y1              =   240
      Y2              =   8520
   End
End
Attribute VB_Name = "frmWKLam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gitop           As Integer
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
Private Sub FormatiereGridWKLam()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim ctmp As String
    Dim lAnzFil As Long
    Dim lFil As Long
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    MSFlexGrid1.Rows = giAnzFil + 1
    MSFlexGrid1.Cols = 15
    
    For lcount = 1 To 12
        MSFlexGrid1.TextMatrix(0, lcount) = UCase(Left(gcMonat(lcount), 3))
    Next lcount
    
    cSQL = "Select * from FILIALEN order by FILIALNR"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        lAnzFil = 0
        Do While Not rsrs.EOF
            lAnzFil = lAnzFil + 1
            If Not IsNull(rsrs!Filialname) Then
                ctmp = rsrs!Filialname
            Else
                ctmp = ""
            End If
            MSFlexGrid1.TextMatrix(lAnzFil, 0) = ctmp
            
            MSFlexGrid1.TextMatrix(lAnzFil, 13) = rsrs!FILIALNR
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    MSFlexGrid1.ColWidth(0) = 1000
    For lcount = 1 To 12
        MSFlexGrid1.ColWidth(lcount) = 550
    Next lcount
    MSFlexGrid1.ColWidth(13) = 0
    MSFlexGrid1.ColWidth(14) = 800
    
    
    For lcount = 1 To 14
        For lFil = 1 To lAnzFil
            If lcount = 13 Then
            
            Else
                MSFlexGrid1.TextMatrix(lFil, lcount) = "0"
            End If
        Next lFil
    Next lcount
    
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatiereGridWKLam"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub FuelleFlexGrid1WKLam(cJahr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr As Integer
    Dim cArtNr As String
    Dim lFil As Long
    Dim lMonat As Long
    Dim lAnzahl As Long
    Dim lJahresAnzahl As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim i As Integer
    
    FormatiereGridWKLam
    
    Label5.Caption = cJahr
    
    cArtNr = gcArtNrFiliale

    If Trim(cArtNr) = "" Then
        Exit Sub
    End If
    
    cSQL = "Select * from UMS_ARTF where "
    cSQL = cSQL & "ARTNR = " & cArtNr & " "
    cSQL = cSQL & "and JAHR = " & cJahr & " "
    cSQL = cSQL & "order by FILIALNR, MONAT "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!FILIALNR) Then
                lFil = rsrs!FILIALNR
            Else
                lFil = 0
            End If
            
            If Not IsNull(rsrs!Monat) Then
                lMonat = rsrs!Monat
            Else
                lMonat = 0
            End If
            
            If Not IsNull(rsrs!ANZAHL) Then
                lAnzahl = rsrs!ANZAHL
            Else
                lAnzahl = 0
            End If
            
           
            
            If lFil <> 0 And lMonat <> 0 Then
                MSFlexGrid1.Col = 13
                
                For i = 1 To MSFlexGrid1.Rows - 1
                
                    MSFlexGrid1.Row = i
                    If lFil = MSFlexGrid1.Text Then
                        MSFlexGrid1.Col = lMonat
                        MSFlexGrid1.Text = Format$(lAnzahl, "#####0")
                        
                        MSFlexGrid1.Col = 13
                    End If
        
                Next i
            End If
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    
    

    
    cSQL = "Select sum(anzahl) as Maxi , Filialnr from UMS_ARTF where "
    cSQL = cSQL & "ARTNR = " & cArtNr & " "
    cSQL = cSQL & "and JAHR = " & cJahr & " "
    cSQL = cSQL & "group by FILIALNR "

    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!FILIALNR) Then
                lFil = rsrs!FILIALNR
            Else
                lFil = 0
            End If

            If Not IsNull(rsrs!maxi) Then
                lAnzahl = rsrs!maxi
            Else
                lAnzahl = 0
            End If

            If lFil <> 0 Then
                MSFlexGrid1.Col = 13
                
                For i = 1 To MSFlexGrid1.Rows - 1
                
                    MSFlexGrid1.Row = i
                    If lFil = MSFlexGrid1.Text Then
                        MSFlexGrid1.Col = 14
                        MSFlexGrid1.Text = Format$(lAnzahl, "#####0")
                        
                        MSFlexGrid1.Col = 13
                    End If
        
                Next i
            End If

            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    
    
    
    
    
    
Exit Sub
LOKAL_ERROR:

    If err.Number = 381 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "FuelleFlexgrid1WKLam"
        Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub LeereDialogWKLam()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    For lcount = 0 To 23
        Label2(lcount).Caption = ""
    Next lcount
    
    Label2(27).Caption = ""

    List1.Clear
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKLam"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseArtikelWKLam()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim rsRs2 As Recordset
    Dim lartnr As Long
    Dim lLinr As Long
    Dim ctmp As String
    Dim lPos As Long
    Dim dWert As Double
    Dim sSQL As String
    Dim sSQL1 As String
    
    lartnr = Val(gcArtNrFiliale)
    sSQL = "select * from artikel where artnr = " & lartnr
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!artnr) Then
            Label2(0).Caption = rsrs!artnr
        End If
        
        If Not IsNull(rsrs!AWM) Then
            ctmp = rsrs!AWM
        Else
            ctmp = "0"
        End If
        
        If ctmp <> "" Then
            If ctmp = "98" Then
                Label2(24).Caption = "Neu"
                Label2(24).BackColor = vbWhite
                Label2(24).ForeColor = vbRed
                
                Label2(0).BackColor = vbWhite
                Label2(0).ForeColor = vbRed
            ElseIf ctmp = "95" Then
                Label2(24).Caption = "nicht lieferbar"
                Label2(24).BackColor = vbBlue
                Label2(24).ForeColor = vbBlack
                
                Label2(0).BackColor = vbBlue
                Label2(0).ForeColor = vbBlack
            ElseIf ctmp = "94" Then
                Label2(24).Caption = "Preisaktion vorbereitet"
                Label2(24).BackColor = glfarbe(0)
                Label2(24).ForeColor = vbBlue
                
                Label2(0).BackColor = glfarbe(0)
                Label2(0).ForeColor = vbBlue
            ElseIf ctmp = "93" Then
                Label2(24).Caption = "Preisaktion jetzt"
                Label2(24).BackColor = vbWhite
                Label2(24).ForeColor = vbGreen
                
                Label2(0).BackColor = vbWhite
                Label2(0).ForeColor = vbGreen
            ElseIf ctmp = "92" Then
                Label2(24).Caption = "lange nicht verkauft"
                Label2(24).BackColor = &H80000012
                Label2(24).ForeColor = vbWhite
                
                Label2(0).BackColor = &H80000012
                Label2(0).ForeColor = vbWhite
            Else
                If CByte(ctmp) < 10 Then
                    Label2(24).BackColor = glfarbe(ctmp)
                    Label2(24).Caption = ermFarbeBez(ctmp)
                    
                    Label2(0).BackColor = glfarbe(ctmp)
                ElseIf CByte(ctmp) > 10 And CByte(ctmp) < 20 Then
                    Label2(24).BackColor = glfarbe2(CInt(ctmp) - 10)
                    Label2(24).Caption = ermFarbeBez(CStr(CInt(ctmp)))
                    
                    Label2(0).BackColor = glfarbe2(CInt(ctmp) - 10)
                Else
                    Label2(24).BackColor = glfarbe(0)
                    Label2(24).Caption = ""
                    
                    Label2(0).BackColor = glfarbe(0)
                End If
            End If
        Else
            Label2(24).BackColor = glfarbe(0)
            Label2(24).Caption = ""
            
            Label2(0).BackColor = glfarbe(0)
        End If
        
        
        
        If Not IsNull(rsrs!BEZEICH) Then
            Label2(1).Caption = rsrs!BEZEICH
        End If
        If Not IsNull(rsrs!AGN) Then
            Label2(2).Caption = rsrs!AGN
        End If
        If Not IsNull(rsrs!PGN) Then
            Label2(26).Caption = rsrs!PGN
        End If
        If Not IsNull(rsrs!LPZ) Then
            Label2(3).Caption = rsrs!LPZ
        End If
        If Not IsNull(rsrs!linr) Then
            Label2(4).Caption = rsrs!linr
            lLinr = rsrs!linr
            
            sSQL1 = "select * from lisrt where linr = " & lLinr
            Set rsRs2 = gdBase.OpenRecordset(sSQL1)
            If Not rsRs2.EOF Then
                If Not IsNull(rsRs2!LIEFBEZ) Then
                    ctmp = rsRs2!LIEFBEZ
                    If InStr(ctmp, "&") > 0 Then
                        lPos = InStr(ctmp, "&")
                        ctmp = Left(ctmp, lPos) & "&" & Mid(ctmp, lPos + 1, Len(ctmp) - lPos)
                    End If
                    Label2(5).Caption = ctmp
                End If
            End If
            rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
        End If
        If Not IsNull(rsrs!LIBESNR) Then
            Label2(6).Caption = rsrs!LIBESNR
        End If
        If Not IsNull(rsrs!MWST) Then
            Label2(7).Caption = rsrs!MWST
        End If
        If Not IsNull(rsrs!EAN) Then
            Label2(8).Caption = rsrs!EAN
        End If
        If Not IsNull(rsrs!EAN2) Then
            Label2(9).Caption = rsrs!EAN2
        End If
        If Not IsNull(rsrs!EAN3) Then
            Label2(10).Caption = rsrs!EAN3
        End If
        If Not IsNull(rsrs!INHALT) Then
            ctmp = rsrs!INHALT
            If ctmp = "0" Then
                ctmp = ""
            End If
            Label2(11).Caption = ctmp
        End If
        If Not IsNull(rsrs!INHALTBEZ) Then
            Label2(12).Caption = rsrs!INHALTBEZ
        End If
        If Not IsNull(rsrs!lekpr) Then
            dWert = rsrs!lekpr
            ctmp = Format$(dWert, "#####0.00")
            Label2(13).Caption = ctmp
        End If
        If Not IsNull(rsrs!ekpr) Then
            dWert = rsrs!ekpr
            ctmp = Format$(dWert, "#####0.00")
            Label2(14).Caption = ctmp
        End If
        If Not IsNull(rsrs!vkpr) Then
            dWert = rsrs!vkpr
            ctmp = Format$(dWert, "#####0.00")
            Label2(15).Caption = ctmp
        End If
        If Not IsNull(rsrs!KVKPR1) Then
            dWert = rsrs!KVKPR1
            ctmp = Format$(dWert, "#####0.00")
            Label2(16).Caption = ctmp
        End If
        If Not IsNull(rsrs!BESTAND) Then
            dWert = rsrs!BESTAND
            ctmp = Format$(dWert, "########0")
            Label2(17).Caption = ctmp
        End If
'        If Not IsNull(rsrs!RKZ) Then
'            ctmp = rsrs!RKZ
'            If ctmp = "J" Then
'                ctmp = "JA"
'            Else
'                ctmp = "NEIN"
'            End If
'            Label2(18).Caption = ctmp
'        End If
'
'        If Not IsNull(rsrs!EXDAT) Then
'            ctmp = rsrs!EXDAT
'            If IsDate(ctmp) Then
'                Label2(23).Caption = DateValue(ctmp)
'            End If
'        End If
    
        If Not IsNull(rsrs!GEFUEHRT) Then
            ctmp = rsrs!GEFUEHRT
            If ctmp = "J" Then
                ctmp = "JA"
            Else
                ctmp = "NEIN"
            End If
            Label2(19).Caption = ctmp
        End If
        If Not IsNull(rsrs!RABATT_OK) Then
            ctmp = rsrs!RABATT_OK
            If ctmp = "J" Then
                ctmp = "JA"
            Else
                ctmp = "NEIN"
            End If
            Label2(20).Caption = ctmp
        End If
        If Not IsNull(rsrs!BONUS_OK) Then
            ctmp = rsrs!BONUS_OK
            If ctmp = "J" Then
                ctmp = "JA"
            Else
                ctmp = "NEIN"
            End If
            Label2(21).Caption = ctmp
        End If
        If Not IsNull(rsrs!PREISSCHU) Then
            ctmp = rsrs!PREISSCHU
            If ctmp = "J" Then
                ctmp = "JA"
            Else
                ctmp = "NEIN"
            End If
            Label2(22).Caption = ctmp
        End If
        
        If Not IsNull(rsrs!NOTIZEN) Then
            Label2(27).Caption = rsrs!NOTIZEN
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    sSQL = "select * from artlief where artnr = " & lartnr
    sSQL = sSQL & " and linr = " & lLinr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!RKZ) Then
            ctmp = rsrs!RKZ
            If ctmp = "J" Then
                ctmp = "JA"
            Else
                ctmp = "NEIN"
            End If
            Label2(18).Caption = ctmp
        End If

        If Not IsNull(rsrs!EXDAT) Then
            ctmp = rsrs!EXDAT
            If IsDate(ctmp) Then
                Label2(23).Caption = DateValue(ctmp)
            End If
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseArtikelWKLam"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub LeseArtikelLieferantenWKLam()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsRs2 As Recordset
    Dim lLinr As Long
    Dim lartnr As Long
    Dim lPos As Long
    Dim ctmp As String
    Dim dWert As Double
    Dim cLBSatz As String
    Dim sSQL1 As String
    
    
    
    lartnr = Val(gcArtNrFiliale)
    cSQL = "Select * from ARTLIEF where ARTNR = " & Trim$(Str$(lartnr))
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!linr) Then
                lLinr = rsrs!linr
            Else
                lLinr = 0
            End If
            ctmp = Format$(lLinr, "#####0")
            ctmp = Space$(6 - Len(ctmp)) & ctmp
            cLBSatz = ctmp & " "
            
            sSQL1 = "select * from lisrt where linr = " & lLinr
            Set rsRs2 = gdBase.OpenRecordset(sSQL1)

            If Not rsRs2.EOF Then
                If Not IsNull(rsRs2!LIEFBEZ) Then
                    ctmp = rsRs2!LIEFBEZ
                    If InStr(ctmp, "&") > 0 Then
                        lPos = InStr(ctmp, "&")
                        ctmp = Left(ctmp, lPos) & "&" & Mid(ctmp, lPos + 1, Len(ctmp) - lPos)
                    End If
                End If
            End If
            rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
            cLBSatz = cLBSatz & ctmp
            List1.AddItem cLBSatz
            
            If Not IsNull(rsrs!LIBESNR) Then
                ctmp = rsrs!LIBESNR
            Else
                ctmp = ""
            End If
            ctmp = ctmp & Space$(13 - Len(ctmp))
            cLBSatz = "BestNr: " & ctmp & " "
            
            If Not IsNull(rsrs!lekpr) Then
                dWert = rsrs!lekpr
            Else
                dWert = 0
            End If
            ctmp = Format$(dWert, "#####0.00")
            ctmp = Space$(9 - Len(ctmp)) & ctmp
            cLBSatz = cLBSatz & "Preis: " & ctmp
            List1.AddItem cLBSatz
            
            
            Dim sRKZ As String
            Dim sExdat As String
            
            If Not IsNull(rsrs!RKZ) Then
                sRKZ = rsrs!RKZ
            Else
                sRKZ = ""
            End If
            
            
            
            ctmp = sRKZ & Space$(2 - Len(sRKZ))
            cLBSatz = "Räumung: " & ctmp & " "
            
            If sRKZ = "J" Then
                If Not IsNull(rsrs!EXDAT) Then
                    sExdat = rsrs!EXDAT
                Else
                    sExdat = ""
                End If
                
                ctmp = sExdat
                cLBSatz = cLBSatz & "Datum: " & ctmp
            End If
            
            
            
            List1.AddItem cLBSatz
            
            
            
            
            
            
            List1.AddItem String$(60, "-")
            List1.AddItem " "
                        
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseArtikelLieferantenWKLam"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LeseBestellungWKLam()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lartnr As Long
    Dim lAnz As Long
    Dim cdat As String
    Dim cLinr As String
    
    lartnr = Val(gcArtNrFiliale)
    
    lAnz = 0
    cSQL = "Select SUM(BESTVOR) as BESTELLT from BESTREST where ARTNR = " & Trim$(Str$(lartnr))
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BESTELLT) Then
            lAnz = rsrs!BESTELLT
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
        
    Label4(0).Caption = Trim$(Str$(lAnz))
    
    
    cdat = ""
    cSQL = "Select max(BEST_datum) as BESTELLT from BESTREST where ARTNR = " & Trim$(Str$(lartnr))
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BESTELLT) Then
            cdat = rsrs!BESTELLT
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
        
    If IsDate(cdat) Then
        Label4(12).Caption = DateValue(cdat)
    Else
        Label4(12).Caption = ""
    End If
    
    Dim lDatum As Long
    
    If IsDate(cdat) Then
    
        lDatum = CLng(DateValue(cdat))
        
        cLinr = ""
        cSQL = "Select Linr from BESTREST where ARTNR = " & Trim$(Str$(lartnr))
        cSQL = cSQL & " and BEST_datum = " & Trim$(Str$(lDatum)) & ""
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            If Not IsNull(rsrs!linr) Then
                cLinr = rsrs!linr
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
            
        If Val(cLinr) > 0 Then
            Label4(14).Caption = ermLiefBez(CLng(cLinr))
        End If
        
    End If
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseBestellungWKLam"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
    
End Sub
Private Sub LeseZentBestellungWKLam()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lartnr As Long
    Dim lAnz As Long
    Dim cdat As String
    
    lartnr = Val(gcArtNrFiliale)
    
    lAnz = 0
    
    cSQL = "Select SUM(BESTVOR) as BESTELLT from ZBREST where ARTNR = " & Trim$(Str$(lartnr))
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BESTELLT) Then
            lAnz = rsrs!BESTELLT
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
        
    Label4(0).Caption = Trim$(Str$(lAnz))
    
    
    cdat = ""
    cSQL = "Select max(BEST_datum) as BESTELLT from ZBREST where ARTNR = " & Trim$(Str$(lartnr))
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BESTELLT) Then
            cdat = rsrs!BESTELLT
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
        
    If IsDate(cdat) Then
        Label4(12).Caption = DateValue(cdat)
    Else
        Label4(12).Caption = ""
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseZentBestellungWKLam"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub



Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iOld As Integer
    Dim iNew As Integer
    
    Dim iOld2 As Integer
    Dim iNew2 As Integer
    
    iOld = CInt(Label9(49).Caption)
    iOld2 = CInt(Label5.Caption)
    Select Case Index
    
        Case Is = 0 '+
            iNew = iOld + 1
            Label9(49).Caption = iNew
            Label9(0).Caption = iNew
            diagrammfuellen gcArtNrFiliale, Label9(49).Caption
        Case Is = 1 '-
            iNew = iOld - 1
            Label9(49).Caption = iNew
            Label9(0).Caption = iNew
            diagrammfuellen gcArtNrFiliale, Label9(49).Caption
        Case Is = 2 '-
            iNew2 = iOld2 - 1
            Label5.Caption = iNew2
            FuelleFlexGrid1WKLam Label5.Caption
        Case Is = 3 '+
            iNew2 = iOld2 + 1
            Label5.Caption = iNew2
            FuelleFlexGrid1WKLam Label5.Caption
            
    End Select
    Label5.Refresh
    Label9(0).Refresh
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim sJahr As String
    Screen.MousePointer = 11
    
    PositionierenWKLam
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    Label9(49).Caption = Year(Now)
    Label9(49).Refresh
    
    Label9(0).Caption = Year(Now)
    Label9(0).Refresh
        
    LeereDialogWKLam
    
    gitop = Shape1(0).Top
    
    If gbFilNr Then
'        FormatiereGridWKLam
        If gcFilNr <> 0 Then
            Frame5.Visible = True
        Else
            Frame5.Visible = False
        End If
    End If
    
    If Trim(gcArtNrFiliale) <> "" Then
    
        LeseArtikelWKLam
        
        Label2(25).Caption = ZeigeArtmerk(gcArtNrFiliale)
        Label4(13).Caption = Label2(25).Caption
        
        LeseArtikelLieferantenWKLam
        
        If Trim$(gcFilNr) = "0" Then
            LeseBestellungWKLam
        Else
        
            If NewTableSuchenDBKombi("ZBREST", gdBase) Then
                LeseZentBestellungWKLam
            Else
                Label4(0).Caption = "keine Info"
                Label4(0).Refresh
                
                Label4(12).Caption = "keine Info"
                Label4(12).Refresh
                
            End If
        End If
        
        sJahr = Year(Now)
        
        If gbFilNr Then
            Label5.Caption = sJahr
            FuelleFlexGrid1WKLam Label5.Caption
        End If
        
        
        If gcArtNrFiliale <> "" Then
            If IsNumeric(gcArtNrFiliale) Then
                diagrammfuellen gcArtNrFiliale, sJahr
            End If
        End If
        
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub diagrammfuellen(sArtnr As String, sJahr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Long
    Dim j               As Long
    Dim iTop            As Long
    Dim myarr(0 To 11)  As Long
    Dim arrUm(0 To 11)  As Single
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim iMax            As Long
    Dim iBuffer         As Long
    Dim iSumAnz         As Long
    Dim siSumUm         As Single
    
    If sArtnr = "" Then
        Exit Sub
    End If
    
    cSQL = "Select * from UMS_ART  where ARTNR = " & sArtnr
    cSQL = cSQL & " and Jahr = " & sJahr
    
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            myarr(rsrs!Monat - 1) = rsrs!ANZAHL
            
            arrUm(rsrs!Monat - 1) = rsrs!UMSATZ
            rsrs.MoveNext
        Loop
    End If
    
    rsrs.Close: Set rsrs = Nothing
    Set rsrs = Nothing
        
    iBuffer = 0
    iMax = 0
    
    For i = 0 To 11
        iBuffer = myarr(i)
        If iBuffer > iMax Then
            iMax = iBuffer
        End If
    Next i
    
        iMax = IIf(iMax = 0, 1, iMax)
    
    For i = 0 To 11
        Shape1(i).Top = gitop
        Shape1(i).Height = (1900 / iMax) * IIf(myarr(i) < 0, 0, myarr(i))
        Shape1(i).Top = gitop - ((1900 / iMax) * myarr(i))
        
        Label10(i).Top = Shape1(i).Top - 350
        Label10(i).Caption = myarr(i)
        Label10(i).Refresh
        

        Label6(i).ForeColor = vbRed
        Label6(i).Top = Shape1(i).Top
        Label6(i).Caption = " " & IIf(arrUm(i) <= 0, "", Format$(arrUm(i), "#####0.00"))
        Label6(i).ToolTipText = IIf(arrUm(i) <= 0, "", Format$(arrUm(i), "#####0.00 €"))
        Label6(i).Refresh
        
        
        
    Next i
    
    For i = 0 To 11
        iSumAnz = iSumAnz + myarr(i)
        siSumUm = siSumUm + arrUm(i)
    Next i
    Label8.Caption = Format$(iSumAnz, "##### Stück")
    Label14.Caption = Format$(siSumUm, "#####0.00 €")
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "diagrammfuellen"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenWKLam()
    On Error GoTo LOKAL_ERROR
    
    Frame6.Height = 3975
    Frame6.Left = 0
    Frame6.Top = 4680
    Frame6.Width = 6735
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKLam"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



'Private Sub Label1_Change(Index As Integer)
'    On Error GoTo LOKAL_ERROR
'
'    If Index = 20 Then
''        MsgBox Label1(20).Caption
'        FuelleFlexGrid1WKLam
'    End If
'
'    Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Label1_Change"
'    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
Private Sub SSCommand1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    Dim lJahr As Long
    Dim lcount As Long
    
    lJahr = Val(Label1(20).Caption)
    
    Select Case Index
        Case Is = 0
            lJahr = lJahr - 1
        Case Is = 1
            lJahr = lJahr + 1
    End Select
    
'    For lCount = 0 To 12
'        Label3(lCount).Caption = "0"
'    Next lCount
'
'    For lCount = 13 To 25
'        Label3(lCount).Caption = "0,00"
'    Next lCount

    
    Label1(20).Caption = Trim$(Str$(lJahr))
    
'    LeseUmsatzWKLam
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 2
            Label2(2).ToolTipText = "Artikelgruppe(AGN) " & Label2(2).Caption & " = " & Ermittleagntext(Label2(2).Caption)
        Case 26
            Label2(26).ToolTipText = "Produktgruppe(PGN) " & Label2(26).Caption & " = " & Ermittlepgntext(Label2(26).Caption)
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label2_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 2
            Label1(2).ToolTipText = "Artikelgruppe(AGN) " & Label2(2).Caption & " = " & Ermittleagntext(Label2(2).Caption)
        Case 23
            Label1(23).ToolTipText = "Produktgruppe(PGN) " & Label2(26).Caption & " = " & Ermittlepgntext(Label2(26).Caption)
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SSCommand2_Click()
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKLam
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand2_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SSCommand2_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case KeyCode
        Case Is = vbKeyEscape
            SSCommand2_Click
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand2_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SSCommand3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            gsARTNR = Label2(0).Caption
            frmWKL64.Show 1
            gsARTNR = ""
        Case 1
            gsARTNR = Label2(0).Caption
            frmWKL63.Show 1
            gsARTNR = ""
        Case 2
            gsARTNR = Label2(0).Caption
            If gsARTNR = "" Then
                Exit Sub
            End If

            gsSEK = Label2(14).Caption
            frmWKL62.Show 1
            
            gsARTNR = ""
        Case 3
            gsARTNR = Label2(0).Caption
            frmWKL67.Show 1
            gsARTNR = ""
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand3_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelinfo ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
