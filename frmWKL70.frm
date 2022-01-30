VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKL70 
   Caption         =   "Artikelsuche"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL70.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   28
      Left            =   9600
      TabIndex        =   104
      Top             =   6480
      Width           =   2040
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
      Caption         =   "Bilder"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'Kein
      Height          =   3255
      Left            =   2520
      TabIndex        =   94
      Top             =   1200
      Visible         =   0   'False
      Width           =   6255
      Begin VB.FileListBox File2 
         Height          =   285
         Left            =   4320
         TabIndex        =   108
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.FileListBox File3 
         Height          =   285
         Left            =   5160
         TabIndex        =   98
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin sevCommand3.Command Command5 
         Height          =   285
         Index           =   31
         Left            =   4440
         TabIndex        =   97
         Top             =   2520
         Width           =   1680
         _ExtentX        =   0
         _ExtentY        =   0
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
         ButtonStyle     =   2
         Caption         =   "übernehmen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   285
         Index           =   30
         Left            =   4440
         TabIndex        =   96
         Top             =   1080
         Width           =   1680
         _ExtentX        =   0
         _ExtentY        =   0
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
         ButtonStyle     =   2
         Caption         =   "suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   285
         Index           =   29
         Left            =   4440
         TabIndex        =   95
         Top             =   2880
         Width           =   1680
         _ExtentX        =   0
         _ExtentY        =   0
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
         ButtonStyle     =   2
         Caption         =   "zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSComDlg.CommonDialog cdlopen 
         Left            =   5760
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   39
         Left            =   120
         TabIndex        =   103
         Top             =   2880
         Width           =   3495
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   38
         Left            =   120
         TabIndex        =   102
         Top             =   1440
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pfad zu den Bildern?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   37
         Left            =   120
         TabIndex        =   101
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Wo sind die Bilder?"
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
         Index           =   36
         Left            =   120
         TabIndex        =   100
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bilder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   35
         Left            =   120
         TabIndex        =   99
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "mit Bestand"
      Height          =   315
      Left            =   9600
      TabIndex        =   93
      Top             =   7080
      Value           =   1  'Aktiviert
      Width           =   1815
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   3
      Left            =   9600
      TabIndex        =   92
      Top             =   7440
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
   Begin VB.CheckBox Check1 
      Caption         =   "Umsatz"
      Height          =   315
      Left            =   9600
      TabIndex        =   67
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Height          =   3015
      Left            =   2760
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   6735
      Begin sevCommand3.Command Command4 
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "-"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   330
         Left            =   6240
         TabIndex        =   14
         Top             =   120
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "+"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "D"
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
         Index           =   11
         Left            =   2760
         TabIndex        =   66
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "N"
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
         Index           =   10
         Left            =   2520
         TabIndex        =   65
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "O"
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
         Index           =   9
         Left            =   2280
         TabIndex        =   64
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "S"
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
         Index           =   8
         Left            =   2040
         TabIndex        =   63
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "A"
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
         Index           =   7
         Left            =   1800
         TabIndex        =   62
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "J"
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
         Index           =   6
         Left            =   1560
         TabIndex        =   61
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "J"
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
         Index           =   5
         Left            =   1320
         TabIndex        =   60
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "M"
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
         Index           =   4
         Left            =   1080
         TabIndex        =   59
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "A"
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
         Index           =   3
         Left            =   840
         TabIndex        =   58
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "M"
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
         Index           =   2
         Left            =   600
         TabIndex        =   57
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "F"
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
         Index           =   1
         Left            =   360
         TabIndex        =   56
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "J"
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
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   11
         Left            =   2760
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   10
         Left            =   2520
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   9
         Left            =   2280
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   8
         Left            =   2040
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   7
         Left            =   1800
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   6
         Left            =   1560
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   5
         Left            =   1320
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   4
         Left            =   1080
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   3
         Left            =   840
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   2
         Left            =   600
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   1
         Left            =   360
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   0
         Left            =   120
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
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
         Index           =   10
         Left            =   2520
         TabIndex        =   54
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
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
         Index           =   9
         Left            =   2280
         TabIndex        =   53
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
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
         Index           =   8
         Left            =   2040
         TabIndex        =   52
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
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
         Index           =   7
         Left            =   1800
         TabIndex        =   51
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
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
         Index           =   6
         Left            =   1560
         TabIndex        =   50
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
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
         Index           =   5
         Left            =   1320
         TabIndex        =   49
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
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
         Index           =   3
         Left            =   840
         TabIndex        =   48
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
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
         Index           =   2
         Left            =   600
         TabIndex        =   47
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
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
         Index           =   1
         Left            =   360
         TabIndex        =   46
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
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
         Index           =   0
         Left            =   120
         TabIndex        =   45
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
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
         Index           =   4
         Left            =   1080
         TabIndex        =   44
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackColor       =   &H0000FFFF&
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
         Index           =   11
         Left            =   2760
         TabIndex        =   43
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C000&
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
         Index           =   49
         Left            =   1200
         TabIndex        =   42
         Top             =   2640
         Width           =   825
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
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
         Index           =   50
         Left            =   4800
         TabIndex        =   41
         Top             =   2640
         Width           =   780
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
         Index           =   12
         Left            =   3720
         TabIndex        =   40
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
         Index           =   13
         Left            =   3960
         TabIndex        =   39
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
         Index           =   14
         Left            =   4200
         TabIndex        =   38
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
         Index           =   15
         Left            =   4440
         TabIndex        =   37
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
         Index           =   16
         Left            =   4680
         TabIndex        =   36
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
         Index           =   17
         Left            =   4920
         TabIndex        =   35
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
         Index           =   18
         Left            =   5160
         TabIndex        =   34
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
         Index           =   19
         Left            =   5400
         TabIndex        =   33
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
         Index           =   20
         Left            =   5640
         TabIndex        =   32
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
         TabIndex        =   31
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
         Index           =   22
         Left            =   6120
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "J"
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
         Index           =   12
         Left            =   3720
         TabIndex        =   28
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "F"
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
         Index           =   13
         Left            =   3960
         TabIndex        =   27
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "M"
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
         Index           =   14
         Left            =   4200
         TabIndex        =   26
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "A"
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
         Index           =   15
         Left            =   4440
         TabIndex        =   25
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "M"
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
         Left            =   4680
         TabIndex        =   24
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "J"
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
         Index           =   17
         Left            =   4920
         TabIndex        =   23
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "J"
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
         Index           =   18
         Left            =   5160
         TabIndex        =   22
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "A"
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
         Index           =   19
         Left            =   5400
         TabIndex        =   21
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "S"
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
         Index           =   20
         Left            =   5640
         TabIndex        =   20
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "O"
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
         Index           =   21
         Left            =   5880
         TabIndex        =   19
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "N"
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
         Index           =   22
         Left            =   6120
         TabIndex        =   18
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "D"
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
         Index           =   23
         Left            =   6360
         TabIndex        =   17
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   12
         Left            =   3720
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   13
         Left            =   3960
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   14
         Left            =   4200
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   15
         Left            =   4440
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   16
         Left            =   4680
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   17
         Left            =   4920
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   18
         Left            =   5160
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   19
         Left            =   5400
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   20
         Left            =   5640
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   21
         Left            =   5880
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   22
         Left            =   6120
         Top             =   2400
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   15
         Index           =   23
         Left            =   6360
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Verkaufszahlen (Menge in Stück)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   120
         Width           =   5055
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   2
      Left            =   9600
      TabIndex        =   12
      Top             =   5520
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
      Caption         =   "Spezialsuche"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   11055
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   8400
         TabIndex        =   90
         Top             =   360
         Width           =   1575
      End
      Begin sevCommand3.Command Command2 
         Height          =   315
         Index           =   3
         Left            =   10080
         TabIndex        =   89
         Top             =   360
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         ButtonStyle     =   2
         Caption         =   "Suche"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   315
         Index           =   2
         Left            =   7440
         TabIndex        =   87
         Top             =   360
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         ButtonStyle     =   2
         Caption         =   "Suche"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   5760
         TabIndex        =   86
         Top             =   360
         Width           =   1575
      End
      Begin sevCommand3.Command Command66 
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   85
         Top             =   120
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
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
         ButtonStyle     =   2
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command66 
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   11
         Top             =   120
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
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
         ButtonStyle     =   2
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3120
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin sevCommand3.Command Command2 
         Height          =   315
         Index           =   1
         Left            =   4800
         TabIndex        =   7
         Top             =   360
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         ButtonStyle     =   2
         Caption         =   "Suche"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   6
         Top             =   360
         Width           =   855
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         ButtonStyle     =   2
         Caption         =   "Suche"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Artikelbezeichnung"
         Height          =   255
         Index           =   3
         Left            =   8400
         TabIndex        =   91
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Lieferantenbestellnummer"
         Height          =   255
         Index           =   2
         Left            =   5760
         TabIndex        =   88
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Produktgruppe"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Marke"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   1
      Left            =   9600
      TabIndex        =   1
      Top             =   6000
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
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   2
      Top             =   7920
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLEX1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6800
      _Version        =   393216
      FocusRect       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'Kein
      Height          =   3015
      Left            =   120
      TabIndex        =   68
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   1560
         TabIndex        =   107
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   855
         Left            =   1560
         MouseIcon       =   "frmWKL70.frx":0442
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   855
         ScaleWidth      =   975
         TabIndex        =   105
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   "Bilder"
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
         Index           =   0
         Left            =   1560
         TabIndex        =   106
         Top             =   960
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   2040
         Top             =   2160
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00C0C000&
         Caption         =   "ArtNr  : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   84
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label lbl4 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Left            =   0
         TabIndex        =   83
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Bestand"
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
         Index           =   4
         Left            =   0
         TabIndex        =   82
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "VKAM"
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
         Index           =   5
         Left            =   0
         TabIndex        =   81
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "VKVM"
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
         Index           =   6
         Left            =   0
         TabIndex        =   80
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "VKLJ"
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
         Index           =   7
         Left            =   0
         TabIndex        =   79
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "VKVJ"
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
         Index           =   8
         Left            =   0
         TabIndex        =   78
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   77
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   76
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Height          =   255
         Index           =   11
         Left            =   840
         TabIndex        =   75
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Height          =   255
         Index           =   12
         Left            =   840
         TabIndex        =   74
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Height          =   255
         Index           =   13
         Left            =   840
         TabIndex        =   73
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Filiale"
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
         Index           =   14
         Left            =   0
         TabIndex        =   72
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbl6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Height          =   255
         Index           =   15
         Left            =   840
         TabIndex        =   71
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "keine:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   70
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Notizen:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   69
         Top             =   2520
         Width           =   2535
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   360
      Left            =   11280
      TabIndex        =   109
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
      Picture         =   "frmWKL70.frx":074C
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin VB.Label Label90 
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
      TabIndex        =   3
      Top             =   7920
      Width           =   9375
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
End
Attribute VB_Name = "frmWKL70"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerAWM  As Byte
Dim SpaltennummerArtnr As Byte
Dim SpaltennummerLIBESNR As Byte
Dim SpaltennummerBEZEICH As Byte

Dim gitop           As Integer
Private Sub Check1_Click()
On Error GoTo LOKAL_ERROR

    Dim cArtNr      As String
    cArtNr = MSHFLEX1.TextMatrix(MSHFLEX1.Row, CLng(SpaltennummerArtnr))
    
    If Check1.Value = vbChecked Then
        Detaildatenermitteln cArtNr, "Umsatz"
    Else
        Detaildatenermitteln cArtNr, "Anzahl"
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Command1_Click()
    On Error GoTo LOKAL_ERROR
    
    gsZSpalte = "Artnr"
    gstab = "ARTSUCH"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1

End Sub
Private Sub Command2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim cValid      As String
    Dim cFeld       As String
    Dim cZeichen    As String
    Dim lcount      As Long
    Dim bTextSuche  As Boolean
    Dim cSuch       As String
    Dim i           As Integer
    Dim sSQL        As String
    
    Screen.MousePointer = 11
    
    loeschNEW "ARTSUCH", gdBase
    CreateTable "ARTSUCH", gdBase
    
    For i = 0 To 100
        gBYTENum(i) = 255555
    Next i
    
    Select Case Index
        Case 1
    
            cValid = "1234567890"
            cFeld = Text1.Text
            
            bTextSuche = False
            
            For lcount = 1 To Len(cFeld)
                cZeichen = Mid(cFeld, lcount, 1)
                If InStr(cValid, cZeichen) = 0 Then
                    bTextSuche = True
                    Exit For
                End If
            Next lcount
            
            If bTextSuche = True Then
                If LoesePGNstringinNum(Trim$(Text1.Text)) = True Then
        
                    SucheTextArtikelWKL70 "PGNNUM"
                    FuellenMShFlex1WKLad "PGNNUM"
                    
                Else
                    
            
                End If
            Else
                If Text1.Text <> "" Then
                    If IsNumeric(Val(Text1.Text)) Then
                        If Val(Trim$(Text1.Text)) < 256 Then
                            gBYTENum(0) = Val(Trim$(Text1.Text))
                            SucheTextArtikelWKL70 "PGNNUM"
                            FuellenMShFlex1WKLad "PGNNUM"
                        End If
                    End If
                End If
            End If
        Case 0
            If Trim$(Text2.Text) <> "" Then
                If LoeseMarkenstringinLPZ12(Trim$(Text2.Text)) = True Then
                    SucheTextArtikelWKL70 "MARKE"
                    FuellenMShFlex1WKLad "MARKE"
                Else
                    anzeige "rot", "Bitte verändern Sie Ihr Suchkriterium", Label90
                    Exit Sub
                End If
            Else
                sSQL = "Delete from  MA" & srechnertab
                gdBase.Execute sSQL, dbFailOnError
            
            End If
           
        Case 2
            
            If Text3.Text <> "" Then
                gcSuch = Trim(Text3.Text)
                SucheTextArtikelWKL70 "LIBESNR"
                FuellenMShFlex1WKLad "LIBESNR"
            End If
            
        Case 3
            
            If Text4.Text <> "" Then
                gcSuch = Trim(Text4.Text)
                SucheTextArtikelWKL70 "BEZEICH"
                FuellenMShFlex1WKLad "BEZEICH"
            End If
            
        
    End Select
    
    If MSHFLEX1.Visible = True Then
        MSHFLEX1.SetFocus
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSHFlexGrid, siEigFak As Single)
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
Private Sub Command4_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cArtNr      As String
    Dim Datum1      As Date
    
    Datum1 = DateValue("01.01." & Label9(49).Caption)
    cArtNr = MSHFLEX1.TextMatrix(glSelect, CLng(SpaltennummerArtnr))
    diagrammfuellenMod3 cArtNr, Datum1
    
    Label9(50).Caption = Label9(49).Caption
    Label9(49).Caption = CInt(Label9(49).Caption) - 1
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub

Private Sub Command3_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cArtNr     As String
    Dim Datum1 As Date
    
    
    Label9(49).Caption = Label9(50).Caption
    Label9(50).Caption = CInt(Label9(50).Caption) + 1
    
    Datum1 = DateValue("01.01." & Label9(50).Caption)
    
    cArtNr = MSHFLEX1.TextMatrix(glSelect, CLng(SpaltennummerArtnr))
    diagrammfuellenMod3 cArtNr, Datum1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub

Private Function LoesePGNstringinNum(cKrit As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim i       As Integer
    
    LoesePGNstringinNum = False
    
    For i = 0 To 100
      gBYTENum(i) = 255555
    Next i
    
    i = 0
    
    sSQL = "Select PGN from PGNDBF where PGNBEZEICH like '" & cKrit & "*' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!PGN) Then
                gBYTENum(i) = rsrs!PGN
                i = i + 1
                LoesePGNstringinNum = True
            
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoesePGNstringinNum"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Function
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim sPfad As String
    Dim i As Integer
    Dim j As Integer
    Dim cArtNr As String
    Dim cLiBesNr As String
    Dim cBezeich As String
    
    
    Select Case Index
    
        Case 31
            Screen.MousePointer = 11
            anzeigeNew "normal", "0", Label2(39)
            
            'liefBestNr
            sPfad = Label2(37).Caption
            
            
'            If Bildspeichern2("613408", sPfad, "680268", File3, Check3.Value) Then
'                anzeigeNew "normal", CInt(Label2(39).Caption) + 1, Label2(39)
'            End If
            
            
            
            
            MSHFLEX1.Redraw = False

            MSHFLEX1.Row = 0
            For i = 1 To MSHFLEX1.Rows - 1

                MSHFLEX1.Row = i

                MSHFLEX1.Col = SpaltennummerArtnr
                cArtNr = MSHFLEX1.Text

                MSHFLEX1.Col = SpaltennummerLIBESNR
                cLiBesNr = Trim(MSHFLEX1.Text)
                
                MSHFLEX1.Col = SpaltennummerBEZEICH
                cBezeich = Trim(MSHFLEX1.Text)

                Dim sBezeichNum As String
                sBezeichNum = ""
                
                Dim sLibesnrNum As String
                sLibesnrNum = ""

                'Punkt ran
                cLiBesNr = cLiBesNr & "."

                If cLiBesNr <> "" Then

                    If IsNumeric(cLiBesNr) = False Then

                        For j = 1 To Len(cLiBesNr)
                            If IsNumeric(Mid(cLiBesNr, j, 1)) = True Then
                                sLibesnrNum = sLibesnrNum & Mid(cLiBesNr, j, 1)
                            Else
                                'endet auf jeden Fall mit einem Punkt also nicht numerisch

                                If Len(sLibesnrNum) > 0 Then
                                    cLiBesNr = CStr(Val(sLibesnrNum))
                                Else
                                    sLibesnrNum = ""
                                End If

                            End If
                        Next j
                    Else
                        cLiBesNr = CStr(Val(cLiBesNr))
                    End If
                End If
                
                'Punkt ran
                cBezeich = cBezeich & "."
                
                If cBezeich <> "" Then
                
                    If IsNumeric(cBezeich) = False Then
                
                        For j = 1 To Len(cBezeich)
                            If IsNumeric(Mid(cBezeich, j, 1)) = True Then
                                sBezeichNum = sBezeichNum & Mid(cBezeich, j, 1)
                            Else
                                'endet auf jeden Fall mit einem Punkt also nicht numerisch
                                
                                If Len(sBezeichNum) > 0 Then
                                    cBezeich = CStr(Val(sBezeichNum))
                                Else
                                    sBezeichNum = ""
                                End If
                                
                            End If
                        Next j
                    Else
                        cBezeich = CStr(Val(cBezeich))
                    End If
                End If

'                If cLiBesNr = "601010" Then
'                    MsgBox "yaeh"
'                End If

'                If Bildspeichern(cArtNr, sPfad, cLiBesNr, File3, Check3.Value) Then
'                    anzeigeNew "normal", CInt(Label2(39).Caption) + 1, Label2(39)
'                End If

                If Bildspeichern2(cArtNr, sPfad, cLiBesNr, cBezeich, File3) Then
                    anzeigeNew "normal", CInt(Label2(39).Caption) + 1, Label2(39)
                End If
            Next i

            Screen.MousePointer = 0

            MSHFLEX1.Redraw = True
            
            anzeigeNew "normal", Label2(39).Caption & " Bilder zugordnet, Fertig!", Label2(39)
        
        Case Is = 30
            'Wo sind die Bilddateien
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .DialogTitle = "Wo sind die Bilder?"
                
                .Filter = "JPEG (*.JPG)| *.JPG|GIF (*.GIF)| *.GIF|PNG (*.PNG)| *.PNG| Bitmapdateien (*.bmp)|*.bmp"
                
                .ShowSave
                
                sPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
                If Right(sPfad, 1) <> "\" Then
                    sPfad = sPfad & "\"
                End If
                Label2(37).Caption = sPfad
                
                File3.Path = sPfad
                File3.Pattern = "*.jpg"
                File3.Refresh
                
                If File3.ListCount = 1 Then
                    Label2(38).Caption = File3.ListCount & " Bild im Verzeichnis"
                Else
                    Label2(38).Caption = File3.ListCount & " Bilder im Verzeichnis"
                End If
            End With
        
        Case 29
            Frame6.Visible = False
        Case 28
            Frame6.Visible = True
            Frame6.BackColor = glH2
        Case 0
            Unload frmWKL70
        Case 1
            If glSelect = 0 Then
                anzeige "rot", "Bitte einen Artikel markieren!", Label90
                Exit Sub
            End If
            MSHFLEX1.Row = glSelect
            MSHFLEX1.Col = SpaltennummerArtnr
            gsARTNR = MSHFLEX1.Text
            gsARTNR = Trim$(gsARTNR)
            
            Unload frmWKL70
        Case 2
            If Frame1.Visible = False Then
                Frame1.Visible = True
                Command5(2).Visible = False
            End If
        Case 3
            loeschNEW "ARTPRINT", gdBase
            CreateTableT2 "ARTPRINT", gdBase
            
            loeschNEW "LIEFPRINT", gdBase
            CreateTable "LIEFPRINT", gdBase
            
            sSQL = "Insert into ARTPRINT select * from Artsuch "
            If Check2.Value = vbChecked Then
                sSQL = sSQL & " where bestand > 0 "
            End If
            sSQL = sSQL & " order by Bezeich "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update ARTPRINT inner join Artikel on ARTPRINT.Artnr = Artikel.artnr "
            sSQL = sSQL & " set ARTPRINT.Inhalt = Artikel.Inhalt "
            sSQL = sSQL & " , ARTPRINT.Inhaltbez = Artikel.Inhaltbez "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Select distinct(linr) from ARTPRINT "
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                If rsrs.RecordCount = 1 Then
                
                    If Not IsNull(rsrs!linr) Then
                
                        sSQL = "Insert into LIEFPRINT select LINR "
                        sSQL = sSQL & ", LIEFBEZ "
                        sSQL = sSQL & ", AWERT  "
                        sSQL = sSQL & ", ZIELEK "
                        sSQL = sSQL & ", FAX "
                        sSQL = sSQL & ", KTEXT "
                        sSQL = sSQL & ", KUERZEL "
                        sSQL = sSQL & ", NOTIZ "
                        sSQL = sSQL & ", PLZ "
                        sSQL = sSQL & ", STADT "
                        sSQL = sSQL & ", STRASSE "
                        sSQL = sSQL & ", TEL from LISRT where linr = " & rsrs!linr
                        gdBase.Execute sSQL, dbFailOnError
                    
                    End If
                
                End If
            End If
            
            reportbildschirm "", "aWKL70a"
            
    End Select
    
err:
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub


Private Sub Command66_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case Is = 0
            Text1_KeyUp vbKeyF2, 0
        Case Is = 1
            Text2_KeyUp vbKeyF2, 0
            
        End Select
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command66_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1

End Sub
Private Sub diagrammfuellenMod3(sArtnr As String, Jahrkl As Date)
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim iTop            As Integer
    Dim myarr(0 To 23)  As Long
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim iMax            As Integer
    Dim iBuffer         As Integer
    
    cSQL = "Select * from UMS_ART  where ARTNR = " & sArtnr
    cSQL = cSQL & " and Jahr = " & Year(Jahrkl) - 1
    
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            myarr(rsrs!Monat - 1) = rsrs!ANZAHL
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    cSQL = "Select * from UMS_ART  where ARTNR = " & sArtnr
    cSQL = cSQL & " and Jahr = " & Year(Jahrkl)

    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            myarr(rsrs!Monat - 1 + 12) = rsrs!ANZAHL
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing: Set rsrs = Nothing
        
    iBuffer = 0
    iMax = 0
    
    For i = 0 To 23
        iBuffer = myarr(i)
        If iBuffer > iMax Then
            iMax = iBuffer
        End If
    Next i
    
        iMax = IIf(iMax = 0, 1, iMax)
    
    For i = 0 To 23
        Shape1(i).Top = gitop
        Shape1(i).Height = (1900 / iMax) * IIf(myarr(i) < 0, 0, myarr(i))
        Shape1(i).Top = gitop - ((1900 / iMax) * myarr(i))
        
        Label10(i).Top = Shape1(i).Top - 250
        Label10(i).Caption = myarr(i)
        Label10(i).Refresh
    Next i
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "diagrammfuellenMod3"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Detaildatenermitteln(cArtNr As String, sArt As String)
    On Error GoTo LOKAL_ERROR
    
    Dim l          As Long
    Dim lVorhanden As Long
    Dim ctmp       As String
    Dim cSQL       As String
    Dim iMonat     As Integer
    Dim rsrs       As Recordset

    cSQL = "Select BESTAND,Bezeich,Notizen from Artikel where ARTNR = " & cArtNr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!BESTAND) Then
            ctmp = rsrs!BESTAND
        Else
            ctmp = "0"
        End If
        lbl6(9).Caption = ctmp
        
        If Not IsNull(rsrs!BEZEICH) Then
            ctmp = rsrs!BEZEICH
        Else
            ctmp = ""
        End If
        lbl4.Caption = ctmp
        
        If Not IsNull(rsrs!NOTIZEN) Then
            ctmp = rsrs!NOTIZEN
        Else
            ctmp = ""
        End If
        Label3(0).Caption = ctmp
        
    Else
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    lbl3.Caption = "ArtNr. : " & cArtNr
    
    Label3(0).Refresh
    lbl3.Refresh
    lbl4.Refresh
    lbl6(9).Refresh
    
    lbl6(12).Caption = vklj(cArtNr, sArt)
    lbl6(13).Caption = vkvj(cArtNr, sArt)
    lbl6(10).Caption = vkam(cArtNr, sArt)
    lbl6(11).Caption = vkvm(cArtNr, sArt)
    
    Select Case sArt
        Case "Anzahl"
            
            
        Case "Umsatz"
            lbl6(12).Caption = Format$(lbl6(12).Caption, "######0.00")
            lbl6(13).Caption = Format$(lbl6(13).Caption, "######0.00")
            lbl6(10).Caption = Format$(lbl6(10).Caption, "######0.00")
            lbl6(11).Caption = Format$(lbl6(11).Caption, "######0.00")
    End Select
    lbl6(12).Refresh
    lbl6(13).Refresh
    lbl6(10).Refresh
    lbl6(11).Refresh
     
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Detaildatenermitteln"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub

Private Sub Command98_Click()

End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Dim iRet As Integer
    Dim lcountBezeich As Long
    Dim lcountLibesnr As Long
    Dim sSQL As String
    
    Screen.MousePointer = 11
    
'    PositionierenWKL61
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    
    gitop = Shape1(0).Top
    Label9(49).Caption = Year(DateValue(Now)) - 1
    Label9(50).Caption = Year(DateValue(Now))
    
    anzeige "normal", "Daten werden ermittelt, bitte warten...", Label90

    Screen.MousePointer = 11
    
    'Grid formatieren
    
    glSelect = 0
    Tabcheck "ARTSUCH"
    FormatGridOverTablay "ARTSUCH"
    
    ermittlespalten
    
    Dim j As Integer
    
    With MSHFLEX1
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
            aBreite(j) = Len(.TextMatrix(0, j)) * 80
        Next j
    End With
    
    Me.Refresh
    
    sSQL = "Delete from  MA" & srechnertab
    gdBase.Execute sSQL, dbFailOnError
    
    If Left(gcSuch, 4) = "LINR" Then
        Command5(2).Visible = False
        Command5(1).Visible = False
        SucheTextArtikelWKL70 "LINR"
        FuellenMShFlex1WKLad "LINR"
    Else
        If gbLibesnrSeek = False Then
            SucheTextArtikelWKL70 "BEZEICH"
            FuellenMShFlex1WKLad "BEZEICH"
        Else
            SucheTextArtikelWKL70 "LIBESNR"
            FuellenMShFlex1WKLad "LIBESNR"
        End If
    End If

    Me.Refresh
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
   
End Sub
Private Sub SucheTextArtikelWKL70(sArt As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSuch As String
    Dim cSQL As String
    Dim sSQL As String
    Dim i As Integer
    Dim cFeld As String
    
    loeschNEW "ARTSUCH", gdBase
    CreateTable "ARTSUCH", gdBase
    
    If gcSuch = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    anzeige "normal", "Artikel werden ermittelt...", Label90
    
    cSQL = "Select "
    cSQL = cSQL & " A.ARTNR "
    cSQL = cSQL & ", A.BEZEICH "
    cSQL = cSQL & ", A.AGN "
    cSQL = cSQL & ", B.LINR "
    cSQL = cSQL & ", A.LPZ "
    cSQL = cSQL & ", B.LIBESNR "
    cSQL = cSQL & ", A.EAN  "
    cSQL = cSQL & ", A.RKZ  "
    cSQL = cSQL & ", A.BESTAND  "
    cSQL = cSQL & ", A.GEFUEHRT "
    cSQL = cSQL & ", A.MWST "
    cSQL = cSQL & ", A.KVKPR1 "
    cSQL = cSQL & ", A.EKPR "
    cSQL = cSQL & ", A.LEKPR "
    cSQL = cSQL & ", A.VKPR "
    cSQL = cSQL & ", A.Preisschu "
    cSQL = cSQL & ", A.Rabatt_OK "
    cSQL = cSQL & ", A.PGN "
    cSQL = cSQL & ", A.AWM "
    cSQL = cSQL & ", A.MinMen "
    cSQL = cSQL & ", A.MinBest "
    cSQL = cSQL & " from ARTIKEL A "
    
    If Datendrin("MA" & srechnertab, gdBase) Then
        cSQL = cSQL & " , Artlief B,MA" & srechnertab & " d where A.artnr = b.artnr and a.artnr = d.artnr "
    Else
        cSQL = cSQL & " , Artlief B where A.artnr = b.artnr  "
    End If
    
    Select Case sArt
    
        Case Is = "LINR"
            cSuch = gcSuch
            cSuch = UCase$(Trim$(Mid$(cSuch, 5, 6)))
            
            cSQL = cSQL & " and B.LINR = " & cSuch

        Case Is = "LIBESNR"
            cSuch = gcSuch
            cSuch = UCase$(Trim$(cSuch))
            
            cSQL = cSQL & " and ucase(B.LIBESNR) like '" & cSuch & "*' "
            
        Case Is = "BEZEICH"
            cSuch = gcSuch
            cSuch = UCase$(Trim$(cSuch))
            
'            cSQL = cSQL & " and A.BEZEICH like '" & cSuch & "*' "
            
            Dim sArray() As String
            sArray = Split(cSuch, " ")
            
            For i = 0 To UBound(sArray)
                cFeld = sArray(i)
                cSQL = cSQL & " and A.BEZEICH like '*" & cFeld & "*' "
            Next i
            
        Case Is = "PGNNUM"
        
            gcSuch = ""
            cSQL = cSQL & " and  (A.PGN = " & gBYTENum(0)
            For i = 1 To UBound(gBYTENum) - 1
                If gBYTENum(i) = 255555 Then
                
                Else
                    cSQL = cSQL & " or A.PGN = " & gBYTENum(i)
                End If
            Next i
            cSQL = cSQL & ")"
        
        Case Is = "MARKE"
            gcSuch = ""
        
    End Select
    
    sSQL = "Insert into ARTSUCH " & cSQL
    
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheTextArtikelWKL70"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
  
End Sub



Private Sub Picture3_Click()
On Error GoTo LOKAL_ERROR
    
    gsARTNR = Picture3.Tag
    frmWKL163.Show 1

    gsARTNR = ""
   
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Picture3_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub

Private Sub Text2_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text2.BackColor = glSelBack1
   
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
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
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = glSelBack1
   
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Text3_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text3.BackColor = glSelBack1
   
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Text3_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text3.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Text4_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text4.BackColor = glSelBack1
   
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyReturn Then
        Command2_Click 3
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Text4_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text4.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub FuellenMShFlex1WKLad(sArt As String)
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
    Dim iRet        As Integer
    Dim cSQL        As String
    
    If NewTableSuchenDBKombi("ARTSUCH", gdBase) = False Then
        Exit Sub
    End If
    
    If sArt = "BEZEICH" Then
    
        loeschNEW "TEMPART", gdBase
        
        
        cSQL = "select distinct(Artnr)as art "
        cSQL = cSQL & ", BEZEICH "
        cSQL = cSQL & ", AGN "
        cSQL = cSQL & ", 0 as LINR "
        cSQL = cSQL & ", LPZ "
        cSQL = cSQL & ", '' as LIBESNR "
        cSQL = cSQL & ", EAN  "
        cSQL = cSQL & ", RKZ  "
        cSQL = cSQL & ", BESTAND  "
        cSQL = cSQL & ", GEFUEHRT "
        cSQL = cSQL & ", MWST "
        cSQL = cSQL & ", KVKPR1 "
        cSQL = cSQL & ", EKPR "
        cSQL = cSQL & ", LEKPR "
        cSQL = cSQL & ", VKPR "
        cSQL = cSQL & ", Preisschu "
        cSQL = cSQL & ", Rabatt_OK "
        cSQL = cSQL & ", PGN "
        cSQL = cSQL & ", AWM "
        cSQL = cSQL & ", MinMen "
        cSQL = cSQL & ", MinBest "
        cSQL = cSQL & " into TEMPART from ARTSUCH"
        gdBase.Execute cSQL, dbFailOnError
        
        loeschNEW "ARTSUCH", gdBase
        
        cSQL = "select Art as ARTNR"
        cSQL = cSQL & ", BEZEICH "
        cSQL = cSQL & ", AGN "
        cSQL = cSQL & ", LINR "
        cSQL = cSQL & ", LPZ "
        cSQL = cSQL & ", LIBESNR "
        cSQL = cSQL & ", EAN  "
        cSQL = cSQL & ", RKZ  "
        cSQL = cSQL & ", BESTAND  "
        cSQL = cSQL & ", GEFUEHRT "
        cSQL = cSQL & ", MWST "
        cSQL = cSQL & ", KVKPR1 "
        cSQL = cSQL & ", EKPR "
        cSQL = cSQL & ", LEKPR "
        cSQL = cSQL & ", VKPR "
        cSQL = cSQL & ", Preisschu "
        cSQL = cSQL & ", Rabatt_OK "
        cSQL = cSQL & ", PGN "
        cSQL = cSQL & ", AWM "
        cSQL = cSQL & ", MinMen "
        cSQL = cSQL & ", MinBest "
        cSQL = cSQL & " into ARTSUCH from TEMPART"
        gdBase.Execute cSQL, dbFailOnError
        
        loeschNEW "TEMPART", gdBase
    End If
    
    Set rsrs = gdBase.OpenRecordset("ARTSUCH", dbOpenTable)
    
    MSHFLEX1.Redraw = False
    MSHFLEX1.Visible = False
    
    counter = 0
    
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveLast
        
        If rsrs.RecordCount > 1000 Then
            iRet = MsgBox("Uppss..." & vbCrLf & "Es wurden mehr als 1000 Datensätze gefunden.(" & rsrs.RecordCount & ")" & vbCrLf & "Wirklich anzeigen?", vbQuestion + vbYesNo, "DATENVOLUMEN")
            If iRet = vbNo Then
                
                Exit Sub
            End If
        End If
        Screen.MousePointer = 11
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If counter = 100 Then
                counter = 0
            End If
            counter = counter + 1
'            pbrZeit.Value = counter
            lrow = lrow + 1
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Col = 0
            
            For i = 0 To byAnzahlSpalten - 1
                MSHFLEX1.Row = 0
                MSHFLEX1.Col = i
                
                If sSpaltenname(i) = MSHFLEX1.Text Then
                    
                    Select Case sSpaltenname(i)
                        Case Is = "Listen - EK", "Listen - VK", "Kassenpreis", "Schnitt - EK"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            MSHFLEX1.Row = lrow

                            MSHFLEX1.Text = Format$(sWert, "####0.00")

                        
                        Case Else
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            MSHFLEX1.Row = lrow
                            MSHFLEX1.Text = sWert
                            
                    End Select
                    
                    FaerbenFlexH MSHFLEX1.TextMatrix(lrow, CLng(SpaltennummerAWM)), MSHFLEX1, 0, CInt(lrow)
                    
            
                    If Len(MSHFLEX1.TextMatrix(lrow, i)) * 80 > aBreite(i) Then
                        aBreite(i) = Len(MSHFLEX1.TextMatrix(lrow, i)) * 80
                    End If
                    
                End If
            Next i
            rsrs.MoveNext
        Loop
    End If
    
    
    
    For i = 0 To byAnzahlSpalten - 1
        MSHFLEX1.Col = i
        MSHFLEX1.ColWidth(i) = aBreite(i) * 1.8
    Next i
    
    
        
    
    rsrs.Close: Set rsrs = Nothing
    If byAnzahlSpalten < 2 Then
    
    Else
        MSHFLEX1.FixedCols = 1
    End If
    
    MSHFLEX1.RowHeight(1) = 0
    lrow = lrow - 1
    
    Text1.Text = gcSuch
    Text2.Text = gcSuch
    Text3.Text = gcSuch
    Text4.Text = gcSuch
    
    Screen.MousePointer = 0

    If lrow > 1 Then
        anzeige "normal", lrow & " Artikel wurden ermittelt.", Label90
        
    ElseIf lrow = 1 Then
        anzeige "normal", lrow & " Artikel wurde ermittelt.", Label90
       
    Else
        anzeige "rot", "Es wurden keine Artikel ermittelt.", Label90
        
        If Left(gcSuch, 4) = "LINR" Then
            Check1.Visible = False
            Exit Sub
        Else

            Frame1.Visible = True
            Frame2.Visible = False
            Frame3.Visible = False
            Check1.Visible = False
            Command5(2).Visible = False
            Exit Sub
        End If
    End If
    
    MSHFLEX1.Row = 2
    MSHFLEX1.Col = 1
    glSelect = 2
    
    Dim cArtNr      As String
    Dim Datum1 As Date
    
    Datum1 = DateValue("01.01." & Label9(50).Caption)
        
    cArtNr = MSHFLEX1.TextMatrix(MSHFLEX1.Row, CLng(SpaltennummerArtnr))
    diagrammfuellenMod3 cArtNr, Datum1
    Frame2.Visible = True
    Frame3.Visible = True
    Check1.Visible = True
    
    If Check1.Value = vbChecked Then
        Detaildatenermitteln cArtNr, "Umsatz"
    Else
        Detaildatenermitteln cArtNr, "Anzahl"
    End If
    
    Tabellenbreiteanpassen MSHFLEX1, 1.5 * gdTabfak
    
    MSHFLEX1.Redraw = True
    MSHFLEX1.Visible = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMShFlex1WKLad"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "AWM"
                SpaltennummerAWM = i
            Case Is = "ARTNR"
                SpaltennummerArtnr = i
            Case Is = "LIBESNR"
                SpaltennummerLIBESNR = i
            Case Is = "BEZEICH"
                SpaltennummerBEZEICH = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_DblClick()
On Error GoTo LOKAL_ERROR
    
    If MSHFLEX1.Row > 1 Then
        Command5_Click 1
    Else
        sortierenHGrid MSHFLEX1
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_Click()
On Error GoTo LOKAL_ERROR

    If MSHFLEX1.Row > 1 Then
        glSelect = MSHFLEX1.Row
    Else
        
    End If
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub

Private Sub MSHFLEX1_EnterCell()
On Error GoTo LOKAL_ERROR
    
    If MSHFLEX1.Row > 1 Then
        glSelect = MSHFLEX1.Row
    Else
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_EnterCell"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    
    If KeyCode = vbKeyF2 Then
    
        lrow = MSHFLEX1.Row
        gsARTNR = MSHFLEX1.TextMatrix(MSHFLEX1.Row, SpaltennummerArtnr)
        If gsARTNR <> "" Then

            frmWKL10.Show 1
            Me.Refresh
            Screen.MousePointer = 11
            MSHFLEX1.Row = lrow
            MSHFLEX1.SetFocus
            
            Screen.MousePointer = 0
        End If
        gsARTNR = ""
    ElseIf KeyCode = vbKeyReturn Then
        Command5_Click 1
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub MSHFLEX1_SelChange()
On Error GoTo LOKAL_ERROR

    Dim cArtNr      As String
    Dim Datum1 As Date
    
    If MSHFLEX1.Row > 1 Then
        Datum1 = DateValue("01.01." & Label9(50).Caption)
        
        cArtNr = MSHFLEX1.TextMatrix(MSHFLEX1.Row, CLng(SpaltennummerArtnr))
        diagrammfuellenMod3 cArtNr, Datum1
        
        If Check1.Value = vbChecked Then
            Detaildatenermitteln cArtNr, "Umsatz"
        Else
            Detaildatenermitteln cArtNr, "Anzahl"
        End If
    Else
        
    End If
    
    
    Bildzeigen MSHFLEX1.TextMatrix(MSHFLEX1.Row, SpaltennummerArtnr), Image1, Picture3, 80
        


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_SelChange"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Bildzeigen(sArt As String, imgx As Image, PicX As PictureBox, iSize As Integer)
On Error GoTo LOKAL_ERROR

    Dim sPfad   As String
    

    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "PICTURE\ARTIKEL"
    
    lbl6(0).Caption = ""
    
    If FileExists(sPfad & "\" & sArt & ".jpg") Then
        imgx.Picture = LoadPicture(sPfad & "\" & sArt & ".jpg")
        
        File1.Path = sPfad
        File1.Pattern = sArt & "*.jpg"
        File1.Refresh
                    
        If File1.ListCount = 1 Then
            lbl6(0).Caption = File1.ListCount & " Bild"
        Else
            lbl6(0).Caption = File1.ListCount & " Bilder"
        End If
    Else
        If FileExists(sPfad & "\" & "keinBild.jpg") Then
            imgx.Picture = LoadPicture(sPfad & "\" & "keinBild.jpg")
        Else
            PicX.Visible = False
            Exit Sub
        End If
    End If
    
    zeigImage_In_Picture_Kasse imgx, PicX, iSize
    PicX.Tag = sArt
    PicX.Visible = True
    
    
    
    
    
Exit Sub
LOKAL_ERROR:

    If err.Number = 481 Then
'        MsgBox "Diese Bild kann nicht gespeichert werden, ungültiges Dateiformat", vbInformation, "Winkiss Hinweis:"
        Kill sPfad & "\" & sArt & ".jpg"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Bildzeigen"
        Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    gF2Prompt.cFeld = ""
    gF2Prompt.cWert = ""
    gF2Prompt.cWert2 = ""
    gF2Prompt.cWahl = ""
    gF2Prompt.bMultiple = False
    
    If KeyCode = vbKeyF2 Then
        
        gF2Prompt.cFeld = "PGN"

        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
            If gF2Prompt.cWahl <> "" Then
                Text1.Text = gF2Prompt.cWahl
                Command2_Click 1
            End If
        End If
    ElseIf KeyCode = vbKeyReturn Then
        Command2_Click 1
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    gF2Prompt.cFeld = ""
    gF2Prompt.cWert = ""
    gF2Prompt.cWert2 = ""
    gF2Prompt.cWahl = ""
    gF2Prompt.bMultiple = False
    
    If KeyCode = vbKeyF2 Then
        
        gF2Prompt.cFeld = "MARKE"

        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
            If gF2Prompt.cWahl <> "" Then
                Text2.Text = gF2Prompt.cWahl
                Command2_Click 0
            End If
        End If
    ElseIf KeyCode = vbKeyReturn Then
        Command2_Click 0
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
