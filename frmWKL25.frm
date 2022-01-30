VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL25 
   BackColor       =   &H00C0C000&
   Caption         =   "Lieferantenzusammenstellung"
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
   Icon            =   "frmWKL25.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Tag             =   "Liefzusa"
   Begin sevCommand3.Command cmdHelp 
      Height          =   555
      Left            =   10320
      TabIndex        =   39
      Top             =   240
      Visible         =   0   'False
      Width           =   1305
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
      ButtonStyle     =   2
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command6 
      Height          =   525
      Index           =   10
      Left            =   9360
      TabIndex        =   37
      Top             =   7800
      Width           =   2175
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   7920
      Visible         =   0   'False
      Width           =   975
      Begin sevCommand3.Command Command6 
         Height          =   525
         Index           =   9
         Left            =   4800
         TabIndex        =   22
         Top             =   3120
         Width           =   2160
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
      Begin sevCommand3.Command Command6 
         Height          =   525
         Index           =   8
         Left            =   6960
         TabIndex        =   19
         Top             =   3120
         Width           =   2160
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
         Caption         =   "Übernehmen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   1800
         Width           =   7815
      End
      Begin VB.Label Label7 
         Caption         =   "Lieferant:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   1440
         Width           =   7815
      End
      Begin VB.Label Label6 
         Caption         =   "ausgewählter Lieferant:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   20
         Top             =   4080
         Width           =   7815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   7815
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   4560
         Width           =   7815
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   1440
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   720
      Top             =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Kein
      Height          =   1215
      Left            =   8760
      TabIndex        =   7
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
      Begin sevCommand3.Command Command6 
         Height          =   525
         Index           =   5
         Left            =   6960
         TabIndex        =   13
         Top             =   3600
         Width           =   2175
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
      Begin sevCommand3.Command Command6 
         Height          =   525
         Index           =   3
         Left            =   6960
         TabIndex        =   9
         Top             =   3000
         Width           =   2175
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
         Caption         =   "Übernehmen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   1320
         TabIndex        =   8
         Top             =   1800
         Width           =   5535
      End
      Begin VB.Label Label14 
         Caption         =   "ausgewählter Lieferant:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   29
         Top             =   4200
         Width           =   7815
      End
      Begin VB.Label Label13 
         Caption         =   "Schritt 1: Wählen Sie ein Vorgang aus!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   28
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label5 
         Caption         =   "Ihre zusammengestellten Lieferanten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   1440
         Width           =   5535
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   4680
         Width           =   7815
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      Height          =   9255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   11415
      Begin sevCommand3.Command Command6 
         Height          =   525
         Index           =   4
         Left            =   7680
         TabIndex        =   11
         Top             =   6840
         Width           =   1575
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
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   5175
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   6240
         TabIndex        =   5
         Top             =   1800
         Width           =   5175
      End
      Begin sevCommand3.Command Command6 
         Height          =   525
         Index           =   0
         Left            =   6120
         TabIndex        =   4
         Top             =   6840
         Width           =   1575
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
      Begin sevCommand3.Command Command4 
         Height          =   495
         Left            =   5400
         TabIndex        =   3
         Top             =   3000
         Width           =   735
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
         Caption         =   "<"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Left            =   5400
         TabIndex        =   2
         Top             =   1800
         Width           =   735
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
         Caption         =   ">"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label11 
         Caption         =   "zugeordnete Lieferanten:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   27
         Top             =   1440
         Width           =   5175
      End
      Begin VB.Label Label12 
         Caption         =   "Lieferant hinzufügen (Doppelklick in der Liste)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   5175
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   11415
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
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
         Left            =   3720
         TabIndex        =   24
         Top             =   960
         Width           =   7695
      End
      Begin VB.Label Label8 
         Caption         =   "ausgewählter Lieferant:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   23
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   6720
         Width           =   5895
      End
   End
   Begin sevCommand3.Command Command2 
      Height          =   495
      Index           =   0
      Left            =   7560
      TabIndex        =   34
      Top             =   3360
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
      Caption         =   "Neu"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command2 
      Height          =   495
      Index           =   1
      Left            =   7560
      TabIndex        =   35
      Top             =   3960
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
      Caption         =   "Bearbeiten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command2 
      Height          =   495
      Index           =   2
      Left            =   7560
      TabIndex        =   36
      Top             =   4560
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
      Caption         =   "Löschen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command2 
      Height          =   495
      Index           =   6
      Left            =   9360
      TabIndex        =   40
      Top             =   960
      Width           =   2175
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
      Caption         =   "Lieferanten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label19 
      Caption         =   $"frmWKL25.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   38
      Top             =   960
      Width           =   9015
   End
   Begin VB.Label Label18 
      Caption         =   "Eine bestehende Zusammenstellung löschen."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   33
      Top             =   4560
      Width           =   7335
   End
   Begin VB.Label Label17 
      Caption         =   "Eine bestehende Zusammenstellung bearbeiten."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   32
      Top             =   3960
      Width           =   7335
   End
   Begin VB.Label Label16 
      Caption         =   "Eine neue Zusammenstellung erstellen."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   31
      Top             =   3360
      Width           =   7335
   End
   Begin VB.Label Label15 
      Caption         =   "Wie möchten Sie vorgehen?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   30
      Top             =   2400
      Width           =   7815
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
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Lieferantenzusammenstellung"
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
End
Attribute VB_Name = "frmWKL25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdHelp_Click()
    On Error GoTo LOKAL_ERROR
    
    zeigeHilfe "KISSHELP", Me.Tag & ".doc", gcPfad
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdHelp_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
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
Private Sub Combo1_Click()
On Error GoTo LOKAL_ERROR

    Label3.ForeColor = glS1
    Label3.Caption = Combo1.Text
    Label3.Refresh

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo1_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Combo1.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo1_GotFocus()
On Error GoTo LOKAL_ERROR
    
    Combo1.BackColor = glSelBack1
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1.Text)

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    AutocompleteCombo KeyCode, Shift, Combo1
    
    Label3.ForeColor = glS1
    Label3.Caption = Combo1.Text
    Label3.Refresh
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List2DelList1()
On Error GoTo LOKAL_ERROR

Dim lCount1         As Long
Dim lCount2         As Long
Dim sVergleich      As String
    
For lCount1 = 0 To List2.ListCount - 1
    sVergleich = List2.list(lCount1)
    For lCount2 = 0 To List1.ListCount - 1
        If List1.list(lCount2) = sVergleich Then
            List1.RemoveItem lCount2
            Exit For
        End If

    Next lCount2
Next lCount1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2DelList1"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub OlinrDetailsAnzeigen()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsOli       As Recordset
    Dim cBez        As String
    Dim cNum        As String
    Dim sOlinr      As String

    List2.Clear
    
    sOlinr = ErmittleLinr(Label9.Caption)
    If sOlinr = "" Then
        Exit Sub
    End If
    
    sSQL = "select * from ueberli where OLinr = " & sOlinr
    Set rsOli = gdBase.OpenRecordset(sSQL)
    
    If Not rsOli.EOF Then
        rsOli.MoveFirst
        Do While Not rsOli.EOF
            If Not IsNull(rsOli!LIEFBEZ) Then
                cBez = Trim(rsOli!LIEFBEZ)
            Else
                cBez = ""
            End If
            
            If Not IsNull(rsOli!linr) Then
                cNum = Trim(rsOli!linr)
            Else
                cNum = ""
            End If
            
            List2.AddItem cBez & Space(35 - Len(cBez)) & cNum & Space(8 - Len(cNum))
            
            rsOli.MoveNext
        Loop
    End If
        
    rsOli.Close
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "OlinrDetailsAnzeigen"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim bFound As Boolean
    Dim lcount As Long
    
    bFound = False
    
    If List1.ListCount = 0 Then
        Exit Sub
    End If
    
    For lcount = 0 To List1.ListCount - 1
        If List1.Selected(lcount) = True Then
            bFound = True
        End If
    Next lcount
    
    If bFound Then
        For lcount = 0 To List1.ListCount - 1
            If List1.Selected(lcount) Then
                List2.AddItem List1.list(lcount)
                List1.RemoveItem lcount
                Exit For
            End If
        Next lcount
    Else
        List2.AddItem List1.list(List1.TopIndex)
        List1.RemoveItem List1.TopIndex
    End If
    
    List1.Refresh
    List2.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Select Case Index
    Case Is = 0 'neu
        'AnzeigeText setzen
        Label4.Caption = "Wählen Sie bitte den Hauptlieferanten für Ihre neue Zusammenstellung aus!"
        Label4.Refresh
        'Ueberlief auflisten
        UeberLieferantenauflisten
        Label3.Caption = ""
        Label3.Refresh
        Frame3.Visible = True
        List2.Clear
        lieferantenauflisten
        LeseLieferanten Combo1, " and HL = true "
        Combo1.SetFocus
        
    Case Is = 2
        Frame2.Visible = True 'Löschen
        'AnzeigeText setzen
        Label13.Caption = "Wählen Sie bitte die zu löschende Lieferantenzusammenstellung aus!"
        Label13.Refresh
        'Ueberlief auflisten
        UeberLieferantenauflisten
        Label1.Caption = ""
        Label1.Refresh
    Case Is = 1
        Frame2.Visible = True 'Bearbeiten
        'AnzeigeText setzen
        Label13.Caption = "Wählen Sie bitte die zu bearbeitende Lieferantenzusammenstellung aus!"
        Label13.Refresh
        'Ueberlief auflisten
        UeberLieferantenauflisten
        Label1.Caption = ""
        Label1.Refresh
    Case Is = 6
        gbFrmComeFrom = True
        Set gfrmComeFrom = frmWKL25
        Unload frmWKL25
        frmWKL17.Show
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub List3_Click()
On Error GoTo LOKAL_ERROR

    Label1.ForeColor = glS1
    Label1.Caption = Trim(List3.list(List3.ListIndex))
    Label1.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List3_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Timer1_Timer()
On Error GoTo LOKAL_ERROR

    Command1_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Timer2_Timer()
On Error GoTo LOKAL_ERROR

    Command4_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer2_Timer"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Timer2.Enabled = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_MouseDown"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Timer2.Enabled = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Timer1.Enabled = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_MouseDown"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Timer1.Enabled = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_Click()
On Error GoTo LOKAL_ERROR
    
    Dim bFound As Boolean
    Dim lcount As Long
    
    If List2.ListCount = 0 Then
        Exit Sub
    End If
    
    
    bFound = False
    
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
                bFound = True
                Exit For
        End If
    Next lcount
    
    If bFound Then
        
        For lcount = 0 To List2.ListCount - 1
            If List2.Selected(lcount) Then
                List1.AddItem List2.list(lcount)
                List2.RemoveItem lcount
                Exit For
            End If
        Next lcount
    
    Else
    
        List1.AddItem List2.list(List2.TopIndex)
        List2.RemoveItem List2.TopIndex
    End If
    
    List1.Refresh
    List2.Refresh
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKL25
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command6_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0 'speichern
            Screen.MousePointer = 11
            DelOlief (Label9.Caption)
            speicherUeberli
            Screen.MousePointer = 0
        Case Is = 9 'zurück zu Frame2
            Frame3.Visible = False
            Frame2.Visible = False
            Frame1.Visible = False
        Case Is = 8 'Übernehmen neu
            'Schritt anzeigen
            If Label3.Caption = "" Or Label3.Caption = "Sie müssen erst ein Lieferant auswählen!" Then
                Label3.ForeColor = vbRed
                Label3.Caption = "Sie müssen erst ein Lieferant auswählen!"
                Label3.Refresh
                Exit Sub
            End If
            Frame3.Visible = False
            Frame1.Visible = True
            Label10.Caption = "Schritt 3: Neue Zusammstellung    Fügen Sie Lieferanten der Zusammenstellung an!"
            Label10.Refresh
            List2.Clear
            
            Label2.Caption = ""
            Label2.Refresh
            
            Label9.ForeColor = glS1
            Label9.Caption = Label3.Caption
            Label9.Refresh
            
            lieferantenauflisten
            
            OlinrDetailsAnzeigen
        Case Is = 3
            Select Case Label13.Caption
                'löschen
                Case Is = "Wählen Sie bitte die zu löschende Lieferantenzusammenstellung aus!"
                    If List3.ListIndex < 0 Then
                    Label1.ForeColor = vbRed
                    Label1.Caption = "Bitte einen Eintrag auswählen!"
                    Label1.Refresh
                    List3.SetFocus
                    
                Else
                    DelOlief (Trim(List3.list(List3.ListIndex)))
                    UeberLieferantenauflisten
                End If
                'Bea
                Case Is = "Wählen Sie bitte die zu bearbeitende Lieferantenzusammenstellung aus!"
                    Screen.MousePointer = 11
                    BeaOlief
                    
                    Label2.Caption = ""
                    Label2.Refresh
                    Screen.MousePointer = 0
            End Select
            
        Case Is = 4 'zurück
            Select Case Label10.Caption
                'neu
                Case Is = "Schritt 3: Neue Zusammstellung    Fügen Sie Lieferanten der Zusammenstellung an!"
                
                    Frame3.Visible = True
                    Frame2.Visible = False
                    Frame1.Visible = False
                'Bea
                Case Is = "Schritt 2: Zusammstellung bearbeiten     Lieferanten entfernen oder anfügen!"
                    Frame3.Visible = False
                    Frame2.Visible = True
                    Frame1.Visible = False
            
            
            End Select

        Case Is = 10   'Schließen
            Unload frmWKL25
        Case Is = 5             'Zurück
            Frame2.Visible = False
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    positionierenwkl25
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    cmdHelp.Visible = istHilfeda(Me.Tag)

    Timer1.Enabled = False
    Timer2.Enabled = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherUeberli()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim lcount      As Long
    Dim sNum        As String
    Dim sBez        As String
    
    Dim sOKuerzel   As String
    Dim sOliefbez   As String
    Dim sOlinr      As String
    
    Dim rsrs        As Recordset
    
    sOlinr = ErmittleLinr(Label9.Caption)
    If Trim(sOlinr) = "" Then
        
        Label2.ForeColor = vbRed
        Label2.Caption = "Diese Zusammenstellung wurde nicht gespeichert."
        Label2.Refresh
        Exit Sub
    End If
    
    sSQL = "Select Liefbez,kuerzel from lisrt where linr = " & sOlinr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.RecordCount = 0 Then
        If Not IsNull(rsrs!LIEFBEZ) Then
            sOliefbez = Trim(rsrs!LIEFBEZ)
        Else
            sOliefbez = ""
        End If
        
        If Not IsNull(rsrs!Kuerzel) Then
            sOKuerzel = Trim(rsrs!Kuerzel)
        Else
            sOKuerzel = ""
        End If

    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    List2.Refresh
    
    If List2.ListCount = 0 Then
        Exit Sub
    End If
    
    For lcount = 0 To List2.ListCount - 1
        sNum = Right(Trim(List2.list(lcount)), 6)
        sNum = Trim(sNum)
        sBez = Trim(Left(List2.list(lcount), 35))
        
        sSQL = "Insert into UEBERLI (olinr,oliefbez,okuerzel,linr, liefbez)"
        sSQL = sSQL & " values "
        sSQL = sSQL & " ('" & sOlinr & "','" & sOliefbez & "','" & sOKuerzel & "','" & sNum & "','" & sBez & "' )"
        gdBase.Execute sSQL, dbFailOnError
    Next lcount

    sSQL = "Delete from ARTLIEF where linr = " & sOlinr
    gdBase.Execute sSQL, dbFailOnError

    For lcount = 0 To List2.ListCount - 1
        sNum = Right(Trim(List2.list(lcount)), 6)
        sNum = Trim(sNum)
        'füllen der Artlief
        'erst del dann addnew
        
        sSQL = "Insert into ARTLIEF Select "
        sSQL = sSQL & "'" & sOlinr & "' as LINR, "
        sSQL = sSQL & " ARTNR, LIBESNR, LEKPR, MINMEN,RKZ,EXDAT "
        sSQL = sSQL & " from Artlief"
        sSQL = sSQL & " where LINR = " & sNum
        gdBase.Execute sSQL, dbFailOnError

    Next lcount
'    List2.Clear
    
    Label2.ForeColor = glS1
    Label2.Caption = "Diese Zusammenstellung wurde gespeichert."
    Label2.Refresh
    Pause (2)
    Label2.ForeColor = glS1
    Label2.Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherUeberli"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub List1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Command1_Click
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List2_dblClick()
    On Error GoTo LOKAL_ERROR
    
    Command4_Click
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub positionierenwkl25()
    On Error GoTo LOKAL_ERROR
    
    With cmdHelp
        .Height = 555
        .Left = 10350
        .Top = 240
        .Width = 1305
    End With
    
    With Frame1
        .Height = 7455
        .Left = 120
        .Top = 960
        .Width = 11655
    End With
    
    With Frame2
        .Height = 7455
        .Left = 120
        .Top = 960
        .Width = 11655
    End With
    
    With Frame3
        .Height = 7455
        .Left = 120
        .Top = 960
        .Width = 11655
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "positionierenwkl25"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lieferantenauflisten()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cBez As String
    Dim cNum As String
    Dim sOlinr As String

    sOlinr = ErmittleLinr(Label9.Caption)
    
    List1.Clear
    
    
    sSQL = " Select * from LISRT where not liefbez is null"
'    sSQL = sSQL & " and HL = -1 "
    
    If sOlinr <> "" Then
    sSQL = sSQL & " and  LINR <> " & sOlinr
    End If
    sSQL = sSQL & " order by Liefbez "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.RecordCount = 0 Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!LIEFBEZ) Then
                cBez = Trim(rsrs!LIEFBEZ)
            Else
                cBez = ""
            End If
            
            If Not IsNull(rsrs!linr) Then
                cNum = Trim(rsrs!linr)
            Else
                cNum = ""
            End If
            
            List1.AddItem cBez & Space(35 - Len(cBez)) & cNum & Space(8 - Len(cNum))
            rsrs.MoveNext
        Loop
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lieferantenauflisten"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DelOlief(sDELstring As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim sOlinr      As String
    
    sOlinr = ErmittleLinr(sDELstring)
    If Trim(sOlinr) = "" Then
'        Label2.ForeColor = vbRed
'        Label2.Caption = "Diese Zusammenstellung wurde nicht gespeichert."
'        Label2.Refresh
        Exit Sub
    End If
    
    sSQL = "Delete from UEBERLI where olinr = " & sOlinr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ARTLIEF where linr = " & sOlinr
    gdBase.Execute sSQL, dbFailOnError

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DelOlief"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub BeaOlief()
    On Error GoTo LOKAL_ERROR
    
    Dim sOliefbez As String
    
    If List3.ListIndex < 0 Then
        Label1.ForeColor = vbRed
        Label1.Caption = "Bitte einen Eintrag auswählen!"
        Label1.Refresh
        List3.SetFocus
    Else
        sOliefbez = Trim(List3.list(List3.ListIndex))
        
        Label9.ForeColor = glS1
        Label9.Caption = Label1.Caption
        Label9.Refresh
        
        Frame2.Visible = False
        Frame1.Visible = True
        lieferantenauflisten
        
        OlinrDetailsAnzeigen
        Label10.Caption = "Schritt 2: Zusammstellung bearbeiten     Lieferanten entfernen oder anfügen!"
        Label10.Refresh
        List2DelList1
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BeaOlief"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub UeberLieferantenauflisten()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim cBez    As String
    Dim cNum    As String
    Dim cKuerz  As String
    
    List3.Clear
    
    sSQL = " Select distinct oliefbez,olinr,okuerzel from UEBERLI where not Oliefbez is null"
    sSQL = sSQL & " order by OLiefbez "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!oliefbez) Then
                cBez = rsrs!oliefbez
            Else
                cBez = ""
            End If
            
            If Not IsNull(rsrs!oKuerzel) Then
                cKuerz = rsrs!oKuerzel
            Else
                cKuerz = ""
            End If
            
            If Not IsNull(rsrs!oLINR) Then
                cNum = rsrs!oLINR
            Else
                cNum = ""
            End If
          
            List3.AddItem cNum & Space(8 - Len(cNum)) & cBez & Space(35 - Len(cBez)) & cKuerz
            rsrs.MoveNext
        Loop
    End If
        
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UeberLieferantenauflisten"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


