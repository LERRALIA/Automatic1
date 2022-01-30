VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWKL45 
   BackColor       =   &H00C0C000&
   Caption         =   "Auswertung der Wareneingänge"
   ClientHeight    =   8595
   ClientLeft      =   1155
   ClientTop       =   1815
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
   Icon            =   "frmWKL45.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command1 
      Height          =   315
      Index           =   7
      Left            =   10095
      TabIndex        =   71
      Top             =   7440
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
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
      Caption         =   "Etiketten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10320
      MultiSelect     =   2  'Erweitert
      TabIndex        =   70
      Top             =   1400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   68
      Text            =   "Text1"
      Top             =   840
      Width           =   615
   End
   Begin sevCommand3.Command Command1 
      Height          =   360
      Index           =   3
      Left            =   3600
      TabIndex        =   56
      Top             =   840
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
   Begin sevCommand3.Command Command1 
      Height          =   360
      Index           =   6
      Left            =   2040
      TabIndex        =   55
      Top             =   840
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
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   720
      MaxLength       =   6
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   7680
      MaxLength       =   13
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
      Height          =   615
      Left            =   5640
      TabIndex        =   46
      Top             =   7560
      Width           =   4455
      Begin VB.OptionButton Option2 
         Caption         =   "Ergebnistabelle"
         Height          =   270
         Index           =   2
         Left            =   0
         TabIndex        =   50
         Top             =   30
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Lieferscheine"
         Height          =   270
         Index           =   1
         Left            =   0
         TabIndex        =   49
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Einkaufsumsätze"
         Height          =   240
         Index           =   0
         Left            =   1320
         TabIndex        =   48
         Top             =   360
         Width           =   1575
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   47
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin sevCommand3.Command Command2 
      Height          =   495
      Left            =   120
      TabIndex        =   37
      Top             =   840
      Visible         =   0   'False
      Width           =   615
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
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5760
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   10095
      TabIndex        =   9
      Top             =   7800
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   661
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
   Begin sevCommand3.Command Command1 
      Height          =   315
      Index           =   2
      Left            =   9120
      TabIndex        =   8
      Top             =   1440
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
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
      Caption         =   "Suchen"
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame0 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      Height          =   3735
      Left            =   4200
      TabIndex        =   23
      Top             =   2880
      Visible         =   0   'False
      Width           =   3615
      Begin sevCommand3.Command Command0 
         Height          =   840
         Index           =   12
         Left            =   2640
         TabIndex        =   22
         Top             =   240
         Width           =   840
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
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   840
         Index           =   11
         Left            =   2640
         TabIndex        =   21
         Top             =   1080
         Width           =   840
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
         Caption         =   "C"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   840
         Index           =   10
         Left            =   2640
         TabIndex        =   20
         Top             =   1920
         Width           =   840
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
         Caption         =   "<"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   840
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   840
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
         Caption         =   "0"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   840
         Index           =   8
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   840
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
         Caption         =   "9"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   840
         Index           =   7
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   840
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
         Caption         =   "8"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   840
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   840
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
         Caption         =   "7"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   840
         Index           =   5
         Left            =   1800
         TabIndex        =   15
         Top             =   1080
         Width           =   840
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
         Caption         =   "6"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   840
         Index           =   4
         Left            =   960
         TabIndex        =   14
         Top             =   1080
         Width           =   840
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
         Caption         =   "5"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   840
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   840
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
         Caption         =   "4"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   840
         Index           =   2
         Left            =   1800
         TabIndex        =   12
         Top             =   1920
         Width           =   840
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
         Caption         =   "3"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   840
         Index           =   1
         Left            =   960
         TabIndex        =   11
         Top             =   1920
         Width           =   840
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
         Caption         =   "2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   840
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   840
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
         Caption         =   "1"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   720
      MaxLength       =   13
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   9720
      TabIndex        =   25
      Top             =   0
      Width           =   1935
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferschein Nr."
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Datum"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferantennummer"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikelnummer"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sortierung nach"
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
         TabIndex        =   32
         Top             =   120
         Width           =   1815
      End
   End
   Begin MSComctlLib.ProgressBar pbrZeit 
      Height          =   255
      Left            =   7080
      TabIndex        =   35
      Top             =   7140
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5175
      Left            =   240
      TabIndex        =   33
      Top             =   1920
      Visible         =   0   'False
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9128
      _Version        =   393216
      Cols            =   15
      FixedCols       =   2
      ForeColorSel    =   8454143
      FocusRect       =   0
      HighLight       =   2
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
   Begin sevCommand3.Command Command1 
      Height          =   360
      Index           =   4
      Left            =   240
      TabIndex        =   57
      Top             =   1400
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
      Picture         =   "frmWKL45.frx":0442
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command20 
      Height          =   360
      Index           =   21
      Left            =   5280
      TabIndex        =   58
      ToolTipText     =   "Kalender"
      Top             =   1395
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
   Begin sevCommand3.Command Command20 
      Height          =   360
      Index           =   20
      Left            =   3600
      TabIndex        =   59
      ToolTipText     =   "Kalender"
      Top             =   1400
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
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
      Height          =   495
      Left            =   5280
      TabIndex        =   60
      Top             =   720
      Width           =   4335
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C000&
         Caption         =   "Vormonat"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   66
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C000&
         Caption         =   "aktueller Monat"
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   65
         Top             =   0
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C000&
         Caption         =   "Gestern"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   64
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C000&
         Caption         =   "Heute"
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C000&
         Caption         =   "aktuelles Jahr"
         Height          =   255
         Index           =   12
         Left            =   2760
         TabIndex        =   62
         Top             =   0
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C000&
         Caption         =   "Vorjahr"
         Height          =   255
         Index           =   14
         Left            =   2760
         TabIndex        =   61
         Top             =   240
         Width           =   1455
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   360
      Index           =   5
      Left            =   4800
      TabIndex        =   67
      Top             =   840
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "AGN"
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
      Index           =   15
      Left            =   4080
      TabIndex        =   69
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "bes. Merkmal"
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
      Index           =   14
      Left            =   7680
      TabIndex        =   52
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   13
      Left            =   4560
      TabIndex        =   45
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   12
      Left            =   4560
      TabIndex        =   44
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nettospanne in €:"
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
      Index           =   11
      Left            =   2880
      TabIndex        =   43
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nettospanne in %:"
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
      Index           =   10
      Left            =   2880
      TabIndex        =   42
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   1560
      TabIndex        =   41
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   1560
      TabIndex        =   40
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Verkaufswert:"
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
      Index           =   7
      Left            =   240
      TabIndex        =   39
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Einkaufswert:"
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
      Index           =   6
      Left            =   240
      TabIndex        =   38
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Linie"
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
      Left            =   2520
      TabIndex        =   36
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblAnzeige 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   34
      Top             =   7200
      Width           =   6735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Lieferschein Nr.:"
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
      Index           =   4
      Left            =   5760
      TabIndex        =   31
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   11640
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "bis Datum:"
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
      Index           =   3
      Left            =   4080
      TabIndex        =   30
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "von Datum:"
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
      Left            =   2280
      TabIndex        =   29
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Lieferant"
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
      Index           =   1
      Left            =   720
      TabIndex        =   28
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
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
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   27
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Auswertungen der Wareneingänge"
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
      Height          =   735
      Left            =   240
      TabIndex        =   26
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "frmWKL45"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iIndex                  As Integer

Dim SpaltennummerArtnr      As Byte
Dim SpaltennummerBEZEICH    As Byte
Dim SpaltennummerLINR       As Byte
Dim SpaltennummerEAN        As Byte
Dim SpaltennummerLIBESNR    As Byte
Dim SpaltennummerEKPR       As Byte
Dim SpaltennummerKVKPR1     As Byte
Dim SpaltennummerAZEIT      As Byte
Dim SpaltennummerADATE      As Byte
Dim SpaltennummerBEDNU      As Byte
Dim SpaltennummerBEDNAME    As Byte
Dim SpaltennummerFILIALNR   As Byte
Dim SpaltennummerBESTALT    As Byte
Dim SpaltennummerBEWEGUNG   As Byte
Dim SpaltennummerBESTNEU    As Byte

Private Sub PositionierenWKL45()
    On Error GoTo LOKAL_ERROR
    
    Frame0.Top = 2880
    Frame0.Height = 3735
    Frame0.Width = 3615
    Frame0.Left = 4200
    
    MSFlexGrid1.Top = 1800
    MSFlexGrid1.Height = 5295
    MSFlexGrid1.Width = 11415
    MSFlexGrid1.Left = 240
    
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL45"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "DRU_EKLJ", gdBase
    loeschNEW "DRU_EINK", gdBase
    
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
Private Sub DruckeJahresWerteWKL45()
    On Error GoTo LOKAL_ERROR
    
    Dim lJahr       As Long
    Dim lLinr       As Long
    Dim cPfad       As String
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim rsJAHR      As Recordset
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
   
    Screen.MousePointer = 11
    
    loeschNEW "DRU_EKLJ", gdBase
    
    cSQL = "Create Table DRU_EKLJ "
    cSQL = cSQL & "( LINR Long"
    cSQL = cSQL & ", LIEFBEZ Text(50)"
    cSQL = cSQL & ", UMS_LFDJ Double"
    cSQL = cSQL & ", UMS_VORJ Double"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index LINR on DRU_EKLJ (LINR)"
    gdBase.Execute cSQL, dbFailOnError
    
    '**************************************
    '* alle Lieferanten einlesen
    '**************************************
    cSQL = "Insert into DRU_EKLJ "
    cSQL = cSQL & "Select LINR, LIEFBEZ, 0 as UMS_LFDJ, 0 as UMS_VORJ "
    cSQL = cSQL & " from LISRT"
    gdBase.Execute cSQL, dbFailOnError
    
    Set rsJAHR = gdBase.OpenRecordset("DRU_EKLJ", dbOpenTable)
    rsJAHR.Index = "LINR"
    
    '********************************************
    '* Einkaufssummen laufendes Jahr ermitteln
    '********************************************
    lJahr = Year(Now)
    cSQL = "Select LINR, SUM(BEWEGUNG * EKPR) as UMS_LFDJ "
    cSQL = cSQL & " from ZUGANG "
    cSQL = cSQL & " where YEAR(ADATE) = " & Trim$(Str$(lJahr))
    cSQL = cSQL & " group by LINR"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!linr) Then
                lLinr = rsrs!linr
            Else
                lLinr = -1
            End If
            rsJAHR.Seek "=", lLinr
            If Not rsJAHR.NoMatch Then
                rsJAHR.Edit
                If Not IsNull(rsrs!UMS_LFDJ) Then
                    rsJAHR!UMS_LFDJ = rsrs!UMS_LFDJ
                Else
                    rsJAHR!UMS_LFDJ = 0
                End If
                rsJAHR.Update
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
        
    '********************************************
    '* Einkaufssummen Vorjahr ermitteln
    '********************************************
    lJahr = Year(Now)
    lJahr = lJahr - 1
    cSQL = "Select LINR, SUM(BEWEGUNG * EKPR) as UMS_VORJ "
    cSQL = cSQL & " from ZUGANG "
    cSQL = cSQL & " where YEAR(ADATE) = " & Trim$(Str$(lJahr))
    cSQL = cSQL & " group by LINR"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!linr) Then
                lLinr = rsrs!linr
            Else
                lLinr = -1
            End If
            rsJAHR.Seek "=", lLinr
            If Not rsJAHR.NoMatch Then
                rsJAHR.Edit
                If Not IsNull(rsrs!UMS_VORJ) Then
                    rsJAHR!UMS_VORJ = rsrs!UMS_VORJ
                Else
                    rsJAHR!UMS_VORJ = 0
                End If
                rsJAHR.Update
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    rsJAHR.Close
    
    Screen.MousePointer = 11
    
    
    
    reportbildschirm "WKL035", "aWKL45"
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeJahreWerteWKL45"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Function fnPruefeDialogEingabeWKL45() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim lVon    As Long
    Dim lBis    As Long
    Dim cVon    As String
    Dim cBis    As String
    
    fnPruefeDialogEingabeWKL45 = 0
    
    
    
    If Trim$(Text1(0).Text) = "" _
    And Trim$(Text1(1).Text) = "" _
    And Trim$(Text1(2).Text) = "" _
    And Trim$(Text1(3).Text) = "" _
    And Trim$(Text1(4).Text) = "" _
    And Trim$(Text1(6).Text) = "" _
    And List1.Visible = False _
    And Trim$(Combo1.Text) = "" Then
        fnPruefeDialogEingabeWKL45 = 1
        Exit Function
    End If
    
    cVon = Trim$(Text1(2).Text)
    If cVon <> "" Then
        If Not IsDate(cVon) Then
            fnPruefeDialogEingabeWKL45 = 2
            Exit Function
        Else
            lVon = DateValue(cVon)
        End If
    End If
    
    cBis = Trim$(Text1(3).Text)
    If cBis <> "" Then
        If Not IsDate(cBis) Then
            fnPruefeDialogEingabeWKL45 = 3
            Exit Function
        Else
            lBis = DateValue(cBis)
        End If
    End If
    
    If cVon <> "" And cBis <> "" Then
        If lVon > lBis Then
            fnPruefeDialogEingabeWKL45 = 4
            Exit Function
        End If
    End If
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeDialogEingabeWKL45"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub FormatierteGridWKL45()
    On Error GoTo LOKAL_ERROR
    
    MSFlexGrid1.Rows = 1
    MSFlexGrid1.Cols = 15
    
    MSFlexGrid1.Row = 0
    
    MSFlexGrid1.Col = 0
    MSFlexGrid1.ColWidth(0) = 800
    MSFlexGrid1.Text = "Art.Nr."
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.ColWidth(1) = 3500
    MSFlexGrid1.Text = "Artikelbezeichnung"
    
    MSFlexGrid1.Col = 2
    MSFlexGrid1.ColWidth(2) = 800
    MSFlexGrid1.Text = "Lief.Nr."

    MSFlexGrid1.Col = 3
    MSFlexGrid1.ColWidth(3) = 1000
    MSFlexGrid1.Text = "Datum"

    MSFlexGrid1.Col = 4
    MSFlexGrid1.ColWidth(4) = 800
    MSFlexGrid1.Text = "Uhrzeit"

    MSFlexGrid1.Col = 5
    MSFlexGrid1.ColWidth(5) = 1200
    MSFlexGrid1.Text = "Alt-Bestand"

    MSFlexGrid1.Col = 6
    MSFlexGrid1.ColWidth(6) = 1200
    MSFlexGrid1.Text = "Bewegung"

    MSFlexGrid1.Col = 7
    MSFlexGrid1.ColWidth(7) = 1200
    MSFlexGrid1.Text = "Neu-Bestand"

    MSFlexGrid1.Col = 8
    MSFlexGrid1.ColWidth(8) = 700
    MSFlexGrid1.Text = "BedNr"

    MSFlexGrid1.Col = 9
    MSFlexGrid1.ColWidth(9) = 2500
    MSFlexGrid1.Text = "Bediener-Name"

    MSFlexGrid1.Col = 10
    MSFlexGrid1.ColWidth(10) = 800
    MSFlexGrid1.Text = "Filiale"

    MSFlexGrid1.Col = 11
    MSFlexGrid1.ColWidth(11) = 1500
    MSFlexGrid1.Text = "EAN"

    MSFlexGrid1.Col = 12
    MSFlexGrid1.ColWidth(12) = 1000
    MSFlexGrid1.Text = "EKPR"

    MSFlexGrid1.Col = 13
    MSFlexGrid1.ColWidth(13) = 1500
    MSFlexGrid1.Text = "Lief.Best.Nr."

    MSFlexGrid1.Col = 14
    MSFlexGrid1.ColWidth(14) = 1500
    MSFlexGrid1.Text = "Einkaufswert"

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatierteGridWKL45"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeereDialogWKL45()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    Text1(5).Text = ""
    Text1(6).Text = ""
    MSFlexGrid1.Rows = 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKL45"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SucheDatenWKL45()
    On Error GoTo LOKAL_ERROR
    
    Dim lMenge      As Long
    Dim lDatVon     As Long
    Dim lDatBis     As Long
    Dim ctmp        As String
    Dim cSQL        As String
    Dim sSQL        As String
    Dim cArtNr      As String
    Dim cLfnr      As String
    Dim cLiefNr     As String
    Dim cLinie      As String
    Dim cDatVon     As String
    Dim cDatBis     As String
    Dim cLS         As String
    
    Dim dWert       As Double
    Dim dEkWert     As Double
    Dim bAnd        As Boolean
    Dim rsrs        As Recordset
    
    Dim i As Integer
    
    
    bAnd = False
    
    cArtNr = Trim$(Text1(0).Text)
    cLfnr = Trim$(Text1(1).Text)
    cDatVon = Trim$(Text1(2).Text)
    cDatBis = Trim$(Text1(3).Text)
    cLS = Trim$(Combo1.Text)
    cLiefNr = Trim$(Text1(4).Text)
    cLinie = Trim$(Text1(5).Text)
    
        
    If cDatVon <> "" Then
        lDatVon = Fix(DateValue(cDatVon))
        cDatVon = Trim$(Str$(lDatVon))
    End If
    
    If cDatBis <> "" Then
        lDatBis = Fix(DateValue(cDatBis))
        cDatBis = Trim$(Str$(lDatBis))
    End If
    
    cSQL = " Select   "
    
    cSQL = cSQL & " Zugang.ARTNR "
    cSQL = cSQL & ", Zugang.BEZEICH "
    cSQL = cSQL & ", Zugang.LINR "
    cSQL = cSQL & ", Zugang.aDate "
    cSQL = cSQL & ", Zugang.Uhrzeit "
    cSQL = cSQL & ", Zugang.Bednu "
    cSQL = cSQL & ", Zugang.Bedname "
    cSQL = cSQL & ", Zugang.Filialnr "
    cSQL = cSQL & ", Zugang.EAN "
    cSQL = cSQL & ", Zugang.Bestandalt "
    cSQL = cSQL & ", Zugang.Bewegung "
    cSQL = cSQL & ", Zugang.Bestandneu "
    cSQL = cSQL & ", Zugang.EKPR "
    cSQL = cSQL & ", Zugang.REK "
    cSQL = cSQL & ", Zugang.Libesnr "
    cSQL = cSQL & ", Zugang.LS "
    cSQL = cSQL & ", 0 as BESTAND "
    cSQL = cSQL & " from ZUGANG "
    
    
    If Text1(6).Text <> "" Or List1.Visible = True And List1.ListCount > 0 Then
        cSQL = cSQL & " inner join Artikel on Zugang.artnr = artikel.artnr "
    End If
    
    
    
    
    
    cSQL = cSQL & " Where "

    If cArtNr <> "" Then
        If Len(cArtNr) <= 6 Then
            cSQL = cSQL & "Zugang.ARTNR = " & cArtNr & " "
            bAnd = True
        ElseIf Len(cArtNr) = 8 And Left(cArtNr, 1) = "2" Then
            cArtNr = Mid(cArtNr, 2, 6)
            cSQL = cSQL & "Zugang.ARTNR = " & cArtNr & " "
            bAnd = True
        Else
            cSQL = cSQL & "Zugang.EAN = '" & cArtNr & "' "
            bAnd = True
        End If
    End If
    
    If cLfnr <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "Zugang.LfNR = " & cLfnr & " "
        bAnd = True
    End If
    
    If cLiefNr <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "Zugang.LINR = " & cLiefNr & " "
        bAnd = True
    End If
    
    If cLS <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "Zugang.LS = '" & cLS & "' "
        bAnd = True
    End If
    
    
    
    'AGN
    If List1.Visible = True And List1.ListCount > 0 Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If

        cSQL = cSQL & "( Artikel.agn=" & Trim$(Left$(List1.list(0), InStr(1, List1.list(0), " ")))
        For i = 1 To List1.ListCount - 1
            cSQL = cSQL & " or Artikel.agn=" & Trim$(Left$(List1.list(i), InStr(1, List1.list(i), " ")))
        Next i
        cSQL = cSQL & " ) "
        bAnd = True
    Else
        'agn
        ctmp = Trim$(Text1(6).Text)
        If ctmp <> "" Then
            If bAnd Then
                cSQL = cSQL & " and "
            End If
            cSQL = cSQL & " artikel.agn  = " & ctmp & " "
            
            bAnd = True
        End If
    End If
    
    If cDatVon <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "Zugang.ADATE >= " & cDatVon & " "
        bAnd = True
    End If
    
    If cDatBis <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "Zugang.ADATE <= " & cDatBis & " "
        bAnd = True
    End If
    
    cSQL = cSQL & "and Zugang.artnr <> 999999 "
    
    If Option1(0).Value = True Then
        cSQL = cSQL & " order by Zugang.ARTNR"
    ElseIf Option1(1).Value = True Then
        cSQL = cSQL & " order by Zugang.ARTNR"
    ElseIf Option1(2).Value = True Then
        cSQL = cSQL & " order by Zugang.ADATE, Zugang.UHRZEIT"
    ElseIf Option1(3).Value = True Then
        cSQL = cSQL & " order by Zugang.LS"
    End If
    
    loeschNEW "Einkauf", gdBase
    
            
    sSQL = "Create Table Einkauf ( "
    sSQL = sSQL & " ARTNR Long"
    sSQL = sSQL & ", BEZEICH Text(35)"
    sSQL = sSQL & ", LINR Long"
    sSQL = sSQL & ", aDate Datetime"
    sSQL = sSQL & ", Uhrzeit Text(5)"
    sSQL = sSQL & ", Bednu Byte"
    sSQL = sSQL & ", Bedname Text(32)"
    sSQL = sSQL & ", Filialnr Byte"
    sSQL = sSQL & ", EAN Text(13)"
    sSQL = sSQL & ", Bestandalt Long"
    sSQL = sSQL & ", Bewegung Long"
    sSQL = sSQL & ", Bestandneu Long"
    sSQL = sSQL & ", EKPR Single"
    sSQL = sSQL & ", REK Single"
    sSQL = sSQL & ", Libesnr Text(13)"
    sSQL = sSQL & ", LS Text(20)"
    sSQL = sSQL & ", EKWERT single"
    sSQL = sSQL & ", KVKPR single"
    sSQL = sSQL & ", LEKPR single"
    sSQL = sSQL & ", lfnr autoincrement"
    sSQL = sSQL & ", BESTAND Long"
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Einkauf " & cSQL
    gdBase.Execute sSQL, dbFailOnError
    
    If cLinie <> "" Then
        sSQL = "Delete Einkauf.* from Einkauf inner join  Artikel on "
        sSQL = sSQL & " Einkauf.ARTNR = Artikel.ARTNR where "
        sSQL = sSQL & " Artikel.LPZ <> " & cLinie
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    sSQL = "Update Einkauf inner join  Artikel on "
    sSQL = sSQL & " Einkauf.ARTNR = Artikel.ARTNR set "
    sSQL = sSQL & " Einkauf.KVKPR = Artikel.KVKPR1 "
    sSQL = sSQL & ", Einkauf.BESTAND = Artikel.BESTAND "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Einkauf inner join ARTLIEF on "
    sSQL = sSQL & " Einkauf.ARTNR = ARTLIEF.ARTNR and "
    sSQL = sSQL & " Einkauf.LINR = ARTLIEF.LINR set "
    sSQL = sSQL & " Einkauf.Libesnr = ARTLIEF.Libesnr "
    sSQL = sSQL & ", Einkauf.LEKPR = ARTLIEF.LEKPR "
    gdBase.Execute sSQL, dbFailOnError
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDatenWKL45"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo1_Click()
    On Error GoTo LOKAL_ERROR
    
    Command1_Click 2
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_Click"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeereForm()
    On Error GoTo LOKAL_ERROR

    MSFlexGrid1.Visible = False
    Label1(8).Caption = ""
    Label1(8).Refresh
    Label1(9).Caption = ""
    Label1(9).Refresh
    Label1(12).Caption = ""
    Label1(12).Refresh
    Label1(13).Caption = ""
    Label1(13).Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_Click"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo1_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Command2.Visible = False
    Frame0.Visible = False
    
    iIndex = 4
    LeereForm
    
    Combo1.BackColor = glSelBack1
    Combo1.SelStart = Len(Combo1.Text)
    Label0.Caption = "4"
    
    lblAnzeige.ForeColor = glS1
    lblAnzeige.Caption = "Hier geben Sie eine Lieferscheinnummer ein!"
    lblAnzeige.Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Command20_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        
        Case Is = 20        ' Kalender
                Text1(2).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
                Text1(3).SetFocus
            
            Case Is = 21        ' Kalender
                Text1(3).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
                'fertig
    End Select
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command20_Click"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iFeldIndex As Integer
    
    iFeldIndex = Val(Label0.Caption)
    
    
    If Index < 10 Or Index > 14 Then
        Select Case iFeldIndex
            Case Is = 0
                If Index < 10 Then
                    Text1(iFeldIndex).Text = Text1(iFeldIndex).Text & Command0(Index).Caption
                End If

            Case Is = 1
                If Index < 10 Then
                    Text1(iFeldIndex).Text = Text1(iFeldIndex).Text & Command0(Index).Caption
                End If
                
            Case 2, 3, 4, 5
                Text1(iFeldIndex).Text = Text1(iFeldIndex).Text & Command0(Index).Caption
                
        End Select
    Else
        Select Case Index
            Case Is = 10        'Backspace
                If Len(Text1(iFeldIndex).Text) > 0 Then
                    Text1(iFeldIndex).Text = Left(Text1(iFeldIndex).Text, Len(Text1(iFeldIndex).Text) - 1)
                End If
            Case Is = 11        'Clear
                Select Case iFeldIndex
                    Case Is = 4
                        Combo1.Text = ""
                    Case Else
                        Text1(iFeldIndex).Text = ""
                End Select
                
            Case Is = 12        'Taste F2
                Text1_KeyUp Val(Label0.Caption), vbKeyF2, 0
            
        
        End Select
    End If
    
    Select Case iFeldIndex
        Case Is = 4
            Combo1.SetFocus
        
        Case Else
            Text1(iFeldIndex).SetFocus
    End Select
    
    
        
        
        

    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub DruckeLieferscheine()
On Error GoTo LOKAL_ERROR

    Dim lDatVon     As Long
    Dim lDatBis     As Long
    Dim cSQL        As String
    Dim sSQL        As String
    Dim cDatVon     As String
    Dim cDatBis     As String
    Dim bAnd        As Boolean
    Dim rsrs        As Recordset
    
    bAnd = False
    
    cDatVon = Trim$(Text1(2).Text)
    cDatBis = Trim$(Text1(3).Text)
    
    If cDatVon <> "" Then
        lDatVon = Fix(DateValue(cDatVon))
        cDatVon = Trim$(Str$(lDatVon))
    End If
    
    If cDatBis <> "" Then
        lDatBis = Fix(DateValue(cDatBis))
        cDatBis = Trim$(Str$(lDatBis))
    End If
    
    cSQL = "Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", LINR "
'    cSQL = cSQL & ", LIEFBEZ "
    cSQL = cSQL & ", aDate "
    cSQL = cSQL & ", Uhrzeit "
    cSQL = cSQL & ", Bednu "
    cSQL = cSQL & ", Bedname "
    cSQL = cSQL & ", Filialnr "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", Bestandalt "
    cSQL = cSQL & ", Bewegung "
    cSQL = cSQL & ", Bestandneu "
    cSQL = cSQL & ", EKPR "
    cSQL = cSQL & ", Libesnr "
    cSQL = cSQL & ", LS from ZUGANG where LS <>'' "
'    cSQL = cSQL & ", EKWERT "
'    cSQL = cSQL & ", VKWERT "
'    cSQL = cSQL & ", KVKPR "
'    cSQL = cSQL & ", LEKPR "
'    cSQL = cSQL & ", Mwst "
'    cSQL = cSQL & ", NSabso "
'    cSQL = cSQL & ", NSrela   "
    bAnd = True
    If cDatVon <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "ADATE >= " & cDatVon & " "
        bAnd = True
    End If
    
    If cDatBis <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "ADATE <= " & cDatBis & " "
        bAnd = True
    End If
    
    loeschNEW "DRU_LS", gdBase
    
    sSQL = "Create Table DRU_LS ( "
    sSQL = sSQL & " ARTNR Long"
    sSQL = sSQL & ", BEZEICH Text(35)"
    sSQL = sSQL & ", LINR Long"
    sSQL = sSQL & ", LIEFBEZ Text(35)"
    sSQL = sSQL & ", aDate Datetime"
    sSQL = sSQL & ", Uhrzeit Text(5)"
    sSQL = sSQL & ", Bednu Byte"
    sSQL = sSQL & ", Bedname Text(32)"
    sSQL = sSQL & ", Filialnr Byte"
    sSQL = sSQL & ", EAN Text(13)"
    sSQL = sSQL & ", Bestandalt Long"
    sSQL = sSQL & ", Bewegung Long"
    sSQL = sSQL & ", Bestandneu Long"
    sSQL = sSQL & ", EKPR Single"
    sSQL = sSQL & ", Libesnr Text(13)"
    sSQL = sSQL & ", LS Text(20)"
    sSQL = sSQL & ", EKWERT single"
    sSQL = sSQL & ", VKWERT single"
    sSQL = sSQL & ", KVKPR single"
    sSQL = sSQL & ", LEKPR single"
    sSQL = sSQL & ", Mwst Text(1)"
    sSQL = sSQL & ", NSabso single"
    sSQL = sSQL & ", NSrela single"
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into DRU_LS " & cSQL
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DRU_LS inner join  LISRT on "
    sSQL = sSQL & " DRU_LS.LINR = LISRT.LINR set "
    sSQL = sSQL & " DRU_LS.Liefbez = LISRT.Liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DRU_LS inner join  Artikel on "
    sSQL = sSQL & " DRU_LS.ARTNR = Artikel.ARTNR set "
    sSQL = sSQL & " DRU_LS.KVKPR = Artikel.KVKPR1 "
    sSQL = sSQL & ", DRU_LS.LEKPR = Artikel.LEKPR "
    sSQL = sSQL & ", DRU_LS.MWST = Artikel.MWST "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DRU_LS  set "
    sSQL = sSQL & " DRU_LS.EKWERT = DRU_LS.Bewegung * DRU_LS.LEKPR"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DRU_LS  set "
    sSQL = sSQL & " DRU_LS.VKWERT = DRU_LS.Bewegung * DRU_LS.KVKPR"
    gdBase.Execute sSQL, dbFailOnError
    
    'Hier gehts los
    
    sSQL = "Select * from DRU_LS where not artnr is null "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    Dim siNS As Single
    Dim sArtnr As String
    Dim iBewegung As Integer
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                sArtnr = rsrs!artnr
            End If
            
            siNS = ErmittleNSProz(sArtnr)
            rsrs.Edit
            rsrs!NSrela = siNS
            rsrs.Update
            
            If Not IsNull(rsrs!BEWEGUNG) Then
                iBewegung = rsrs!BEWEGUNG
            Else
                iBewegung = 0
            End If
            
            siNS = 0
            siNS = ErmittleNSEuro(sArtnr)
            siNS = siNS * iBewegung
    
            rsrs.Edit
            rsrs!NSabso = siNS
            rsrs.Update
            
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeLieferscheine"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim sSQL As String
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0     '** Beenden **
            Unload frmWKL45
        Case Is = 1     '** Drucken **
            If Option2(0).Value Then 'Jahreswerte
                DruckeJahresWerteWKL45
            ElseIf Option2(1).Value Then 'Lieferscheinauswertung

                DruckeLieferscheine
                reportbildschirm "dWKL045b", "aWKL45b"
            ElseIf Option2(2).Value Then
                If MSFlexGrid1.Visible Then
                    DruckeEinkaufDatenWK45
                Else
                    lblAnzeige.ForeColor = vbRed
                    lblAnzeige.Caption = "Keine Druckdaten vorhanden!"
                    lblAnzeige.Refresh
                End If
            End If
        Case Is = 2     '** Suchen **
            iRet = fnPruefeDialogEingabeWKL45()
            Select Case iRet
                Case Is = 0
                
                    lblAnzeige.ForeColor = glS1
                    lblAnzeige.Caption = "Artikel werden ermittelt..."
                    lblAnzeige.Refresh
                    Tabcheck "Einkauf"
                    FormatGridOverTablay "Einkauf"
    
                    Dim j As Integer
                    
                    With MSFlexGrid1
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
            
                    Me.Refresh
                    SucheDatenWKL45
                    FuellenMShFlex1WKL45 ""
                    
                    ermittlespalten
                    
                    Tabellenbreiteanpassen MSFlexGrid1, 1.1 * gdTabfak
                    
                Case Is = 2     '** Von-Datum falsch **
                    lblAnzeige.ForeColor = vbRed
                    lblAnzeige.Caption = "Das eingegebene VON-Datum ist falsch!"
                    lblAnzeige.Refresh
                    Text1(2).SetFocus
                Case Is = 3     '** Bis-Datum falsch **
                    MsgBox "Das eingegebene BIS-Datum ist falsch!", vbCritical, "STOP!"
                    Text1(3).SetFocus
                Case Is = 4     '** Bis-Datum > Von-Datum **
                    lblAnzeige.ForeColor = vbRed
                    lblAnzeige.Caption = "Das VON-Datum ist größer als das BIS-Datum!"
                    lblAnzeige.Refresh
                    Text1(2).SetFocus
            End Select
        Case 3
            Screen.MousePointer = 0
            Text1_KeyUp 5, vbKeyF2, 0
        Case 4      'Tabellencreator
            Screen.MousePointer = 0
            gsZSpalte = "Artnr"
            gstab = "Einkauf"
            Screen.MousePointer = 0
            frmWKL36.Show 1
        Case 5
            Screen.MousePointer = 0
            Text1_KeyUp 6, vbKeyF2, 0
        Case 6
            Screen.MousePointer = 0
            Text1_KeyUp 4, vbKeyF2, 0
        Case 7
            'Etiketten abstellen
            
            Screen.MousePointer = 11
            loeschNEW "LSTEETI", gdBase
            CreateTableT2 "LSTEETI", gdBase
            
            sSQL = "Insert into LSTEETI select Artnr "
            sSQL = sSQL & ", BEZEICH "
            sSQL = sSQL & ", BEWEGUNG as BESTAND "
            sSQL = sSQL & ", BEWEGUNG as ANZAHL "
            sSQL = sSQL & ", KVKPR as VKPR "
            sSQL = sSQL & ", LIBESNR "
            sSQL = sSQL & ", EAN "
            sSQL = sSQL & ", 1 as LPZ "
            sSQL = sSQL & ", LINR "
            sSQL = sSQL & ", '" & gcFilNr & "' as FILNR "
            sSQL = sSQL & " from Einkauf where bewegung > 0"
            
            gdBase.Execute sSQL, dbFailOnError

            gsETILS = "aus Lieferschein"
            frmWKL30.Show 1
            
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "ARTNR"
                SpaltennummerArtnr = i
            Case Is = "BEZEICH"
                SpaltennummerBEZEICH = i
            Case Is = "LINR"
                SpaltennummerLINR = i
            Case Is = "LIBESNR"
                SpaltennummerLIBESNR = i
            Case Is = "EKPR"
                SpaltennummerEKPR = i
            Case Is = "KVKPR"
                SpaltennummerKVKPR1 = i
            Case Is = "EAN"
                SpaltennummerEAN = i
            Case Is = "UHRZEIT"
                SpaltennummerAZEIT = i
            Case Is = "ADATE"
                SpaltennummerADATE = i
            Case Is = "BEDNU"
                SpaltennummerBEDNU = i
            Case Is = "BEDNAME"
                SpaltennummerBEDNAME = i
            Case Is = "FILIALNR"
                SpaltennummerFILIALNR = i
            Case Is = "BESTANDALT"
                SpaltennummerBESTALT = i
            Case Is = "BEWEGUNG"
                SpaltennummerBEWEGUNG = i
            Case Is = "BESTANDNEU"
                SpaltennummerBESTNEU = i
            
        End Select
    Next i
    
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuellenMShFlex1WKL45(sOrder As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim counter     As Long
    Dim i           As Integer
    Dim sSQL        As String
    
    sSQL = "Select * from Einkauf " & sOrder
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        Screen.MousePointer = 0
        lblAnzeige.ForeColor = vbRed
        lblAnzeige.Caption = "Es wurden keine Artikel ermittelt."
        lblAnzeige.Refresh
        
        Option2(2).Visible = False
        Option2(2).Value = False
        Exit Sub
    End If
    
    Label1(8).ForeColor = vbRed
    Label1(8).Caption = "wird ermittelt..."
    Label1(8).Refresh
    
    Label1(9).ForeColor = vbRed
    Label1(9).Caption = "wird ermittelt..."
    Label1(9).Refresh
    
    Label1(12).ForeColor = vbRed
    Label1(12).Caption = "wird ermittelt..."
    Label1(12).Refresh
    
    Label1(13).ForeColor = vbRed
    Label1(13).Caption = "wird ermittelt..."
    Label1(13).Refresh
    
    With MSFlexGrid1
    .Redraw = False
    
    pbrZeit.Visible = True
    pbrZeit.Max = 1000
    counter = 0
    
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If counter = 1000 Then
                counter = 0
            End If
            counter = counter + 1
            pbrZeit.Value = counter
            lrow = lrow + 1
            
            .Rows = lrow + 1
            .Col = 0
            
            For i = 0 To byAnzahlSpalten - 1
                .Row = 0
                .Col = i
                
                If sSpaltenname(i) = .Text Then
                    
                    Select Case sSpaltenname(i)
                    
                        Case Is = "Schnitt - EK", "le.Re. - EK"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00")

                        Case Else
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
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
        .ColWidth(i) = aBreite(i) * 1.5
    Next i
        
    
    rsrs.Close: Set rsrs = Nothing
    pbrZeit.Visible = False
    If byAnzahlSpalten < 2 Then
    
    Else
        .FixedCols = 1
    End If
    
    .RowHeight(1) = 0
    lrow = lrow - 1
    
    lblAnzeige.ForeColor = glS1
    If lrow > 1 Then
        lblAnzeige.Caption = lrow & " Artikel wurden ermittelt."
    ElseIf lrow = 1 Then
        lblAnzeige.Caption = lrow & " Artikel wurde ermittelt."
    End If
    lblAnzeige.Refresh
    
    .Redraw = True
    .Visible = True
    End With
    
    Option2(2).Visible = True
    Option2(2).Value = True
    
    
    Label1(8).Caption = ErmittleEkWert
    Label1(8).Refresh
    Label1(8).ForeColor = glS1
    Label1(8).Refresh
    
    Label1(12).Caption = ErmittleNettospanneProzent
    Label1(12).Refresh
    Label1(12).ForeColor = glS1
    Label1(12).Refresh
    
    Label1(9).Caption = ErmittleVkWert
    Label1(9).Refresh
    Label1(9).ForeColor = glS1
    Label1(9).Refresh
    
    Label1(13).Caption = ErmittleNettospanneEuro
    Label1(13).Refresh
    Label1(13).ForeColor = glS1
    Label1(13).Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMShFlex1WKL45"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

    
End Sub
Private Function ErmittleEkWert() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim dtmp As Double
    
    ErmittleEkWert = ""
    
    sSQL = "Select sum(ekpr * Bewegung)as Ekwert from einkauf"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!EKWERT) Then
            dtmp = rsrs!EKWERT
        Else
            dtmp = 0
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    ErmittleEkWert = Format$(dtmp, "###,###,##0.00 €")
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErmittleEkWert"
    Fehler.gsFehlertext = "Bei der Ermittlung des Einkaufswertes ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ErmittleVkWert() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim dtmp As Double
    
    ErmittleVkWert = ""
    
    sSQL = "Select sum(kvkpr * Bewegung)as Vkwert from einkauf"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!VKWERT) Then
            dtmp = rsrs!VKWERT
        Else
            dtmp = 0
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    ErmittleVkWert = Format$(dtmp, "###,###,##0.00 €")
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErmittleVkWert"
    Fehler.gsFehlertext = "Bei der Ermittlung des Verkaufswertes ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ErmittleNettospanneProzent() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim sArtnr      As String
    Dim siNSSum     As Single
    Dim siNS        As Single
    Dim lcount      As Long
    Dim dErg        As Double
    Dim counter     As Long
    
    ErmittleNettospanneProzent = ""
    lcount = 0
    siNSSum = 0
    siNS = 0
    
    sSQL = "Select * from Einkauf where not artnr is null "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
    
        pbrZeit.Visible = True
        pbrZeit.Max = 1000
        counter = 0
    
        Do While Not rsrs.EOF
            
            If counter = 1000 Then
                counter = 0
            End If
            counter = counter + 1
            pbrZeit.Value = counter
            
            If Not IsNull(rsrs!artnr) Then
                sArtnr = rsrs!artnr
            End If
            siNS = 0
            siNS = ErmittleNSProz(sArtnr)
            siNSSum = siNSSum + siNS
            lcount = lcount + 1
            
        rsrs.MoveNext
        Loop
        
        pbrZeit.Visible = False
        
    Else
        Exit Function
    End If
    rsrs.Close: Set rsrs = Nothing
    
    dErg = siNSSum / lcount
    
    ErmittleNettospanneProzent = Format$(dErg, "###0.00")
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErmittleVkWert"
    Fehler.gsFehlertext = "Bei der Ermittlung der Nettospanne in Prozent ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ErmittleNettospanneEuro() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim sArtnr      As String
    Dim siNSSum     As Single
    Dim siNS        As Single
    Dim iBewegung   As Integer
    Dim dErg        As Double
    Dim counter     As Long
    
    ErmittleNettospanneEuro = ""
    siNSSum = 0
    siNS = 0
    
    sSQL = "Select * from Einkauf where not artnr is null "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
    
        pbrZeit.Visible = True
        pbrZeit.Max = 1000
        counter = 0
        
        Do While Not rsrs.EOF
            
            If counter = 1000 Then
                counter = 0
            End If
            counter = counter + 1
            pbrZeit.Value = counter
            
            If Not IsNull(rsrs!artnr) Then
                sArtnr = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!BEWEGUNG) Then
                If Val(rsrs!BEWEGUNG) > 1000 Then
                    iBewegung = 0
                Else
                    iBewegung = rsrs!BEWEGUNG
                End If
            Else
                iBewegung = 0
            End If
            
            siNS = 0
            siNS = ErmittleNSEuro(sArtnr)
            siNS = siNS * iBewegung
            siNSSum = siNSSum + siNS
            
        rsrs.MoveNext
        Loop
        pbrZeit.Visible = False
    Else
        Exit Function
    End If
    rsrs.Close: Set rsrs = Nothing
    
    dErg = siNSSum
    
    ErmittleNettospanneEuro = Format$(dErg, "###,###,##0.00")
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErmittleNettospanneEuro"
    Fehler.gsFehlertext = "Bei der Ermittlung der Nettospanne in Euro ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub DruckeEinkaufDatenWK45()
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld       As String
    Dim cArtNr      As String
    Dim cLinr       As String
    Dim cDatVon     As String
    Dim cDatBis     As String
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cRowArtnr   As String
    Dim lrow        As Long
    Dim lRows       As Long
    Dim ctmp        As String
    
    If SpaltennummerLINR = 0 Or SpaltennummerBEZEICH = 0 Or SpaltennummerADATE = 0 Or SpaltennummerAZEIT = 0 Or SpaltennummerBEDNU = 0 Or SpaltennummerBEDNAME = 0 _
    Or SpaltennummerFILIALNR = 0 Or SpaltennummerEAN = 0 Or SpaltennummerBESTALT = 0 Or SpaltennummerBEWEGUNG = 0 _
    Or SpaltennummerBESTNEU = 0 Or SpaltennummerEKPR = 0 Or SpaltennummerLIBESNR = 0 Or SpaltennummerKVKPR1 = 0 Then
        MsgBox "Die Druckvorschau kann nicht erstellt werden, bitte alle verfügbaren Spalten auswählen (Tabellenansicht)", vbOKOnly, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    loeschNEW "DRU_EINK", gdBase
    
    cSQL = "Create Table DRU_EINK ("
    cSQL = cSQL & " SUARTNR Text(13)"
    cSQL = cSQL & ", SULINR Text(6)"
    cSQL = cSQL & ", ARTNR Text(6)"
    cSQL = cSQL & ", LINR Text(6)"
    cSQL = cSQL & ", DATVON Text(10)"
    cSQL = cSQL & ", DATBIS Text(10)"
    cSQL = cSQL & ", DRUTEXT Text(185)"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(13)"
    cSQL = cSQL & ", ADATE DateTime"
    cSQL = cSQL & ", ZEIT Text(5)"
    cSQL = cSQL & ", BEDNU Long"
    cSQL = cSQL & ", BEDNAME Text(32)"
    cSQL = cSQL & ", FILIALNR Long"
    cSQL = cSQL & ", BEST_ALT Long"
    cSQL = cSQL & ", BEWEGUNG Long"
    cSQL = cSQL & ", BEST_NEU Long"
    cSQL = cSQL & ", EKPR Double"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", KVKPR Double"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    cArtNr = Text1(0).Text
    cLinr = Text1(4).Text
    cDatVon = Text1(2).Text
    cDatBis = Text1(3).Text
    
    lRows = MSFlexGrid1.Rows
'    lCols = MSFlexGrid1.Cols
    MSFlexGrid1.Redraw = False
    

    For lrow = 1 To lRows - 1
        MSFlexGrid1.Row = lrow
        
        cSQL = "Insert into DRU_EINK "
        
        cSQL = cSQL & "( SUARTNR"
        cSQL = cSQL & ", SULINR "
        cSQL = cSQL & ", DATVON "
        cSQL = cSQL & ", DATBIS "
        
        cSQL = cSQL & ", ARTNR "
        cSQL = cSQL & ", BEZEICH "
        cSQL = cSQL & ", LINR "
        cSQL = cSQL & ", aDate "
        cSQL = cSQL & ", Zeit"
        cSQL = cSQL & ", Bednu "
        cSQL = cSQL & ", Bedname "
        cSQL = cSQL & ", Filialnr "
        cSQL = cSQL & ", EAN "
        cSQL = cSQL & ", best_alt"
        cSQL = cSQL & ", Bewegung "
        cSQL = cSQL & ", best_neu"
        cSQL = cSQL & ", EKPR "
        cSQL = cSQL & ", Libesnr "
        cSQL = cSQL & ", KVKPR "
        cSQL = cSQL & " )"
        cSQL = cSQL & " values ("
        cSQL = cSQL & "'" & cArtNr & "', "
        cSQL = cSQL & "'" & cLinr & "', "
        cSQL = cSQL & "'" & cDatVon & "', "
        cSQL = cSQL & "'" & cDatBis & "', "
        
        cRowArtnr = MSFlexGrid1.TextMatrix(lrow, SpaltennummerArtnr)      'ARTNR
        
        
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerArtnr)      'ARTNR
        cSQL = cSQL & "" & ctmp & ", "
                
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerBEZEICH)      'BEZEICH
        cSQL = cSQL & "'" & ctmp & "', "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerLINR)      'LINR
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerADATE)      'DATUM
        cSQL = cSQL & "'" & ctmp & "', "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerAZEIT)      'UHRZEIT
        cSQL = cSQL & "'" & ctmp & "', "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerBEDNU)      'Bednu
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerBEDNAME)      'Bedname
        cSQL = cSQL & "'" & ctmp & "', "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerFILIALNR)      'FILIALNR
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerEAN)      'EAN
        cSQL = cSQL & "'" & ctmp & "', "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerBESTALT)      'BESTALT
        cSQL = cSQL & "" & ctmp & ", "
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerBEWEGUNG)      'BEWEGUNG
        cSQL = cSQL & "" & ctmp & ", "
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerBESTNEU)      'BESTNEU
        cSQL = cSQL & "" & ctmp & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerEKPR)     'EKPR
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerLIBESNR)      'LIBESNR
        cSQL = cSQL & "'" & ctmp & "', "
        
        
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, SpaltennummerKVKPR1)     'KVKPR1
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & " "
        
        
        cSQL = cSQL & ") "
        
        
        
        
        
        
        
        If cRowArtnr = "" Then
            ctmp = ""
            cFeld = ""
            cSQL = ""
        Else

'            MsgBox cSQL
            
            cRowArtnr = ""
            gdBase.Execute cSQL, dbFailOnError
            ctmp = ""
            cFeld = ""
            cSQL = ""
        End If
        
    Next lrow
    MSFlexGrid1.Redraw = True
    
    
    
    
    
    
    
    
'    cSQL = "Insert into DRU_EINK Select "
'    cSQL = cSQL & " ARTNR "
'    cSQL = cSQL & ", BEZEICH "
'    cSQL = cSQL & ", LINR "
'    cSQL = cSQL & ", aDate "
'    cSQL = cSQL & ", Uhrzeit as Zeit"
'    cSQL = cSQL & ", Bednu "
'    cSQL = cSQL & ", Bedname "
'    cSQL = cSQL & ", Filialnr "
'    cSQL = cSQL & ", EAN "
'    cSQL = cSQL & ", Bestandalt as best_alt"
'    cSQL = cSQL & ", Bewegung "
'    cSQL = cSQL & ", Bestandneu as best_neu"
'    cSQL = cSQL & ", EKPR "
'    cSQL = cSQL & ", Libesnr "
'    cSQL = cSQL & ", KVKPR "
'    cSQL = cSQL & ", '" & cArtNr & "' as SUARTNR "
'    cSQL = cSQL & ", '" & cLinr & "' as SULINR "
'    cSQL = cSQL & ", '" & cDatVon & "' as DATVON "
'    cSQL = cSQL & ", '" & cDatBis & "' as DATBIS "
'    cSQL = cSQL & "  from Einkauf "
'    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Druckvorschau wird erstellt...", lblAnzeige
    Screen.MousePointer = 11
    
    If Modul6.FindFile(gcDBPfad, "aWKL45s.rpt") Then
        reportbildschirm "spez5b", "aWKL45s"
    Else
        reportbildschirm "WKL011", "aWKL45a"
    End If
    anzeige "normal", "", lblAnzeige
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeEinkaufDatenWK45"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Command2.Visible = False
    Frame0.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad       As String
    Dim cBName      As String
    
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "Picture"
    

    If Frame0.Visible = False Then      'Tastatur einblenden
        LeereForm
        Frame0.Visible = True
        Select Case iIndex
            Case Is = 4
                Combo1.SetFocus
            Case Else
                Text1(iIndex).SetFocus
        End Select
        
    Else                                'Tastatur ausblenden
        Frame0.Visible = False

    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Command2.Visible = False
    Frame0.Visible = False
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
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
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad       As String
    Dim cBName      As String

    Screen.MousePointer = 11
    
    PositionierenWKL45
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    iIndex = 0
    LeereDialogWKL45
    LeseLieferschein "Zugang", Combo1
    

    Text1(0).Height = 330
    Text1(1).Height = 330
    Text1(2).Height = 330
    Text1(3).Height = 330
    
    Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "Picture"
    
    cBName = "tastatur.jpg"
    If Modul6.FindFile(cPfad, cBName) Then
        Set Command2.Picture = LoadPicture(cPfad & "\" & cBName)
    Else
        Command2.FontSize = 8
        Command2.Caption = "Tastatur Ein"
        Set Command2.Picture = Nothing
    End If
    
    FormatierteGridWKL45
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
   
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    
    AutocompleteCombo KeyCode, Shift, Combo1
    Command1_Click 2
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim lcol As Long
    Dim lrow As Long
    lcol = MSFlexGrid1.Col
    lrow = MSFlexGrid1.Row
    
    If MSFlexGrid1.Row > 1 Then
        
    Else
    
        If MSFlexGrid1.TextMatrix(0, MSFlexGrid1.Col) = "Lieferantenbestnr." Then
            'sortier anders
            
            cSQL = "Update Einkauf set LIBESNR = '' where LIBESNR is null"
            gdBase.Execute cSQL, dbFailOnError
                
            If byteSortReihen = 1 Then
                
                
                FuellenMShFlex1WKL45 " order by val(LIBESNR) asc "
                    
                Tabellenbreiteanpassen MSFlexGrid1, 1.1 * gdTabfak
                
                byteSortReihen = 2
                
            ElseIf byteSortReihen = 2 Then
            
                FuellenMShFlex1WKL45 " order by val(LIBESNR) desc "
                    
                Tabellenbreiteanpassen MSFlexGrid1, 1.1 * gdTabfak
                
                byteSortReihen = 1
                
            End If
        Else
            sortierenGrid MSFlexGrid1
        End If
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Command2.Visible = False
    Frame0.Visible = False
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Option1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Command2.Visible = False
    Frame0.Visible = False
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Option1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyReturn Then
        Command1_Click 2
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case Is = 2    'vormonat
        
            If Month(DateValue(Now)) = 1 Then
                Text1(2).Text = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
                Text1(3).Text = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Else
                Text1(2).Text = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                Select Case Month(DateValue(Now)) - 1
                    Case 1, 3, 5, 7, 8, 10, 12
                        Text1(3).Text = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                    
                    Case 2
                        If Year(DateValue(Now)) = 2016 Then
                            Text1(3).Text = Format("29." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        Else
                            Text1(3).Text = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        End If
                        
                        If Year(DateValue(Now)) = 2020 Then
                            Text1(3).Text = Format("29." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        Else
                            Text1(3).Text = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        End If
                    
                    Case Else
                        Text1(3).Text = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                End Select
            End If
                
        Case Is = 5     'ak monat
            Text1(2).Text = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YYYY")
            Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
        Case Is = 6     'gestern
            Text1(2).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
            Text1(3).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
        Case Is = 7     'heute
            Text1(2).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
        Case 12 'aktuelles Jahr
            Text1(2).Text = Format("01.01." & Year(DateValue(Now)), "DD.MM.YYYY")
            Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
        Case 14 'Vorjahr
            Text1(2).Text = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Text1(3).Text = Format("31.12." & Year(Now) - 1, "DD.MM.YYYY")
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option3_Click"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    Case Is = 0
        Frame0.Visible = True
        lblAnzeige.ForeColor = glS1
        lblAnzeige.Caption = "Hier geben Sie eine Artikelnummer/EAN ein!"
        lblAnzeige.Refresh
    
        
    Case Is = 3, 2
        Command2.Visible = False
        Frame0.Visible = False
        lblAnzeige.ForeColor = glS1
        lblAnzeige.Caption = "Hier geben Sie ein Datum ein!"
        lblAnzeige.Refresh
    Case Else
        Command2.Visible = False
        Frame0.Visible = False
    End Select
    
    iIndex = Index
    LeereForm
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = Len(Text1(Index).Text)
    Label0.Caption = Trim$(Str$(Index))

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cZeichen = Chr$(KeyAscii)
    If cZeichen = "," Then
        cZeichen = "."
    End If
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
    Select Case Index
        Case 0, 1, 4, 5, 6
            cValid = "0123456789" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
        Case 2, 3
            cValid = "0123456789." & Chr$(8)
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
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim lcount As Long

    If KeyCode = vbKeyReturn Then
        Command1_Click 2
    End If
    
    If KeyCode = vbKeyF2 Then
    
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        
        Select Case Index
            Case Is = 0
                
                gF2Prompt.bMultiple = False
                    
                ctmp = Text1(4).Text
                ctmp = Trim$(ctmp)
                If ctmp = "" Then
                    lblAnzeige.ForeColor = vbRed
                    lblAnzeige.Caption = "Bitte geben Sie erst einen Lieferanten ein!"
                    lblAnzeige.Refresh
                    Text1(4).SetFocus
                    Exit Sub
                End If
                
                gF2Prompt.cFeld = "ARTNR"
                gF2Prompt.cWert = ctmp
                        
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                End If
                
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                End If
                Text1(Index).SetFocus
            Case Is = 4
                gF2Prompt.bMultiple = False
                gF2Prompt.cFeld = "LINR"
                
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                End If
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                End If
            Case 5
                ctmp = Text1(4).Text
                ctmp = Trim$(ctmp)
                If ctmp = "" Then
                    lblAnzeige.ForeColor = vbRed
                    lblAnzeige.Caption = "Bitte geben Sie erst einen Lieferanten ein!"
                    lblAnzeige.Refresh
                    Text1(4).SetFocus
                    Exit Sub
                End If
                gF2Prompt.bMultiple = False
                gF2Prompt.cFeld = "LPZ"
                gF2Prompt.cWert = ctmp
                    
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                End If
                Text1(5).Text = Trim(gF2Prompt.cWahl)
            Case Is = 6
'                gF2Prompt.bMultiple = False
'                gF2Prompt.cFeld = "AGN"
'
'                If gF2Prompt.cFeld <> "" Then
'                    frmWK00a.Show 1
'                End If
'                If gF2Prompt.cWahl <> "" Then
'                    Text1(Index).Text = gF2Prompt.cWahl
'                End If
                
                
                
                gF2Prompt.bMultiple = True
                gF2Prompt.cFeld = "AGN"
                
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    
                    
                    
                    List1.Visible = False
                    List1.Clear
                    For lcount = 0 To 100
                        If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                            List1.Visible = True
                            Text1(Index).Text = ""
                            
                            If gF2Prompt.cArray(lcount) <> "" Then
                                List1.AddItem gF2Prompt.cArray(lcount)
                            End If
                        
                        Else
                            If gF2Prompt.cArray(lcount) <> "" Then
                               
                                List1.AddItem gF2Prompt.cArray(lcount)
                                Text1(Index).Text = Left(gF2Prompt.cArray(lcount), InStr(1, gF2Prompt.cArray(lcount), " "))
                            End If
                            
                        End If
                    Next lcount
    
                End If
                
                
                
                
        End Select
    End If
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Auswertung der Wareneingänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


