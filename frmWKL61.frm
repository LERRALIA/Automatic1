VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWKL61 
   Caption         =   "Terminpreise"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL61.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   6480
      TabIndex        =   79
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame5"
      Height          =   8295
      Left            =   240
      TabIndex        =   41
      Top             =   120
      Width           =   11295
      Begin VB.TextBox Text5 
         Height          =   315
         Index           =   6
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   89
         Top             =   1200
         Width           =   1335
      End
      Begin sevCommand3.Command Command1 
         Height          =   310
         Index           =   0
         Left            =   9600
         TabIndex        =   81
         Top             =   120
         Width           =   1095
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
         Caption         =   "Import"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.PictureBox picprogress 
         Height          =   255
         Left            =   8280
         ScaleHeight     =   195
         ScaleWidth      =   1515
         TabIndex        =   80
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "nur mit Bestand"
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
         Left            =   8280
         TabIndex        =   77
         Top             =   480
         Width           =   1575
      End
      Begin sevCommand3.Command Command1 
         Height          =   315
         Index           =   9
         Left            =   4560
         TabIndex        =   76
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
      Begin VB.TextBox Text5 
         Height          =   315
         Index           =   5
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   74
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Index           =   3
         Left            =   6480
         TabIndex        =   71
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Index           =   4
         Left            =   6480
         TabIndex        =   70
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin sevCommand3.Command Command5 
         Height          =   310
         Index           =   12
         Left            =   8280
         TabIndex        =   57
         Top             =   120
         Width           =   1215
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
         Caption         =   "Anfügen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   11
         Left            =   9480
         TabIndex        =   55
         Top             =   5640
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
         Caption         =   "Entfernen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'Kein
         Height          =   1575
         Left            =   120
         TabIndex        =   45
         Top             =   5040
         Width           =   9255
         Begin VB.ComboBox cboRegalEndlos 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            TabIndex        =   93
            Text            =   "Combo1"
            Top             =   120
            Width           =   2295
         End
         Begin sevCommand3.Command Command3 
            Height          =   345
            Index           =   1
            Left            =   7800
            TabIndex        =   78
            Top             =   1200
            Width           =   1335
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
            Caption         =   "Drucken"
            Enabled         =   0   'False
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command3 
            Height          =   345
            Index           =   3
            Left            =   3120
            TabIndex        =   50
            Top             =   1200
            Width           =   1575
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
            Caption         =   "Berechnung"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command3 
            Height          =   345
            Index           =   2
            Left            =   4800
            TabIndex        =   49
            Top             =   1200
            Width           =   1215
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
            Caption         =   "Runden"
            Enabled         =   0   'False
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.TextBox Text5 
            Height          =   315
            Index           =   1
            Left            =   2040
            TabIndex        =   51
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox Text5 
            Height          =   315
            Index           =   2
            Left            =   2040
            TabIndex        =   47
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox Text5 
            Height          =   315
            Index           =   0
            Left            =   2040
            TabIndex        =   46
            Top             =   840
            Width           =   855
         End
         Begin sevCommand3.Command Command3 
            Height          =   345
            Index           =   0
            Left            =   6120
            TabIndex        =   48
            Top             =   1200
            Width           =   1575
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
            Caption         =   "Übernehmen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command3 
            Height          =   345
            Index           =   4
            Left            =   7800
            TabIndex        =   91
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
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
            Caption         =   "Etiketten"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command3 
            Height          =   315
            Index           =   5
            Left            =   4440
            TabIndex        =   94
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
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
            Caption         =   "Drucken"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command3 
            Height          =   345
            Index           =   6
            Left            =   6360
            TabIndex        =   96
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
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
            Caption         =   "aktive grün"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "Regaletikett, endlos"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   92
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "abrunden auf:"
            Height          =   255
            Index           =   2
            Left            =   3120
            TabIndex        =   20
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "aufrunden auf:"
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   69
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "Rundungskriterium:"
            Height          =   255
            Index           =   4
            Left            =   4800
            TabIndex        =   68
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "auf"
            Height          =   255
            Index           =   5
            Left            =   4320
            TabIndex        =   67
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "ab"
            Height          =   255
            Index           =   6
            Left            =   4320
            TabIndex        =   66
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "RK"
            Height          =   255
            Index           =   7
            Left            =   6720
            TabIndex        =   65
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "EK"
            Height          =   255
            Index           =   8
            Left            =   6720
            TabIndex        =   64
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "Berechnungsgrundlage:"
            Height          =   255
            Index           =   9
            Left            =   4560
            TabIndex        =   63
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "Abschlag in %"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   54
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "Abschlag in Euro"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   53
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "Preis ersetzen in Euro"
            Height          =   255
            Index           =   11
            Left            =   0
            TabIndex        =   52
            Top             =   840
            Width           =   1935
         End
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   9
         Left            =   9480
         TabIndex        =   42
         Top             =   6120
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
         Height          =   3375
         Left            =   120
         TabIndex        =   44
         Top             =   1560
         Visible         =   0   'False
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5953
         _Version        =   393216
         FocusRect       =   0
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin sevCommand3.Command Command98 
         Height          =   360
         Left            =   3120
         TabIndex        =   84
         Top             =   120
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
         Picture         =   "frmWKL61.frx":0442
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   315
         Index           =   1
         Left            =   4560
         TabIndex        =   90
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferant"
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
         Index           =   15
         Left            =   3600
         TabIndex        =   88
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbl6 
         Caption         =   "merkeRow"
         Height          =   255
         Index           =   0
         Left            =   8880
         TabIndex        =   87
         Top             =   6840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
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
         Left            =   120
         TabIndex        =   82
         Top             =   840
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Image Image2 
         Height          =   915
         Left            =   10800
         MouseIcon       =   "frmWKL61.frx":0AD4
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL61.frx":0DDE
         ToolTipText     =   "Klicken Sie hier, wenn Sie Daten aus dem MDE - Gerät einlesen möchten"
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "AGN"
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
         Index           =   14
         Left            =   3600
         TabIndex        =   75
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Artnr/EAN"
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
         Index           =   12
         Left            =   4920
         TabIndex        =   73
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Verkaufspreis"
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
         Index           =   13
         Left            =   4800
         TabIndex        =   72
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Aktionpreise zuweisen"
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
         TabIndex        =   43
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Frame4"
      Height          =   1935
      Left            =   8760
      TabIndex        =   12
      Top             =   5640
      Width           =   3135
      Begin VB.CheckBox Check2 
         Caption         =   "alle alten Aktionen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6720
         TabIndex        =   95
         Top             =   5760
         Visible         =   0   'False
         Width           =   2655
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   15
         Left            =   9480
         TabIndex        =   19
         Top             =   4320
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
         Caption         =   "Auswerten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   14
         Left            =   9480
         TabIndex        =   62
         Top             =   4800
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   8
         Left            =   9480
         TabIndex        =   39
         Top             =   5280
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   7
         Left            =   9480
         TabIndex        =   29
         Top             =   5760
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   4
         Left            =   9480
         TabIndex        =   13
         Top             =   6240
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
      Begin MSComctlLib.TreeView List3 
         Height          =   3855
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6800
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Anzahl Artikel:"
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
         Index           =   27
         Left            =   9480
         TabIndex        =   59
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   9480
         TabIndex        =   58
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   10200
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Aktions Nr.:"
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
         Index           =   14
         Left            =   7560
         TabIndex        =   36
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   7560
         TabIndex        =   35
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Bis:"
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
         Index           =   12
         Left            =   7560
         TabIndex        =   34
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   7560
         TabIndex        =   33
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Von:"
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
         Index           =   10
         Left            =   7560
         TabIndex        =   32
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   7560
         TabIndex        =   31
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   8
         Left            =   7560
         TabIndex        =   30
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Name der Aktion:"
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
         Index           =   7
         Left            =   7560
         TabIndex        =   28
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Beschreibung der Aktion:"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   27
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Preisaktionen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "Terminpreise auswählen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   7815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Frame2"
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   7800
      Width           =   1215
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   13
         Left            =   2640
         TabIndex        =   56
         Top             =   4320
         Width           =   3015
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
         Caption         =   "Artikel bearbeiten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   6
         Left            =   480
         TabIndex        =   25
         Top             =   4320
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   23
         Top             =   2400
         Width           =   7935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         MaxLength       =   30
         TabIndex        =   21
         Top             =   1320
         Width           =   7935
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   2
         Left            =   9480
         TabIndex        =   10
         Top             =   6120
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
      Begin MSComCtl2.DTPicker Text1 
         Height          =   495
         Index           =   0
         Left            =   480
         TabIndex        =   15
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         CalendarTitleBackColor=   12615680
         Format          =   144179201
         UpDown          =   -1  'True
         CurrentDate     =   38425
      End
      Begin MSComCtl2.DTPicker Text1 
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   12615680
         Format          =   144179201
         UpDown          =   -1  'True
         CurrentDate     =   38425
      End
      Begin sevCommand3.Command Command0 
         Height          =   480
         Index           =   0
         Left            =   2400
         TabIndex        =   85
         ToolTipText     =   "Kalender"
         Top             =   1320
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   847
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
         Image           =   20
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   480
         Index           =   1
         Left            =   2400
         TabIndex        =   86
         ToolTipText     =   "Kalender"
         Top             =   2400
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   847
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
         Image           =   20
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Anzahl Artikel:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   3600
         TabIndex        =   61
         Top             =   3840
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   10440
         TabIndex        =   60
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   10440
         TabIndex        =   40
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   480
         X2              =   11520
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Label Label1 
         Caption         =   "Beschreibung der Aktion:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3600
         TabIndex        =   24
         Top             =   1920
         Width           =   6855
      End
      Begin VB.Label Label1 
         Caption         =   "Name der Aktion:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3600
         TabIndex        =   22
         Top             =   840
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   18
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "bis:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   17
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Terminpreise erstellen bzw. bearbeiten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   480
         TabIndex        =   11
         Top             =   360
         Width           =   9255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF80FF&
      Caption         =   "Frame1"
      Height          =   7935
      Left            =   10680
      TabIndex        =   3
      Top             =   0
      Width           =   1215
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Terminpreisaktionen auswerten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   1080
         Value           =   -1  'True
         Width           =   6615
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Terminpreisaktionen bearbeiten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   480
         TabIndex        =   6
         Top             =   1680
         Width           =   6615
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Terminpreise löschen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   480
         TabIndex        =   5
         Top             =   2280
         Width           =   6615
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   1
         Left            =   9480
         TabIndex        =   4
         Top             =   6120
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
         Caption         =   "weiter"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Wählen Sie einen Bearbeitungsschritt aus!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   7815
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
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "nähere Informationen hier: (bitte anklicken)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   7680
      MouseIcon       =   "frmWKL61.frx":13C1
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   83
      ToolTipText     =   "hier alle Neuigkeiten lesen"
      Top             =   480
      Width           =   3975
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
      TabIndex        =   2
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
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Terminpreise"
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
      Width           =   5055
   End
End
Attribute VB_Name = "frmWKL61"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SpaltennummerAktion         As Byte
Dim SpaltennummerArtnr          As Byte
Dim SpaltennummerBEZEICH        As Byte
Dim SpaltennummerEAN            As Byte
Dim SpaltennummerBESTAND        As Byte
Dim SpaltennummerAWM            As Byte
Dim SpaltennummerKVKN           As Byte
Dim SpaltennummerKVKA           As Byte

Dim gbAnfügen                   As Boolean

Dim gbAender                    As Boolean
Dim mdeErr As Boolean

Private Sub Check1_Click()
    On Error GoTo LOKAL_ERROR
    
    If Check1.Value = vbChecked Then
'        Check1.Caption = "Markierung für alle Artikel zurücksetzen"
        flex "alle Artikel markieren"
        
    Else
'        Check1.Caption = "alle Artikel markieren"
        flex "Markierung für alle Artikel zurücksetzen"
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdGo_Click"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub flex(krit As String)
On Error GoTo LOKAL_ERROR
    
'    Dim lCount  As Long
    With MSHFLEX1
    .Redraw = False
    For lcount = 1 To .Rows - 1
        .Col = SpaltennummerAktion
        .Row = lcount
        
        Select Case krit
            Case "alle Artikel markieren"
                
                .Text = "Ja"

            Case "Markierung für alle Artikel zurücksetzen"
                .Text = "Nein"
                
            
        End Select
        
    Next lcount
    
    .Redraw = True
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "flex"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Command0_Click(Index As Integer)

    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case Is = 0
            Text1(0).Value = Format(Datumschreiben11a(3700, 260), "DD.MM.YY")
            Text1(1).Value = Text1(0).Value
            
        Case Is = 1
            Text1(1).Value = Format(Datumschreiben11a(5600, 260), "DD.MM.YY")
            'fertig
        
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim sSQL    As String
Dim rsrs As Recordset

anzeige "normal", "", Label9

Select Case Index

    Case 0  'Übernehmen
        InsertTerminPreisWK61 CInt(Label1(16).Caption), Text1(0).Value, Text1(1).Value
        Command3(1).Enabled = True
    Case 1  'drucken
        drucken CInt(Label1(16).Caption)
    Case 2  'Runden
    
        Set rsrs = gdBase.OpenRecordset("ARTT23")
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
            rsrs.Edit
            rsrs!KVKN = Runden(CDbl(rsrs!KVKN))
            rsrs.Update
            
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
       
        FuellenMShFlex1WKL61
        ermittlespalten
        Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
        
        FaerbenHGrid MSHFLEX1, CInt(SpaltennummerAWM), CInt(SpaltennummerArtnr)
        
        FaerbeHBestvor MSHFLEX1, SpaltennummerKVKN, SpaltennummerArtnr
        
        If MSHFLEX1.Visible = True Then
            MSHFLEX1.Col = 1
            MSHFLEX1.Row = 2
            MSHFLEX1.SetFocus
        End If
        
    Case 3  'Berechnung
        Berechnen
    Case 4 'Etiketten
        
        Screen.MousePointer = 11
        
        
        'erst Übernehmen
        InsertTerminPreisWK61 CInt(Label1(16).Caption), Text1(0).Value, Text1(1).Value
        
        loeschNEW "LSTEETI", gdBase
        CreateTableT2 "LSTEETI", gdBase
        
        sSQL = "Insert into LSTEETI select Artnr "
        sSQL = sSQL & ", '' as  BEZEICH "
        sSQL = sSQL & ", 1 as BESTAND "
        sSQL = sSQL & ", 1 as ANZAHL "
        sSQL = sSQL & ", KVKPR1NEU as VKPR "
        
        sSQL = sSQL & ", '' as LIBESNR "
        sSQL = sSQL & ", '' as EAN "
        sSQL = sSQL & ", 0 as LPZ "
        sSQL = sSQL & ", 0 as LINR "
        
        sSQL = sSQL & ", '" & gcFilNr & "' as FILNR "
        sSQL = sSQL & " from PRSTERM "
        sSQL = sSQL & " where preisnr = " & CInt(Label1(16).Caption)
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update LSTEETI inner join Artikel on LSTEETI.Artnr = artikel.artnr "
        sSQL = sSQL & " set LSTEETI.BEZEICH = artikel.BEZEICH"
        sSQL = sSQL & ", LSTEETI.EAN = artikel.EAN"
        sSQL = sSQL & ", LSTEETI.LPZ = artikel.LPZ"
        gdBase.Execute sSQL, dbFailOnError
        
        Set rsrs = gdBase.OpenRecordset("LSTEETI")
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
            rsrs.Edit
            rsrs!linr = ermLiefLinrmitkleinstenLEKPR(rsrs!artnr, gdBase)
            rsrs.Update
            
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        
        sSQL = "Update LSTEETI inner join Artlief on LSTEETI.Artnr = Artlief.artnr and LSTEETI.linr = Artlief.linr "
        sSQL = sSQL & " set LSTEETI.LIBESNR = Artlief.LIBESNR"
        gdBase.Execute sSQL, dbFailOnError
        
        gsETILS = "aus Lieferschein"
        frmWKL30.Show 1
        
    Case 5 'spezial Regaletiketten endlos für Terminpreise
    
    
        'erst Übernehmen
        InsertTerminPreisWK61 CInt(Label1(16).Caption), Text1(0).Value, Text1(1).Value
        
        loeschNEW "LSTEETI", gdBase
        CreateTableT2 "LSTEETI", gdBase
        
        sSQL = "Insert into LSTEETI select Artnr "
        sSQL = sSQL & ", '' as  BEZEICH "
        sSQL = sSQL & ", 1 as BESTAND "
        sSQL = sSQL & ", 1 as ANZAHL "
        sSQL = sSQL & ", KVKPR1NEU as VKPR "
        
        sSQL = sSQL & ", '' as LIBESNR "
        sSQL = sSQL & ", '' as EAN "
        sSQL = sSQL & ", 0 as LPZ "
        sSQL = sSQL & ", 0 as LINR "
        
        sSQL = sSQL & ", '" & gcFilNr & "' as FILNR "
        sSQL = sSQL & " from PRSTERM "
        sSQL = sSQL & " where preisnr = " & CInt(Label1(16).Caption)
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update LSTEETI inner join Artikel on LSTEETI.Artnr = artikel.artnr "
        sSQL = sSQL & " set LSTEETI.BEZEICH = artikel.BEZEICH"
        sSQL = sSQL & ", LSTEETI.EAN = artikel.EAN"
        sSQL = sSQL & ", LSTEETI.LPZ = artikel.LPZ"
        gdBase.Execute sSQL, dbFailOnError
        
        Set rsrs = gdBase.OpenRecordset("LSTEETI")
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
            rsrs.Edit
            rsrs!linr = ermLiefLinrmitkleinstenLEKPR(rsrs!artnr, gdBase)
            rsrs.Update
            
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        
        sSQL = "Update LSTEETI inner join Artlief on LSTEETI.Artnr = Artlief.artnr and LSTEETI.linr = Artlief.linr "
        sSQL = sSQL & " set LSTEETI.LIBESNR = Artlief.LIBESNR"
        gdBase.Execute sSQL, dbFailOnError
    
        Dim cArtNr As String
        Dim cSpezpreis As String
        ReDim acArtNr(0 To 0) As String
        ReDim acSpezPreis(0 To 0) As String
        Dim lAnzahl As Long
        lAnzahl = -1
        
        
        Set rsrs = gdBase.OpenRecordset("LSTEETI", dbOpenTable)
    
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
                cSpezpreis = rsrs!vkpr
    
                lAnzahl = lAnzahl + 1
                ReDim Preserve acArtNr(0 To lAnzahl) As String
                ReDim Preserve acSpezPreis(0 To lAnzahl) As String
            
                acArtNr(lAnzahl) = cArtNr
                acSpezPreis(lAnzahl) = cSpezpreis
               
            End If
            
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        
        
        Dim iTage As Integer
        Set rsrs = gdBase.OpenRecordset("E30")
        If Not rsrs.EOF Then
            iTage = 60
            If Not IsNull(rsrs!VKTAGE) Then
                iTage = rsrs!VKTAGE
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        
        
        Select Case cboRegalEndlos.Text
            Case "50 x 40"
                DruckeTLPRegaletikett50x40Variante3 acArtNr(), acSpezPreis(), lAnzahl, iTage, Label4
                reportbildschirmToPrinterETI "aWKL318e", gcEtikettenDrucker, True

        End Select
        
    Case 6
        
        sSQL = "Update Artikel inner join prsterm on Artikel.Artnr = prsterm.artnr  "
        sSQL = sSQL & " set Artikel.awm = '93' "
        sSQL = sSQL & " where prsterm.status = 1 "
        gdBase.Execute sSQL, dbFailOnError
        
    
    
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Berechnen()
    On Error GoTo LOKAL_ERROR

    Dim dABschlag       As Double
    Dim dNettospanne    As Double
    Dim dABschlagEuro   As Double
    Dim dErsetzungEuro  As Double
    Dim sSQL            As String
    Dim rsNs            As Recordset
    Dim dEK             As Double
    Dim cMWST           As String
    Dim dKVKN           As Double
    Dim cKVKN           As String
    Dim sKalkkrit       As String
    Dim lrow            As Long
    Dim ctmp            As String
    Dim i               As Integer
    Dim j               As Integer
    
    j = 0
    For i = 0 To 2
        If Text5(i).Text <> "" Then
            j = j + 1
        End If
    Next
    
    If j > 1 Then
        anzeigeNew "rot", "Bitte nur ein Berechnungskriterium eingeben!", Label9
        Exit Sub
        
    ElseIf j < 1 Then
        anzeigeNew "rot", "Bitte ein Berechnungskriterium eingeben!", Label9
        Exit Sub
        Text5(2).SetFocus
    End If


    
    dABschlagEuro = 0
    If Text5(2).Text <> "" Then
        If IsNumeric(Text5(2).Text) Then
            ctmp = fnMoveComma2Point$(Text5(2).Text)
            dABschlagEuro = Val(ctmp)
            
            If dABschlagEuro = 0 Then
                Text5(2).SetFocus
                anzeigeNew "rot", "Bitte einen Abschlagswert > 0 eingeben!", Label9
                Exit Sub
            End If
        Else
            anzeigeNew "rot", "Bitte einen Abschlagswert eingeben!", Label9
            Exit Sub
        End If
    End If
    
    If dABschlagEuro > 0 Then
        sSQL = "Update ARTT23 set KVKN = KVKA - " & ctmp
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    dABschlag = 0
    If Text5(1).Text <> "" Then
        If IsNumeric(Text5(1).Text) Then
            ctmp = fnMoveComma2Point$(Text5(1).Text)
            dABschlag = Val(ctmp)
            
            If dABschlag = 0 Then
                Text5(1).SetFocus
                anzeigeNew "rot", "Bitte ein Abschlagswert > 0 eingeben!", Label9
                Exit Sub
            End If
        Else
            anzeigeNew "rot", "Bitte einen Abschlagswert eingeben!", Label9
            Exit Sub
        End If
    End If
    
    If dABschlag <> 0 Then
        sSQL = "Update ARTT23 set KVKN = KVKA - ((kvka * " & ctmp & " )/100)"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    dErsetzungEuro = 0
    If Text5(0).Text <> "" Then
        If IsNumeric(Text5(0).Text) Then
            ctmp = fnMoveComma2Point$(Text5(0).Text)
            dErsetzungEuro = Val(ctmp)
            
            If dErsetzungEuro = 0 Then
                Text5(0).SetFocus
                anzeigeNew "rot", "Bitte einen Ersetzungswert > 0 eingeben!", Label9
                Exit Sub
            End If
        
        Else
            anzeigeNew "rot", "Bitte einen Ersetzungswert eingeben!", Label9
            Exit Sub
        End If
    End If
    
    If dErsetzungEuro > 0 Then
        sSQL = "Update ARTT23 set KVKN = " & ctmp
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    FuellenMShFlex1WKL61
    ermittlespalten
    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
    
    FaerbenHGrid MSHFLEX1, CInt(SpaltennummerAWM), CInt(SpaltennummerArtnr)
    
    FaerbeHBestvor MSHFLEX1, SpaltennummerKVKN, SpaltennummerArtnr
    
    If MSHFLEX1.Visible = True Then
        MSHFLEX1.Col = 1
        MSHFLEX1.Row = 2
        MSHFLEX1.SetFocus
        Command3(2).Enabled = True
    End If
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Berechnen"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub InsertTerminPreisWK61(giPreisNr As Integer, cVon As String, cBis As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim rsrs1        As Recordset
    Dim cFeld       As String
    Dim lartnr      As Long
    Dim dKVkPr1Alt  As Double
    Dim dKVkPr1Neu  As Double
    Dim lDatVon     As Long
    Dim lDatBis     As Long
    Dim cBez        As String
    Dim cEAN        As String
    Dim cBest       As String
    Dim j           As Integer
    
    
    loeschNEW "DRUAKTION", gdBase
    CreateTable "DRUAKTION", gdBase
    
    
    lDatVon = DateValue(cVon)
    lDatBis = DateValue(cBis)
    
    Screen.MousePointer = 11
    
    Set rsrs1 = gdBase.OpenRecordset("DRUAKTION", dbOpenTable)
    
    With MSHFLEX1
        .Redraw = False
        
        For j = 2 To .Rows - 1
            .Row = j
            .Col = SpaltennummerArtnr
            If IsNumeric(.Text) Then
                lartnr = .Text
                
                .Col = SpaltennummerBEZEICH
                cBez = .Text
                
                .Col = SpaltennummerEAN
                cEAN = .Text
                
                .Col = SpaltennummerBESTAND
                cBest = .Text
                
                .Col = SpaltennummerKVKN
                dKVkPr1Neu = 0
                If .Text <> "" Then
                    If IsNumeric(.Text) Then
                        dKVkPr1Neu = .Text
                    End If
                End If
                
                .Col = SpaltennummerKVKA
                dKVkPr1Alt = .Text
                
                cSQL = "Select * from PRSTERM where ARTNR = " & lartnr
                cSQL = cSQL & " and preisnr = " & giPreisNr
                Set rsrs = gdBase.OpenRecordset(cSQL)
                If Not rsrs.EOF Then
                    rsrs.Edit
                Else
                    rsrs.AddNew
                    rsrs!Status = 0
                    setzeFarbeinWK lartnr, "94"
                End If
                
                rsrs!artnr = lartnr
                rsrs!KVKPR1ALT = dKVkPr1Alt
                rsrs!KVKPR1NEU = dKVkPr1Neu
                rsrs!DAT_VON = lDatVon
                rsrs!DAT_BIS = lDatBis
                rsrs!FILIALE = gcFilNr
                rsrs!Preisnr = giPreisNr
                
                rsrs!Pos = j - 1
                rsrs.Update
                rsrs.Close: Set rsrs = Nothing
                
                rsrs1.AddNew
                rsrs1!artnr = lartnr
                rsrs1!BEZEICH = cBez
                rsrs1!EAN = cEAN
                rsrs1!BESTAND = cBest
                rsrs1!KVKPR1ALT = dKVkPr1Alt
                rsrs1!KVKPR1NEU = dKVkPr1Neu
                rsrs1!DAT_VON = lDatVon
                rsrs1!DAT_BIS = lDatBis
                rsrs1!FILIALE = gcFilNr
                rsrs1!Preisnr = giPreisNr
                rsrs1!Status = 0
                rsrs1!Pos = j - 1
                rsrs1.Update
                    
            End If
        Next j
        
        
        .Redraw = True
    End With
    
    rsrs1.Close: Set rsrs1 = Nothing
  
    Screen.MousePointer = 0
    
    anzeigeNew "normal", "", Label9

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InsertTerminPreisWK61"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub drucken(giPreisNr As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String

    loeschNEW "DRUAKTIONK", gdBase
    CreateTable "DRUAKTIONK", gdBase
    
    cSQL = "Insert into DRUAKTIONK Select * from PREISTERM where "
    cSQL = cSQL & " preisnr = " & giPreisNr
    gdBase.Execute cSQL, dbFailOnError
    
    anzeigeNew "normal", "Druckvorschau wird erstellt...", Label9
    
    Screen.MousePointer = 0
    
    reportbildschirm "", "awkl61a"
    
    anzeigeNew "normal", "", Label9
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucken"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub setzaktivenpreisZurück(cART As String, cPreisNr As String)
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim rsArt        As Recordset
    Dim cFeld       As String
    Dim lartnr      As Long
    Dim dKVkPr1Alt  As Double
    Dim dKVkPr1Neu  As Double
    Dim lDatVon     As Long
    Dim lDatBis     As Long
    Dim cBez        As String
    Dim cBest       As String
    
    Dim cRabattOk   As String
    Dim cBonusOk    As String
    Dim cPreisSchu  As String

    cSQL = "Select * from PRSTERM where artnr = " & cART & " and Preisnr =  " & cPreisNr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                lartnr = rsrs!artnr
            Else
                lartnr = -1
            End If
            
            If Not IsNull(rsrs!KVKPR1ALT) Then
                dKVkPr1Alt = rsrs!KVKPR1ALT
            Else
                dKVkPr1Alt = -1
            End If
            
            If Not IsNull(rsrs!RABATT_OK) Then
                cRabattOk = rsrs!RABATT_OK
            Else
                cRabattOk = "J"
            End If
            
            If Not IsNull(rsrs!BONUS_OK) Then
                cBonusOk = rsrs!BONUS_OK
            Else
                cBonusOk = "J"
            End If
            
            If Not IsNull(rsrs!PREISSCHU) Then
                cPreisSchu = rsrs!PREISSCHU
            Else
                cPreisSchu = "N"
            End If
            
            If lartnr <> -1 Then

                Set rsArt = gdBase.OpenRecordset("Select * from Artikel where artnr = " & lartnr)
                If Not rsArt.EOF Then
                    dVkPr = dKVkPr1Alt
                    
                    rsArt.Edit
                    rsArt!AWM = ermMerkFarbe(rsArt!artnr, "93")
                    DELMerkFarbe rsArt!artnr
                    rsArt!KVKPR1 = dKVkPr1Alt
                    rsArt!RABATT_OK = cRabattOk
                    rsArt!BONUS_OK = cBonusOk
                    rsArt!PREISSCHU = cPreisSchu
                    
                    BeginTrans
                    bTrans = True
                    rsArt.Update
                    CommitTrans
                    
                    schreibeWKEtidru CStr(lartnr), CLng(rsArt!BESTAND), Val(gcFilNr)
                    bTrans = False
                End If
                rsArt.Close: Set rsArt = Nothing
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "setzaktivenpreisZurück"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim cValid As String
Dim cFeld As String
Dim cZeichen As String
Dim lcount As Long
Dim bTextSuche As Boolean
Dim byteErg As Byte
Dim iRet As Integer

anzeige "normal", "", Label9

Select Case Index

    Case 0
        Unload frmWKL61
    Case 1
        Zeigeauswahlframe
    Case 2 'neue Terminpreise zurück
        
        Frame2.Visible = False
        Label1(16).Caption = ""
        Label1(16).Visible = True
        Zeigeauswahlframe
        
    Case 3 'Terminpreise bearbeiten zurück
        Label1(28).Caption = ermanzArtikel(CInt(Label1(16).Caption))
        Frame1.Visible = True
    Case 4 'Terminpreise löschen zurück
        Frame4.Visible = False
        Frame1.Visible = True

    Case 6 'TerminpreisAktionen speichern
        Termpreisaktionspeichern
    Case 7 'TerminpreisAktionen löschen
    
        If Check2.Value = vbChecked Then
            LoeschAlleAlten
            ZeigDelframe
        Else
            If List3.SelectedItem Is Nothing Then
                anzeige "rot", "Bitte markieren Sie eine Preisaktion!", Label9
            Else
                If Termpreisaktiondel = False Then
                    ZeigDelframe
                    anzeige "rot", "Diese Preisaktion kann nicht gelöscht werden z.Z. aktiv", Label9
                Else
                    ZeigDelframe
                End If
            End If
        End If
    Case 14 'neu
        Frame4.Visible = False
        Frame2.Visible = True

        
        Text1(0).Value = DateValue(Now)
        Text1(1).Value = DateValue(Now)
        Text2.Text = ""
        Text3.Text = ""
        Label1(28).Caption = ""
        Me.Refresh
    Case 8 'TerminpreisAktionen bearbeiten
        If List3.SelectedItem Is Nothing Then
            anzeige "rot", "Bitte markieren Sie eine Preisaktion!", Label9
        Else
            TermpreisaktionBEA
        End If
    Case 9 'Preise zuordnen zurück
        iRet = MsgBox("Möchten Sie wirklich zurück?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbYes Then
            Label1(28).Caption = ermanzArtikel(CInt(Label1(16).Caption))
            Frame5.Visible = False
        End If
    Case 10
        Frame5.Visible = True
        leereDialogF5
        SucheArtikel
        
        Label4.Caption = "Aktionspreise zuordnen: " & Trim(Text2.Text) & Space(3) & Label1(16).Caption
    Case 12
        If Text5(3).Text = "" And Text5(5).Text = "" And Text5(6).Text = "" Then
           Exit Sub
        End If
        
        If Text5(5).Text = "" And Text5(6).Text = "" Then
            Screen.MousePointer = 11
        
            cValid = "1234567890"
            cFeld = Text5(3).Text
            bTextSuche = False
            
            For lcount = 1 To Len(cFeld)
                cZeichen = Mid(cFeld, lcount, 1)
                If InStr(cValid, cZeichen) = 0 Then
                    bTextSuche = True
                    Exit For
                End If
            Next lcount
            
            If bTextSuche Then
                gcSuch = Text5(3).Text
                gsARTNR = ""
                frmWKL70.Show 1
                Me.Refresh
                If gsARTNR <> "" Then
                    Text5(3).Text = gsARTNR
                    gsARTNR = ""
                    Command5_Click 12
                End If
            Else
                byteErg = SucheArtikel
                If byteErg = 0 Then
                    Aktuelleanzeigen
                ElseIf byteErg = 2 Then
                    anzeige "rot", "keinen Artikel gefunden", Label9
                ElseIf byteErg = 1 Then
                    anzeige "rot", "Artikel schon enthalten", Label9
                End If
            End If
            
            Text5(3).Text = ""
            Text5(3).SetFocus
            Screen.MousePointer = 0
        
        Else
            If IsNumeric(Text5(5).Text) Or IsNumeric(Text5(6).Text) Then
                byteErg = SucheArtikel
                If byteErg = 0 Then
                    Aktuelleanzeigen
                ElseIf byteErg = 2 Then
                    anzeige "rot", "keinen Artikel gefunden", Label9
                ElseIf byteErg = 1 Then
                    anzeige "rot", "Artikel schon enthalten", Label9
                End If
                
            End If
            
            Text5(3).Text = ""
            Text5(3).SetFocus
            Screen.MousePointer = 0
            'oder agn weg gehen
        End If
    Case 13 ' Artikel anzeigen
    
        Command5_Click 6 'speichern
    
        Frame5.Visible = True
        leereDialogF5
        
        If Label1(16).Caption = "" Then Label1(16).Caption = "0"
            
        Artikelanzeigen CLng(Label1(16).Caption)
        
        
        Label4.Caption = "Aktionspreise zuordnen: " & Trim(Text2.Text) & Space(3) & Label1(16).Caption
        Text5(3).SetFocus
    
    Case 11
    
        If MSHFLEX1.RowSel > 1 Then
            Screen.MousePointer = 11
            FlexGrid_Update MSHFLEX1
            Screen.MousePointer = 0
        End If
        

        gcArtNr = ""
    Case 15
        If List3.SelectedItem Is Nothing Then
            anzeige "rot", "Bitte markieren Sie eine Preisaktion!", Label9
        Else
            TermpreisaktionAuswert
        End If
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub FlexGrid_Update(oGrid As MSHFlexGrid)
On Error GoTo LOKAL_ERROR

    Dim cART As String
    Dim nRow As Long
    Dim nCol As Long
    Dim nRowSel As Long
    Dim nColSel As Long
    Dim nDelRow As Long
  
    Dim lBig As Long
    
    Dim sSQL As String
    Dim rsrs As Recordset
  
    
    With oGrid
        ' aktuelle Selektion merken
        nRow = .Row
        nCol = .Col
        nRowSel = .RowSel
        nColSel = .ColSel
        
        If nRow > nRowSel Then
            lBig = nRow
            nDelRow = nRowSel - 1
        Else
            lBig = nRowSel
            nDelRow = nRow - 1
        End If
        
        Do While nDelRow < lBig
            nDelRow = nDelRow + 1
            
            If nDelRow > 1 Then
                gcArtNr = ""
                gcArtNr = MSHFLEX1.TextMatrix(nDelRow, SpaltennummerArtnr)
                cART = ""
                cART = MSHFLEX1.TextMatrix(nDelRow, SpaltennummerArtnr)
        
                If gcArtNr <> "" Then
                    If IsNumeric(gcArtNr) Then
        
                        sSQL = " Select * from PRSTERM where Preisnr = " & Label1(16).Caption
                        sSQL = sSQL & " and artnr = " & gcArtNr
                        sSQL = sSQL & " and STATUS = 1 "
                        Set rsrs = gdBase.OpenRecordset(sSQL)
                        If Not rsrs.EOF Then
                            rsrs.MoveLast
                            If rsrs.RecordCount > 0 Then
                                setzaktivenpreisZurück cART, Label1(16).Caption
                            End If
                        End If
                        rsrs.Close: Set rsrs = Nothing
        
                        sSQL = "Delete from ARTT23 where artnr = " & gcArtNr
                        gdBase.Execute sSQL, dbFailOnError
        
                        sSQL = "Delete from PRSTERM where Preisnr = " & Label1(16).Caption
                        sSQL = sSQL & " and artnr = " & gcArtNr
                        gdBase.Execute sSQL, dbFailOnError
                        
                        MSHFLEX1.TextMatrix(nDelRow, SpaltennummerArtnr) = "entfernt"
                    End If
                End If
            End If
        Loop

    End With
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FlexGrid_Update"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function SucheArtikel() As Byte
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim cSQLInsert  As String

    Dim rsrs        As Recordset
    Dim rsrs1       As Recordset
    Dim cFeld       As String
    
    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim lAnzFelder  As Long
    Dim lagn        As Long
    Dim lLinr       As Long

    Dim lcol        As Long
    Dim dWert       As Double
    Dim iRet        As Integer
    Dim cEAN        As String
    Dim cArtNr      As String
    Dim cEigNr      As String
    

    SucheArtikel = 1
    
    Me.Refresh

    Screen.MousePointer = 11
    anzeige "normal", "Daten werden ermittelt...", Label9

    lagn = 0
    
    If IsNumeric(Text5(5).Text) Then
        lagn = CLng(Text5(5).Text)
    End If
    
    lLinr = 0
    
    If IsNumeric(Text5(6).Text) Then
        lLinr = CLng(Text5(6).Text)
    End If
    
    
    cSQLInsert = "Insert into ARTT23 "
    
    
    cSQL = " Select A.ARTNR "
    cSQL = cSQL & ", A.BEZEICH"
    cSQL = cSQL & ", A.AGN"
    cSQL = cSQL & ", A.PGN"
    cSQL = cSQL & ", A.VKPR"
    cSQL = cSQL & ", A.MWST"
    
    

    cSQL = cSQL & ", B.LINR"

    
    
    cSQL = cSQL & ", A.LIBESNR"
    cSQL = cSQL & ", A.EAN"
    cSQL = cSQL & ", A.RKZ"
    cSQL = cSQL & ", A.LPZ"
    cSQL = cSQL & ", A.NOTIZEN"
    cSQL = cSQL & ", A.BESTAND"
    cSQL = cSQL & ", A.AWM"
    cSQL = cSQL & ", A.EAN2"
    cSQL = cSQL & ", A.EAN3"
    cSQL = cSQL & ", A.INHALT"
    cSQL = cSQL & ", A.INHALTBEZ"
    cSQL = cSQL & ", A.GRUNDPREIS"
    cSQL = cSQL & ", A.MINBEST "
    cSQL = cSQL & ", A.RABATT_OK"
    cSQL = cSQL & ", A.GEFUEHRT"
    cSQL = cSQL & ", A.EKPR as SEK"
    
    cSQL = cSQL & ", B.LEKPR "
    
    cSQL = cSQL & ", A.KVKPR1 as KVKA"
    cSQL = cSQL & ", A.PREISSCHU "
    cSQL = cSQL & "  "
    cSQL = cSQL & " from ARTIKEL A, Artlief B "
    

    
    cSQL = cSQL & " where ( A.SYNSTATUS = 'E' or A.SYNSTATUS = 'A' or A.SYNSTATUS is null ) "
    
   
        
        cSQL = cSQL & " and A.ARTNR = B.ARTNR "

    

    cFeld = Text5(3).Text
    cFeld = Trim$(cFeld)

    If cFeld <> "" Then
        If IsNumeric(cFeld) = True Then
            
            cEAN = cFeld
            If Len(cFeld) <= 6 Then
                cArtNr = cFeld
            Else
                cArtNr = ""
            End If
            If Left(cFeld, 1) = "2" Or Left(cFeld, 1) = "0" And Len(cFeld) = 8 Then
                cEigNr = Mid(cFeld, 2, 6)
            Else
                cEigNr = ""
            End If

            cSQL = cSQL & " and ("
            If cEAN <> "" Then
                If InStr(cEAN, "*") > 0 Then
                    cSQL = cSQL & "A.EAN like '" & cEAN & "' "
                Else
                    cSQL = cSQL & "A.EAN = '" & cEAN & "' "
                End If
                If InStr(cEAN, "*") > 0 Then
                    cSQL = cSQL & "or A.EAN2 like '" & cEAN & "' "
                Else
                    cSQL = cSQL & "or A.EAN2 = '" & cEAN & "' "
                End If
                If InStr(cEAN, "*") > 0 Then
                    cSQL = cSQL & "or A.EAN3 like '" & cEAN & "' "
                Else
                    cSQL = cSQL & "or A.EAN3 = '" & cEAN & "' "
                End If
            End If
            If cArtNr <> "" Then
                If InStr(cArtNr, "*") > 0 Then
                    cSQL = cSQL & " or A.ARTNR like '" & cArtNr & "' "
                Else
                    cSQL = cSQL & " or A.ARTNR = " & cArtNr & " "
                End If
            End If
            If cEigNr <> "" Then
                cSQL = cSQL & " or A.ARTNR = " & cEigNr & " "
            End If
            cSQL = cSQL & ") "
        Else
            Text5(3).SetFocus
            anzeige "rot", "Artikelnummer oder EAN - Code ?", Label9

            Exit Function
        End If
    End If
    
    If lagn <> 0 Then
        cSQL = cSQL & " and A.agn = " & lagn
    End If
    
    If lLinr <> 0 Then
        cSQL = cSQL & " and B.linr = " & lLinr
    End If
    
    If Check1.Value = vbChecked Then
        cSQL = cSQL & " and A.bestand > 0 "
    End If
    cSQL = cSQL & " and A.artnr not in (select artnr from ARTT23) order by B.LINR, A.LPZ, A.BEZEICH "
    

    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        gdBase.Execute cSQLInsert & cSQL, dbFailOnError
        
        
        DuplikateDelTabelle "ARTT23", gdBase, ""
        
        
        SucheArtikel = 0
    Else
        rsrs.Close: Set rsrs = Nothing
        
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
        If Not rsrs1.EOF Then
            SucheArtikel = 1
        Else
            SucheArtikel = 2
        End If
        rsrs1.Close: Set rsrs1 = Nothing
    End If
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheArtikel"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub Aktuelleanzeigen()
On Error GoTo LOKAL_ERROR

    Tabcheck "ARTTERM"
    FormatGridOverTablay "ARTTERM"

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

    FuellenMShFlex1WKL61
    ermittlespalten
    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak

    FaerbenHGrid MSHFLEX1, CInt(SpaltennummerAWM), CInt(SpaltennummerArtnr)

    FaerbeHBestvor MSHFLEX1, SpaltennummerKVKN, SpaltennummerArtnr

    If MSHFLEX1.Visible = True Then
        MSHFLEX1.Col = 1 'SpaltennummerKVKN
        MSHFLEX1.Row = 2
        MSHFLEX1.SetFocus
    End If

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Aktuelleanzeigen"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Artikelanzeigen(iPreisNr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cFeld       As String
    Dim cwhere      As String
    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim lAnzFelder  As Long
    
    Dim lcol        As Long
    Dim dWert       As Double
    Dim iRet        As Integer
    Dim cEAN        As String
    Dim cArtNr      As String
    Dim cEigNr      As String
    Dim bytePGN     As Byte
    
    Dim cJoin       As String
    Dim cdabapfad   As String
    
    Screen.MousePointer = 11
    anzeige "normal", "Daten werden ermittelt...", Label9
    
    loeschNEW "Artt23", gdBase
    CreateTable "ARTT23", gdBase
    
    '*******für die aktiven
    
    cSQL = "Insert into ARTT23 Select A.ARTNR"
    cSQL = cSQL & ", A.BEZEICH"
    cSQL = cSQL & ", A.AGN"
    cSQL = cSQL & ", A.PGN"
    cSQL = cSQL & ", A.VKPR"
    cSQL = cSQL & ", A.MWST"
    cSQL = cSQL & ", A.LINR"
    cSQL = cSQL & ", A.LIBESNR"
    cSQL = cSQL & ", A.EAN"
    cSQL = cSQL & ", A.RKZ"
    cSQL = cSQL & ", A.LPZ"
    cSQL = cSQL & ", A.NOTIZEN"
    cSQL = cSQL & ", A.BESTAND"
    cSQL = cSQL & ", A.AWM"
    cSQL = cSQL & ", A.EAN2"
    cSQL = cSQL & ", A.EAN3"
    cSQL = cSQL & ", A.INHALT"
    cSQL = cSQL & ", A.INHALTBEZ"
    cSQL = cSQL & ", A.GRUNDPREIS"
    cSQL = cSQL & ", A.MINBEST "
    cSQL = cSQL & ", A.RABATT_OK"
    cSQL = cSQL & ", A.GEFUEHRT"
    cSQL = cSQL & ", A.EKPR as SEK"
    cSQL = cSQL & ", B.KVKPR1NEU as KVKN "
    cSQL = cSQL & ", B.KVKPR1ALT as KVKA "
    cSQL = cSQL & ", A.PREISSCHU "
   
    cSQL = cSQL & " from ARTIKEL A,PRSTERM B "
    cSQL = cSQL & " where A.artnr = B.ARTNR and B.PREISNR = " & iPreisNr
    cSQL = cSQL & " and B.Status = 1 "
    
    cSQL = cSQL & " order by B.POS desc "
    gdBase.Execute cSQL, dbFailOnError
    
    '*******für die aktiven Ende
    
    cSQL = "Insert into ARTT23 Select A.ARTNR"
    cSQL = cSQL & ", A.BEZEICH"
    cSQL = cSQL & ", A.AGN"
    cSQL = cSQL & ", A.PGN"
    cSQL = cSQL & ", A.VKPR"
    cSQL = cSQL & ", A.MWST"
    cSQL = cSQL & ", A.LINR"
    cSQL = cSQL & ", A.LIBESNR"
    cSQL = cSQL & ", A.EAN"
    cSQL = cSQL & ", A.RKZ"
    cSQL = cSQL & ", A.LPZ"
    cSQL = cSQL & ", A.NOTIZEN"
    cSQL = cSQL & ", A.BESTAND"
    cSQL = cSQL & ", A.AWM"
    cSQL = cSQL & ", A.EAN2"
    cSQL = cSQL & ", A.EAN3"
    cSQL = cSQL & ", A.INHALT"
    cSQL = cSQL & ", A.INHALTBEZ"
    cSQL = cSQL & ", A.GRUNDPREIS"
    cSQL = cSQL & ", A.MINBEST "
    cSQL = cSQL & ", A.RABATT_OK"
    cSQL = cSQL & ", A.GEFUEHRT"
    cSQL = cSQL & ", A.EKPR as SEK"
    cSQL = cSQL & ", B.KVKPR1NEU as KVKN "
    cSQL = cSQL & ", A.KVKPR1 as KVKA "
    cSQL = cSQL & ", A.PREISSCHU "
   
    cSQL = cSQL & " from ARTIKEL A,PRSTERM B "
    cSQL = cSQL & " where A.artnr = B.ARTNR and B.PREISNR = " & iPreisNr
    cSQL = cSQL & " and (B.Status = 99 or B.Status = 0)"
    
    cSQL = cSQL & " order by B.POS desc "
    gdBase.Execute cSQL, dbFailOnError
    
    Tabcheck "ARTTERM"
    FormatGridOverTablay "ARTTERM"
    
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
    
    FuellenMShFlex1WKL61
    ermittlespalten
    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
    
    FaerbenHGrid MSHFLEX1, CInt(SpaltennummerAWM), CInt(SpaltennummerArtnr)
    
    FaerbeHBestvor MSHFLEX1, SpaltennummerKVKN, SpaltennummerArtnr
    
    If MSHFLEX1.Visible = True Then
        MSHFLEX1.Col = 1 'SpaltennummerKVKN
        MSHFLEX1.Row = 2
        MSHFLEX1.SetFocus
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Artikelanzeigen"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FaerbeHBestvor(gridx As MSHFlexGrid, spaltebestvor As Byte, spalteartnr As Byte)
    On Error GoTo LOKAL_ERROR

    Dim j           As Integer
    With gridx
        .Redraw = False
        For j = 1 To .Rows - 1
            .Row = j
            .Col = spaltebestvor
            .CellBackColor = vbGreen
        Next j
        .Redraw = True
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "FaerbeHBestvor"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuellenMShFlex1WKL61()
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
    
    sSQL = "Select * from ARTT23 order by pos desc"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    MSHFLEX1.Redraw = False
    MSHFLEX1.Visible = False
    
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        

            lrow = lrow + 1
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Col = 0
            
            For i = 0 To byAnzahlSpalten - 1
                MSHFLEX1.Row = 0
                MSHFLEX1.Col = i
                
                If sSpaltenname(i) = MSHFLEX1.Text Then
                    
                    Select Case sSpaltenname(i)
                        Case Is = "Listen - EK", "KVK regulär", "KVK Aktion", "Schnitt - EK"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            MSHFLEX1.Row = lrow
                            MSHFLEX1.Text = Format$(sWert, "####0.00")

                        Case Is = "Diff KVK in Euro", "Diff KVK in %"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            MSHFLEX1.Row = lrow
                            MSHFLEX1.Text = Format$(sWert, "###0.00")

                        Case Else
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            MSHFLEX1.Row = lrow
                            MSHFLEX1.Text = sWert
                        
                    End Select
                    
            
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
    
    Screen.MousePointer = 0
    
    If lrow > 1 Then
        
        anzeige "normal", lrow & " Artikel wurden ermittelt.", Label9
    ElseIf lrow = 1 Then
        anzeige "normal", lrow & " Artikel wurde ermittelt.", Label9
    Else
'        anzeige "rot", "Es wurden keine Artikel ermittelt.", Label9
'
'        Exit Sub
    End If
   
    
    MSHFLEX1.Redraw = True
    MSHFLEX1.Visible = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMShFlex1WKL61"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Zeigeauswahlframe()
    On Error GoTo LOKAL_ERROR
    
    Frame1.Visible = False
    
    If Option2(0).Value = True Then         'Auswerten
        ZeigDelframe
        Command5(8).Visible = False
        Command5(14).Visible = False
        Command5(7).Visible = False
        Check2.Visible = False
        
        Command5(15).Visible = True
        
    ElseIf Option2(1).Value = True Then     'Bearbeiten/Neu
       
        ZeigDelframe
        Command5(7).Visible = False
        Command5(15).Visible = False
        Check2.Visible = False
        
        Command5(8).Visible = True
        Command5(14).Visible = True
        
    ElseIf Option2(2).Value = True Then     'Löschen
        
        ZeigDelframe
        Command5(8).Visible = False
        Command5(14).Visible = False
        Command5(15).Visible = False
        
        Command5(7).Visible = True
        Check2.Visible = True
        

    End If
    
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zeigeauswahlframe"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigDelframe()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim iPreisNr    As Integer
    Dim iStatus     As Integer
    Dim sPreisname  As String
    Dim lcount      As Long
    Dim cSatz       As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim lHeute      As Long
    Dim anzart      As Long
    Dim rsDe        As Recordset
    
    lHeute = DateValue(Now)
    Frame4.Visible = True
    List3.Nodes.Clear
    
    sSQL = " Select * from Preisterm"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Preisnr) Then
                iPreisNr = rsrs!Preisnr
                
                anzart = ermanzArtikel(iPreisNr)
                If Not IsNull(rsrs!preisname) Then
                    sPreisname = rsrs!preisname
                    
                    cSatz = CStr(iPreisNr) & Space(6 - Len(CStr(iPreisNr))) & Left(sPreisname, 30)
                    cSatz = cSatz & Space(37 - Len(cSatz)) & CStr(anzart) & Space(5 - Len(CStr(anzart))) & " Artikel"
                    
                    sSQL = " Select max(Status) as disStatus from Prsterm where Preisnr = " & iPreisNr
                    Set rsDe = gdBase.OpenRecordset(sSQL)
                    If Not rsDe.EOF Then
                        
                        rsDe.MoveFirst
                        Do While Not rsDe.EOF
                            If Not IsNull(rsDe!disStatus) Then
                                iStatus = rsDe!disStatus
                                Select Case iStatus
                                
                                    Case 0
                                        cSatz = cSatz & " vorbereitet"
                                        List3.Nodes.Add Text:=cSatz
'                                        List3.Nodes(List3.Nodes.Count).BackColor = vbWhite
                                    Case 1
                                        cSatz = cSatz & " z.Z. aktiv"
                                        List3.Nodes.Add Text:=cSatz
'                                        List3.Nodes(List3.Nodes.Count).BackColor = vbRed
                                    
                                    Case 99
                                        cSatz = cSatz & " ausgelaufen"
                                        List3.Nodes.Add Text:=cSatz
'                                        List3.Nodes(List3.Nodes.Count).BackColor = vbBlue
                                End Select
                            Else
                                cSatz = cSatz & " vorbereitet"
                                List3.Nodes.Add Text:=cSatz
                        
                            End If
                        
                        rsDe.MoveNext
                        Loop
                    
                    End If
                    rsDe.Close: Set rsDe = Nothing
                    
                    

                End If
            End If
        
        rsrs.MoveNext
        Loop
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If lcount > 0 Then
        List3.Nodes(1).Selected = True
        List3_NodeClick List3.Nodes(1)
    End If
    
    List3.SetFocus
    
    anzeige "normal", lcount & " Preisaktionen werden angezeigt.", Label9
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigDelframe"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Function ermanzArtikel(iPreisNr As Integer) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim lcount      As Long
    
    ermanzArtikel = 0

    sSQL = " Select * from PRSTERM where Preisnr = " & iPreisNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        ermanzArtikel = rsrs.RecordCount
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermanzArtikel"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Termpreisaktionspeichern()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim sPreisname  As String
    Dim sPreisbesch As String
    Dim iPreisNr    As Integer
    Dim cVon        As String
    Dim cBis        As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim lHeute      As Long
    
    sPreisname = Trim(Text2.Text)
    If Label1(16).Caption <> "" Then
        iPreisNr = CInt(Label1(16).Caption)
    Else
        iPreisNr = ermMaxPreisnr
        
        If sPreisname = "" Then
            anzeige "rot", "Vergeben Sie einen Namen für diese Aktion!", Label9
            Text2.SetFocus
            
            Exit Sub
        End If
        
        If checksPreisname(sPreisname) = False Then
            anzeige "rot", "Vergeben Sie einen noch nicht vergebenen Namen für diese Aktion!", Label9
            Text2.SetFocus
            
            Exit Sub
        End If
    End If

    sPreisbesch = Trim(Text3.Text)
    
    
    cVon = Text1(0).Value
    cBis = Text1(1).Value
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)

    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    If Label1(16).Caption <> "" Then
        sSQL = "Update PREISTERM  set "
        sSQL = sSQL & " Preisname = '" & sPreisname & "'"
        sSQL = sSQL & " , Preisbesch = '" & sPreisbesch & "'"
        sSQL = sSQL & " , von = " & lVon & " "
        sSQL = sSQL & " , bis =  " & lBis & " "
        sSQL = sSQL & " where preisnr = " & iPreisNr
    Else
        sSQL = "Insert into PREISTERM (Preisname,Preisbesch,preisnr,von,bis)"
        sSQL = sSQL & " values  "
        sSQL = sSQL & " ( '" & sPreisname & "'"
        sSQL = sSQL & " , '" & sPreisbesch & "'"
        sSQL = sSQL & " , " & iPreisNr & " "
        sSQL = sSQL & " , " & lVon & " "
        sSQL = sSQL & " , " & lBis & " ) "
        
        Label1(16).Caption = iPreisNr
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update PRSTERM set "
    sSQL = sSQL & "  DAT_VON = " & lVon & " "
    sSQL = sSQL & " , DAT_BIS =  " & lBis & " "
    sSQL = sSQL & " where preisnr = " & iPreisNr
    gdBase.Execute sSQL, dbFailOnError
    
    lHeute = Fix(Now)
    
    If lHeute <= lVon Or lHeute < lBis Then
        'nur dann kann man eine schon ausgelaufene reaktivieren
    
        sSQL = "Update PRSTERM set "
        sSQL = sSQL & " Status = 0 "
        sSQL = sSQL & " where preisnr = " & iPreisNr
        sSQL = sSQL & " and Status = 99 "
        gdBase.Execute sSQL, dbFailOnError
    
    End If

    anzeige "normal", "Daten wurden gespeichert.", Label9
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Termpreisaktionspeichern"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Termpreisaktiondel() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim iPreisNr    As Integer
    Dim idisStatus  As Integer
    Dim lHeute      As Long
    Dim rsrs        As DAO.Recordset
   
    Termpreisaktiondel = False
    
    iPreisNr = CInt(Trim(Label1(15).Caption))
    
    lHeute = DateValue(Now)
    
    sSQL = " Select max(Status) as disStatus from Prsterm where preisnr = " & iPreisNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
  
        If Not IsNull(rsrs!disStatus) Then
            idisStatus = rsrs!disStatus
        End If
        
        Select Case idisStatus
            Case 0
                Termpreisaktiondel = True
            
            Case 1
                Termpreisaktiondel = False
            Case 99
                Termpreisaktiondel = True
        End Select
        
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Termpreisaktiondel = True Then
        sSQL = "Delete from PREISTERM where preisnr = " & iPreisNr
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Delete from PRSTERM where preisnr = " & iPreisNr
        gdBase.Execute sSQL, dbFailOnError
    End If
   
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Termpreisaktiondel"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub LoeschAlleAlten()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim iPreisNr    As Integer
    Dim rsrs        As DAO.Recordset
   
    sSQL = " Select distinct(preisnr) as dispreisnr from Prsterm where Status = 99 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!dispreisnr) Then
                iPreisNr = rsrs!dispreisnr
            End If
            
            sSQL = "Delete from PREISTERM where preisnr = " & iPreisNr
            gdBase.Execute sSQL, dbFailOnError
    
            sSQL = "Delete from PRSTERM where preisnr = " & iPreisNr
            gdBase.Execute sSQL, dbFailOnError
            
            
            rsrs.MoveNext
        Loop
            
    End If
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoeschAlleAlten"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TermpreisaktionBEA()
    On Error GoTo LOKAL_ERROR
    
    
    Dim cLBSatz As String
    
    cLBSatz = List3.SelectedItem.Text
    cLBSatz = Left(cLBSatz, 6)
    cLBSatz = Trim$(cLBSatz)
    
    
    Text1(0).Value = DateValue(Now)
    Text1(1).Value = DateValue(Now)
    Text2.Text = ""
    Text3.Text = ""
    Label1(16).Caption = ""
    Label1(16).Visible = True
    
    Frame4.Visible = False
    Frame2.Visible = True
    
    
    
    zeigePreistermaktiontoBea cLBSatz
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TermpreisaktionBEA"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TermpreisaktionAuswert()
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    
    cLBSatz = List3.SelectedItem.Text
    cLBSatz = Left(cLBSatz, 6)
    cLBSatz = Trim$(cLBSatz)

    AuswertungAktion cLBSatz
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TermpreisaktionAuswert"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command98_Click()
On Error GoTo LOKAL_ERROR

    gsZSpalte = "Artnr"
    gsZSpalte1 = "KVKN"
    gstab = "ARTTERM"
    frmWKL36.Show 1
    'fertig
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command98_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    PositionierenWKL61
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    anzeige "normal", "", Label9
    
    If Not NewTableSuchenDBKombi("PREISTERM", gdBase) Then
        CreateTable "PREISTERM", gdBase
    End If
    
    Label1(16).Caption = ""
    gbAnfügen = False
    
    fülleCboEtikettenSpezTermin cboRegalEndlos
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub leereDialogF5()
On Error GoTo LOKAL_ERROR
    
    MSHFLEX1.Visible = False
    
    Text5(0).Text = ""
    Text5(1).Text = ""
    Text5(2).Text = ""
    
    anzeige "normal", "", Label9
    
    Label3(5).Caption = giAufrunden
    Label3(6).Caption = giAbrunden
    Label3(7).Caption = giRundkrit
    Label3(8).Caption = IIf(gsSpanne = "LEK", "List - EK", "Schnitt - EK")
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "leereDialogF5"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
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
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "TERMERR", gdBase
    loeschNEW "DRUAKTIONA", gdBase
    loeschNEW "DRUAKTIONK", gdBase
    loeschNEW "DRUAKTION", gdBase
    loeschNEW "Artt23", gdBase
    
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
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "ETIMERK"
            SpaltennummerAktion = i
            Case Is = "ARTNR"
                SpaltennummerArtnr = i
            Case Is = "BEZEICH"
                SpaltennummerBEZEICH = i
            Case Is = "EAN"
                SpaltennummerEAN = i
            Case Is = "BESTAND"
                SpaltennummerBESTAND = i
            Case Is = "AWM"
                SpaltennummerAWM = i
            Case Is = "KVKN"
                SpaltennummerKVKN = i
            Case Is = "KVKA"
                SpaltennummerKVKA = i

        End Select
        
    Next i
    
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FormatMShFlex1WKL61()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsRL As Recordset
    
    Dim i As Byte
    Dim j As Byte
    
    sSQL = "Select * from TABLay" & srechnertab & " where ANZEIGE = 'J' and Tabname = 'ARTTERM' order by Reihenf"
    Set rsRL = gdBase.OpenRecordset(sSQL)
    
    If rsRL.RecordCount = 0 Then
    
    Else
        byAnzahlSpalten = rsRL.RecordCount
        ReDim sSpaltenname(byAnzahlSpalten)
        ReDim sSpaltenbez(byAnzahlSpalten)
        ReDim aBreite(byAnzahlSpalten)
        rsRL.MoveFirst
        i = 0
        Do While Not rsRL.EOF
            sSpaltenname(i) = rsRL!Spaltenna
            sSpaltenbez(i) = rsRL!Spaltenbez
            i = i + 1
            rsRL.MoveNext
        Loop
    End If
    rsRL.Close
    
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
'            aBreite(j) = TextWidth(.TextMatrix(0, j)) ' * 1.8
        Next j
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatMShFlex1WKL61"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub
Private Function ermittlePGN(SPGNBEZ As String) As Byte
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    ermittlePGN = 255
    If SPGNBEZ <> "" Then
        sSQL = " Select PGN from PGNDBF where PGNBEZEICH = '" & SPGNBEZ & "'"
        Set rs = gdBase.OpenRecordset(sSQL)
        If Not rs.EOF Then
            If Not IsNull(rs!PGN) Then
                ermittlePGN = rs!PGN
            End If
        End If
        rs.Close: Set rs = Nothing
    End If
    
    
    
   
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlePGN"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function


Private Sub Image2_Click()
On Error GoTo LOKAL_ERROR
    
    MDElesen
    If mdeErr Then
        reportbildschirm "", "aWKL46e" 'Error artikel mde
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Image2_Click"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MDElesen()
    On Error GoTo LOKAL_ERROR
    
    If MDEeinlesenOhneLinr(Label9, txtStatus, picprogress, frmWKL61) = False Then
        anzeigeNew "rot", "Es konnten keine Daten aus dem MDE - Gerät ausgelesen werden.", Label9
    Else
        anzeigeNew "normal", "", Label9
        MdeVerarbeitung
        
        Aktuelleanzeigen
        
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
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MdeVerarbeitung()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsMDE       As Recordset
    Dim rsRS2       As Recordset
    Dim rsFilBu     As Recordset
    Dim rsArt       As Recordset
    Dim seekEAN     As String
    Dim lMenge      As Long
    Dim lscanfolge  As Long
    Dim sArtnr      As String
    
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
                    
                    Set rsRS2 = gdBase.OpenRecordset("select * from ARTT23 where artnr = " & sArtnr)
                    If Not rsRS2.EOF Then
                        
                    Else
     
                        rsRS2.AddNew
                        rsRS2!artnr = rsArt!artnr
                        rsRS2!BEZEICH = rsArt!BEZEICH
                        rsRS2!AGN = rsArt!AGN
                        rsRS2!PGN = rsArt!PGN
                        rsRS2!vkpr = rsArt!vkpr
                        rsRS2!MWST = rsArt!MWST
                        rsRS2!linr = rsArt!linr
                        rsRS2!LIBESNR = rsArt!LIBESNR
                        rsRS2!EAN = rsArt!EAN
                        rsRS2!RKZ = rsArt!RKZ
                        rsRS2!LPZ = rsArt!LPZ
                        rsRS2!NOTIZEN = rsArt!NOTIZEN
                        rsRS2!BESTAND = rsArt!BESTAND
                        rsRS2!AWM = rsArt!AWM
                        rsRS2!EAN2 = rsArt!EAN2
                        rsRS2!EAN3 = rsArt!EAN3
                        rsRS2!INHALT = rsArt!INHALT
                        rsRS2!INHALTBEZ = rsArt!INHALTBEZ
                        rsRS2!GRUNDPREIS = rsArt!GRUNDPREIS
                        rsRS2!MINBEST = rsArt!MINBEST
                        rsRS2!RABATT_OK = rsArt!RABATT_OK
                        rsRS2!GEFUEHRT = rsArt!GEFUEHRT
                        rsRS2!sEK = rsArt!ekpr
                        rsRS2!KVKA = rsArt!KVKPR1
                        rsRS2!PREISSCHU = rsArt!PREISSCHU
                        
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
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

If Index = 17 Then
    URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/171-preisaktion.html"
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_DblClick(Index As Integer)
On Error GoTo LOKAL_ERROR

If Index = 23 Then
    Label1(Index).Caption = "alle Farben"
    Label1(Index).Tag = ""
    Label1(Index).BackColor = Label1(20).BackColor
    Label1(Index).ForeColor = Label1(20).ForeColor
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSHFLEX1_DblClick()
On Error GoTo LOKAL_ERROR

    If MSHFLEX1.Row > 1 Then

    Else
        sortierenHGrid MSHFLEX1
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_Click"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    Dim lcol As Long
    Dim lrow As Long
    
    lcol = MSHFLEX1.Col
    lrow = MSHFLEX1.Row
    
    
    
    
    lbl6(0).Caption = lrow
    
    
    
    
    cZeichen = Chr$(KeyAscii)
    
    Select Case lcol
         Case Is = SpaltennummerKVKN
            gbAenderKVK = True
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSHFLEX1.Row = lrow
                MSHFLEX1.Col = lcol
                cValid = MSHFLEX1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSHFLEX1.Text = cValid
                End If
            End If
    
        Case Is = SpaltennummerBESTAND
            gbAender = True
            cValid = "1234567890-" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSHFLEX1.Row = lrow
                MSHFLEX1.Col = lcol
                cValid = MSHFLEX1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSHFLEX1.Text = cValid
                End If
            End If
    
     End Select
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil  Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long
    
    lrow = MSHFLEX1.Row
    lcol = MSHFLEX1.Col
    

    
    If KeyCode <> vbKeyDown And KeyCode <> vbKeyUp And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft Then
    
        Select Case lcol
            Case Is = SpaltennummerKVKN, SpaltennummerBESTAND
        
                If iKeypress = 0 And KeyCode <> vbKeyBack And KeyCode <> vbKeyF2 Then
                    If KeyCode = 187 Or KeyCode = 189 Then
                    
                    Else
                        MSHFLEX1.Row = lrow
                        MSHFLEX1.Col = lcol
                        MSHFLEX1.Text = ""
                    
                    End If
                    
                ElseIf iKeypress > 0 And KeyCode = 46 Then
                
                    MSHFLEX1.Row = lrow
                    MSHFLEX1.Col = lcol
                    MSHFLEX1.Text = ""
                
                End If
                iKeypress = iKeypress + 1
        End Select
    End If
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
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
            
            MSHFLEX1.TopRow = lrow
            MSHFLEX1.Col = SpaltennummerBESTVOR
            MSHFLEX1.Row = lrow
            MSHFLEX1.SetFocus
            
            Screen.MousePointer = 0
        End If
        gsARTNR = ""
    End If
    
    MSHFLEX1.Redraw = True

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten. "
    
    Fehlermeldung1

End Sub
Private Sub MSHFLEX1_LeaveCell()
On Error GoTo LOKAL_ERROR
    
    iKeypress = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_LeaveCell"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_LostFocus()
On Error GoTo LOKAL_ERROR
    
    MSHFLEX1_LeaveCell
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_SelChange()
On Error GoTo LOKAL_ERROR

    Dim lColmerker          As Long
    Dim lRowmerker          As Long
    Dim cartnrzuSpeichern   As String
    Dim lBest               As Long
    
    If MSHFLEX1.Row > 1 Then
    
        MSHFLEX1.Redraw = False
        
        If gbAender Then
            lColmerker = MSHFLEX1.Col
            lRowmerker = MSHFLEX1.Row
            
            lBest = Val(MSHFLEX1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerBESTAND))
            If lBest > 1000 Then
                MsgBox MSHFLEX1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerBESTAND) & " Dieser Wert wird nicht gespeichert.", vbInformation, "Winkiss Hinweis:"
                gbAender = False
                MSHFLEX1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerBESTAND) = 0
                MSHFLEX1.Redraw = True
                Exit Sub
            ElseIf lBest < -1000 Then
                MsgBox MSHFLEX1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerBESTAND) & " Dieser Wert wird nicht gespeichert.", vbInformation, "Winkiss Hinweis:"
                gbAender = False
                MSHFLEX1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerBESTAND) = 0
                MSHFLEX1.Redraw = True
                Exit Sub
            End If
            
            cartnrzuSpeichern = MSHFLEX1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Bestandsveraenderung cartnrzuSpeichern, lBest, "Terminpreis Bea"
            
            MSHFLEX1.Col = lColmerker
            MSHFLEX1.Row = lRowmerker
            
            gbAender = False
            
        End If
        
        MSHFLEX1.Redraw = True
    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLex1_SelChange"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option2_dblClick(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Command5_Click 1
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option2_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Text2.BackColor = glSelBack1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.BackColor = glSelBack1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    Select Case Index
    
        Case 11
            If KeyCode = vbKeyF2 Then
                frmWKL49.Show 1
            End If
        
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case 0
            Excelimport
            
            Aktuelleanzeigen
            
            If mdeErr Then
                anzeigeNew "normal", "nicht erkannte Artikel werden angezeigt...", Label4
                reportbildschirm "", "aWKL61e" 'Error artikel mde
                anzeigeNew "normal", "", Label4
            End If
            
        Case 1
            Text5_KeyUp 6, vbKeyF2, 0
    
        Case 9
            Text5_KeyUp 5, vbKeyF2, 0

    End Select

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Excelimport()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPfad As String
    Dim cDatname As String
    Dim dbExcel As Database
    Dim rsrs As Recordset
    Dim rsRS2 As Recordset
    Dim rsArt As Recordset
    Dim gsExcel50 As String
    Dim rsFilBu     As Recordset
    
    Dim seekEAN As String
    
   
    gsExcel50 = "Excel 5.0;"
    
     
    
    If pfadseekExcel_Angebot = False Then
        anzeige "rot2", "Abbruch durch Benutzer", Label4
        Exit Sub
    End If
    
    Screen.MousePointer = 11

    anzeige "normal", "", Label4
    cPfad = Label2.Caption
    
    Set dbExcel = OpenDatabase(cPfad, 0, 0, gsExcel50)

    If Not NewTableSuchenDBKombi("ARTT23", gdBase) Then
        CreateTable "ARTT23", gdBase
    End If
    Set rsRS2 = gdBase.OpenRecordset("ARTT23")
    
    loeschNEW "TERMERR", gdBase
    CreateTableT2 "TERMERR", gdBase
    
    Set rsFilBu = gdBase.OpenRecordset("TERMERR")
    
    
    If InStr(UCase(cPfad), "BUDNI") > 0 Then
    
        Dim sLinr As String
        Dim rsLi As DAO.Recordset
        sSQL = "select LINR from LISRT where FORMAT = 'EDIBUDNI'"
        Set rsLi = gdBase.OpenRecordset(sSQL)
        If Not rsLi.EOF Then
            sLinr = Trim(rsLi!linr)
        End If
        rsLi.Close: Set rsLi = Nothing
    
        'Import für Budni-Aktionartikel
        
        Dim sBudArtnr As String
        Dim sArtnr As String
        Dim rsArtlief As DAO.Recordset
        Set rsrs = dbExcel.OpenRecordset("Prospekt$")
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!BUDNINR) Then
                    sBudArtnr = rsrs!BUDNINR
                    
                    sArtnr = "0"
                    
                    sSQL = "select * from artlief where libesnr = '" & sBudArtnr & "'"
                    sSQL = sSQL & " and linr = " & sLinr
                    Set rsArtlief = gdBase.OpenRecordset(sSQL)
                    If Not rsArtlief.EOF Then
                        sArtnr = Trim(rsArtlief!artnr)
                    Else
                    
                        mdeErr = True
                        rsFilBu.AddNew
    
                        rsFilBu!EAN = ""
                        rsFilBu!NAN = Left(sBudArtnr, 13)
                        rsFilBu!Lieferant = sLinr
                        rsFilBu!ARTIKELTEXT = ""
                        rsFilBu!VPE = 0
                        rsFilBu!LVP = rsrs!VK
    
                        rsFilBu.Update
                    
                    End If
                    rsArtlief.Close: Set rsArtlief = Nothing
                    
                    If Val(sArtnr) > 0 Then
                    
                        sSQL = "select * from artikel where artnr = " & sArtnr
                        Set rsArt = gdBase.OpenRecordset(sSQL)
                        If Not rsArt.EOF Then
                            sArtnr = Trim(rsArt!artnr)
                        
                            Set rsRS2 = gdBase.OpenRecordset("select * from ARTT23 where artnr = " & sArtnr)
                            If Not rsRS2.EOF Then
                                
                            Else
                            
                                rsRS2.AddNew
                                rsRS2!artnr = rsArt!artnr
                                rsRS2!BEZEICH = rsArt!BEZEICH
                                rsRS2!AGN = rsArt!AGN
                                rsRS2!PGN = rsArt!PGN
                                rsRS2!vkpr = rsArt!vkpr
                                rsRS2!MWST = rsArt!MWST
                                rsRS2!linr = rsArt!linr
                                rsRS2!LIBESNR = rsArt!LIBESNR
                                rsRS2!EAN = rsArt!EAN
                                rsRS2!RKZ = rsArt!RKZ
                                rsRS2!LPZ = rsArt!LPZ
                                rsRS2!NOTIZEN = rsArt!NOTIZEN
                                rsRS2!BESTAND = rsArt!BESTAND
                                rsRS2!AWM = rsArt!AWM
                                rsRS2!EAN2 = rsArt!EAN2
                                rsRS2!EAN3 = rsArt!EAN3
                                rsRS2!INHALT = rsArt!INHALT
                                rsRS2!INHALTBEZ = rsArt!INHALTBEZ
                                rsRS2!GRUNDPREIS = rsArt!GRUNDPREIS
                                rsRS2!MINBEST = rsArt!MINBEST
                                rsRS2!RABATT_OK = rsArt!RABATT_OK
                                rsRS2!GEFUEHRT = rsArt!GEFUEHRT
                                rsRS2!sEK = rsArt!ekpr
                                rsRS2!KVKA = rsArt!KVKPR1
                                rsRS2!PREISSCHU = rsArt!PREISSCHU
                                
                                rsRS2!KVKN = rsrs!VK
                                
                                rsRS2.Update
                            
                            End If
                            rsRS2.Close: Set rsRS2 = Nothing
                        
                        
                    
                        Else 'hier die unbekannten
                        
                            mdeErr = True
                            rsFilBu.AddNew
    
                            rsFilBu!EAN = ""
                            rsFilBu!NAN = Left(sBudArtnr, 13)
                            rsFilBu!Lieferant = sLinr
                            rsFilBu!ARTIKELTEXT = ""
                            rsFilBu!VPE = 0
                            rsFilBu!LVP = rsrs!VK
    
                            rsFilBu.Update
                            
                        End If
                        rsArt.Close: Set rsArt = Nothing
                    
                    
                    
                    End If
                Else 'hier die unbekannten
                    

                End If
                    
            rsrs.MoveNext
            Loop
            
        End If
        rsrs.Close: Set rsrs = Nothing
        
    ElseIf InStr(UCase(cPfad), "SORTIMENT HZ") > 0 Then 'Rewe Drogerie
    
        lAnzZ = 0
        Set rsrs = dbExcel.OpenRecordset("ANGEBOT$")
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!EAN) Then
                    seekEAN = Trim(Val(rsrs!EAN))
                    seekEAN = checkean(seekEAN)
                    
                    
                    If Len(seekEAN) = 11 Then
                        seekEAN = "0" & seekEAN
                
                        sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                    ElseIf Len(seekEAN) = 8 Then
    '                    If Left(seekEAN, 1) = "2" Then
    '                        seekEAN = Mid$(seekEAN, 2, 6)
    '                        sSQL = "select * from artikel where artnr = " & seekEAN
    '                    Else
                            sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                            sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                            sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
    '                    End If
                    Else
                        sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                    End If
    
                    
                    Set rsArt = gdBase.OpenRecordset(sSQL)
                    If Not rsArt.EOF Then
                        sArtnr = Trim(rsArt!artnr)
                        
                        Set rsRS2 = gdBase.OpenRecordset("select * from ARTT23 where artnr = " & sArtnr)
                        If Not rsRS2.EOF Then
                            
                        Else
         
                            rsRS2.AddNew
                            rsRS2!artnr = rsArt!artnr
                            rsRS2!BEZEICH = rsArt!BEZEICH
                            rsRS2!AGN = rsArt!AGN
                            rsRS2!PGN = rsArt!PGN
                            rsRS2!vkpr = rsArt!vkpr
                            rsRS2!MWST = rsArt!MWST
                            rsRS2!linr = rsArt!linr
                            rsRS2!LIBESNR = rsArt!LIBESNR
                            rsRS2!EAN = rsArt!EAN
                            rsRS2!RKZ = rsArt!RKZ
                            rsRS2!LPZ = rsArt!LPZ
                            rsRS2!NOTIZEN = rsArt!NOTIZEN
                            rsRS2!BESTAND = rsArt!BESTAND
                            rsRS2!AWM = rsArt!AWM
                            rsRS2!EAN2 = rsArt!EAN2
                            rsRS2!EAN3 = rsArt!EAN3
                            rsRS2!INHALT = rsArt!INHALT
                            rsRS2!INHALTBEZ = rsArt!INHALTBEZ
                            rsRS2!GRUNDPREIS = rsArt!GRUNDPREIS
                            rsRS2!MINBEST = rsArt!MINBEST
                            rsRS2!RABATT_OK = rsArt!RABATT_OK
                            rsRS2!GEFUEHRT = rsArt!GEFUEHRT
                            rsRS2!sEK = rsArt!ekpr
                            rsRS2!KVKA = rsArt!KVKPR1
                            rsRS2!PREISSCHU = rsArt!PREISSCHU
                            
                            rsRS2!KVKN = rsrs!VK
                            
                            rsRS2.Update
                        
                        End If
                          
                        rsRS2.Close: Set rsRS2 = Nothing
                    
                    Else 'hier die unbekannten
                    
                        mdeErr = True
                        rsFilBu.AddNew
                        
                        rsFilBu!EAN = Left(seekEAN, 13)
                        rsFilBu!NAN = Left(rsrs!NAN, 13)
'                        rsFilBu!Lieferant = Left(rsrs!Lieferant, 50)
'                        rsFilBu!ARTIKELTEXT = Left(rsrs!ARTIKELTEXT, 50)
'                        rsFilBu!VPE = rsrs!VPE
                        
                        rsFilBu!Lieferant = ""
                        rsFilBu!ARTIKELTEXT = ""
                        rsFilBu!VPE = 0
                        
                        
                        rsFilBu!LVP = rsrs!VK
                        
                        rsFilBu.Update
                        
                    End If
                    rsArt.Close: Set rsArt = Nothing
                Else 'hier die unbekannten
                    
                    mdeErr = True
                    rsFilBu.AddNew
                    
                    rsFilBu!EAN = ""
                    rsFilBu!NAN = Left(rsrs!NAN, 13)
'                    rsFilBu!Lieferant = Left(rsrs!Lieferant, 50)
'                    rsFilBu!ARTIKELTEXT = Left(rsrs!ARTIKELTEXT, 50)
'                    rsFilBu!VPE = rsrs!VPE

                    rsFilBu!Lieferant = ""
                    rsFilBu!ARTIKELTEXT = ""
                    rsFilBu!VPE = 0

                    rsFilBu!LVP = rsrs!VK
                    
                    rsFilBu.Update
                        
                End If
                    
            rsrs.MoveNext
            Loop
            
        End If
        rsrs.Close: Set rsrs = Nothing
    
    
    
    
    
    
    Else
    

        lAnzZ = 0
        Set rsrs = dbExcel.OpenRecordset("Prospekt$")
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!EAN) Then
                    seekEAN = Trim(Val(rsrs!EAN))
                    seekEAN = checkean(seekEAN)
                    
                    
                    If Len(seekEAN) = 11 Then
                        seekEAN = "0" & seekEAN
                
                        sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                    ElseIf Len(seekEAN) = 8 Then
    '                    If Left(seekEAN, 1) = "2" Then
    '                        seekEAN = Mid$(seekEAN, 2, 6)
    '                        sSQL = "select * from artikel where artnr = " & seekEAN
    '                    Else
                            sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                            sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                            sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
    '                    End If
                    Else
                        sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                    End If
    
                    
                    Set rsArt = gdBase.OpenRecordset(sSQL)
                    If Not rsArt.EOF Then
                        sArtnr = Trim(rsArt!artnr)
                        
                        Set rsRS2 = gdBase.OpenRecordset("select * from ARTT23 where artnr = " & sArtnr)
                        If Not rsRS2.EOF Then
                            
                        Else
         
                            rsRS2.AddNew
                            rsRS2!artnr = rsArt!artnr
                            rsRS2!BEZEICH = rsArt!BEZEICH
                            rsRS2!AGN = rsArt!AGN
                            rsRS2!PGN = rsArt!PGN
                            rsRS2!vkpr = rsArt!vkpr
                            rsRS2!MWST = rsArt!MWST
                            rsRS2!linr = rsArt!linr
                            rsRS2!LIBESNR = rsArt!LIBESNR
                            rsRS2!EAN = rsArt!EAN
                            rsRS2!RKZ = rsArt!RKZ
                            rsRS2!LPZ = rsArt!LPZ
                            rsRS2!NOTIZEN = rsArt!NOTIZEN
                            rsRS2!BESTAND = rsArt!BESTAND
                            rsRS2!AWM = rsArt!AWM
                            rsRS2!EAN2 = rsArt!EAN2
                            rsRS2!EAN3 = rsArt!EAN3
                            rsRS2!INHALT = rsArt!INHALT
                            rsRS2!INHALTBEZ = rsArt!INHALTBEZ
                            rsRS2!GRUNDPREIS = rsArt!GRUNDPREIS
                            rsRS2!MINBEST = rsArt!MINBEST
                            rsRS2!RABATT_OK = rsArt!RABATT_OK
                            rsRS2!GEFUEHRT = rsArt!GEFUEHRT
                            rsRS2!sEK = rsArt!ekpr
                            rsRS2!KVKA = rsArt!KVKPR1
                            rsRS2!PREISSCHU = rsArt!PREISSCHU
                            
                            rsRS2!KVKN = rsrs!LVP
                            
                            rsRS2.Update
                        
                        End If
                          
                        rsRS2.Close: Set rsRS2 = Nothing
                    
                    Else 'hier die unbekannten
                    
                        mdeErr = True
                        rsFilBu.AddNew
                        
                        rsFilBu!EAN = Left(seekEAN, 13)
                        rsFilBu!NAN = Left(rsrs!NAN, 13)
                        rsFilBu!Lieferant = Left(rsrs!Lieferant, 50)
                        rsFilBu!ARTIKELTEXT = Left(rsrs!ARTIKELTEXT, 50)
                        rsFilBu!VPE = rsrs!VPE
                        rsFilBu!LVP = rsrs!LVP
                        
                        rsFilBu.Update
                        
                    End If
                    rsArt.Close: Set rsArt = Nothing
                Else 'hier die unbekannten
                    
                    mdeErr = True
                    rsFilBu.AddNew
                    
                    rsFilBu!EAN = ""
                    rsFilBu!NAN = Left(rsrs!NAN, 13)
                    rsFilBu!Lieferant = Left(rsrs!Lieferant, 50)
                    rsFilBu!ARTIKELTEXT = Left(rsrs!ARTIKELTEXT, 50)
                    rsFilBu!VPE = rsrs!VPE
                    rsFilBu!LVP = rsrs!LVP
                    
                    rsFilBu.Update
                        
                End If
                    
            rsrs.MoveNext
            Loop
            
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    rsFilBu.Close: Set rsFilBu = Nothing
    
    Screen.MousePointer = 0

    dbExcel.Close
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3125 Or err.Number = 3011 Then
        anzeige "rot", "Die Excelliste hat nicht das erwartete Format", Label4
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Excelimport"
        Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Private Function pfadseekExcel_Angebot() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    Dim sExcelpfad  As String
    
    pfadseekExcel_Angebot = False

    sTitle = "Speichern des Pfades"
    
    sFilter = "Excel - Dateien (*.xls)|*.xls"
    
    sOldpfad = gcDBPfad & "\IN"
    sExcelpfad = pfadaendernKomplett(sTitle, sFilter, sOldpfad)
    
    If UCase(Right(sExcelpfad, 3)) = "XLS" Then
        pfadseekExcel_Angebot = True
        Label2.Caption = sExcelpfad
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "pfadseekExcel_Angebot"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub PositionierenWKL61()
On Error GoTo LOKAL_ERROR
    
    With Frame1
        .Height = 6615
        .Width = 11535
        .Top = 960
        .Left = 120
        .BorderStyle = 0
    End With
    
    With Frame2
        .Height = 6615
        .Width = 11535
        .Top = 960
        .Left = 120
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame4
        .Height = 6615
        .Width = 11535
        .Top = 960
        .Left = 120
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame5
        .Height = 6615
        .Width = 11535
        .Top = 960
        .Left = 120
        .BorderStyle = 0
        .Visible = False
    End With
    
    MSHFLEX1.Top = 1560
    MSHFLEX1.Left = 120
    MSHFLEX1.Width = 11415
    MSHFLEX1.Height = 3375
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL1"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub zeigePreistermaktion(sAktionNr As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    sSQL = " Select * from Preisterm where preisnr = " & sAktionNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Preisnr) Then
            Label1(15).Caption = Trim(rsrs!Preisnr)
        End If
        
        If Not IsNull(rsrs!preisname) Then
            Label1(9).Caption = Trim(rsrs!preisname)
        End If
        
        If Not IsNull(rsrs!preisbesch) Then
            Label1(8).Caption = Trim(rsrs!preisbesch)
        End If
        
        If Not IsNull(rsrs!Von) Then
            Label1(11).Caption = Trim(rsrs!Von)
        End If
        
        If Not IsNull(rsrs!Bis) Then
            Label1(13).Caption = Trim(rsrs!Bis)
        End If
        
        Label1(26).Caption = CStr(ermanzArtikel(CInt(rsrs!Preisnr)))
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigePreistermaktion"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigePreistermaktiontoBea(sAktionNr As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    sSQL = " Select * from Preisterm where preisnr = " & sAktionNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Preisnr) Then
            
            Label1(16).Caption = Trim(rsrs!Preisnr)
        End If
        
        If Not IsNull(rsrs!preisname) Then
            Text2.Text = Trim(rsrs!preisname)
        End If
        
        If Not IsNull(rsrs!preisbesch) Then
            Text3.Text = Trim(rsrs!preisbesch)
        End If
        
        If Not IsNull(rsrs!Von) Then
            Text1(0).Value = Trim(rsrs!Von)
        End If
        
        If Not IsNull(rsrs!Bis) Then
            Text1(1).Value = Trim(rsrs!Bis)
        End If
        
        Label1(28).Caption = CStr(ermanzArtikel(CInt(rsrs!Preisnr)))
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigePreistermaktiontoBea"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub AuswertungAktion(sAktionNr As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsrs1       As Recordset
    Dim lDiff       As Long
    Dim Ldiffe      As Long
    Dim Ldiffm      As Long
    
    Dim lDatVon     As Long
    Dim lDatBis     As Long
    
    Dim lDatvonZR   As Long
    Dim lDatbisZR   As Long
    
    
    Dim cDatVon     As String
    Dim cDatBis     As String
    
    Dim cDatvonZR   As String
    Dim cDatbisZR   As String
    
    Dim sMW         As String
    Dim lcount      As Long
    
    
    Screen.MousePointer = 11
    loeschNEW "DRUAKTIONA", gdBase
    CreateTable "DRUAKTIONA", gdBase
    
    loeschNEW "DRUAKTIONK", gdBase
    CreateTable "DRUAKTIONK", gdBase
    


    cSQL = "Select * from PREISTERM where "
    cSQL = cSQL & " preisnr = " & sAktionNr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!Bis) Then
            lDatBis = CLng(rsrs!Bis)
        End If
        
        If Not IsNull(rsrs!Von) Then
            lDatVon = CLng(rsrs!Von)
        End If
        
        lDiff = lDatBis - lDatVon
        Ldiffe = CInt(lDiff / 7)
        Ldiffm = lDiff Mod 7
        If Ldiffm > 0 Then
            Ldiffm = Ldiffe + 1
        End If
        
        lDiff = Ldiffm * 7
        
        lDatvonZR = lDatVon - lDiff
        lDatbisZR = lDatBis - lDiff
        
        cDatVon = Trim$(Str$(lDatVon))
        cDatBis = Trim$(Str$(lDatBis))
        
        cDatvonZR = Trim$(Str$(lDatvonZR))
        cDatbisZR = Trim$(Str$(lDatbisZR))
   
    End If
    rsrs.Close: Set rsrs = Nothing
    

    Set rsrs1 = gdBase.OpenRecordset("DRUAKTIONA", dbOpenTable)
    
    cSQL = "Select * from PRSTERM where "
    cSQL = cSQL & " preisnr = " & sAktionNr
    cSQL = cSQL & " and Status > 0 "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
    
        rsrs1.AddNew
        rsrs1!artnr = rsrs!artnr
        anzeige "normal", "in " & lcount & " sec fertig", Label9
        lcount = lcount - 1
        rsrs1!KVKPR1ALT = rsrs!KVKPR1ALT
        rsrs1!KVKPR1NEU = rsrs!KVKPR1NEU
        rsrs1!DAT_VON = rsrs!DAT_VON
        rsrs1!DAT_BIS = rsrs!DAT_BIS
        
        rsrs1!DAT_VONZR = lDatvonZR
        rsrs1!Dat_bisZR = lDatbisZR
        rsrs1!FILIALE = rsrs!FILIALE
        rsrs1!Preisnr = sAktionNr
        
        sMW = ermMWST(rsrs!artnr)
        
        rsrs1!ekpr = ermavgEK(rsrs!artnr, cDatVon, cDatBis, rsrs1!FILIALE)
        rsrs1!Menge = ermMENGE(rsrs!artnr, cDatVon, cDatBis, rsrs1!FILIALE)
        rsrs1!Preis = ermUMSATZ(rsrs!artnr, cDatVon, cDatBis, rsrs1!FILIALE)
        rsrs1!ERTRAG = ermERTRAG(rsrs!artnr, cDatVon, cDatBis, rsrs1!FILIALE, sMW)
        
        
        rsrs1!ekprZR = ermavgEK(rsrs!artnr, cDatvonZR, cDatbisZR, rsrs1!FILIALE)
        rsrs1!MengeZR = ermMENGE(rsrs!artnr, cDatvonZR, cDatbisZR, rsrs1!FILIALE)
        rsrs1!PreisZR = ermUMSATZ(rsrs!artnr, cDatvonZR, cDatbisZR, rsrs1!FILIALE)
        rsrs1!ERTRAGZR = ermERTRAG(rsrs!artnr, cDatvonZR, cDatbisZR, rsrs1!FILIALE, sMW)
        
        rsrs1.Update
       
        
        rsrs.MoveNext
        Loop
        
    End If
    rsrs1.Close: Set rsrs1 = Nothing
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Update DRUAKTIONA inner join ARTIKEL on DRUAKTIONA.artnr = artikel.artnr "
    sSQL = sSQL & " SET DRUAKTIONA.Bezeich = Artikel.Bezeich "
    sSQL = sSQL & " , DRUAKTIONA.BESTAND = Artikel.BESTAND "
    sSQL = sSQL & " , DRUAKTIONA.MWST = Artikel.MWST "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DRUAKTIONA "
    sSQL = sSQL & " set "
    sSQL = sSQL & " NSP = ((((Preis/(100 + " & gdMWStV & "))* 100) - (EKPR * Menge))* 100) / ((Preis/(100 + " & gdMWStV & "))* 100)"
    sSQL = sSQL & " where MWST = 'V' "
    sSQL = sSQL & " and PREIS <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DRUAKTIONA "
    sSQL = sSQL & " set "
    sSQL = sSQL & " NSP = ((((Preis/(100 + " & gdMWStE & "))* 100) - (EKPR * Menge))* 100) / ((Preis/(100 + " & gdMWStE & "))* 100)"
    sSQL = sSQL & " where MWST = 'E' "
    sSQL = sSQL & " and Preis <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DRUAKTIONA "
    sSQL = sSQL & " set "
    sSQL = sSQL & " NSP = ((((Preis/(100 + " & gdMWStO & "))* 100) - (EKPR * Menge))* 100) / ((Preis/(100 + " & gdMWStO & "))* 100)"
    sSQL = sSQL & " where MWST = 'O' "
    sSQL = sSQL & " and Preis <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update DRUAKTIONA "
    sSQL = sSQL & " set "
    sSQL = sSQL & " NSPZR = ((((Preiszr/(100 + " & gdMWStV & "))* 100) - (EKPRzr * Mengezr))* 100) / ((Preiszr/(100 + " & gdMWStV & "))* 100)"
    sSQL = sSQL & " where MWST = 'V' "
    sSQL = sSQL & " and PREISzr <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DRUAKTIONA "
    sSQL = sSQL & " set "
    sSQL = sSQL & " NSPZR = ((((Preiszr/(100 + " & gdMWStE & "))* 100) - (EKPRzr * Mengezr))* 100) / ((Preiszr/(100 + " & gdMWStE & "))* 100)"
    sSQL = sSQL & " where MWST = 'E' "
    sSQL = sSQL & " and Preiszr <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DRUAKTIONA "
    sSQL = sSQL & " set "
    sSQL = sSQL & " NSPZR = ((((Preiszr/(100 + " & gdMWStO & "))* 100) - (EKPRzr * Mengezr))* 100) / ((Preiszr/(100 + " & gdMWStO & "))* 100)"
    sSQL = sSQL & " where MWST = 'O' "
    sSQL = sSQL & " and Preiszr <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into DRUAKTIONK Select * from PREISTERM where "
    sSQL = sSQL & " preisnr = " & sAktionNr
    gdBase.Execute sSQL, dbFailOnError
    
    anzeigeNew "normal", "Druckvorschau wird erstellt...", Label9
    
    Screen.MousePointer = 0
    
    reportbildschirm "", "awkl61b"
    
    anzeigeNew "normal", "", Label9
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AuswertungAktion"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List3_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    
    cLBSatz = List3.SelectedItem.Text
    cLBSatz = Left(cLBSatz, 6)
    cLBSatz = Trim$(cLBSatz)
    
    Label1(15).Caption = ""
    Label1(8).Caption = ""
    Label1(9).Caption = ""
    Label1(11).Caption = ""
    Label1(13).Caption = ""
    Label1(26).Caption = ""
    
    Label1(15).Refresh
    Label1(8).Refresh
    Label1(9).Refresh
    Label1(11).Refresh
    Label1(13).Refresh
    
    zeigePreistermaktion cLBSatz
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List3_Click"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text5_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If Index = 3 Then
        If KeyCode = vbKeyReturn Then
            Command5_Click 12
        End If
    End If
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            Case Is = 5
                gF2Prompt.cFeld = "AGN"
                
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    If gF2Prompt.cWahl <> "" Then
                        Text5(Index).Text = gF2Prompt.cWahl
                    End If
                End If
            Case Is = 6
                gF2Prompt.cFeld = "LINR"
                
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    If gF2Prompt.cWahl <> "" Then
                        Text5(Index).Text = gF2Prompt.cWahl
                    End If
                End If
        End Select
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    Select Case Index
        Case 0, 1, 2, 4
            cValid = "1234567890,-" & Chr$(8)
        Case 3
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß"
        Case 5, 6
            cValid = "1234567890" & Chr$(8)
    End Select
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    
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
    Fehler.gsFunktion = "Text5_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text5_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text5(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text5_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text5(Index).SelStart = 0
    Text5(Index).SelLength = Len(Text5(Index).Text)
    Text5(Index).BackColor = glSelBack1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
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
