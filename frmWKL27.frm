VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWKL27 
   BackColor       =   &H00C0C000&
   Caption         =   "Einlesen der Zentraldaten"
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
   Icon            =   "frmWKL27.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Height          =   7815
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   11775
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   1920
         TabIndex        =   26
         Top             =   7080
         Width           =   1095
      End
      Begin sevCommand3.Command Command6 
         Height          =   165
         Index           =   1
         Left            =   3120
         TabIndex        =   25
         Top             =   7080
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
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
         ToolTip         =   "Vor"
         ToolTipTitle    =   "Vor"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   165
         Index           =   2
         Left            =   3120
         TabIndex        =   24
         Top             =   7320
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
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
         ToolTip         =   "Zurück"
         ToolTipTitle    =   "Zurück"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   0
         Left            =   8160
         TabIndex        =   23
         Top             =   7080
         Width           =   1695
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
         Caption         =   "Protokoll"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.PictureBox picprogress 
         Height          =   375
         Left            =   6960
         ScaleHeight     =   315
         ScaleWidth      =   4635
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtstatus 
         Height          =   315
         Left            =   4440
         TabIndex        =   20
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin sevCommand3.Command Command8 
         Height          =   375
         Left            =   4200
         TabIndex        =   19
         Top             =   7080
         Width           =   1695
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
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   3
         Left            =   5160
         TabIndex        =   16
         Top             =   1560
         Width           =   1695
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
         Caption         =   "Einlesen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   1080
         TabIndex        =   15
         Top             =   1560
         Width           =   3975
      End
      Begin VB.FileListBox File2 
         Height          =   300
         Left            =   5760
         Pattern         =   "Y*.lzh"
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame44 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H0000FFFF&
         Height          =   2175
         Left            =   1080
         TabIndex        =   7
         Top             =   4080
         Visible         =   0   'False
         Width           =   10575
         Begin VB.Label Label2 
            Alignment       =   2  'Zentriert
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   375
            Index           =   5
            Left            =   4200
            TabIndex        =   13
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Zentriert
            BackStyle       =   0  'Transparent
            Caption         =   "von"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   375
            Index           =   4
            Left            =   3360
            TabIndex        =   12
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Zentriert
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   11
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Sätze:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Warte"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   9
            Top             =   480
            Width           =   8895
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Aktion:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.TextBox txtKinPfad 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   10575
      End
      Begin sevCommand3.Command cmdStandardUp 
         Height          =   375
         Left            =   9960
         TabIndex        =   4
         Top             =   240
         Width           =   1695
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
         Caption         =   "Standard"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdUpdate 
         Height          =   375
         Left            =   8160
         TabIndex        =   3
         Top             =   240
         Width           =   1695
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
         Caption         =   "Ändern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   7
         Left            =   9960
         TabIndex        =   2
         Top             =   7080
         Width           =   1695
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
      Begin MSComctlLib.ProgressBar pbrAbschluss 
         Height          =   375
         Left            =   1080
         TabIndex        =   18
         Top             =   6360
         Visible         =   0   'False
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
      End
      Begin sevCommand3.Command Command6 
         Height          =   405
         Index           =   20
         Left            =   3600
         TabIndex        =   28
         ToolTipText     =   "Kalender"
         Top             =   7080
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
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
      Begin sevCommand3.Command Command1 
         Height          =   375
         Left            =   6000
         TabIndex        =   29
         Top             =   7080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
         Caption         =   "in den Etikettenpool"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "ab:"
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
         Left            =   600
         TabIndex        =   27
         Top             =   7200
         Width           =   1215
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   "Datei wird entpackt..."
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
         Index           =   0
         Left            =   6960
         TabIndex        =   22
         Top             =   1320
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   "Pfad zu den Kassendateien"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   7095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   5400
         TabIndex        =   6
         Top             =   2160
         Width           =   6135
      End
   End
   Begin VB.Label lblUeberschrift 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Kassendateien aus der Zentrale einlesen"
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
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   11895
   End
End
Attribute VB_Name = "frmWKL27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub PositionierenWKL27()
    On Error GoTo LOKAL_ERROR
    
    Frame6.Top = 840
    Frame6.Left = 0
    Frame6.Height = 7695
    Frame6.Width = 12000
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL27"
    Fehler.gsFehlertext = "Im Programmteil Kassendateien einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click()
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    If IsDate(Text1(0).Text) = False Then


    Else
        If NewTableSuchenDBKombi("ETIPROTS", gdBase) Then
            
            cSQL = "insert into etidru select artnr, bezeich, vkprneu as vkpr "
            cSQL = cSQL & ",bestand,anzahl, libesnr, lpz,ean,linr,filnr, '' as pcname  from etiprots "
            cSQL = cSQL & " where WEDATE = " & CLng(DateValue(Text1(0).Text))
            cSQL = cSQL & " and Bestand > 0 "
            gdBase.Execute cSQL, dbFailOnError
            
        Else
            MsgBox "Es sind keine Daten vorhanden.", vbInformation, "Winkiss Hinweis:"
        End If
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassendateien einlesen ist ein Fehler aufgetreten."
    
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
Private Sub cmdStandardUp_Click()
     On Error GoTo LOKAL_ERROR
   
    txtKinPfad.Text = gcDBPfad & "\In"
    gsKinPfad = gcDBPfad & "\In"
    
    speicherpfad
    Dateienladen

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassendateien einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdUpdate_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    
    sTitle = "Speichern des Kassendateipfades"
    sFilter = "LZH - Dateien (*.lzh)| Y*.lzh| "
    sOldpfad = txtKinPfad.Text
    gsKinPfad = pfadaendern(sTitle, sFilter, sOldpfad)
    
    txtKinPfad.Text = gsKinPfad
    speicherpfad
    Dateienladen
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdUpdate_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassendateien einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command6_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Dim lDat As Long
    
    Select Case Index
        Case 0
            Dim cPfad As String
            cPfad = gcDBPfad
            If Right(cPfad, 1) <> "\" Then
                cPfad = cPfad & "\"
            End If
            
            zeigeHilfe "LPROTOK", "YProtokoll.txt", cPfad
        Case 1
        
            If IsDate(Text1(0).Text) = False Then
                Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
            
                If IsDate(Text1(0).Text) = True Then
                    lDat = CLng(DateValue(Text1(0).Text))
                End If
                
                lDat = lDat + 1
                Text1(0).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 2
        
            If IsDate(Text1(0).Text) = False Then
                Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(0).Text) = True Then
                    lDat = CLng(DateValue(Text1(0).Text))
                End If
                
                lDat = lDat - 1
                Text1(0).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case Is = 3     'einlesen
            If List9.ListCount = 0 Then
                'Datenholen
            Else
                einlesen
            End If
            
        Case Is = 7
            Unload frmWKL27
        Case 20
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassendateien einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command8_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    
    If IsDate(Text1(0).Text) = False Then

        If NewTableSuchenDBKombi("ETIPROT", gdBase) Then
            If gbPAEBON = False Then
            
                
                
                
                Vorarbeit_27d
                reportbildschirm "dWKL", "aWKL27d"
                
            Else
                DruckenPAEBON
            End If
        Else
            MsgBox "Es sind keine Druckdaten vorhanden.", vbInformation, "Winkiss Hinweis:"
        End If
    Else
        If NewTableSuchenDBKombi("ETIPROTS", gdBase) Then
            
            loeschNEW "ETIPROTP", gdBase
            CreateTable "ETIPROTP", gdBase
            
            cSQL = "Insert into ETIPROTP Select * from ETIPROTS "
            cSQL = cSQL & " where WEDATE = " & CLng(DateValue(Text1(0).Text))
            gdBase.Execute cSQL, dbFailOnError
            
            reportbildschirm "dWKL", "aWKL27e"
            
        Else
            MsgBox "Es sind keine Druckdaten vorhanden.", vbInformation, "Winkiss Hinweis:"
        End If
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil Kassendateien einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckenPAEBON()
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    Dim cDatum      As String
    Dim czeit       As String
    Dim cArtNr      As String
    Dim cBezeich    As String
    Dim cNPreis     As String
    Dim rsrs        As Recordset
    Dim iAnzSätze   As Integer
    
    Dim i           As Integer
    Dim lcount As Long
    
    lcount = 0
    
    cDatum = DateValue(Now)
    czeit = TimeValue(Now)
    
    Set rsrs = gdBase.OpenRecordset("ETIPROT")
    If Not rsrs.EOF Then
        rsrs.MoveLast
        iAnzSätze = rsrs.RecordCount
        ReDim cZeilen(0 To (iAnzSätze * 3) + 5) As String
        
        cZeilen(0) = "Preisänderungen"
        cZeilen(1) = "-----------------"
        cZeilen(2) = "insgesamt: " & iAnzSätze
        cZeilen(3) = "Datum: " & cDatum
        cZeilen(4) = "Zeit:  " & czeit
        cZeilen(5) = vbCrLf
        
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!VKPRNEU) Then
                cNPreis = Format(rsrs!VKPRNEU, "######.00")
            End If
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            End If
        
            cZeilen(6 + lcount) = "Artnr: " & cArtNr & Space(12 - Len(cNPreis)) & cNPreis & " " & gcWaehrung
            cZeilen(6 + lcount + 1) = cBezeich
            cZeilen(6 + lcount + 2) = ""
            
            lcount = lcount + 3
            
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    'Drucke den Beleg
    
    DruckeArbeitszeitBelegWK20d cZeilen(), (iAnzSätze * 3) + 5
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckenPAEBON"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Vorarbeit_27d()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    
    Screen.MousePointer = 11
    
    cSQL = "Update ETIPROT inner join Artikel on ETIPROT.artnr = artikel.artnr "
    cSQL = cSQL & " set ETIPROT.NOTIZEN = artikel.NOTIZEN , ETIPROT.LVK = artikel.VKPR, ETIPROT.RABATT_OK = artikel.RABATT_OK "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = " Select * from ETIPROT "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                rsrs.Edit
                rsrs!lastzu = ErmlzZugang(rsrs!artnr)
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Vorarbeit_27d"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Activate()
    On Error GoTo LOKAL_ERROR
    
    Dim rsEtiprot As Recordset

    If gbfrm27 Then
        frmWKL27.Refresh
        
        If gbNacht = True Then
            schreibeProtokollNachtAblauf "Einlesen der Kassendateien beginnt"
        End If
        
        einlesen
        
        If gbNacht = True Then
            schreibeProtokollNachtAblauf "Einlesen der Kassendateien endet"
        End If
    
        If NewTableSuchenDBKombi("ETIPROT", gdBase) Then
            If Datendrin("Etiprot", gdBase) Then
                Command8.BackColor = vbRed
            End If
        End If
        
            Command8.BackColor = Command6(0).BackColor
            If NewTableSuchenDBKombi("Etiprot", gdBase) Then
                Set rsEtiprot = gdBase.OpenRecordset("Etiprot", dbOpenTable)
                If Not rsEtiprot.EOF Then
                    Command8.BackColor = vbRed
                
                    If gbFTPautomatic = True Then
                    
                        If gbDruck27 Then
                            If gbPAEBON Then
                                DruckenPAEBON
                            Else
                                'wieder aufgenommen für bensch am 19.08.2010
                                
                                Vorarbeit_27d
                                reportbildschirmToPrinter "aWKL27d"
                            End If
                            
                        End If
                        Label5.Caption = "Es liegen Preisänderungen vor."
                        Label5.Refresh
                    
                        Unload frmWKL27
                    Else
                        If gbPAEBON Then
                            DruckenPAEBON
                        Else
                            Vorarbeit_27d
                            reportbildschirm "dWKL", "aWKL27d"
                        End If
                        
                        Label5.Caption = "Es liegen Preisänderungen vor."
                        Label5.Refresh
                    End If
                
                Else
                    Unload frmWKL27
                End If
                rsEtiprot.Close
            Else
                Unload frmWKL27
            End If
        End If
    
    gbfrm27 = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Activate"
    Fehler.gsFehlertext = "Im Programmteil Kassendateien einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    gbPreisAender = False
    PositionierenWKL27
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    Frame6.Visible = True
    
    txtKinPfad.Text = gsKinPfad
    Dateienladen
    
    Command8.BackColor = Command6(0).BackColor
    
    If NewTableSuchenDBKombi("ETIPROT", gdBase) Then
        If Datendrin("Etiprot", gdBase) Then
            Command8.BackColor = vbRed
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kassendateien einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Dateienladen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim iRet As Integer
    Dim cPfad As String
    Dim lcount As Long
    Dim lRet As Long
    Dim lfail As Long
    Dim cQuelle As String
    Dim cZiel As String
    Dim cdatei As String
    Dim ctmp As String
    Dim cDatum As String

    Screen.MousePointer = 11
 
    With File2
        .Path = gsKinPfad
        .Refresh
    End With
    List9.Clear
    For lcount = 0 To File2.ListCount - 1
        cdatei = File2.list(lcount)
        cdatei = UCase$(cdatei)
        cPfad = gsKinPfad
        If Right(cPfad, 1) <> "\" Then
            cPfad = cPfad & "\"
        End If
        ctmp = cPfad & cdatei
        cDatum = FileDateTime(ctmp)
        cdatei = cdatei & Space$(12 - Len(cdatei)) & " " & cDatum
        List9.AddItem cdatei
    Next lcount
    
    If List9.ListCount = 0 Then
        'Dateien holen
        
        Dim bmerke As Boolean
        bmerke = gbFTPautomatic
        
        gbFTPautomatic = True
        
        giKissFtpMode = 10 'FTPMODE= 10 , Kombimodus Kassendateien holen und schicken
        frmWKL38.Show 1    ' Programmupdates,Stammdaten holen
        
        gbFTPautomatic = bmerke
        
        With File2
            .Path = gsKinPfad
            .Refresh
        End With
        List9.Clear
        For lcount = 0 To File2.ListCount - 1
            cdatei = File2.list(lcount)
            cdatei = UCase$(cdatei)
            cPfad = gsKinPfad
            If Right(cPfad, 1) <> "\" Then
                cPfad = cPfad & "\"
            End If
            ctmp = cPfad & cdatei
            cDatum = FileDateTime(ctmp)
            cdatei = cdatei & Space$(12 - Len(cdatei)) & " " & cDatum
            List9.AddItem cdatei
        Next lcount
        
    End If
    

    If List9.ListCount = 0 Then
        
        Screen.MousePointer = 0
        Label5.Font.Bold = True
        Label5.ForeColor = vbRed
        Label5.Caption = "Es ist keine Datei vorhanden."
        Label5.Refresh
    
    Else
        Screen.MousePointer = 0
'        einlesen
    End If
        
    Exit Sub
LOKAL_ERROR:
    If err.Number = 68 Then
        List9.Clear
        Label5.Font.Bold = True
        Label5.ForeColor = vbRed
        Label5.Caption = "Das Öffnen von der Diskette ist gescheitert."
        Label5.Refresh
        Screen.MousePointer = 0
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Dateienladen"
        Fehler.gsFehlertext = "Im Programmteil Kassendateien einlesen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Sub
Private Sub einlesen()
    On Error GoTo LOKAL_ERROR
    
    Dim cQuelle As String
    Dim iRet As Integer
    Dim cPfad As String
    Dim cPfad1 As String
    Dim t   As Integer
    Dim Task$, hProcess&, Result&
    Dim i As Integer
    Dim j As Integer
    Dim stabelle As String
    Dim lcount As Long
    Dim sdatname As String
    
    Dim cdatei As String
    Dim ctmp As String
    Dim cDatum As String
    Dim iMaxi As Integer
    Dim byMax   As Byte
    Dim cPfad23 As String
    
    Screen.MousePointer = 11
    
    cPfad = gsKinPfad      'Kassendateieneingangspfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad1 = gsKinPfad      'Kassendateieneingangspfad
    If Right$(cPfad1, 1) = "\" Then
        cPfad1 = Left(cPfad1, Len(cPfad1) - 1)
    End If

    File2.Path = cPfad1
    File2.Pattern = "Y*.lzh"
    File2.Refresh
    
    Label5.Font.Bold = True
    Label5.ForeColor = vbBlue
    Label5.Caption = "bitte warten..."
    Label5.Refresh
    
    If File2.ListCount > 0 Then
        cDatum = DateValue(Now) & " " & TimeValue(Now)
        byMax = File2.ListCount - 1
        
        i = 0
        
        For j = 0 To byMax
step:
            iMaxi = lfnrErmitteln("Y")
            iMaxi = iMaxi + 1
            
            stabelle = File2.list(i)
            cQuelle = cPfad & stabelle
            sdatname = Left(stabelle, 8)
            gcDateidatum = FileDateTime(cQuelle)
            

                
            If Val(CLng(Right(sdatname, 5))) >= iMaxi Then
            
                Label5.Font.Bold = True
                Label5.ForeColor = vbBlue
                Label5.Caption = stabelle & " wird eingelesen..."
                Label5.Refresh
                
                
                Kill cPfad & "ZF.mdb"
                
                Dim cPfad2 As String
                cPfad2 = ShortPath(cPfad)
                

                If Not FileExists(cPfad & "ZF.mdb") Then
                
                    lbl6(0).Visible = True
                    lbl6(0).Caption = stabelle & " wird entpackt..."
                    lbl6(0).Refresh
                    
                    picprogress.Visible = True
                    ShowProgress picprogress, 0, 0, 0
                    
                    Zip_Unzip "XYC6T349G6", cPfad1, cPfad & stabelle, txtStatus
                    
                    lbl6(0).Caption = "erfolgreich"
                    lbl6(0).Refresh
                    
                    Pause 1
                    picprogress.Visible = False
                    lbl6(0).Visible = False
                    lbl6(0).Caption = ""
                    lbl6(0).Refresh
                    
                End If
                
                
    
                '*************************************************************
                '* Nach Filial-Kassendateien aus dem Programm ZENTRALE suchen!
                '*************************************************************
                    
                frmWKL27.Frame44.Visible = True
                
                If Modul6.FindFile(cPfad1, "ZF.mdb") Then
                    schreibeProtokollKassStop "Kassenabschluss"
    
                    If SucheFilialKassenDateienMOD6(picprogress, txtStatus, cPfad, 6, pbrAbschluss, frmWKL27) Then
                        lfnrSchreiben Val(CLng(Right(sdatname, 5))), sdatname, cDatum
                        schreibeYProtokoll "Datei: " & stabelle & " wurde eingelesen."
                    End If
                    
                    cPfad23 = gcDBPfad               'Datenbankpfad
                    If Right(cPfad23, 1) <> "\" Then
                        cPfad23 = cPfad23 & "\"
                    End If
                    Kill cPfad23 & "KASSSTOP.TXT"
                Else
                    schreibeProtokoll "Verarbeite Kassendatei: " & cPfad & "ZF.mdb nicht gefunden."
                End If
                
                Kill cPfad & stabelle
                i = i + 1

            Else
                Kill cPfad & stabelle
                i = i + 1
            End If
        Next j
        
        File2.Path = cPfad1
        File2.Pattern = "Y*.lzh"
        File2.Refresh

        List9.Clear
        For lcount = 0 To File2.ListCount - 1
            cdatei = File2.list(lcount)
            cdatei = UCase$(cdatei)
            ctmp = cPfad & cdatei
            cDatum = FileDateTime(ctmp)
            cdatei = cdatei & Space$(12 - Len(cdatei)) & " " & cDatum
            List9.AddItem cdatei
        Next lcount
        
        If byMax = File2.ListCount - 1 Then
        
            Label5.Font.Bold = True
            Label5.ForeColor = vbRed
            Label5.Caption = "Hotline anfrufen - " & iMaxi & " wird erwartet"
            Label5.Refresh
        ElseIf File2.ListCount = 0 Then
        
            Frame44.Visible = False
            Label5.Font.Bold = True
            Label5.ForeColor = vbBlue
            
            If gbPreisAender Then
                Label5.Caption = "Es liegen Preisänderungen vor."
                Label5.Refresh
            Else
                Label5.Caption = "Einlesen der Kassendateien beendet!"
                Label5.Refresh
            End If
            
        End If
    Else
        Label5.Font.Bold = True
        Label5.ForeColor = vbRed
        Label5.Caption = "Keine Kassendateien vorhanden."
        Label5.Refresh
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Screen.MousePointer = 0
        
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "einlesen"
        Fehler.gsFehlertext = "Im Programmteil Kassendateien einlesen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Command5_Click()
    On Error GoTo LOKAL_ERROR
    
    giKassenDatei = vbNo
    Unload frmWKL27
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kassendateien einlesen ist ein Fehler aufgetreten."
    
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
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kassendateien einlesen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
