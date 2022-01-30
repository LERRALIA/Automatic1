VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWKL38 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "KISSNET..."
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   682
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.FileListBox File1 
      Height          =   300
      Left            =   120
      TabIndex        =   53
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar pbrFile 
      Height          =   300
      Left            =   3000
      TabIndex        =   51
      Top             =   3720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      TabIndex        =   47
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
      Begin sevCommand3.Command Command6 
         Height          =   495
         Left            =   1680
         TabIndex        =   50
         Top             =   2280
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
         Caption         =   "Trennen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   495
         Left            =   120
         TabIndex        =   49
         Top             =   1320
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
         Caption         =   "Verbinden"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      TabIndex        =   40
      Top             =   2040
      Visible         =   0   'False
      Width           =   7335
      Begin sevCommand3.Command Command9 
         Height          =   255
         Index           =   2
         Left            =   5520
         TabIndex        =   46
         Top             =   900
         Visible         =   0   'False
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
         Caption         =   "Einlesen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command9 
         Height          =   255
         Index           =   1
         Left            =   5520
         TabIndex        =   45
         Top             =   540
         Visible         =   0   'False
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
         Caption         =   "Einlesen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command9 
         Height          =   255
         Index           =   0
         Left            =   5520
         TabIndex        =   44
         Top             =   180
         Visible         =   0   'False
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
         Caption         =   "Einlesen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Check2"
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
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   840
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Check2"
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
         Index           =   1
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Check2"
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
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   5175
      End
   End
   Begin sevCommand3.Command cmdStart 
      Height          =   735
      Left            =   120
      TabIndex        =   33
      Top             =   3720
      Width           =   2415
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
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
      Caption         =   "Start"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command8 
      Height          =   735
      Left            =   7680
      TabIndex        =   32
      Top             =   3720
      Width           =   2415
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
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
      Caption         =   "Abbrechen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.Frame fraLocal 
         Caption         =   " Local System "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   0
         TabIndex        =   18
         Top             =   120
         Width           =   4695
         Begin VB.PictureBox picBack 
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   120
            ScaleHeight     =   269
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   229
            TabIndex        =   29
            Top             =   600
            Width           =   3495
            Begin MSComctlLib.ListView lvLocal 
               Height          =   3915
               Left            =   0
               TabIndex        =   30
               Top             =   0
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   6906
               View            =   3
               LabelEdit       =   1
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               _Version        =   393217
               Icons           =   "ilList"
               SmallIcons      =   "ilList"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Comic Sans MS"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Key             =   "Name"
                  Text            =   "Name"
                  Object.Width           =   4022
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Key             =   "Size"
                  Text            =   "Size"
                  Object.Width           =   1587
               EndProperty
            End
         End
         Begin VB.TextBox txtCurPath 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   3495
         End
         Begin sevCommand3.Command cmdlMkDir 
            Height          =   375
            Left            =   3720
            TabIndex        =   27
            Top             =   600
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
            Caption         =   "MkDir"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdlExec 
            Height          =   375
            Left            =   3720
            TabIndex        =   26
            Top             =   960
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
            Caption         =   "Exec"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdlRename 
            Height          =   375
            Left            =   3720
            TabIndex        =   25
            Top             =   1320
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
            Caption         =   "Rename"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdlDelete 
            Height          =   375
            Left            =   3720
            TabIndex        =   24
            Top             =   1680
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
            Caption         =   "Delete"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdlRefresh 
            Height          =   375
            Left            =   3720
            TabIndex        =   23
            Top             =   2040
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
            Caption         =   "Refresh"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdRetrieve 
            Height          =   495
            Left            =   3900
            TabIndex        =   22
            Top             =   2640
            Width           =   495
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
            Caption         =   "<--"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdPut 
            Height          =   495
            Left            =   3900
            TabIndex        =   21
            Top             =   3240
            Width           =   495
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
            Caption         =   "-->"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdPutNew 
            Height          =   615
            Left            =   3900
            TabIndex        =   20
            Top             =   3840
            Width           =   495
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
            Caption         =   "-->"
            PictureAlign    =   2
            Version3        =   -1  'True
         End
         Begin VB.TextBox txtPattern 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3720
            TabIndex        =   19
            Text            =   "*.*"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblLocFiles 
            Alignment       =   1  'Rechts
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   31
            Top             =   4470
            Width           =   855
         End
      End
      Begin VB.Frame fraRemote 
         Caption         =   " Remote System "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   4920
         TabIndex        =   9
         Top             =   0
         Width           =   5295
         Begin sevCommand3.Command cmdrMkDir 
            Height          =   375
            Left            =   4320
            TabIndex        =   15
            Top             =   600
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
            Caption         =   "MkDir"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdrExec 
            Height          =   375
            Left            =   4320
            TabIndex        =   14
            Top             =   960
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
            Caption         =   "View"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdrRename 
            Height          =   375
            Left            =   4320
            TabIndex        =   13
            Top             =   1320
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
            Caption         =   "Rename"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdrDelete 
            Height          =   375
            Left            =   4320
            TabIndex        =   12
            Top             =   1680
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
            Caption         =   "Delete"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdrRefresh 
            Height          =   375
            Left            =   4320
            TabIndex        =   11
            Top             =   2040
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
            Caption         =   "Refresh"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.TextBox txtRemPath 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   5055
         End
         Begin MSComctlLib.ListView lvRemote 
            Height          =   4095
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   7223
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   393217
            Icons           =   "ilList"
            SmallIcons      =   "ilList"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "Name"
               Text            =   "Name"
               Object.Width           =   3987
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "Size"
               Text            =   "Size"
               Object.Width           =   2646
            EndProperty
         End
         Begin VB.Label lblNumFiles 
            Alignment       =   1  'Rechts
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   17
            Top             =   4440
            Width           =   855
         End
      End
      Begin VB.OptionButton chkASCII 
         Caption         =   "ASCII"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   4920
         Width           =   855
      End
      Begin VB.OptionButton chkBinary 
         Caption         =   "Binary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   7
         Top             =   4920
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   5160
         Width           =   10095
         Begin VB.TextBox txtStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertikal
            TabIndex        =   6
            Top             =   180
            Width           =   9915
         End
      End
      Begin sevCommand3.Command cmdConnect 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   6240
         Width           =   1575
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
         Caption         =   "Connect"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdCancel 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   6240
         Width           =   1575
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
         Caption         =   "Cancel"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdExit 
         Height          =   375
         Left            =   8640
         TabIndex        =   2
         Top             =   6240
         Width           =   1575
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
         Caption         =   "Exit"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox lstTemp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   6240
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ImageList ilList 
         Left            =   3480
         Top             =   6120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   13
         ImageHeight     =   13
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWKL38.frx":0000
               Key             =   "Directory"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWKL38.frx":00FA
               Key             =   "Up"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWKL38.frx":01F4
               Key             =   "File"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWKL38.frx":02EE
               Key             =   "Drive"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ProgressBar pbrFiles 
      Height          =   300
      Left            =   3000
      TabIndex        =   52
      Top             =   4080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
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
      Left            =   480
      TabIndex        =   54
      Top             =   1680
      Width           =   8895
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   480
      TabIndex        =   34
      Top             =   120
      Width           =   8895
   End
   Begin VB.Label lblTimeLeft 
      BackStyle       =   0  'Transparent
      Caption         =   "verbleibende Zeit: 00:00:00 h  0 Kbps"
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
      Left            =   2820
      TabIndex        =   39
      Top             =   4425
      Width           =   4275
   End
   Begin VB.Label lblFilesProg 
      Alignment       =   1  'Rechts
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6960
      TabIndex        =   38
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblFileProg 
      Alignment       =   1  'Rechts
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6960
      TabIndex        =   37
      Top             =   3735
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   150
      Left            =   2760
      Picture         =   "frmWKL38.frx":03E8
      Top             =   3720
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   195
      Left            =   2760
      Picture         =   "frmWKL38.frx":0542
      Top             =   4080
      Width           =   225
   End
   Begin VB.Label lblFileCount 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "(0/0)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5880
      TabIndex        =   36
      Top             =   3375
      Width           =   1065
   End
   Begin VB.Label lblCurrentFile 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   240
      Left            =   3180
      TabIndex        =   35
      Top             =   3345
      Width           =   2940
   End
   Begin VB.Image imgUpload 
      Height          =   240
      Left            =   2760
      Picture         =   "frmWKL38.frx":07F4
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDownload 
      Height          =   240
      Left            =   2760
      Picture         =   "frmWKL38.frx":2F96
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmWKL38"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iDatanzahl      As Integer
Dim sHosti          As String
Dim sUseri          As String
Dim sPassi          As String
Dim glFilesizeFTP   As Long
Dim glFilesizeLOKAL As Long
Dim lerrorZaehler   As Long

Private sDrives As String

Const SW_SHOWNORMAL = 1
Const FO_DELETE = &H3
Const FO_RENAME = &H4
Const FOF_ALLOWUNDO = &H40
Const FOF_NOCONFIRMATION = &H10
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Declare Function SHFileOperation Lib "Shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public WithEvents rfFile As cRemoteFile
Attribute rfFile.VB_VarHelpID = -1

Public WithEvents rfConnection As cConnection
Attribute rfConnection.VB_VarHelpID = -1
Public cFiles As New Collection, cAttrs As New Collection, cSize As New Collection, cRemAttrs As New Collection
Public nTotal As Long, DriveCol As New Collection, sCurPath As String

Private Sub Check2_Click(Index As Integer)

End Sub

Private Sub Command9_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    
    Select Case Check2(Index).Caption
        Case Is = "Neue Stammdaten sind übertragen."
            If glLevel >= DlgZugriff(1).dZugriff Then
                frmWKL11.Show 1
            Else
                MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
            End If
            
        Case Is = "Neues Programmupdate ist übertragen."
            
            frmWKL53.cmdUpdEinlesen_Click
            
        Case Is = "Neue Kassendatei ist übertragen."
            If glLevel >= DlgZugriff(5).dZugriff Then
                frmWKL27.Show 1
            Else
                MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
            End If
    End Select
        
    
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 400 Then 'Formular wird schon angezeigt
        Unload frmWKL38
        
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command9_Click"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    PositionierenWKL38
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    'Umstellung FTP Adresse
    Dim lHeute  As Long
    Dim lBis    As Long
    Dim lVon    As Long
    Dim cSQL    As String
    
    lHeute = DateValue(Now)
    lVon = DateValue("11.10.2010")
    lBis = DateValue("19.10.2010")
    
    If lHeute >= lVon Then
        If lHeute < lBis Then
            cSQL = "Update stammftp set FTPAD = '80.86.85.121'"
            gdBase.Execute cSQL, dbFailOnError
        End If
    End If
    'Ende Umstellung
    
    lvLocal.Move 0, 0, picBack.ScaleWidth, picBack.ScaleHeight
    GetDrives
    SetEnabled False
    
    Set rfFile = New cRemoteFile
    Set rfConnection = New cConnection
    UploadFlag = FTP_TRANSFER_TYPE_BINARY
    
    txtPattern.Text = "*.*" 'GetSetting("KPD FTP", "Pattern", "Pattern", "*.*")
    Frame1.Visible = False
    
    Select Case giKissFtpMode

        Case Is = 2
            txtPattern.Text = "senden.*"
    End Select
    
    FillLocalListView gsKinPfad & "\" 'GetSetting("KPD FTP", "Path", "LastPath", "C:\Windows\")
    
    Screen.MousePointer = 0
    Select Case giKissFtpMode
        Case Is = 1
            Label16.Caption = "Möchten Sie jetzt nach neuen Stammdaten bzw. Programmupdates suchen?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 2
            Label16.Caption = "Möchten Sie Ihre Bestellung abschicken?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
            Command8.Visible = False
        Case Is = 3
            Label16.Caption = "Möchten Sie jetzt nach neuen Stammdaten suchen?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 4
            Label16.Caption = "Möchten Sie jetzt nach neuen Programmupdates suchen?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 5
            Label16.Caption = "Jetzt werden Ihre Statistik - Dateien abgeschickt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 6
            Label16.Caption = "Jetzt werden Ihre Kassendateien an die Zentrale abgeschickt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 7
            Label16.Caption = "Jetzt werden Ihre Kassendateien an die Zentrale und die Statistik - Dateien abgeschickt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 8
            Label16.Caption = "Möchten Sie jetzt nach neuen Kassendateien suchen?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 9
            Label16.Caption = "Jetzt werden Ihre Kassendateien an die Zentrale geschickt bzw. neue Kassendateien geholt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 10
            Label16.Caption = "Möchten Sie jetzt nach neuen Kassendateien, Stammdaten und Programmupdates suchen?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 11
            Label16.Caption = "Möchten Sie Ihre Bestellung abschicken?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 12
            Label16.Caption = "Möchten Sie Ihre Bestellung abschicken?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 13
            Label16.Caption = "Möchten Sie Warenverteilungsdateien abholen?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 14
            Label16.Caption = "Möchten Sie Warenverteilungsdateien abschicken?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 15
            Label16.Caption = "Möchten Sie Warenverteilungsdateien abholen?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 16
            Label16.Caption = "Möchten Sie Warenverteilungsdateien abschicken?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 17
            Label16.Caption = "Möchten Sie jetzt nach neuen Stammdaten suchen?" & vbCrLf
            Label16.Caption = Label16.Caption & glLiNr & " " & ermLiefBez(glLiNr) & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 18
            Label16.Caption = "Möchten Sie jetzt nach den Tagesstammdaten suchen?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 19
            Label16.Caption = "Möchten Sie Ihre Bestellung abschicken?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 20
            Label16.Caption = "Möchten Sie Lagerdateien abholen?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 21
            Label16.Caption = "Jetzt wird die Datenbankdatei bereitgstellt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 22
            Label16.Caption = "Jetzt wird die Datenbankdatei abgeholt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 23
            Label16.Caption = "Möchten Sie Ihre Email abschicken?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 26
            Label16.Caption = "Jetzt wird die Biedro-Bestellung abgeschickt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 27
            Label16.Caption = "Jetzt wird die Loreal-Bestellung abgeschickt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 34
            Label16.Caption = "Jetzt wird die Lüning-Bestellung abgeschickt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 35
            Label16.Caption = "Jetzt wird die Budnikowsky-Bestellung abgeschickt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 36
            Label16.Caption = "Jetzt werden elektronische Lieferscheine abgeholt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 37
            Label16.Caption = "Jetzt werden Lüning Stammdaten abgeholt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 38
            Label16.Caption = "Jetzt wird die Pural-Bestellung abgeschickt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 39
            Label16.Caption = "Jetzt wird die Biogarten-Bestellung abgeschickt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 40
            Label16.Caption = "Jetzt werden Bela Stammdaten abgeholt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 43
            Label16.Caption = "Jetzt werden die Coupon-Einlösungen abgeschickt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 45
            Label16.Caption = "Jetzt wird die Rinklin-Bestellung abgeschickt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 46
            Label16.Caption = "Möchten Sie Ihre Email abschicken?" & vbCrLf
            Label16.Caption = Label16.Caption & "Dann drücken Sie 'Start'!"
        Case Is = 47
            Label16.Caption = "Jetzt wird die Menson-Bestellung abgeschickt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
        Case Is = 48
            Label16.Caption = "Jetzt wird die Boerlind-Bestellung abgeschickt." & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Start'!"
    End Select
    
    Label16.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub chkASCII_Click()
    On Error GoTo LOKAL_ERROR
    
    UploadFlag = FTP_TRANSFER_TYPE_ASCII

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "chkASCII_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub chkBinary_Click()
    On Error GoTo LOKAL_ERROR
    
    UploadFlag = FTP_TRANSFER_TYPE_BINARY
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "chkBinary_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdConnect_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim iStep As Integer
    
    iStep = 0

    If cmdConnect.Caption = "Connect" Then
        iStep = 1
        cmdConnect.Caption = "Disconnect"
        iStep = 2
        
'        rfConnection.Disconnect
        frmWKL38.rfConnection.CreateConnection True, sHosti, sUseri, sPassi
        
        iStep = 3
        GetStatus
        iStep = 4
        
        frmWKL38.FillRemoteListView
        iStep = 5
    Else
        iStep = 6
        rfConnection.Disconnect
        iStep = 7
        GetStatus
        iStep = 8
        
        lvRemote.ListItems.Clear
        iStep = 9
        txtRemPath.Text = ""
        iStep = 10
        cmdConnect.Caption = "Connect"
        iStep = 11
    End If
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 91 Then
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "cmdConnect_Click"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten. " & iStep
        
        Fehlermeldung1
    End If
End Sub
Private Sub cmdExit_Click()
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKL38
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdExit_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdStart_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    cmdStart.Visible = False
    Command8.Visible = False
    
    Screen.MousePointer = 11
    
    If onlineFrage Then
        Pause (1)
        If ftpconnectFrage = False Then
            Label16.Caption = "Sie wurden mit unserem FTP - Server nicht verbunden. Versuchen Sie es zu einem späteren Zeitpunkt nocheinmal!"
            Label16.Refresh
            
            Pause (2)
            
            Verbindungtrennen

            Command8.Visible = True
            Command8.Caption = "Beenden"
        Else
        
            Screen.MousePointer = 11
            Select Case giKissFtpMode
                Case Is = 1, 3, 4   'Stammdaten / Programmupdates
                
                    StadaUndProUpdates
                    
                    LeseLIZENZ
                    
                    If gbLizenz = True Then
                        If gbNOWOCHENDATEN = False Then
                            If giWochendat = 0 Then
                                giWochendat = ermNextWochendatei
                            End If
                            WochenStamdaholen
                        End If
                    End If
                    
                    Protokolleeren
                    
                    If gbFtpZENT Then
                        If gbWVNOT = False Then
                            WarenverteilungvonZentraleX
                        End If
                    End If
                    
                    Mailtraffic1
                    
                Case Is = 5         'Stat leeren
                    statleeren
                    theBigFTPFehler = True
                Case Is = 2         'Bestellungen
                    bestellungabschicken
                Case Is = 6         'Kassendateien an die Zentrale senden
                    If KassendateienanZentrale Then
                        theBigFTPFehler = True
                    Else
                        theBigFTPFehler = False
                    End If
                    
                Case Is = 7         'Kassendateien und Stat
                    statleeren
    
                    If KassendateienanZentrale Then
                        theBigFTPFehler = True
                    Else
                        theBigFTPFehler = False
                    End If
                    
                Case Is = 8         'Kassendateien
                    KassendateienvonZentrale
                Case Is = 9         'Kassendateien holen und schicken Kombi
                    KassendateienvonZentrale
                    
                    If gbWVNOT = False Then
                        WarenverteilungvonZentraleX
                    End If
    
                    If KassendateienanZentrale Then
                        theBigFTPFehler = True
                    Else
                        theBigFTPFehler = False
                    End If
                    
                Case Is = 10        'Kombi Stammdaten schicken und holen und ProUpdates zuerst, dann kassendateien
                    StadaUndProUpdates
                    
                    LeseLIZENZ
                    
                    If gbLizenz = True Then
                        If gbNOWOCHENDATEN = False Then
                            If giWochendat = 0 Then
                                giWochendat = ermNextWochendatei
                            End If
                            WochenStamdaholen
                        End If
                    End If
                    
                    Protokolleeren
                    KassendateienvonZentrale
                    
                    If KassendateienanZentrale Then
                        theBigFTPFehler = True
                    Else
                        theBigFTPFehler = False
                    End If
                    
                    If gbWVNOT = False Then
                        WarenverteilungvonZentraleX
                    End If
                    Mailtraffic1
                    
                Case Is = 11        'Rewe Bestellungen nicht mehr zu uns, sondern auf den Rewe- Server
                    ReweZuUns
                Case Is = 12        'Coty Bestellungen zu uns
                    CotyZuUns
                Case Is = 13        'Warenverteilungen holen
                    WarenverteilungvonZentrale
                Case Is = 14        'Warenverteilungen abschicken
                    WarenverteilungenanZentrale
                Case Is = 15        'Expressverteilungen holen
                    If gbWVNOT = False Then
                        WarenverteilungvonZentraleX
                    End If
                Case Is = 16        'Expressverteilungen abschicken
                    WarenverteilungenanZentraleX
                Case Is = 17        'Stammdaten pro Lieferant holen
                    LieferantenStamdaholen
                Case Is = 18        'Stammdaten pro Lieferant holen
                    TagesStamdaholen
                Case Is = 19        'Bestellemail verschicken
                    BestellEmailverschickenSSL
                Case Is = 20       'lagerdateienabholen
                    Lagerdateienabholen
                Case Is = 21        'End.zip wegschicken
                    EndZipweg
                Case Is = 22        'End.zip holen
                    EndZipHol
                Case Is = 23        'Email verschicken
                    BestellEmailverschicken
                Case Is = 24        'Jede Bestellungen zu uns
                    JedeBestZuUns
                Case Is = 25        'VEDES_DSL_ZuVedes
                    VEDES_DSL_ZuVedes
                Case Is = 26        'Biedro Bestellungen zu uns
                    BiedroZuUns
                Case Is = 27        'LOREAL Bestellungen zu uns
                    LOREALZuUns
                Case Is = 28        'BBI Bestellungen zu uns
                    BBIZuUns
                Case Is = 29        'Cospar GFK zu uns
                    CosparGfkZuUns
                Case Is = 30        'RFS zu uns
                    RFSZuUns
                Case Is = 31        'Bela Bestellungen zu uns / später zu Bela
                    BelaBestellungen
                Case Is = 32        'Grund FT zu uns
                    FTZuUns
                Case Is = 33        'ERNST Bestellungen zu uns
                    ErnstBestellungen
                Case Is = 34        'Lüning Bestellungen zu uns
                    LueningZuUns
                Case Is = 35        'Budni Bestellungen zu Budni
                    BudniBestellungen
                Case Is = 36        'Budni Lieferavis holen
                    BudniLieferavis
                Case Is = 37        'tägliche Lüning Stammdaten abholen
                    Lüning_Stada_holen
                Case Is = 38        'Pural Bestellungen zu uns
                    PuralZuUns
                Case Is = 39        'Biogarten Bestellungen zu uns
                    BiogartenZuUns
                Case Is = 40        'Bela Stammdaten abholen
                    Bela_Stada_holen
                Case Is = 41        'VEDES zu uns
                    VEDESZuUns
                Case Is = 42        'Carnot Bestellungen zu uns
                    CarnotZuUns
                Case Is = 43        'Dronova Couponeinlösungen zu uns
                    DronovCouponEinlZuUns
                Case Is = 44        'out leeren
                    outLeeren 'StadaUndProUpdates
                Case Is = 45        'mit hinterlegten Userdaten FTP , Rinklin-Bestellung hochladen
                    RinklinBestellungen
                Case Is = 46        'Email verschicken, neu mit SSL
                    BestellEmailverschickenSSL
                    
                Case Is = 47        'mit hinterlegten Userdaten FTP , Menson-Bestellung hochladen
                    MensonBestellungen
                Case Is = 48        'mit hinterlegten Userdaten FTP , Boerlind-Bestellung hochladen
                    BoerBestellungen
                    
                
                End Select
            End If
    
    Else
        Label16.Caption = "Nicht verbunden"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
    If gbFTPautomatic Then
        Verbindungtrennen
        Unload frmWKL38
        Screen.MousePointer = 0
        Exit Sub
    Else
        If Frame3.Visible = True Then
            For i = 0 To 2
            
            
                If Check2(i).Visible = True Then
                    Select Case Check2(i).Caption
                        Case Is = "Stammdaten werden übertragen..."
                            Check2(i).Caption = "Neue Stammdaten sind übertragen."
                        Case Is = "Programmupdate wird übertragen..."
                            Check2(i).Caption = "Neues Programmupdate ist übertragen."
                        Case Is = "Kassendateien werden übertragen..."
                            Check2(i).Caption = "Neue Kassendatei ist übertragen."
                    End Select
                    Command9(i).Visible = True
                Else
                    Command9(i).Visible = False
                End If
            Next i
        End If
        
    End If
    
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdStart_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub BestellEmailverschicken()
'    On Error GoTo LOKAL_ERROR
'
'    ' =============================
'    ' Hier beginnt der Sendevorgang
'    ' =============================
'
'    ' Anmelden am Server
'    Screen.MousePointer = 11
'    Label16.Caption = "Anmelden an: " & sevSMTP1.ServerName
'    DoEvents
'
'    mailocxcheck
''    systemdatcheck "sevMail32.ocx"
'
'    With sevSMTP1
'        .SenderName = gcBestellEmail.SenderName
'        .ReplyTo = gcBestellEmail.ReplyTo
'        .SenderEMail = gcBestellEmail.SenderEMail
'
'        gbCCfromBestlief = True
'        If gbCCfromBestlief = True Then
'            .CC = gcBestellEmail.CC
'        End If
'
'        If gcBestellEmail.BCC <> "" Then
'            .BCC = gcBestellEmail.BCC
'        End If
'
'        .Recipient = gcBestellEmail.Recipient
'        .SMTPAUTH = gcBestellEmail.SMTPAUTH
'        .ServerName = gcBestellEmail.ServerName
'        .ServerPort = gcBestellEmail.ServerPort
'        .Username = gcBestellEmail.Username
'        .Password = gcBestellEmail.Password
'        .Subject = gcBestellEmail.Subject
'        .Message = gcBestellEmail.Message
'        .AutoZIP = gcBestellEmail.AutoZIP
'
'        .AttachmentClear
'        If gcBestellEmail.Attachment1 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment1
'        End If
'
'        If gcBestellEmail.Attachment2 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment2
'        End If
'
'        If gcBestellEmail.Attachment3 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment3
'        End If
'
'        If gcBestellEmail.Attachment4 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment4
'        End If
'
'        If gcBestellEmail.Attachment5 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment5
'        End If
'
'        If gcBestellEmail.Attachment6 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment6
'        End If
'    End With
'
'    ' Anmeldung erfolgreich?
'    If sevSMTP1.Connect() = True Then
'        Label16.Caption = "Angemeldet an " & sevSMTP1.ServerName
'        DoEvents
'
''        picprogemail.Visible = True
'
'        ' Sendevorgang starten
'        Screen.MousePointer = vbNormal
'
'''        ' autom. Speichern des Seitenquelltextes
'''        sevSMTP1.SaveMailTo = Send_GetNextFile()
'
'        Dim lngBytesSent As Long
'
'        lngBytesSent = sevSMTP1.SendMail()
'        If lngBytesSent < 0 Then
'            ' Fehler aufgetreten!
'            Label16.Caption = "Fehler!"
'            Label1.Caption = sevSMTP1.SMTPErrorText
'
'            MsgBox "LastReponse: " & sevSMTP1.LastResponse & vbCrLf & _
'              "SMTPError: " & CStr(sevSMTP1.SMTPError) & " - " & sevSMTP1.SMTPErrorText
'
'        Else
'            Label16.Caption = "Nachricht versandt (" & CStr(lngBytesSent) & " Bytes)"
'        End If
'
'        ' Abmelden
'        sevSMTP1.Disconnect
''        picprogemail.Visible = False
'    Else
'        Label16.Caption = "Fehler!"
'        Label1.Caption = sevSMTP1.SMTPErrorText & vbCrLf & sevSMTP1.LastResponse
'    End If
'
'    Pause (4)
'
'    Command8.Visible = True
'    Command8.Caption = "Beenden"
'
'    Screen.MousePointer = 0
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "BestellEmailverschicken"
'    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
End Sub
Private Sub BestellEmailverschickenSSL()
    On Error GoTo LOKAL_ERROR

    ' =============================
    ' Hier beginnt der Sendevorgang
    ' =============================
    
    ' Anmelden am Server
    Screen.MousePointer = 11
    Label16.Caption = "Anmelden an: " & gcBestellEmail.ServerName
    DoEvents
    
    
    mailDLLcheck
    
    
    Dim From As String
    Dim to_addr As String
    Dim cc_addr As String
    Dim ServerAddr As String
    
    
    From = gcBestellEmail.SenderName '(textFrom.Text)
    to_addr = gcBestellEmail.Recipient
    cc_addr = gcBestellEmail.CC
    ServerAddr = gcBestellEmail.ServerName
   
    
    
    
    'Declare and create easendmail mail object instance
    Dim oSmtp As EASendMailObjLib.Mail
    Set oSmtp = New EASendMailObjLib.Mail
    'The license code for EASendMail ActiveX Object,
    'for evaluation usage, please use "TryIt" as the license code.
    
    
    'ES-D1508812687-00538-A396A7A3E3AC9A9B-U1EB72UA1E5C16FC
    
    oSmtp.LicenseCode = "ES-D1508812687-00538-A396A7A3E3AC9A9B-U1EB72UA1E5C16FC"
    'oSmtp.LogFileName = App.Path & "\smtp.txt" 'enable smtp log
    oSmtp.ServerAddr = gcBestellEmail.ServerName
    oSmtp.ServerPort = gcBestellEmail.ServerPort
    
    
    
    oSmtp.Protocol = 0 'lstProtocol.ListIndex
    
    If gcBestellEmail.ServerName <> "" Then
        
        oSmtp.ServerPort = CLng(gcBestellEmail.ServerPort)
        
        
        oSmtp.Username = Trim(gcBestellEmail.Username)
        oSmtp.Password = Trim(gcBestellEmail.Password)
        
        If gcBestellEmail.SSL = True Then
       
            oSmtp.SSL_init
            'If SSL port is 465 or other port rather than 25 or 587 port, please use
            'oSmtp.ServerPort = 465
            'oSmtp.SSL_starttls = 0
        End If
    End If
    
    
    
    oSmtp.Charset = "utf-8" 'm_arCharset(lstCharset.ListIndex, 1)
'    Dim name, addr As String
'    fnParseAddr From, name, addr
    
    'Using this email to be replied to another address
    'oSmtp.ReplyTo = ReplyAddress
    
    oSmtp.From = gcBestellEmail.SenderName ' name
    oSmtp.FromAddr = gcBestellEmail.SenderEMail 'addr
    
    'add digital signature
    oSmtp.SignerCert.Unload
'    If chkSign.Value = Checked Then
'        If Not oSmtp.SignerCert.FindSubject(addr, CERT_SYSTEM_STORE_CURRENT_USER, "my") Then
'            MsgBox oSmtp.SignerCert.GetLastError() & ":" & addr
'            btnSend.Enabled = True
'        Exit Sub
'        End If
'        If Not oSmtp.SignerCert.HasPrivateKey Then
'            MsgBox "Signer certificate has not private key, this certificate can not be used to sign email!"
'            btnSend.Enabled = True
'            Exit Sub
'        End If
'    End If
    
    oSmtp.AddRecipientEx to_addr, 0  ' 0, Normal recipient, 1, cc, 2, bcc
    oSmtp.AddRecipientEx cc_addr, 0
    
    Dim recipients As String
    recipients = to_addr & "," & cc_addr
    fnTrim recipients, ","
    
    Dim i, Count As Integer
    'encrypt email by recipients certificate
    oSmtp.RecipientsCerts.Clear
'    If chkEncrypt.Value = Checked Then
'        Dim arAddr
'        arAddr = SplitEx(recipients, ",")   'split the multiple address to an array
'        Count = UBound(arAddr)
'        For i = LBound(arAddr) To Count
'            addr = arAddr(i)
'            fnTrim addr, " ,;"
'            If addr <> "" Then
'                'find the encrypting certificate for every recipients
'                Dim oEncryptCert As New EASendMailObjLib.Certificate
'                If Not oEncryptCert.FindSubject(addr, CERT_SYSTEM_STORE_CURRENT_USER, "AddressBook") Then
'                    If Not oEncryptCert.FindSubject(addr, CERT_SYSTEM_STORE_CURRENT_USER, "my") Then
'                        MsgBox oEncryptCert.GetLastError() & ":" & addr
'                        btnSend.Enabled = True
'                        Exit Sub
'                    End If
'                End If
'                oSmtp.RecipientsCerts.Add oEncryptCert
'            End If
'        Next
'    End If
    
    
    Dim m_arAttachment() As String
    Dim iCount As Integer
    iCount = 0
    
    
    ReDim m_arAttachment(iCount)
    
    If gcBestellEmail.Attachment1 <> "" Then
        m_arAttachment(iCount) = gcBestellEmail.Attachment1
    End If
    
    iCount = iCount + 1: ReDim Preserve m_arAttachment(iCount)
    
    
    
    
    If gcBestellEmail.Attachment2 <> "" Then
        m_arAttachment(iCount) = gcBestellEmail.Attachment2
    End If
    
    iCount = iCount + 1: ReDim Preserve m_arAttachment(iCount)
    
    
    If gcBestellEmail.Attachment3 <> "" Then
        m_arAttachment(iCount) = gcBestellEmail.Attachment3
    End If
    
    iCount = iCount + 1: ReDim Preserve m_arAttachment(iCount)
    
    If gcBestellEmail.Attachment4 <> "" Then
        m_arAttachment(iCount) = gcBestellEmail.Attachment4
    End If
    
    iCount = iCount + 1: ReDim Preserve m_arAttachment(iCount)
    
    
    If gcBestellEmail.Attachment5 <> "" Then
        m_arAttachment(iCount) = gcBestellEmail.Attachment5
    End If
    
    iCount = iCount + 1: ReDim Preserve m_arAttachment(iCount)
    
    If gcBestellEmail.Attachment6 <> "" Then
        m_arAttachment(iCount) = gcBestellEmail.Attachment6
    End If
    
    iCount = iCount + 1: ReDim Preserve m_arAttachment(iCount)
    
    
    
    iCount = UBound(m_arAttachment)
    For i = 0 To iCount - 1
        If oSmtp.AddAttachment(m_arAttachment(i)) <> 0 Then
'            MsgBox oSmtp.GetLastErrDescription() & ":" & m_arAttachment(i)
'            btnSend.Enabled = True
'            Exit Sub
        End If
    Next
    
    
    
    
    
    
    
    
    
    Dim Subject As String
    Dim Bodytext As String
    
    Subject = gcBestellEmail.Subject
    Bodytext = gcBestellEmail.Message
    
'    Bodytext = Replace(Bodytext, "[$from]", From)
'    Bodytext = Replace(Bodytext, "[$to]", recipients)
'    Bodytext = Replace(Bodytext, "[$subject]", Subject)
'
    oSmtp.Subject = Subject
    oSmtp.Bodytext = Bodytext
    
    'oSmtp.BodyFormat = 1    ' Using HTML FORMAT to send mail
    
'''    If InStr(1, recipients, ",", 1) > 1 And ServerAddr = "" Then
'''        'To send email without specified smtp server, we have to send the emails one by one
'''        ' to multiple recipients. That is because every recipient has different smtp server.
'''        DirectSend oSmtp, recipients
'''''        btnSend.Enabled = True
'''''        textStatus.Caption = ""
'''        Exit Sub
'''    End If
'''
    
    
    
    
    If oSmtp.SendMail() = 0 Then
        Label16.Caption = "Nachricht erfolgreich versendet"
    Else
        Label16.Caption = oSmtp.GetLastErrDescription()  'Get last error description
    End If
    Label16.Refresh
    
    
    'Ende neu
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
'''    With sevSMTP1
'''        .SenderName = gcBestellEmail.SenderName
'''        .ReplyTo = gcBestellEmail.ReplyTo
'''        .SenderEMail = gcBestellEmail.SenderEMail
'''
'''        gbCCfromBestlief = True
'''        If gbCCfromBestlief = True Then
'''            .CC = gcBestellEmail.CC
'''        End If
'''
'''        If gcBestellEmail.BCC <> "" Then
'''            .BCC = gcBestellEmail.BCC
'''        End If
'''
'''        .Recipient = gcBestellEmail.Recipient
'''        .SMTPAUTH = gcBestellEmail.SMTPAUTH
'''        .ServerName = gcBestellEmail.ServerName
'''        .ServerPort = gcBestellEmail.ServerPort
'''        .Username = gcBestellEmail.Username
'''        .Password = gcBestellEmail.Password
'''        .Subject = gcBestellEmail.Subject
'''        .Message = gcBestellEmail.Message
'''        .AutoZIP = gcBestellEmail.AutoZIP
'''
'''        .AttachmentClear
'''        If gcBestellEmail.Attachment1 <> "" Then
'''            .AttachmentAdd gcBestellEmail.Attachment1
'''        End If
'''
'''        If gcBestellEmail.Attachment2 <> "" Then
'''            .AttachmentAdd gcBestellEmail.Attachment2
'''        End If
'''
'''        If gcBestellEmail.Attachment3 <> "" Then
'''            .AttachmentAdd gcBestellEmail.Attachment3
'''        End If
'''
'''        If gcBestellEmail.Attachment4 <> "" Then
'''            .AttachmentAdd gcBestellEmail.Attachment4
'''        End If
'''
'''        If gcBestellEmail.Attachment5 <> "" Then
'''            .AttachmentAdd gcBestellEmail.Attachment5
'''        End If
'''
'''        If gcBestellEmail.Attachment6 <> "" Then
'''            .AttachmentAdd gcBestellEmail.Attachment6
'''        End If
'''    End With
'''
'''
'''
'''
'''
'''
'''    ' Anmeldung erfolgreich?
'''    If sevSMTP1.Connect() = True Then
'''        Label16.Caption = "Angemeldet an " & sevSMTP1.ServerName
'''        DoEvents
'''
'''        ' Sendevorgang starten
'''        Screen.MousePointer = vbNormal
'''
'''        Dim lngBytesSent As Long
'''
'''        lngBytesSent = sevSMTP1.SendMail()
'''        If lngBytesSent < 0 Then
'''            ' Fehler aufgetreten!
'''            Label16.Caption = "Fehler!"
'''            Label1.Caption = sevSMTP1.SMTPErrorText
'''
'''            MsgBox "LastReponse: " & sevSMTP1.LastResponse & vbCrLf & _
'''              "SMTPError: " & CStr(sevSMTP1.SMTPError) & " - " & sevSMTP1.SMTPErrorText
'''
'''        Else
'''            Label16.Caption = "Nachricht versandt (" & CStr(lngBytesSent) & " Bytes)"
'''        End If
'''
'''        ' Abmelden
'''        sevSMTP1.Disconnect
'''
'''    Else
'''        Label16.Caption = "Fehler!"
'''        Label1.Caption = sevSMTP1.SMTPErrorText & vbCrLf & sevSMTP1.LastResponse
'''    End If
'''
    Pause (4)
    
    Command8.Visible = True
    Command8.Caption = "Beenden"
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BestellEmailverschickenSSL"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Funktion1()
    On Error GoTo LOKAL_ERROR
    
    Label16.Caption = "Neue Stammdaten und Programmupdates werden gesucht"
    Label16.Refresh
    
    Pause (1)
    
    If WechsleInsUnterverzFTP("OUT") = False Then
        Exit Sub
    End If
    
    Pause (1)
    
    WechsleInsUnterverzLOKAL gsKinPfad & "\"
    
    Pause (1)
    
    UebertrageAllVonFTP
    Pause (3)
    DelAllFTP
    
    
    
    Label16.Caption = "Protokolle werden übertragen..."
    Label16.Refresh
    
    Pause (1)

    Dim sDabaProtoPfad As String
    
    sDabaProtoPfad = gcDBPfad               'Datenbankpfad
    If Right(sDabaProtoPfad, 1) <> "\" Then
        sDabaProtoPfad = sDabaProtoPfad & "\"
    End If
    sDabaProtoPfad = sDabaProtoPfad & "Protok\"
    
  
    If WechsleInsUnterverzFTP("Protokol") = False Then
        Exit Sub
    End If
    
    
    Pause (1)
    WechsleInsUnterverzLOKAL sDabaProtoPfad
    Pause (1)
    UebertrageAllVonLOKAL
    Pause (1)
    DelAllLOKAL
    
    Dim sDabaMailPfad As String
    Dim sDabaMailoutPfad As String
    
    sDabaMailPfad = gcDBPfad               'Datenbankpfad
    If Right(sDabaMailPfad, 1) <> "\" Then
        sDabaMailPfad = sDabaMailPfad & "\"
    End If
    sDabaMailPfad = sDabaMailPfad & "Mail\"
    
    sDabaMailoutPfad = gcDBPfad               'Datenbankpfad
    If Right(sDabaMailoutPfad, 1) <> "\" Then
        sDabaMailoutPfad = sDabaMailoutPfad & "\"
    End If
    sDabaMailoutPfad = sDabaMailoutPfad & "Mailout\"
                
    Label16.Caption = "Neue Emails werden gesucht..."
    Label16.Refresh
    
    Pause (1)
    
    If WechsleInsUnterverzFTP("Mail") = False Then
        Exit Sub
    End If
    
   
    Pause (1)
    WechsleInsUnterverzLOKAL sDabaMailPfad
    Pause (1)
    UebertrageAlleMailsVonFTP
    Pause (3)
    DelAllFTP
    Pause (1)
    
    Label16.Caption = "vorhandene Emails werden versendet..."
    Label16.Refresh
    
    Pause (1)
    
    'Mails rausschicken
   
    If WechsleInsUnterverzFTP("Mailin") = False Then
        Exit Sub
    End If
    
    Pause (1)
    WechsleInsUnterverzLOKAL sDabaMailoutPfad
    Pause (1)
    UebertrageAllVonLOKAL
    
    Pause (1)
    DelAllLOKAL
    
    Label16.Caption = "KISSNET... wird getrennt..."
    Label16.Refresh

    
    
    
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Funktion1"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command8_Click()
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKL38
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdlDelete_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim msg As VbMsgBoxResult, cnt As Long
    msg = MsgBox("Are you sure you want to delete this item?", vbQuestion + vbYesNo, App.Title)
    If msg = vbYes Then
        For cnt = 2 To lvLocal.ListItems.Count - DriveCol.Count
            If lvLocal.ListItems.Item(cnt).Selected = True Then
                If lvLocal.ListItems.Item(cnt).Text = ".." Or Right$(lvLocal.ListItems.Item(cnt).Text, 2) <> ":\" Then
                    Dim FO As SHFILEOPSTRUCT
                    FO.pFrom = sCurPath + lvLocal.ListItems.Item(cnt).Text
                    FO.fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
                    FO.wFunc = FO_DELETE
                    SHFileOperation FO
                End If
            End If
        Next cnt
        FillLocalListView sCurPath
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdlDelete_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdlExec_Click()
    On Error GoTo LOKAL_ERROR
    
    If lvLocal.SelectedItem <> ".." Then ShellExecute 0, vbNullString, sCurPath + lvLocal.SelectedItem, vbNullString, sCurPath, SW_SHOWNORMAL

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdlExec_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdlMkDir_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sRet As String
    sRet = InputBox("Enter new local directory name:")
    If sRet <> "" Then
        MkDir sCurPath + sRet
        FillLocalListView sCurPath
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdlMkDir_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdlRefresh_Click()
    On Error GoTo LOKAL_ERROR
    
    FillLocalListView sCurPath
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdlRefresh_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdlRename_Click()
    On Error GoTo LOKAL_ERROR
    
    If lvLocal.SelectedItem = ".." Or Right$(lvLocal.SelectedItem, 2) = ":\" Then Exit Sub
    Dim sRet As String
    sRet = InputBox("Enter a new name for " + lvLocal.SelectedItem)
    If sRet <> "" Then
        Dim FO As SHFILEOPSTRUCT
        FO.pFrom = sCurPath + lvLocal.SelectedItem
        FO.pTo = sCurPath + sRet
        FO.fFlags = FOF_NOCONFIRMATION
        FO.wFunc = FO_RENAME
        SHFileOperation FO
        FillLocalListView sCurPath
        FillLocalListView sCurPath
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdlRename_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub cmdPut_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cnt As Long, bOk As Boolean
    For cnt = 1 To lvLocal.ListItems.Count
        If lvLocal.ListItems.Item(cnt).Selected = True Then
            If lvLocal.ListItems.Item(cnt).Text <> ".." And Right$(lvLocal.ListItems.Item(cnt).Text, 2) <> ":\" Then
                If (GetAttr(sCurPath + lvLocal.ListItems.Item(cnt).Text) And vbDirectory) <> vbDirectory Then
                    AddToCollection FOP_UPLOAD, lvLocal.ListItems.Item(cnt).Text, sCurPath, FileLen(sCurPath + lvLocal.ListItems.Item(cnt).Text)
                    bOk = True
                End If
            End If
        End If
    Next cnt
    If bFOBusy = False And bOk Then meStartFO
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdPut_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function FindRemoteFileSize(ByVal sInput As String) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cnt As Long
    sInput = LCase$(sInput)
    FindRemoteFileSize = -1
    For cnt = 1 To cFiles.Count
        If LCase$(cFiles.Item(cnt)) = sInput Then
            FindRemoteFileSize = cSize(cnt)
            Exit For
        End If
    Next cnt
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FindRemoteFileSize"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub cmdPutNew_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cnt As Long, bOk As Boolean, ret As Long
    For cnt = 1 To lvLocal.ListItems.Count
        If lvLocal.ListItems.Item(cnt).Selected = True Then
            If lvLocal.ListItems.Item(cnt).Text <> ".." And Right$(lvLocal.ListItems.Item(cnt).Text, 2) <> ":\" Then
                If (GetAttr(sCurPath + lvLocal.ListItems.Item(cnt).Text) And vbDirectory) <> vbDirectory Then
                    ret = FindRemoteFileSize(lvLocal.ListItems.Item(cnt).Text)
                    If ret <> -1 Then
                        If FileLen(sCurPath + lvLocal.ListItems.Item(cnt).Text) <> ret Then
                            AddToCollection FOP_UPLOAD, lvLocal.ListItems.Item(cnt).Text, sCurPath, FileLen(sCurPath + lvLocal.ListItems.Item(cnt).Text)
                            bOk = True
                        End If
                    Else
                        AddToCollection FOP_UPLOAD, lvLocal.ListItems.Item(cnt).Text, sCurPath, FileLen(sCurPath + lvLocal.ListItems.Item(cnt).Text)
                        bOk = True
                    End If
                End If
            End If
        End If
    Next cnt
    If bFOBusy = False And bOk Then meStartFO
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdPutNew_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdrDelete_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim msg As VbMsgBoxResult, cnt As Long
    msg = MsgBox("Are you sure you want to delete these items?", vbQuestion + vbYesNo, App.Title)
    If msg = vbYes Then
        For cnt = 1 To cRemAttrs.Count
            If lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) <> vbDirectory Then
                rfFile.RemoteFile = lvRemote.ListItems.Item(cnt).Text
                rfFile.DeleteFile rfConnection
                GetStatus
            ElseIf lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) = vbDirectory And lvRemote.ListItems.Item(cnt).Text <> ".." Then
                rfConnection.RemoveDirectory lvRemote.ListItems.Item(cnt).Text
                GetStatus
            End If
        Next cnt
        FillRemoteListView
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdrDelete_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdRetrieve_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cnt As Long
    Dim bOk As Boolean
    
    For cnt = 1 To lvRemote.ListItems.Count
        If lvRemote.ListItems.Item(cnt).Selected = True Then
            If lvRemote.ListItems.Item(cnt).Text <> ".." Then
                If (cRemAttrs.Item(cnt) And vbDirectory) <> vbDirectory Then
                    AddToCollection FOP_DOWNLOAD, lvRemote.ListItems.Item(cnt).Text, sCurPath, Val(lvRemote.ListItems.Item(cnt).SubItems(1))
                    bOk = True
                    
                End If
            End If
        End If
    Next cnt
    If bFOBusy = False And bOk Then meStartFO
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdRetrieve_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdrExec_Click()
    On Error GoTo LOKAL_ERROR
    
    If (cRemAttrs.Item(lvRemote.SelectedItem.Index) And vbDirectory) <> vbDirectory Then
        Dim strTemp As String
        strTemp = String(100, 0)
        GetTempPath 100, strTemp
        strTemp = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
        rfFile.RemoteFile = lvRemote.SelectedItem
        rfFile.GetFile rfConnection, strTemp + lvRemote.SelectedItem
        'Open strTemp + lvRemote.SelectedItem For Binary As #1
        '    Put #1, , rfFile.FileData
        'Close
        ShellExecute 0, vbNullString, strTemp + lvRemote.SelectedItem, vbNullString, vbNullString, SW_SHOWNORMAL
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdrExec_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdrMkDir_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim ret As String
    ret = InputBox("Enter new remote directory name:")
    If ret <> "" Then
        rfConnection.CreateDirectory ret
        GetStatus
        FillRemoteListView
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdrMkDir_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdrRefresh_Click()
    On Error GoTo LOKAL_ERROR
    
    FillRemoteListView
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdrRefresh_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdrRename_Click()
    On Error GoTo LOKAL_ERROR
    
    If lvRemote.SelectedItem.Index = 1 Then Exit Sub
    Dim ret As String
    ret = InputBox("Enter the new name for " + lvRemote.SelectedItem.Text)
    If ret <> "" Then
        rfFile.RemoteFile = lvRemote.SelectedItem.Text
        rfFile.RenameFile rfConnection, ret
        GetStatus
        FillRemoteListView
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdrRename_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub UpdateProgress()
    On Error GoTo LOKAL_ERROR
    
    If Me.Visible = False Then Exit Sub
    
    Dim t           As Long
    Dim iSec        As Double
    Dim iMin        As Integer
    Dim iHour       As Integer
    
    Dim sRestZeit   As String
    
    lblCurrentFile.Caption = ActiveFile
    If foItems = 0 Then
        lblFileCount.Caption = "(0/0)"
    Else
        lblFileCount.Caption = "(" + CStr(ActiveIndex) + "/" + CStr(foItems) + ")"
    End If

    If ActiveFileBytesTotal <> 0 Then
        pbrFile.value = Int(ActiveFileBytesSent / ActiveFileBytesTotal * 100)
        lblFileProg.Caption = CStr(Int(ActiveFileBytesSent / ActiveFileBytesTotal * 100)) + "%"
    Else
        lblFileProg.Caption = ""
    End If
    If TotalFileSize <> 0 Then
        pbrFiles.value = Int(SentBytes / TotalFileSize * 100)
        lblFilesProg.Caption = CStr(Int(SentBytes / TotalFileSize * 100)) + "%"
'        If Left(lblFilesProg.Caption, 2) = "99" Then
''            MsgBox "Stop 99"
'        End If
    Else
        lblFilesProg.Caption = ""
    End If
    If SentBytes <> 0 Then
        t = GetTickCount - StartT
        If t <> 0 Then
            OldSpeed = (OldSpeed + ((SentBytes / 1000) / (t / 1000))) / 2
            iSec = Int(((TotalFileSize - SentBytes) / 1000) / OldSpeed)
            
            iHour = 0
            iMin = 0
            
            Select Case iSec
                Case Is > 3600
                    
                    iHour = Int(iSec / 3600)
                    iSec = iSec Mod 3600
                    Select Case iSec
                        Case Is > 60
                    
                            iMin = Int(iSec / 60)
                            iSec = iSec Mod 60
                        Case Else
'                            iMin = Int(iSec / 60)
'                            iSec = iSec Mod 60
                    End Select
                        
                    
                Case Is > 60
                    iHour = 0
                    iMin = Int(iSec / 60)
                    iSec = iSec Mod 60
                Case Is < 61
                    iHour = 0
                    iMin = 0
            End Select
            
            sRestZeit = Format$(iHour, "00") & ":" & Format$(iMin, "00") & ":" & Format$(iSec, "00")
            
            lblTimeLeft.Caption = "verbleibende Zeit: " & sRestZeit & " h  " + Format(OldSpeed, "#.##") + "Kbps"
        End If
    End If
    If ActiveProcedure = FOP_UPLOAD Then
        imgUpload.Visible = True
        imgDownload.Visible = False
    ElseIf ActiveProcedure = FOP_DOWNLOAD Then
        imgUpload.Visible = False
        imgDownload.Visible = True
    Else
        imgUpload.Visible = False
        imgDownload.Visible = False
    End If
    DoEvents
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UpdateProgress"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub PositionierenWKL38()
    On Error GoTo LOKAL_ERROR
    
'    Label16.Height = 2055
'    Label16.Left = 1080
'    Label16.Top = 240
'    Label16.Width = 6615

'    Frame1.Top = 0
'    Frame1.Left = 0
'    Frame1.Width = 10000
'    Frame1.Height = 4000

'    Frame3.Top = 2400
'    Frame3.Left = 960
'    Frame3.Width = 7095
'    Frame3.Height = 1575
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL38"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function isdfueda() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim S&
    Dim LN&
    Dim X%
    Dim R(255) As RASENTRYNAME95
    Dim lRet As Long

    Screen.MousePointer = 11
    
    isdfueda = False
    
    If gsDFU = "keine DFÜ vorhanden" Then
        Exit Function
    End If
    
    '### Namen der bestehenden DFÜ-Verbindungen einlesen
    R(0).dwSize = 264
    S = 256 * R(0).dwSize
    lRet = RasEnumEntries(vbNullString, vbNullString, R(0), S, LN)

    List2.Clear
    
    If lRet = 0 Then
        If LN <> 0 Then
            '### Es besteht mindestens eine DFÜ-Verbindung
            For X = 0 To LN - 1
                
                ConName = StrConv(R(X).szEntryName(), vbUnicode)
                
                Select Case giKissFtpMode
                    Case Is = 2
                        If Trim(Left$(ConName, 6)) = "Esüdro" Then
                            List2.AddItem Trim(Left$(ConName, 6))
                            isdfueda = True
                        End If
                    Case Else
'                        Änderung 15 11 05
                        If Trim(Left$(ConName, Len(gsDFU))) = gsDFU Then
                            List2.AddItem gsDFU

                            isdfueda = True
                        End If
                End Select
            Next X
        End If
    End If
    
    If List2.ListCount > 0 Then
        isdfueda = True
    End If
    
    
    
    
    Screen.MousePointer = 0
    
     Exit Function
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "isdfueda"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
End Function
Private Sub Command7_Click()
On Error GoTo LOKAL_ERROR
    Dim iRet As Long
    Dim iNochmal As Byte
    
    gbVerbindungstarten = False
    
    iRet = 1
    ConID = 5
    
    If List2.ListIndex = -1 Then
        MsgBox "Kein Eintrag"
        Exit Sub
    End If
    
    iNochmal = 0
    
nochmal:
    iRet = InternetDial(Me.hwnd, ConName, DIAL_FORCE_UNATTENDED, ConID, 0)
    
    If iRet <> 0 Then
        
        Select Case iRet
            Case Is = 631 'Anwahl abgebrochen
                Command8.Visible = True
                Command8.Caption = "Beenden"
                Exit Sub
            Case Is = 678
                Label16.Caption = "Besetzt - später wiederholen"
                Label16.Refresh
                Screen.MousePointer = 0
                Command8.Visible = True
                Command8.Caption = "Beenden"
                Exit Sub
            Case Is = 680
                Label16.Caption = "Anschluss kann nicht geöffnet werden." & vbCrLf & "Hotline anrufen!"
                Label16.Refresh
                Screen.MousePointer = 0
                Command8.Visible = True
                Command8.Caption = "Beenden"
                Exit Sub
            Case Else
                If iNochmal > 3 Then
                    Label16.Caption = "Fehlernummer: " & iRet & vbCrLf & "Hotline anrufen!"
                    Label16.Refresh
                    Screen.MousePointer = 0
                    Command8.Visible = True
                    Command8.Caption = "Beenden"
                    Exit Sub
                Else
                    iNochmal = iNochmal + 1
                    GoTo nochmal
                End If
        End Select
    Else
        gbVerbindungstarten = True
    End If
    
    

    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub InternetVerbinden()
On Error GoTo LOKAL_ERROR

    Dim iRet As Long
    Dim iNochmal As Byte
    Dim sDfuName As String

    gbVerbindungstarten = False
    If isdfueda = False Then
    
        Label16.Caption = "Keine DFÜ - Verbindung vorhanden."
        Label16.Refresh
        Pause (4)
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
        Exit Sub
    
    End If
    
    Select Case giKissFtpMode
        Case Is = 2
            sDfuName = "Esüdro"
        Case Else
            'Änderung
            sDfuName = gsDFU '"KISSHANN"
    End Select
    

    iRet = 1
    ConID = 4

    iNochmal = 0

nochmal:

    If gbFTPautomatic = False Then
        iRet = InternetDial(Me.hwnd, sDfuName, &H2000, ConID, 0)
    ElseIf gbFTPautomatic = True Then
        iRet = InternetDial(Me.hwnd, sDfuName, DIAL_FORCE_UNATTENDED, ConID, 0)
    End If
    If iRet <> 0 Then

        Select Case iRet
            Case Is = 631 'Anwahl abgebrochen
                Command8.Visible = True
                Command8.Caption = "Beenden"
                Exit Sub
            Case Is = 678
                Label16.Caption = "Besetzt - später wiederholen"
                Label16.Refresh
                Screen.MousePointer = 0
                Command8.Visible = True
                Command8.Caption = "Beenden"
                Exit Sub
            Case Is = 680
                Label16.Caption = "Anschluss kann nicht geöffnet werden." & vbCrLf & "Hotline anrufen!"
                Label16.Refresh
                Screen.MousePointer = 0
                Command8.Visible = True
                Command8.Caption = "Beenden"
                Exit Sub
            Case Else
                If iNochmal > 3 Then
                    Label16.Caption = "Fehlernummer: " & iRet & vbCrLf & "Hotline anrufen!"
                    Label16.Refresh
                    Screen.MousePointer = 0
                    Command8.Visible = True
                    Command8.Caption = "Beenden"
                    Exit Sub
                Else
                    iNochmal = iNochmal + 1
                    GoTo nochmal
                End If
        End Select
    ElseIf iRet = 0 Then
        gbVerbindungstarten = True
    End If



    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InternetVerbinden"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Activate()
On Error GoTo LOKAL_ERROR

If gbFTPautomatic Then
    cmdStart_Click
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Activate"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub StadaProtokollschreiben(sfilen As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Not tableSuchenDBKombi("FTPPRO", 1) Then
        sSQL = "Create Table FTPPRO (DatName Text(20)"
        sSQL = sSQL & " , loadDate datetime )"
        
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    sSQL = "Insert into FTPPRO  (DatName,loadDate) values ( '" & sfilen & "',datevalue(now))"
    gdBase.Execute sSQL, dbFailOnError

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "StadaProtokollschreiben"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub List2_Click()
    On Error GoTo LOKAL_ERROR
    
    ConName = List2.list(List2.ListIndex)
    DFÜname = ConName
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_Click"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function WechsleInsUnterverzFTP(sUnterverzeichnis As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
   
    WechsleInsUnterverzFTP = False
    
    
    MousePointer = 11
    
    Label16.Caption = Label16.Caption & "..."
    Label16.Refresh
    
    PauseSi 0.3
    Label16.Caption = Label16.Caption & "..."
    Label16.Refresh
    
    PauseSi 0.3
    Label16.Caption = Label16.Caption & "..."
    Label16.Refresh
    
    PauseSi 0.3
    Label16.Caption = Label16.Caption & "..."
    Label16.Refresh
    
    PauseSi 0.3
    Label16.Caption = Label16.Caption & "..."
    Label16.Refresh
    
    PauseSi 0.3
    Label16.Caption = Label16.Caption & "..."
    Label16.Refresh
    
    PauseSi 0.3
    Label16.Caption = Label16.Caption & "..."
    Label16.Refresh
    
    PauseSi 0.3
    Label16.Caption = Label16.Caption & "..."
    Label16.Refresh
    
    PauseSi 0.3
    Label16.Caption = Label16.Caption & "..."
    Label16.Refresh
    
    PauseSi 0.3
    Label16.Caption = Label16.Caption & "..."
    Label16.Refresh
    
    anzeige "erfolg", Label16.Caption, Label16

    
    lvRemote.Refresh

    For i = 1 To lvRemote.ListItems.Count
''        'neu 24.08.06
''        schreibeProtokoll "Das Unterverzeichnis: " & UCase(Trim(lvRemote.ListItems.Item(i))) & " wurde gefunden."
''        'neu 24.08.06
        
        If UCase(Trim(lvRemote.ListItems.Item(i))) = UCase(Trim(sUnterverzeichnis)) Then
            lvRemote.ListItems.Item(i).Selected = True
            WechsleInsUnterverzFTP = True
            
            Exit For
        End If
    Next i
    
    'dblClick auf Liste
    lvRemote_DblClick
               
    MousePointer = 0
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WechsleInsUnterverzFTP"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub WechsleInsUnterverzLOKAL(sUnterverzPfad As String)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim bOk As Boolean
    
    MousePointer = 11
    
    FillLocalListView sUnterverzPfad
         
    MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WechsleInsUnterverzLOKAL"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DelAllFTP()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim cnt As Long
    
    'selectiere
    For i = 1 To lvRemote.ListItems.Count
        If Left(UCase(lvRemote.ListItems.Item(i)), 1) = "D" Then
            lvRemote.ListItems.Item(i).Selected = False
        Else
            lvRemote.ListItems.Item(i).Selected = True
        End If
    Next i
    
    

    'lösche
    
    For cnt = 1 To cRemAttrs.Count
        If lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) <> vbDirectory Then
            rfFile.RemoteFile = lvRemote.ListItems.Item(cnt).Text
            rfFile.DeleteFile rfConnection
            StadaProtokollschreiben lvRemote.ListItems.Item(cnt).Text
            GetStatus
        ElseIf lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) = vbDirectory And lvRemote.ListItems.Item(cnt).Text <> ".." Then
'            rfConnection.RemoveDirectory lvRemote.ListItems.Item(cnt).Text
'            GetStatus
        End If
    Next cnt
    'Liste neu füllen
    FillRemoteListView

    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 35600 Then
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DelAllFTP"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub DeleinteilfromFTP()
    On Error GoTo LOKAL_ERROR
    
    Dim i           As Integer
    Dim cnt         As Long
    Dim cFil        As String
    Dim sfilename   As String
    
    'selectiere
    
    For i = 1 To lvRemote.ListItems.Count
        sfilename = lvRemote.ListItems.Item(i)
        cFil = Trim(gcFilNr)
        If Len(cFil) = 1 Then
            cFil = "0" & cFil
        End If
        
        Select Case UCase(Left(sfilename, 4))
            Case Is = "WV" & cFil
                lvRemote.ListItems.Item(i).Selected = True
        End Select
    Next i
    
    

    'lösche
    
    For cnt = 1 To cRemAttrs.Count
        If lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) <> vbDirectory Then
            rfFile.RemoteFile = lvRemote.ListItems.Item(cnt).Text
            rfFile.DeleteFile rfConnection
            StadaProtokollschreiben lvRemote.ListItems.Item(cnt).Text
            GetStatus
        ElseIf lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) = vbDirectory And lvRemote.ListItems.Item(cnt).Text <> ".." Then
'            rfConnection.RemoveDirectory lvRemote.ListItems.Item(cnt).Text
'            GetStatus
        End If
    Next cnt
    'Liste neu füllen
    FillRemoteListView

    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DeleinteilfromFTP"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub DeleinteilfromFTP_Budni(sKUNDNR As String)
    On Error GoTo LOKAL_ERROR
    
    Dim i           As Integer
    Dim cnt         As Long
    Dim cFil        As String
    Dim sfilename   As String
    
    'selectiere
    
    For i = 1 To lvRemote.ListItems.Count
        sfilename = lvRemote.ListItems.Item(i)
        sKUNDNR = String$(10 - Len(sKUNDNR), "0") & sKUNDNR
    
    
        
        Select Case Left(sfilename, 28)
        
            Case Is = "BUDNIDESADV_Kunde " & sKUNDNR
                lvRemote.ListItems.Item(i).Selected = True
                
        End Select
    Next i
    
    

    'lösche
    
    For cnt = 1 To cRemAttrs.Count
        If lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) <> vbDirectory Then
            rfFile.RemoteFile = lvRemote.ListItems.Item(cnt).Text
            rfFile.DeleteFile rfConnection
            StadaProtokollschreiben lvRemote.ListItems.Item(cnt).Text
            GetStatus
        ElseIf lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) = vbDirectory And lvRemote.ListItems.Item(cnt).Text <> ".." Then
'            rfConnection.RemoveDirectory lvRemote.ListItems.Item(cnt).Text
'            GetStatus
        End If
    Next cnt
    'Liste neu füllen
    FillRemoteListView

    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DeleinteilfromFTP_Budni"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub DeleinteilfromFTP_Lüning(sKUNDNR As String)
    On Error GoTo LOKAL_ERROR
    
    Dim i           As Integer
    Dim cnt         As Long
    Dim cFil        As String
    Dim sfilename   As String
    
    'selectiere
    
    For i = 1 To lvRemote.ListItems.Count
        sfilename = lvRemote.ListItems.Item(i)
        sKUNDNR = String$(6 - Len(sKUNDNR), "0") & sKUNDNR
    
        Select Case Left(sfilename, 7)
            Case Is = "A" & sKUNDNR
                lvRemote.ListItems.Item(i).Selected = True
                
        End Select
    Next i
    
    'lösche
    
    For cnt = 1 To cRemAttrs.Count
        If lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) <> vbDirectory Then
            rfFile.RemoteFile = lvRemote.ListItems.Item(cnt).Text
            rfFile.DeleteFile rfConnection
            StadaProtokollschreiben lvRemote.ListItems.Item(cnt).Text
            GetStatus
        ElseIf lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) = vbDirectory And lvRemote.ListItems.Item(cnt).Text <> ".." Then
'            rfConnection.RemoveDirectory lvRemote.ListItems.Item(cnt).Text
'            GetStatus
        End If
    Next cnt
    'Liste neu füllen
    FillRemoteListView

    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DeleinteilfromFTP_Lüning"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub DeleinteilfromFTPX()
    On Error GoTo LOKAL_ERROR
    
    Dim i           As Integer
    Dim cnt         As Long
    Dim cFil        As String
    Dim sfilename   As String
    
    'selectiere
    
    For i = 1 To lvRemote.ListItems.Count
        sfilename = lvRemote.ListItems.Item(i)
        cFil = Trim(gcFilNr)
        If Len(cFil) = 1 Then
            cFil = "0" & cFil
        End If
        
        Select Case UCase(Left(sfilename, 3))
            Case Is = "N" & cFil
                lvRemote.ListItems.Item(i).Selected = True
        End Select
    Next i
    
    

    'lösche
    
    For cnt = 1 To cRemAttrs.Count
        If lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) <> vbDirectory Then
            rfFile.RemoteFile = lvRemote.ListItems.Item(cnt).Text
            rfFile.DeleteFile rfConnection
            StadaProtokollschreiben lvRemote.ListItems.Item(cnt).Text
            GetStatus
        ElseIf lvRemote.ListItems.Item(cnt).Selected = True And (cRemAttrs.Item(cnt) And vbDirectory) = vbDirectory And lvRemote.ListItems.Item(cnt).Text <> ".." Then
'            rfConnection.RemoveDirectory lvRemote.ListItems.Item(cnt).Text
'            GetStatus
        End If
    Next cnt
    'Liste neu füllen
    FillRemoteListView

    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DeleinteilfromFTPX"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub DelAllLOKAL()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim cnt As Long
    
    'selectiere
    For i = 1 To lvLocal.ListItems.Count
        lvLocal.ListItems.Item(i).Selected = True
    Next i
    
    'lösche
    For cnt = 2 To lvLocal.ListItems.Count - DriveCol.Count
        If lvLocal.ListItems.Item(cnt).Selected = True Then
            If lvLocal.ListItems.Item(cnt).Text = ".." Or Right$(lvLocal.ListItems.Item(cnt).Text, 2) <> ":\" Then
                Dim FO As SHFILEOPSTRUCT
                FO.pFrom = sCurPath + lvLocal.ListItems.Item(cnt).Text
                FO.fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
                FO.wFunc = FO_DELETE
                SHFileOperation FO
            End If
        End If
    Next cnt
    
    'liste neu füllen
    FillLocalListView sCurPath

    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DelAllLOKAL"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Verbindungtrennen()
    On Error GoTo LOKAL_ERROR
    
    Dim t           As Integer
    Dim sDfuName    As String
    
    If gbDSL = True Then
        Exit Sub
        
    End If
    
    Select Case giKissFtpMode
        Case Is = 2
            sDfuName = "Esüdro"
        Case Else
            'Änderung
            
            If Trim$(gsDFU) = "" Then gsDFU = "KISSHANN"
            sDfuName = gsDFU ' "KISSHANN"
    End Select
    
    If Trim(sDfuName) <> "" Then
        
        t = 2
        Do Until t = 5
            If DFÜStatus Then
                Call HangUp(sDfuName)
            Else

                t = 5
            End If
        Loop
        
    End If
    
    If DFÜStatus = False Then
        Label16.Caption = "Die Verbindung wurde getrennt."
        Label16.Refresh
    Else
        Label16.Caption = "Sie sind noch mit dem Internet verbunden."
        Label16.Refresh
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Verbindungtrennen"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ftpconnectFrage() As Boolean
    On Error GoTo LOKAL_ERROR
    
    ftpconnectFrage = False
    
    If giKissFtpMode = 2 Or giKissFtpMode = 19 Or giKissFtpMode = 23 Or giKissFtpMode = 46 Then
        ftpconnectFrage = True
        Exit Function
    End If
    
    Label16.Caption = "FTP - Server wird verbunden...(Test), Bitte warten..."
    Label16.Refresh
    
    sHosti = gsStammFTPAdresse
    sUseri = gsStammFTPUSER
    sPassi = gsStammFTPPASS
    
    If frmWKL38.rfConnection.CreateConnection(True, sHosti, sUseri, sPassi) = True Then
        
        ftpconnectFrage = True
        Label16.Caption = "FTP - Server ist erreichbar"
        Label16.Refresh
        

    Else
        
        Label16.Caption = ""
        Label16.Refresh
        ftpconnectFrage = False
    End If
    
Exit Function
LOKAL_ERROR:
    If err.Number = 91 Then
        Exit Function
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ftpconnectFrage"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function
Private Function onlineFrage() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim t As Integer
    
    If gbDSL = True Then
        onlineFrage = True
        Exit Function
    End If
    
    
    onlineFrage = False
    
    Label16.Caption = "DFÜ - Verbindung wird gesucht..."
    Label16.Refresh
    
    If DFÜStatus Then
        onlineFrage = True
        Exit Function
    Else
    '************************************      nicht online
        Screen.MousePointer = 0
        InternetVerbinden
        
        Screen.MousePointer = 11
        
        If gbVerbindungstarten Then
            t = 0
            Do Until t = 10
                If DFÜStatus Then
                    Label16.Caption = "Verbindung ist hergestellt"
                    Label16.Refresh
                    onlineFrage = True
                    Exit Function
                
                    t = 5
                Else
                    Pause (1)
                    t = t + 1
                    onlineFrage = False
                End If
            Loop
        Else
            Label16.Caption = "Es wurde keine Verbindung hergestellt."
            Label16.Refresh
            onlineFrage = False
            Exit Function
        End If
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "onlineFrage"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    LogtoEnd Me
    
    If rfConnection.Connected Then rfConnection.Disconnect
'    SaveSetting "KPD FTP", "Pattern", "Pattern", txtPattern.Text
'    SaveSetting "KPD FTP", "Path", "LastPath", sCurPath
    Set frmWKL38 = Nothing
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub


Private Sub lvLocal_DblClick()
    On Error GoTo LOKAL_ERROR
    
    If Right$(lvLocal.SelectedItem, 2) = ":\" Then
        FillLocalListView lvLocal.SelectedItem
    ElseIf lvLocal.SelectedItem = ".." Then
        FillLocalListView RemoveLastDir(sCurPath)
    ElseIf (GetAttr(sCurPath + lvLocal.SelectedItem) And vbDirectory) = vbDirectory Then
        FillLocalListView sCurPath + lvLocal.SelectedItem + "\"
    Else
        Dim cnt As Long, bOk As Boolean
        For cnt = 1 To lvLocal.ListItems.Count
            If lvLocal.ListItems.Item(cnt).Selected = True Then
                If lvLocal.ListItems.Item(cnt).Text <> ".." And Right$(lvLocal.ListItems.Item(cnt).Text, 2) <> ":\" Then
                    If (GetAttr(sCurPath + lvLocal.ListItems.Item(cnt).Text) And vbDirectory) <> vbDirectory Then
                        AddToCollection FOP_UPLOAD, lvLocal.ListItems.Item(cnt).Text, sCurPath, FileLen(sCurPath + lvLocal.ListItems.Item(cnt).Text)
                        bOk = True
                    End If
                End If
            End If
        Next cnt
        If bFOBusy = False And bOk Then meStartFO
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lvLocal_DblClick"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub lvRemote_DblClick()
    On Error GoTo LOKAL_ERROR
    
    If bFOBusy Then
        MsgBox "Unable to execute command...", vbExclamation + vbOKOnly, App.Title
        Exit Sub
    End If
    Dim ret As Long
    ret = GetRemoteIndex
    If ret <> -1 Then
        If (cRemAttrs.Item(ret) And vbDirectory) = vbDirectory Then
            rfConnection.SetNewDirectory lvRemote.SelectedItem
            GetStatus
            FillRemoteListView
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lvRemote_DblClick"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function GetRemoteIndex() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim bOk As Boolean
    For GetRemoteIndex = 1 To lvRemote.ListItems.Count
        If lvRemote.ListItems.Item(GetRemoteIndex).Selected Then
            bOk = True
            Exit For
        End If
    Next GetRemoteIndex
    If bOk = False Then GetRemoteIndex = -1
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GetRemoteIndex"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Sub FillLocalListView(sPath As String)
    On Error GoTo LOKAL_ERROR
    
    Dim ret As String, cnt As Long, Tel As Long
    If IsDriveAvailable(sPath) = False Then
        MsgBox "Drive not ready!", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If
    lvLocal.Visible = False
    lvLocal.ListItems.Clear
    lvLocal.ListItems.Add , , "..", , "Up"
    lstTemp.Clear
    ret = Dir(sPath, vbDirectory)
    While ret <> ""
        If (GetAttr(sPath + ret) And vbDirectory) = vbDirectory And ret <> ".." And ret <> "." Then lstTemp.AddItem ret
        ret = Dir()
    Wend
    Tel = lstTemp.ListCount
    For cnt = 0 To lstTemp.ListCount - 1
        lvLocal.ListItems.Add , , lstTemp.list(cnt), , "Directory"
    Next cnt
    lstTemp.Clear
    ret = Dir(sPath + txtPattern.Text, vbNormal)
    While ret <> ""
        If (GetAttr(sPath + ret) And vbDirectory) <> vbDirectory Then lstTemp.AddItem ret
        ret = Dir()
    Wend
    Tel = Tel + lstTemp.ListCount
    For cnt = 0 To lstTemp.ListCount - 1
        lvLocal.ListItems.Add , , lstTemp.list(cnt), , "File"
        lvLocal.ListItems.Item(lvLocal.ListItems.Count).SubItems(1) = FileLen(sPath + lstTemp.list(cnt))
    Next cnt
    For cnt = 1 To DriveCol.Count
        lvLocal.ListItems.Add , , DriveCol.Item(cnt), , "Drive"
    Next cnt
    lvLocal.Visible = True
    txtCurPath.Text = sPath
    lblLocFiles.Caption = CStr(Tel)
    sCurPath = sPath
    Putfocus lvLocal.hwnd
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FillLocalListView"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub GetDrives()
    On Error GoTo LOKAL_ERROR
    
    Dim LDs As Long, cnt As Long
    LDs = GetLogicalDrives
    sDrives = "Available drives:"
    For cnt = 0 To 25
        If (LDs And 2 ^ cnt) <> 0 Then
            DriveCol.Add Chr$(65 + cnt) + ":\"
        End If
    Next cnt
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GetDrives"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function IsDriveAvailable(sDrive As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    If GetVolumeInformation(Left$(sDrive, 3), vbNullString, 0, ByVal 0&, 0, 0, vbNullString, 0) <> 0 Then IsDriveAvailable = True

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "IsDriveAvailable"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function RemoveLastDir(ByVal sInput As String) As String
    On Error GoTo LOKAL_ERROR
    
   Dim cnt As Long
   
    RemoveLastDir = sInput
    If Right$(sInput, 1) = "\" Then sInput = Left$(sInput, Len(sInput) - 1)
    For cnt = 0 To Len(sInput) - 1
        If Mid$(sInput, Len(sInput) - cnt, 1) = "\" Then
            RemoveLastDir = Left$(sInput, Len(sInput) - cnt)
            Exit For
        End If
    Next
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "RemoveLastDir"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Sub FillRemoteListView()
    On Error GoTo LOKAL_ERROR
    
    Dim cnt As Long
    'While lvRemote.ListItems.Count > 0
        'lvRemote.ListItems.Remove lvRemote.ListItems.Count
    'Wend
    lvRemote.ListItems.Clear
    rfConnection.EnumFiles cFiles, cAttrs, cSize
    rfConnection.ClearCollection cRemAttrs
    lvRemote.ListItems.Add , , "..", , "Up"
    cRemAttrs.Add vbDirectory
    lstTemp.Clear
    For cnt = 1 To cFiles.Count
        If (cAttrs(cnt) And vbDirectory) = vbDirectory Then
            lstTemp.AddItem cFiles(cnt)
            cRemAttrs.Add vbDirectory
        End If
    Next cnt
    For cnt = 0 To lstTemp.ListCount - 1
        lvRemote.ListItems.Add , , lstTemp.list(cnt), , "Directory"
    Next cnt
    lstTemp.Clear
    For cnt = 1 To cFiles.Count
        If (cAttrs(cnt) And vbDirectory) <> vbDirectory Then
            lstTemp.AddItem cFiles(cnt) + "/" + CStr(cSize(cnt))
            cRemAttrs.Add vbNormal
        End If
    Next cnt
    For cnt = 0 To lstTemp.ListCount - 1
        lvRemote.ListItems.Add , , Left$(lstTemp.list(cnt), InStr(1, lstTemp.list(cnt), "/") - 1), , "File"
        lvRemote.ListItems.Item(lvRemote.ListItems.Count).SubItems(1) = Right$(lstTemp.list(cnt), Len(lstTemp.list(cnt)) - InStr(1, lstTemp.list(cnt), "/"))
    Next cnt
    txtRemPath.Text = rfConnection.GetCurrentDirectory
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FillRemoteListView"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub meStartFO()
    On Error GoTo LOKAL_ERROR
    
    StartFO

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "meStartFO"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub SetEnabled(bEnabled As Boolean)
    On Error GoTo LOKAL_ERROR
    
    lvRemote.Enabled = bEnabled
    cmdrMkDir.Enabled = bEnabled
    cmdrExec.Enabled = bEnabled
    cmdrRename.Enabled = bEnabled
    cmdrDelete.Enabled = bEnabled
    cmdrRefresh.Enabled = bEnabled
    cmdRetrieve.Enabled = bEnabled
    cmdPut.Enabled = bEnabled
    cmdCancel.Enabled = bEnabled
    cmdPutNew.Enabled = bEnabled
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SetEnabled"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub rfConnection_StatusChanged(NewStatus As tNewStatus, sOptionalInfo As String)
    On Error GoTo LOKAL_ERROR
    
    If NewStatus = nsConnected Then
        SetEnabled True
    ElseIf NewStatus = nsDisconnected Then
        SetEnabled False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "rfConnection_StatusChanged"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub StadaUndProUpdates()
    On Error GoTo LOKAL_ERROR
                
    Label16.Caption = "Neue Stammdaten und Programmupdates werden gesucht"
    Label16.Refresh
    
    sHosti = gsStammFTPAdresse
    sUseri = gsStammFTPUSER
    sPassi = gsStammFTPPASS

    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    
    Pause (1)
    
    If WechsleInsUnterverzFTP("OUT") = False Then
        Exit Sub
    End If
    
    Pause (1)
    
    WechsleInsUnterverzLOKAL gsKinPfad & "\"
    
    Pause (3)
    
    UebertrageAllVonFTP
    
    Pause (5)
    
    DelAllFTP
    
    Pause (2)
    

    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"

    If giKissFtpMode <> 10 And giKissFtpMode <> 1 And giKissFtpMode <> 3 And giKissFtpMode <> 4 Then
        Verbindungtrennen
    End If

    SchreibeLastFTP
    
    If giKissFtpMode <> 10 And giKissFtpMode <> 1 And giKissFtpMode <> 3 And giKissFtpMode <> 4 Then
        If gbFTPautomatic = False Then
            Pause (1)
            Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
            Label16.Refresh
    
            Command8.Visible = True
            Command8.Caption = "Beenden"
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "StadaUndProUpdates"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub outLeeren()
    On Error GoTo LOKAL_ERROR
                
    Label16.Caption = "Neue Stammdaten und Programmupdates werden gesucht"
    Label16.Refresh
    
    sHosti = gsStammFTPAdresse
    sUseri = gsStammFTPUSER
    sPassi = gsStammFTPPASS

    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    
    Pause (1)
    
    If WechsleInsUnterverzFTP("OUT") = False Then
        Exit Sub
    End If
    
    Pause (1)
    
    WechsleInsUnterverzLOKAL gsKinPfad & "\"
    
    Pause (1)
    
    UebertrageAllVonFTP
    
    Pause (1)
    
    DelAllFTP
    
    Pause (1)
    

    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"

     Verbindungtrennen

    If gbFTPautomatic = False Then
        Pause (1)
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh

        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "outLeeren"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Mailtraffic()
    On Error GoTo LOKAL_ERROR
    
    Dim sDabaMailPfad As String
    Dim sDabaMailoutPfad As String
    
    sDabaMailPfad = gcDBPfad               'Datenbankpfad
    If Right(sDabaMailPfad, 1) <> "\" Then
        sDabaMailPfad = sDabaMailPfad & "\"
    End If
    sDabaMailPfad = sDabaMailPfad & "Mail\"
    
    sDabaMailoutPfad = gcDBPfad               'Datenbankpfad
    If Right(sDabaMailoutPfad, 1) <> "\" Then
        sDabaMailoutPfad = sDabaMailoutPfad & "\"
    End If
    sDabaMailoutPfad = sDabaMailoutPfad & "Mailout\"
                
    Label16.Caption = "Neue Emails werden gesucht..."
    Label16.Refresh
    
    sHosti = gsStammFTPAdresse
    sUseri = gsStammFTPUSER
    sPassi = gsStammFTPPASS
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    'Mails vom FTP Server holen
    
    Pause (1)
    
    If WechsleInsUnterverzFTP("Mail") = False Then
        Exit Sub
    End If
    
   
    Pause (1)
    WechsleInsUnterverzLOKAL sDabaMailPfad
    Pause (1)
    UebertrageAlleMailsVonFTP
    Pause (3)
    DelAllFTP
    Pause (1)
    
    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"
    
    Label16.Caption = "Programmfehler werden versendet..."
    Label16.Refresh
    
    Pause (1)
    
    sHosti = gsStammFTPAdresse
    sUseri = gsStammFTPUSER
    sPassi = gsStammFTPPASS
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    
    'Mails rausschicken
   
    If WechsleInsUnterverzFTP("Mailin") = False Then
        Exit Sub
    End If
    
    Pause (1)
    WechsleInsUnterverzLOKAL sDabaMailoutPfad
    Pause (1)
    UebertrageAllVonLOKAL
    
    Pause (1)
    DelAllLOKAL
    
    Label16.Caption = "KISSNET... wird getrennt..."
    Label16.Refresh

    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"

    Verbindungtrennen

    If gbFTPautomatic = False Then
        Pause (1)
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh

        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Mailtraffic"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Mailtraffic1()
    On Error GoTo LOKAL_ERROR
    
    Dim sDabaMailoutPfad As String
    
    sDabaMailoutPfad = gcDBPfad               'Datenbankpfad
    If Right(sDabaMailoutPfad, 1) <> "\" Then
        sDabaMailoutPfad = sDabaMailoutPfad & "\"
    End If
    sDabaMailoutPfad = sDabaMailoutPfad & "Mailout\"
                

    Label16.Caption = "Programmfehler werden versendet..."
    Label16.Refresh
    
    Pause (1)
    
    sHosti = gsStammFTPAdresse
    sUseri = gsStammFTPUSER
    sPassi = gsStammFTPPASS
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    
    'Mails rausschicken
   
    If WechsleInsUnterverzFTP("Mailin") = False Then
        Exit Sub
    End If
    
    Pause (1)
    WechsleInsUnterverzLOKAL sDabaMailoutPfad
    Pause (1)
    UebertrageAllVonLOKAL
    
    Pause (1)
    DelAllLOKAL
    
    Label16.Caption = "KISSNET... wird getrennt..."
    Label16.Refresh

    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"

    Verbindungtrennen

    If gbFTPautomatic = False Then
        Pause (1)
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh

        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Mailtraffic1"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub UebertrageAllVonFTP() 'nur für
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim sfilename       As String
    Dim iMaxi           As Integer
    Dim sSQL            As String
    
    'selectiere Datei/en
    iMaxi = lfnrErmitteln("Y")
    
    For i = 1 To lvRemote.ListItems.Count
    
next1:
        
        sfilename = lvRemote.ListItems.Item(i)
        
        glFilesizeFTP = 0
        
        If UCase(Left(sfilename, 1)) = "D" Then
        
        Else
            Select Case UCase(Left(sfilename, 2))
                
                Case Is = "MA"
                    If giKissFtpMode = 1 Or giKissFtpMode = 10 Or giKissFtpMode = 3 Or giKissFtpMode = 4 Then
                        
                        glFilesizeFTP = Val(lvRemote.ListItems.Item(i).SubItems(1))
                        
                        lvRemote.ListItems.Item(i).Selected = True
                        Frame3.Visible = True
                        If Check2(0).Visible = False Then
                            Check2(0).Visible = True
                            Check2(0).Caption = "Stammdaten werden übertragen..."
                            Check2(0).value = vbChecked
                        End If
                    End If
                Case Is = "WK"
                    If giKissFtpMode = 1 Or giKissFtpMode = 10 Or giKissFtpMode = 3 Or giKissFtpMode = 4 Then
                        
                        glFilesizeFTP = Val(lvRemote.ListItems.Item(i).SubItems(1))
                        
                        lvRemote.ListItems.Item(i).Selected = True
                        Frame3.Visible = True
                        If Check2(1).Visible = False Then
                            Check2(1).Visible = True
                            Check2(1).Caption = "Programmupdate wird übertragen..."
                            Check2(1).value = vbChecked
                            gsUpdDatName = sfilename
                        Else
                            gsUpdDatName = sfilename
                        End If
                    End If
                    
    
                Case Is = "Y0"
                    If giKissFtpMode = 8 Or giKissFtpMode = 9 Or giKissFtpMode = 10 Then
                    
                        glFilesizeFTP = Val(lvRemote.ListItems.Item(i).SubItems(1))
                        
                        If Val(Mid(sfilename, 2, 7)) > iMaxi Then
                        
                            lvRemote.ListItems.Item(i).Selected = True
                            Frame3.Visible = True
                            If Check2(2).Visible = False Then
                                Check2(2).Visible = True
                                Check2(2).Caption = "Kassendateien werden übertragen..."
                                Check2(2).value = vbChecked
                            End If
                            sSQL = "Update DBEINSTE Set FILKASDAT = true "
                            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
                            gbFilKasDat = True
                        Else
                            If i < lvRemote.ListItems.Count Then
                                i = i + 1
                                GoTo next1
                            Else
                                Exit For
                            End If
                        End If
                    End If
    
                Case Else
                glFilesizeFTP = Val(lvRemote.ListItems.Item(i).SubItems(1))
                lvRemote.ListItems.Item(i).Selected = True
            End Select
                
            'Klicke auf den Übertragungsbutton
            If glFilesizeFTP > 0 Then
                cmdRetrieve_Click
                Pause (3)
            End If
             
            If glFilesizeFTP > 0 Then
                glFilesizeLOKAL = 0
                glFilesizeLOKAL = FileLen(txtCurPath.Text & sfilename)
                If glFilesizeLOKAL = glFilesizeFTP Then
                
                Else
                
                   Kill txtCurPath.Text & sfilename
                   
                   If lerrorZaehler > 2 Then
                       Exit For
                   Else
                       lerrorZaehler = lerrorZaehler + 1
                       Pause (3)
                       cmdConnect.Caption = "Connect"
                       cmdStart_Click
                       
                   End If
                End If
            End If
                
        End If
    Next i
    
    
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 35600 Then
        Exit Sub
    ElseIf err.Number = 53 Then
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "UebertrageAllVonFTP"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        Fehlermeldung1
    End If
End Sub
Private Sub UebertrageAllVonFTPwv() 'nur für
    On Error GoTo LOKAL_ERROR
    
    
    Dim i               As Integer
    Dim sfilename       As String
    Dim cFil As String
    
    Dim sSQL            As String
    
    
    
    For i = 1 To lvRemote.ListItems.Count
    
        sfilename = lvRemote.ListItems.Item(i)
    
        cFil = Trim(gcFilNr)
        If Len(cFil) = 1 Then
            cFil = "0" & cFil
        End If
        
        Select Case UCase(Left(sfilename, 4))
        
            Case Is = "WV" & cFil
                lvRemote.ListItems.Item(i).Selected = True
                
        End Select
        
    Next i

    'Klicke auf den Übertragungsbutton
    cmdRetrieve_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertrageAllVonFTPwv"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub UebertrageAllVonFTP_Budni(sKUNDNR As String) 'nur für
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim sfilename       As String
    

    For i = 1 To lvRemote.ListItems.Count
    
        sfilename = lvRemote.ListItems.Item(i)
        
        sKUNDNR = String$(10 - Len(sKUNDNR), "0") & sKUNDNR
    
    
        
        Select Case Left(sfilename, 28)
        
            Case Is = "BUDNIDESADV_Kunde " & sKUNDNR
                lvRemote.ListItems.Item(i).Selected = True
                
        End Select
        
    Next i

    'Klicke auf den Übertragungsbutton
    cmdRetrieve_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertrageAllVonFTP_Budni"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub UebertrageAllVonFTP_Lüning(sKUNDNR As String) 'nur für
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim sfilename       As String
    

    For i = 1 To lvRemote.ListItems.Count
    
        sfilename = lvRemote.ListItems.Item(i)
        
        sKUNDNR = String$(6 - Len(sKUNDNR), "0") & sKUNDNR
    
    
        
        Select Case Left(sfilename, 7)
        
            Case Is = "A" & sKUNDNR
                lvRemote.ListItems.Item(i).Selected = True
                
        End Select
        
    Next i

    'Klicke auf den Übertragungsbutton
    cmdRetrieve_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertrageAllVonFTP_Lüning"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub UebertrageAllVonFTP_Bela() 'nur für
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim sfilename       As String
    
    Dim sKennNr As String
    sKennNr = ermKenn_fromBela
    

    For i = 1 To lvRemote.ListItems.Count
    
        sfilename = lvRemote.ListItems.Item(i)
        
        If UCase(sfilename) = "KISSURLAD." & sKennNr Then
        
            lvRemote.ListItems.Item(i).Selected = True
        
        End If
    Next i

    'Klicke auf den Übertragungsbutton
    cmdRetrieve_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertrageAllVonFTP_Bela"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub UebertrageAllVonFTPEndzip() 'nur für
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim sfilename       As String
    
    For i = 1 To lvRemote.ListItems.Count
        sfilename = lvRemote.ListItems.Item(i)
        
        Select Case UCase(Left(sfilename, 3))
        
            Case Is = "END"
                lvRemote.ListItems.Item(i).Selected = True
        End Select
    Next i

    'Klicke auf den Übertragungsbutton
    cmdRetrieve_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertrageAllVonFTPEndzip"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub UebertrageAllVonFTPLagerdat() 'nur für
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim sfilename       As String
    
    For i = 1 To lvRemote.ListItems.Count
        sfilename = lvRemote.ListItems.Item(i)
        
        Select Case UCase(Left(sfilename, 1))
        
            Case Is = "F"
                lvRemote.ListItems.Item(i).Selected = True
        End Select
    Next i

    'Klicke auf den Übertragungsbutton
    cmdRetrieve_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertrageAllVonFTPLagerdat"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub UebertrageAllVonFTPwvN() 'nur für
    On Error GoTo LOKAL_ERROR
    
    
    Dim i               As Integer
    Dim sfilename       As String
    Dim cFil As String
    
    Dim sSQL            As String
    
    
    
    For i = 1 To lvRemote.ListItems.Count
    
        sfilename = lvRemote.ListItems.Item(i)
    
        cFil = Trim(gcFilNr)
        If Len(cFil) = 1 Then
            cFil = "0" & cFil
        End If
        
        Select Case UCase(Left(sfilename, 3))
        
            Case Is = "N" & cFil
                lvRemote.ListItems.Item(i).Selected = True
                
        End Select
        
    Next i

    'Klicke auf den Übertragungsbutton
    cmdRetrieve_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertrageAllVonFTPwvN"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub UebertrageAllVonFTPKissLief() 'nur für
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim sfilename       As String
    
    For i = 1 To lvRemote.ListItems.Count
        sfilename = lvRemote.ListItems.Item(i)
    
        If Len(sfilename) >= 8 Then
            If Val(Mid(sfilename, 2, 6)) = glLiNr Then
                lvRemote.ListItems.Item(i).Selected = True
            End If
        End If
    Next i

    'Klicke auf den Übertragungsbutton
    cmdRetrieve_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertrageAllVonFTPKissLief"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub UebertrageAllVonFTPKissMast() 'nur für
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim sfilename       As String
    Dim cWoche          As String
    
    For j = giWochendat To 53
    
        cWoche = j
        If Len(cWoche) = 1 Then
            cWoche = "0" & cWoche
        End If
        
        For i = 1 To lvRemote.ListItems.Count
            sfilename = lvRemote.ListItems.Item(i)
        
            If Mid(sfilename, 10, 2) = cWoche Then
                lvRemote.ListItems.Item(i).Selected = True
            
                eintragenWochendatei j, sfilename
            End If
        Next i
    Next j

    'Klicke auf den Übertragungsbutton
    cmdRetrieve_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertrageAllVonFTPKissMast"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub UebertrageAllVonFTPTagesSTAMMDA() 'nur für
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim sfilename       As String
    
    For i = 1 To lvRemote.ListItems.Count
        sfilename = lvRemote.ListItems.Item(i)
    
        If UCase(Mid(sfilename, 2, 7)) = "AKTUELL" Then
            lvRemote.ListItems.Item(i).Selected = True
        End If
    Next i

    'Klicke auf den Übertragungsbutton
    cmdRetrieve_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertrageAllVonFTPTagesSTAMMDA"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub UebertrageAlleMailsVonFTP() 'nur für Mail abholen
    On Error GoTo LOKAL_ERROR

    Dim i               As Integer
    
    For i = 1 To lvRemote.ListItems.Count
        lvRemote.ListItems.Item(i).Selected = True
    Next i
    
    'Klicke auf den Übertragungsbutton
    cmdRetrieve_Click

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertrageAlleMailsVonFTP"
    Fehler.gsFehlertext = "Im Programmteil KISSNET... Datenübertragung ist ein Fehler aufgetreten."
    Fehlermeldung1

End Sub
Private Sub UebertrageAllVonLOKAL() 'nur für
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    
    'selectiere Datei/en

    For i = 1 To lvLocal.ListItems.Count
        lvLocal.ListItems.Item(i).Selected = True
    Next i
    
    'Klicke auf den Übertragungsbutton
    cmdPut_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertrageAllVonLOKAL"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub UebertrageAllVonLOKAL_Kassendateien() 'nur für
    On Error GoTo LOKAL_ERROR
    
    Dim i           As Integer
    Dim cdatei      As String
    Dim cDatum      As String
    Dim czeit       As String
    
    'selectiere Datei/en

    For i = 1 To lvLocal.ListItems.Count
        lvLocal.ListItems.Item(i).Selected = True
        
        If gbErrPrint Then
            If Left(lvLocal.ListItems.Item(i), 1) = "F" And Right(UCase(lvLocal.ListItems.Item(i)), 3) = "LZH" Then
            'dann drucken wir, wenn es erwünscht ist pro Datei
            
                cdatei = lvLocal.ListItems.Item(i)
                ReDim cZeilen(0 To 6) As String
                
                cDatum = DateValue(Now)
                czeit = TimeValue(Now)
                
                'Drucke den Beleg
        
                cZeilen(0) = "Dateiübertragung"
                cZeilen(1) = "-----------------"
                cZeilen(2) = "Kassendatei: " & cdatei
                cZeilen(3) = "wurde übertragen"
                cZeilen(4) = ""
                cZeilen(5) = "Datum: " & cDatum
                cZeilen(6) = "Zeit:  " & czeit
                
                DruckeArbeitszeitBelegWK20d cZeilen(), 6
            End If
        End If
    Next i
    
    'Klicke auf den Übertragungsbutton
    cmdPut_Click
    
    
    
    
    
    
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebertrageAllVonLOKAL_Kassendateien"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub bestellungabschicken()
    On Error GoTo LOKAL_ERROR
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    Pause (1)
    sHosti = FTP_Server
    sUseri = FTP_User
    sPassi = FTP_PassW
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    
    Pause (1)
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    txtPattern.Text = "senden.*"
    Pause (1)
    WechsleInsUnterverzLOKAL "C:\"
    Pause (1)
    UebertrageAllVonLOKAL
    Pause (1)
    gbErfolg = True
    
'    gbErfolg = False
'    If CheckDateiVorhanden(gsdat) Then
'        Label16.Caption = "Die Datei " & gsdat & " wurde erfolgreich übertragen."
'        Label16.Refresh
'        Pause (3)
'        gbErfolg = True
'    Else
'        Label16.Caption = "Die Datei " & gsdat & " wurde nicht übertragen."
'        Label16.Refresh
'        Pause (3)
'    End If

'    Label16.Caption = "FTP - Server wird getrennt..."
'    Label16.Refresh
    
    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"
    
    Verbindungtrennen 'Dfü trennen
'    Pause (1)
    
    If gbErfolg = True Then
        Label16.Caption = "Ihre Bestellung wurde erfolgreich übertragen."
        Label16.Refresh
    Else
        Label16.Caption = "Ihre Bestellung konnte nicht übertragen werden. Versuchen Sie es zu einem späteren Zeitpunkt nochmal."
        Label16.Refresh
    End If
    Command8.Visible = True
    Command8.Caption = "Beenden"
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "bestellungabschicken"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub KassendateienvonZentrale()
    On Error GoTo LOKAL_ERROR
                
    Label16.Caption = "Neue Kassendateien werden gesucht"
    Label16.Refresh
    
    sHosti = gsZenFTPAdresse
    sUseri = gsZenFTPUSER
    sPassi = gsZenFTPPASS
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    
    Pause (1)
    
    If WechsleInsUnterverzFTP("ZENOUT") = False Then
        Exit Sub
    End If
    Pause (1)
    
    WechsleInsUnterverzLOKAL gsKinPfad & "\"
    
    Pause (1)
    
    UebertrageAllVonFTP
    
'    Label16.Caption = "FTP - Server wird getrennt..."
'    Label16.Refresh
    
    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"
    
    If giKissFtpMode <> 9 And giKissFtpMode <> 10 Then
        Verbindungtrennen
        
        If gbFTPautomatic = False Then
            Pause (1)
        
            Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
            Label16.Refresh
            
            Command8.Visible = True
            Command8.Caption = "Beenden"
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KassendateienvonZentrale"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Lagerdateienabholen()
    On Error GoTo LOKAL_ERROR
                
    Label16.Caption = "Neue Kassendateien werden gesucht"
    Label16.Refresh
    
    sHosti = gsStammFTPAdresse
    sUseri = gsStammFTPUSER
    sPassi = gsStammFTPPASS
    
'    sHosti = gsZenFTPAdresse
'    sUseri = gsZenFTPUSER
'    sPassi = gsZenFTPPASS
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    
    Pause (3)
    
    If WechsleInsUnterverzFTP("ZENIN") = False Then
        Exit Sub
    End If
    Pause (3)
    
    WechsleInsUnterverzLOKAL gsKinPfad & "\"
    
    Pause (3)
    
    UebertrageAllVonFTPLagerdat
    
'    UebertrageAllVonFTP
    
    Pause (5)
    
    DelAllFTP
    
    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
        Pause (1)
    
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Lagerdateienabholen"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WarenverteilungvonZentrale()
    On Error GoTo LOKAL_ERROR
                
    Label16.Caption = "Neue Kassendateien werden gesucht"
    Label16.Refresh

    Pause (1)
    
    If CInt(gcFilNr) = 1 Then
        sHosti = gsStammFTPAdresse
        sUseri = gsStammFTPUSER
        sPassi = gsStammFTPPASS
    ElseIf CInt(gcFilNr) > 1 Then
        sHosti = gsZenFTPAdresse
        sUseri = gsZenFTPUSER
        sPassi = gsZenFTPPASS
    End If
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"

    Pause (1)
    If WechsleInsUnterverzFTP("WV") = False Then
        Exit Sub
    End If
    
    
    WechsleInsUnterverzLOKAL gsKinPfad & "\"
   
    UebertrageAllVonFTPwv
    
    DeleinteilfromFTP
    
    
'    Label16.Caption = "FTP - Server wird getrennt..."
'    Label16.Refresh
    
    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"
    

    Verbindungtrennen

    If gbFTPautomatic = False Then
        Pause (1)

        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh

        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If

    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WarenverteilungvonZentrale"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WarenverteilungvonZentraleX()
    On Error GoTo LOKAL_ERROR
                
    Label16.Caption = "Expressdateien werden gesucht"
    Label16.Refresh

    Pause (1)
    
    sHosti = gsZenFTPAdresse
    sUseri = gsZenFTPUSER
    sPassi = gsZenFTPPASS

    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    
    Pause (3)
    If WechsleInsUnterverzFTP("WV") = False Then
        Exit Sub
    End If
    
    Pause (3)
    
    WechsleInsUnterverzLOKAL gsKinPfad & "\"
    
    Pause (3)
   
    UebertrageAllVonFTPwvN
    
    Pause (3)
    
    DeleinteilfromFTPX
    
    Pause (3)
    
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"
    
    If giKissFtpMode <> 9 And giKissFtpMode <> 10 And giKissFtpMode <> 1 Then
        Verbindungtrennen
        
        If gbFTPautomatic = False Then
            Pause (1)
        
            Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
            Label16.Refresh
            
            Command8.Visible = True
            Command8.Caption = "Beenden"
        End If
    End If
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WarenverteilungvonZentraleX"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LieferantenStamdaholen()
    On Error GoTo LOKAL_ERROR
                
    Label16.Caption = "Lieferantenstammdaten werden gesucht"
    Label16.Refresh

    Pause (1)
    
    sHosti = gsStammFTPAdresse
    sUseri = "kisslie1" 'gsStammFTPUSER
    sPassi = "sd45UC2" 'gsStammFTPPASS
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    
    Pause (3)
    If WechsleInsUnterverzFTP("IN") = False Then
        Exit Sub
    End If
    
    Pause (3)
    
    WechsleInsUnterverzLOKAL gsKinPfad & "\"
    
    Pause (3)
   
    UebertrageAllVonFTPKissLief
    
    Pause (3)
    
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"
    
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
        Pause (1)
    
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LieferantenStamdaholen"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WochenStamdaholen()
    On Error GoTo LOKAL_ERROR
                
    Label16.Caption = "Wochenstammdaten werden gesucht"
    Label16.Refresh

    Pause (1)
    
    sHosti = gsStammFTPAdresse
    sUseri = "kissmas1" 'gsStammFTPUSER
    sPassi = "sd45UC2" 'gsStammFTPPASS
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    
    Pause (3)
    If WechsleInsUnterverzFTP("IN") = False Then
        Exit Sub
    End If
    
    Pause (3)
    
    WechsleInsUnterverzLOKAL gsKinPfad & "\"
    
    Pause (3)
   
    UebertrageAllVonFTPKissMast
    
    Pause (3)
    
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"
    
    
    If giKissFtpMode <> 10 And giKissFtpMode <> 1 And giKissFtpMode <> 3 And giKissFtpMode <> 4 Then
        Verbindungtrennen
    End If
    
    If giKissFtpMode <> 10 And giKissFtpMode <> 1 And giKissFtpMode <> 3 And giKissFtpMode <> 4 Then
        If gbFTPautomatic = False Then
            Pause (1)
            Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
            Label16.Refresh
    
            Command8.Visible = True
            Command8.Caption = "Beenden"
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WochenStamdaholen"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub TagesStamdaholen()
    On Error GoTo LOKAL_ERROR
                
    Label16.Caption = "Tagesstammdaten werden gesucht"
    Label16.Refresh

    Pause (1)
    
    sHosti = gsStammFTPAdresse
    sUseri = "kisslie1" 'gsStammFTPUSER
    sPassi = "sd45UC2" ' gsStammFTPPASS
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    
    Pause (3)
    If WechsleInsUnterverzFTP("IN") = False Then
        Exit Sub
    End If
    
    Pause (3)
    
    WechsleInsUnterverzLOKAL gsKinPfad & "\"
    
    Pause (3)
   
    UebertrageAllVonFTPTagesSTAMMDA
    
    Pause (3)
    
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"
    
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
        Pause (1)
    
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TagesStamdaholen"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub UebertrageEinzelDat(sdat As String, sLoTransPfa As String)
    On Error GoTo LOKAL_ERROR
    
    Dim bOk As Boolean
    
    AddToCollection FOP_UPLOAD, sdat, sLoTransPfa, FileLen(sLoTransPfa + sdat)
    bOk = True

    If bFOBusy = False And bOk Then StartFOEinzelDatfromLoToRe sLoTransPfa + sdat
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "UebertrageEinzelDat"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Sub
Private Sub EinzeldatUmbenennenR(sOriName As String)
    On Error GoTo LOKAL_ERROR
    
    If sOriName <> "" Then
        rfFile.RemoteFile = lvRemote.SelectedItem.Text
        rfFile.RenameFile rfConnection, sOriName
        GetStatus
        FillRemoteListView
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EinzeldatUmbenennenR"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub EinzeldatUmbenennenL(sDumminame As String)
    On Error GoTo LOKAL_ERROR



    If sDumminame <> "" Then
        Dim FO As SHFILEOPSTRUCT
        FO.pFrom = sCurPath + lvLocal.SelectedItem
        FO.pTo = sCurPath + sDumminame
        FO.fFlags = FOF_NOCONFIRMATION
        FO.wFunc = FO_RENAME
        SHFileOperation FO

    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EinzeldatUmbenennenL"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Function UebertrageAllesEinzelnvonLokal(sLOPfa As String, sRePfa As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim l               As Integer
    Dim Wert1           As Long
    Dim iMax            As Integer
    Dim sOriName        As String
    Dim sDumminame      As String
    Dim oldpath         As String
    Dim newpath         As String
    Dim lfail           As Long
    Dim lRet            As Long
    Dim sLoTransPfa     As String
    Dim bFound          As Boolean
    
    UebertrageAllesEinzelnvonLokal = True
    bFound = False
    sLoTransPfa = gcDBPfad    'Datenbankpfad
    If Right$(sLoTransPfa, 1) <> "\" Then
        sLoTransPfa = sLoTransPfa & "\"
    End If
    sLoTransPfa = sLoTransPfa & "TRANSOUT\"
    Kill sLoTransPfa & "*.*"
    
    File1.Path = sLOPfa
    File1.Refresh

    For i = 0 To File1.ListCount - 1
    
        Randomize
        Wert1 = 0
        sDumminame = "D"
        sOriName = ""
        Wert1 = Int((9999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99999
        sDumminame = sDumminame & CStr(Wert1)

        sOriName = File1.list(i)

        'Jetzt erste Datei umbenennen und ins Transout kopieren
        oldpath = sLOPfa & File1.list(i)
        newpath = sLoTransPfa & sDumminame
        lRet = CopyFile(oldpath, newpath, lfail)
        
        Pause (1)
        
        UebertrageEinzelDat sDumminame, sLoTransPfa

        For j = 1 To lvRemote.ListItems.Count
            
            If lvRemote.ListItems.Item(j) = sDumminame Then
                bFound = True
                
                lvRemote.ListItems.Item(j).Selected = True

                Pause (9)
                EinzeldatUmbenennenR sOriName
                Pause (9)
                
                For l = 1 To lvRemote.ListItems.Count
                
                    If UCase(lvRemote.ListItems.Item(l)) = UCase(sOriName) Then
                    
                        If gbErrPrint Then
                            Dim cDatum      As String
                            Dim czeit       As String
                            ReDim cZeilen(0 To 6) As String
                            
                            cDatum = DateValue(Now)
                            czeit = TimeValue(Now)
                            
                            'Drucke den Beleg
            
                            cZeilen(0) = "Dateiübertragung"
                            cZeilen(1) = "-----------------"
                            cZeilen(2) = "Diese Datei: " & sOriName
                            cZeilen(3) = "wurde übertragen"
                            cZeilen(4) = ""
                            cZeilen(5) = "Datum: " & cDatum
                            cZeilen(6) = "Zeit:  " & czeit
                            
                            DruckeArbeitszeitBelegWK20d cZeilen(), 6
                        End If
                        
                        Kill sLoTransPfa & sDumminame
                        Kill sLOPfa & sOriName
                        bFound = True
                        Exit For
                        
                    Else
                        bFound = False
                    End If
                Next l
            Else
                lvRemote.ListItems.Item(j).Selected = False
                bFound = False
            End If
        Next j
        
        If bFound = False Then
            UebertrageAllesEinzelnvonLokal = False
            Exit For

        End If

    Next i
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 35600 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "UebertrageAllesEinzelnvonLokal"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Function
Private Function UebertrageAllesEinzelnvonLokal_forBela(sLOPfa As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim l               As Integer
    Dim Wert1           As Long
    Dim iMax            As Integer
    Dim sOriName        As String
    Dim sDumminame      As String
    Dim oldpath         As String
    Dim newpath         As String
    Dim lfail           As Long
    Dim lRet            As Long
    Dim sLoTransPfa     As String
    Dim bFound          As Boolean
    
    UebertrageAllesEinzelnvonLokal_forBela = True
    bFound = False
    sLoTransPfa = gcDBPfad    'Datenbankpfad
    If Right$(sLoTransPfa, 1) <> "\" Then
        sLoTransPfa = sLoTransPfa & "\"
    End If
    sLoTransPfa = sLoTransPfa & "TRANSOUT\"
    Kill sLoTransPfa & "*.*"
    
    File1.Path = sLOPfa
    File1.Refresh

    For i = 0 To File1.ListCount - 1
    
        Randomize
        Wert1 = 0
        sDumminame = "D"
        sOriName = ""
        Wert1 = Int((9999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99999
        sDumminame = sDumminame & CStr(Wert1)

        sOriName = File1.list(i)

        'Jetzt erste Datei umbenennen und ins Transout kopieren
        oldpath = sLOPfa & File1.list(i)
        newpath = sLoTransPfa & sDumminame
        lRet = CopyFile(oldpath, newpath, lfail)
        
        Pause (1)
        
        UebertrageEinzelDat sDumminame, sLoTransPfa

        For j = 1 To lvRemote.ListItems.Count
            
            If lvRemote.ListItems.Item(j) = sDumminame Then
                bFound = True
                
                lvRemote.ListItems.Item(j).Selected = True

                Pause (9)
                EinzeldatUmbenennenR sOriName
                Pause (9)
                
                For l = 1 To lvRemote.ListItems.Count
                
                    If UCase(lvRemote.ListItems.Item(l)) = UCase(sOriName) Then
                    
'''''                        If gbErrPrint Then
'''''                            Dim cDatum      As String
'''''                            Dim czeit       As String
'''''                            ReDim cZeilen(0 To 6) As String
'''''
'''''                            cDatum = DateValue(Now)
'''''                            czeit = TimeValue(Now)
'''''
'''''                            'Drucke den Beleg
'''''
'''''                            cZeilen(0) = "Dateiübertragung"
'''''                            cZeilen(1) = "-----------------"
'''''                            cZeilen(2) = "Diese Datei: " & sOriName
'''''                            cZeilen(3) = "wurde übertragen"
'''''                            cZeilen(4) = ""
'''''                            cZeilen(5) = "Datum: " & cDatum
'''''                            cZeilen(6) = "Zeit:  " & czeit
'''''
'''''                            DruckeArbeitszeitBelegWK20d cZeilen(), 6
'''''                        End If
                        
                        Kill sLoTransPfa & sDumminame
                        Kill sLOPfa & sOriName
                        bFound = True
                        Exit For
                        
                    Else
                        bFound = False
                    End If
                Next l
            Else
                lvRemote.ListItems.Item(j).Selected = False
                bFound = False
            End If
        Next j
        
        If bFound = False Then
            UebertrageAllesEinzelnvonLokal_forBela = False
            Exit For

        End If

    Next i
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 35600 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "UebertrageAllesEinzelnvonLokal_forBela"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Function
Private Function UebertrageAllesEinzelnvonLokal_forRinklin(sLOPfa As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim l               As Integer
    Dim Wert1           As Long
    Dim iMax            As Integer
    Dim sOriName        As String
    Dim sDumminame      As String
    Dim oldpath         As String
    Dim newpath         As String
    Dim lfail           As Long
    Dim lRet            As Long
    Dim sLoTransPfa     As String
    Dim bFound          As Boolean
    
    UebertrageAllesEinzelnvonLokal_forRinklin = True
    bFound = False
    sLoTransPfa = gcDBPfad    'Datenbankpfad
    If Right$(sLoTransPfa, 1) <> "\" Then
        sLoTransPfa = sLoTransPfa & "\"
    End If
    sLoTransPfa = sLoTransPfa & "TRANSOUT\"
    Kill sLoTransPfa & "*.*"
    
    File1.Path = sLOPfa
    File1.Refresh

    For i = 0 To File1.ListCount - 1
    
        Randomize
        Wert1 = 0
        sDumminame = "E"
        sOriName = ""
        Wert1 = Int((9999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99999
        sDumminame = sDumminame & CStr(Wert1)

        sOriName = File1.list(i)

        'Jetzt erste Datei umbenennen und ins Transout kopieren
        oldpath = sLOPfa & File1.list(i)
        newpath = sLoTransPfa & sDumminame
        lRet = CopyFile(oldpath, newpath, lfail)
        
        Pause (1)
        
        UebertrageEinzelDat sDumminame, sLoTransPfa

        For j = 1 To lvRemote.ListItems.Count
            
            If lvRemote.ListItems.Item(j) = sDumminame Then
                bFound = True
                
                lvRemote.ListItems.Item(j).Selected = True

                Pause (9)
                EinzeldatUmbenennenR sOriName
                Pause (9)
                
                For l = 1 To lvRemote.ListItems.Count
                
                    If UCase(lvRemote.ListItems.Item(l)) = UCase(sOriName) Then
                        
                        Kill sLoTransPfa & sDumminame
                        Kill sLOPfa & sOriName
                        bFound = True
                        Exit For
                        
                    Else
                        bFound = False
                    End If
                Next l
            Else
                lvRemote.ListItems.Item(j).Selected = False
                bFound = False
            End If
        Next j
        
        If bFound = False Then
            UebertrageAllesEinzelnvonLokal_forRinklin = False
            Exit For

        End If

    Next i
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 35600 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "UebertrageAllesEinzelnvonLokal_forRinklin"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Function
Private Function UebertrageAllesEinzelnvonLokal_forMenson(sLOPfa As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim l               As Integer
    Dim Wert1           As Long
    Dim iMax            As Integer
    Dim sOriName        As String
    Dim sDumminame      As String
    Dim oldpath         As String
    Dim newpath         As String
    Dim lfail           As Long
    Dim lRet            As Long
    Dim sLoTransPfa     As String
    Dim bFound          As Boolean
    
    UebertrageAllesEinzelnvonLokal_forMenson = True
    bFound = False
    sLoTransPfa = gcDBPfad    'Datenbankpfad
    If Right$(sLoTransPfa, 1) <> "\" Then
        sLoTransPfa = sLoTransPfa & "\"
    End If
    sLoTransPfa = sLoTransPfa & "TRANSOUT\"
    Kill sLoTransPfa & "*.*"
    
    File1.Path = sLOPfa
    File1.Refresh

    For i = 0 To File1.ListCount - 1
    
        Randomize
        Wert1 = 0
        sDumminame = "E"
        sOriName = ""
        Wert1 = Int((9999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99999
        sDumminame = sDumminame & CStr(Wert1)

        sOriName = File1.list(i)

        'Jetzt erste Datei umbenennen und ins Transout kopieren
        oldpath = sLOPfa & File1.list(i)
        newpath = sLoTransPfa & sDumminame
        lRet = CopyFile(oldpath, newpath, lfail)
        
        Pause (1)
        
        UebertrageEinzelDat sDumminame, sLoTransPfa

        For j = 1 To lvRemote.ListItems.Count
            
            If lvRemote.ListItems.Item(j) = sDumminame Then
                bFound = True
                
                lvRemote.ListItems.Item(j).Selected = True

                Pause (9)
                EinzeldatUmbenennenR sOriName
                Pause (9)
                
                For l = 1 To lvRemote.ListItems.Count
                
                    If UCase(lvRemote.ListItems.Item(l)) = UCase(sOriName) Then
                        
                        Kill sLoTransPfa & sDumminame
                        Kill sLOPfa & sOriName
                        bFound = True
                        Exit For
                        
                    Else
                        bFound = False
                    End If
                Next l
            Else
                lvRemote.ListItems.Item(j).Selected = False
                bFound = False
            End If
        Next j
        
        If bFound = False Then
            UebertrageAllesEinzelnvonLokal_forMenson = False
            Exit For

        End If

    Next i
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 35600 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "UebertrageAllesEinzelnvonLokal_forMenson"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Function
Private Function UebertrageAllesEinzelnvonLokal_forBudni(sLOPfa As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim l               As Integer
    Dim Wert1           As Long
    Dim iMax            As Integer
    Dim sOriName        As String
    Dim sDumminame      As String
    Dim oldpath         As String
    Dim newpath         As String
    Dim lfail           As Long
    Dim lRet            As Long
    Dim sLoTransPfa     As String
    Dim bFound          As Boolean
    
    UebertrageAllesEinzelnvonLokal_forBudni = True
    bFound = False
    sLoTransPfa = gcDBPfad    'Datenbankpfad
    If Right$(sLoTransPfa, 1) <> "\" Then
        sLoTransPfa = sLoTransPfa & "\"
    End If
    sLoTransPfa = sLoTransPfa & "TRANSOUT\"
    Kill sLoTransPfa & "*.*"
    
    File1.Path = sLOPfa
    File1.Refresh

    For i = 0 To File1.ListCount - 1
    
        Randomize
        Wert1 = 0
        sDumminame = "D"
        sOriName = ""
        Wert1 = Int((9999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99999
        
        sOriName = File1.list(i)
        sDumminame = sDumminame & CStr(Wert1) & sOriName

'        sOriName = File1.list(i)

        'Jetzt erste Datei umbenennen und ins Transout kopieren
        oldpath = sLOPfa & File1.list(i)
        newpath = sLoTransPfa & sDumminame
        lRet = CopyFile(oldpath, newpath, lfail)
        
        Pause (1)
        
        UebertrageEinzelDat sDumminame, sLoTransPfa

        For j = 1 To lvRemote.ListItems.Count
            
            If lvRemote.ListItems.Item(j) = sDumminame Then
                bFound = True
                
                lvRemote.ListItems.Item(j).Selected = True

                Pause (9)
                EinzeldatUmbenennenR sOriName
                Pause (9)
                
                For l = 1 To lvRemote.ListItems.Count
                
                    If UCase(lvRemote.ListItems.Item(l)) = UCase(sOriName) Then
                    
'''''                        If gbErrPrint Then
'''''                            Dim cDatum      As String
'''''                            Dim czeit       As String
'''''                            ReDim cZeilen(0 To 6) As String
'''''
'''''                            cDatum = DateValue(Now)
'''''                            czeit = TimeValue(Now)
'''''
'''''                            'Drucke den Beleg
'''''
'''''                            cZeilen(0) = "Dateiübertragung"
'''''                            cZeilen(1) = "-----------------"
'''''                            cZeilen(2) = "Diese Datei: " & sOriName
'''''                            cZeilen(3) = "wurde übertragen"
'''''                            cZeilen(4) = ""
'''''                            cZeilen(5) = "Datum: " & cDatum
'''''                            cZeilen(6) = "Zeit:  " & czeit
'''''
'''''                            DruckeArbeitszeitBelegWK20d cZeilen(), 6
'''''                        End If
                        
                        Kill sLoTransPfa & sDumminame
                        Kill sLOPfa & sOriName
                        gbBudni_Bestellung_erfolgreich = True
                        bFound = True
                        Exit For
                        
                    Else
                        bFound = False
                    End If
                Next l
            Else
                lvRemote.ListItems.Item(j).Selected = False
                bFound = False
            End If
        Next j
        
        If bFound = False Then
            UebertrageAllesEinzelnvonLokal_forBudni = False
            Exit For

        End If

    Next i
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 35600 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "UebertrageAllesEinzelnvonLokal_forBudni"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Function
Private Function UebertrageAllesEinzelnvonLokal_forLuening(sLOPfa As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim l               As Integer
    Dim Wert1           As Long
    Dim iMax            As Integer
    Dim sOriName        As String
    Dim sDumminame      As String
    Dim oldpath         As String
    Dim newpath         As String
    Dim lfail           As Long
    Dim lRet            As Long
    Dim sLoTransPfa     As String
    Dim bFound          As Boolean
    
    UebertrageAllesEinzelnvonLokal_forLuening = True
    bFound = False
    sLoTransPfa = gcDBPfad    'Datenbankpfad
    If Right$(sLoTransPfa, 1) <> "\" Then
        sLoTransPfa = sLoTransPfa & "\"
    End If
    sLoTransPfa = sLoTransPfa & "TRANSOUT\"
    Kill sLoTransPfa & "*.*"
    
    File1.Path = sLOPfa
    File1.Refresh

    For i = 0 To File1.ListCount - 1
    
        Randomize
        Wert1 = 0
        sDumminame = "D"
        sOriName = ""
        Wert1 = Int((9999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99999
        sDumminame = sDumminame & CStr(Wert1)

        sOriName = File1.list(i)

        'Jetzt erste Datei umbenennen und ins Transout kopieren
        oldpath = sLOPfa & File1.list(i)
        newpath = sLoTransPfa & sDumminame
        lRet = CopyFile(oldpath, newpath, lfail)
        
        Pause (1)
        
        UebertrageEinzelDat sDumminame, sLoTransPfa

        For j = 1 To lvRemote.ListItems.Count
            
            If lvRemote.ListItems.Item(j) = sDumminame Then
                bFound = True
                
                lvRemote.ListItems.Item(j).Selected = True

                Pause (9)
                EinzeldatUmbenennenR sOriName
                Pause (9)
                
                For l = 1 To lvRemote.ListItems.Count
                
                    If UCase(lvRemote.ListItems.Item(l)) = UCase(sOriName) Then
                    
'''''                        If gbErrPrint Then
'''''                            Dim cDatum      As String
'''''                            Dim czeit       As String
'''''                            ReDim cZeilen(0 To 6) As String
'''''
'''''                            cDatum = DateValue(Now)
'''''                            czeit = TimeValue(Now)
'''''
'''''                            'Drucke den Beleg
'''''
'''''                            cZeilen(0) = "Dateiübertragung"
'''''                            cZeilen(1) = "-----------------"
'''''                            cZeilen(2) = "Diese Datei: " & sOriName
'''''                            cZeilen(3) = "wurde übertragen"
'''''                            cZeilen(4) = ""
'''''                            cZeilen(5) = "Datum: " & cDatum
'''''                            cZeilen(6) = "Zeit:  " & czeit
'''''
'''''                            DruckeArbeitszeitBelegWK20d cZeilen(), 6
'''''                        End If
                        
                        Kill sLoTransPfa & sDumminame
                        Kill sLOPfa & sOriName
                        bFound = True
                        Exit For
                        
                    Else
                        bFound = False
                    End If
                Next l
            Else
                lvRemote.ListItems.Item(j).Selected = False
                bFound = False
            End If
        Next j
        
        If bFound = False Then
            UebertrageAllesEinzelnvonLokal_forLuening = False
            Exit For

        End If

    Next i
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 35600 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "UebertrageAllesEinzelnvonLokal_forLuening"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Function
Private Function UebertrageAllesEinzelnvonLokal_forRewe(sLOPfa As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim l               As Integer
    Dim Wert1           As Long
    Dim iMax            As Integer
    Dim sOriName        As String
    Dim sDumminame      As String
    Dim oldpath         As String
    Dim newpath         As String
    Dim lfail           As Long
    Dim lRet            As Long
    Dim sLoTransPfa     As String
    Dim bFound          As Boolean
    
    UebertrageAllesEinzelnvonLokal_forRewe = True
    bFound = False
    sLoTransPfa = gcDBPfad    'Datenbankpfad
    If Right$(sLoTransPfa, 1) <> "\" Then
        sLoTransPfa = sLoTransPfa & "\"
    End If
    sLoTransPfa = sLoTransPfa & "TRANSOUT\"
    Kill sLoTransPfa & "*.*"
    
    File1.Path = sLOPfa
    File1.Refresh

    For i = 0 To File1.ListCount - 1
    
        Randomize
        Wert1 = 0
        sDumminame = "D"
        sOriName = ""
        Wert1 = Int((9999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99999
        sDumminame = sDumminame & CStr(Wert1)

        sOriName = File1.list(i)

        'Jetzt erste Datei umbenennen und ins Transout kopieren
        oldpath = sLOPfa & File1.list(i)
        newpath = sLoTransPfa & sDumminame
        lRet = CopyFile(oldpath, newpath, lfail)
        
        Pause (1)
        
        UebertrageEinzelDat sDumminame, sLoTransPfa

        For j = 1 To lvRemote.ListItems.Count
            
            If lvRemote.ListItems.Item(j) = sDumminame Then
                bFound = True
                
                lvRemote.ListItems.Item(j).Selected = True

                Pause (9)
                EinzeldatUmbenennenR sOriName
                Pause (9)
                
                For l = 1 To lvRemote.ListItems.Count
                
                    If UCase(lvRemote.ListItems.Item(l)) = UCase(sOriName) Then
                    
'''''                        If gbErrPrint Then
'''''                            Dim cDatum      As String
'''''                            Dim czeit       As String
'''''                            ReDim cZeilen(0 To 6) As String
'''''
'''''                            cDatum = DateValue(Now)
'''''                            czeit = TimeValue(Now)
'''''
'''''                            'Drucke den Beleg
'''''
'''''                            cZeilen(0) = "Dateiübertragung"
'''''                            cZeilen(1) = "-----------------"
'''''                            cZeilen(2) = "Diese Datei: " & sOriName
'''''                            cZeilen(3) = "wurde übertragen"
'''''                            cZeilen(4) = ""
'''''                            cZeilen(5) = "Datum: " & cDatum
'''''                            cZeilen(6) = "Zeit:  " & czeit
'''''
'''''                            DruckeArbeitszeitBelegWK20d cZeilen(), 6
'''''                        End If
                        
                        Kill sLoTransPfa & sDumminame
                        Kill sLOPfa & sOriName
                        bFound = True
                        Exit For
                        
                    Else
                        bFound = False
                    End If
                Next l
            Else
                lvRemote.ListItems.Item(j).Selected = False
                bFound = False
            End If
        Next j
        
        If bFound = False Then
            UebertrageAllesEinzelnvonLokal_forRewe = False
            Exit For

        End If

    Next i
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 35600 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "UebertrageAllesEinzelnvonLokal_forRewe"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Function
Private Function UebertrageAllesEinzelnvonLokal_forPural(sLOPfa As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim l               As Integer
    Dim Wert1           As Long
    Dim iMax            As Integer
    Dim sOriName        As String
    Dim sDumminame      As String
    Dim oldpath         As String
    Dim newpath         As String
    Dim lfail           As Long
    Dim lRet            As Long
    Dim sLoTransPfa     As String
    Dim bFound          As Boolean
    
    UebertrageAllesEinzelnvonLokal_forPural = True
    bFound = False
    sLoTransPfa = gcDBPfad    'Datenbankpfad
    If Right$(sLoTransPfa, 1) <> "\" Then
        sLoTransPfa = sLoTransPfa & "\"
    End If
    sLoTransPfa = sLoTransPfa & "TRANSOUT\"
    Kill sLoTransPfa & "*.*"
    
    File1.Path = sLOPfa
    File1.Refresh

    For i = 0 To File1.ListCount - 1
    
        Randomize
        Wert1 = 0
        sDumminame = "D"
        sOriName = ""
        Wert1 = Int((9999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99999
        sDumminame = sDumminame & CStr(Wert1)

        sOriName = File1.list(i)

        'Jetzt erste Datei umbenennen und ins Transout kopieren
        oldpath = sLOPfa & File1.list(i)
        newpath = sLoTransPfa & sDumminame
        lRet = CopyFile(oldpath, newpath, lfail)
        
        Pause (1)
        
        UebertrageEinzelDat sDumminame, sLoTransPfa

        For j = 1 To lvRemote.ListItems.Count
            
            If lvRemote.ListItems.Item(j) = sDumminame Then
                bFound = True
                
                lvRemote.ListItems.Item(j).Selected = True

                Pause (9)
                EinzeldatUmbenennenR sOriName
                Pause (9)
                
                For l = 1 To lvRemote.ListItems.Count
                
                    If UCase(lvRemote.ListItems.Item(l)) = UCase(sOriName) Then
                        
                        Kill sLoTransPfa & sDumminame
                        Kill sLOPfa & sOriName
                        bFound = True
                        Exit For
                        
                    Else
                        bFound = False
                    End If
                Next l
            Else
                lvRemote.ListItems.Item(j).Selected = False
                bFound = False
            End If
        Next j
        
        If bFound = False Then
            UebertrageAllesEinzelnvonLokal_forPural = False
            Exit For

        End If

    Next i
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 35600 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "UebertrageAllesEinzelnvonLokal_forPural"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Function
Private Function UebertrageAllesEinzelnvonLokal_forCarnot(sLOPfa As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim l               As Integer
    Dim Wert1           As Long
    Dim iMax            As Integer
    Dim sOriName        As String
    Dim sDumminame      As String
    Dim oldpath         As String
    Dim newpath         As String
    Dim lfail           As Long
    Dim lRet            As Long
    Dim sLoTransPfa     As String
    Dim bFound          As Boolean
    
    UebertrageAllesEinzelnvonLokal_forCarnot = True
    bFound = False
    sLoTransPfa = gcDBPfad    'Datenbankpfad
    If Right$(sLoTransPfa, 1) <> "\" Then
        sLoTransPfa = sLoTransPfa & "\"
    End If
    sLoTransPfa = sLoTransPfa & "TRANSOUT\"
    Kill sLoTransPfa & "*.*"
    
    File1.Path = sLOPfa
    File1.Refresh

    For i = 0 To File1.ListCount - 1
    
        Randomize
        Wert1 = 0
        sDumminame = "D"
        sOriName = ""
        Wert1 = Int((9999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99999
        sDumminame = sDumminame & CStr(Wert1)

        sOriName = File1.list(i)

        'Jetzt erste Datei umbenennen und ins Transout kopieren
        oldpath = sLOPfa & File1.list(i)
        newpath = sLoTransPfa & sDumminame
        lRet = CopyFile(oldpath, newpath, lfail)
        
        Pause (1)
        
        UebertrageEinzelDat sDumminame, sLoTransPfa

        For j = 1 To lvRemote.ListItems.Count
            
            If lvRemote.ListItems.Item(j) = sDumminame Then
                bFound = True
                
                lvRemote.ListItems.Item(j).Selected = True

                Pause (9)
                EinzeldatUmbenennenR sOriName
                Pause (9)
                
                For l = 1 To lvRemote.ListItems.Count
                
                    If UCase(lvRemote.ListItems.Item(l)) = UCase(sOriName) Then
                        
                        Kill sLoTransPfa & sDumminame
                        Kill sLOPfa & sOriName
                        bFound = True
                        Exit For
                        
                    Else
                        bFound = False
                    End If
                Next l
            Else
                lvRemote.ListItems.Item(j).Selected = False
                bFound = False
            End If
        Next j
        
        If bFound = False Then
            UebertrageAllesEinzelnvonLokal_forCarnot = False
            Exit For

        End If

    Next i
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 35600 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "UebertrageAllesEinzelnvonLokal_forCarnot"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Function
Private Function UebertrageAllesEinzelnvonLokal_forDronovaCouponEinL(sLOPfa As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim l               As Integer
    Dim Wert1           As Long
    Dim iMax            As Integer
    Dim sOriName        As String
    Dim sDumminame      As String
    Dim oldpath         As String
    Dim newpath         As String
    Dim lfail           As Long
    Dim lRet            As Long
    Dim sLoTransPfa     As String
    Dim bFound          As Boolean
    
    UebertrageAllesEinzelnvonLokal_forDronovaCouponEinL = True
    bFound = False
    sLoTransPfa = gcDBPfad    'Datenbankpfad
    If Right$(sLoTransPfa, 1) <> "\" Then
        sLoTransPfa = sLoTransPfa & "\"
    End If
    sLoTransPfa = sLoTransPfa & "TRANSOUT\"
    Kill sLoTransPfa & "*.*"
    
    File1.Path = sLOPfa
    File1.Refresh

    For i = 0 To File1.ListCount - 1
    
        Randomize
        Wert1 = 0
        sDumminame = "A"
        sOriName = ""
        Wert1 = Int((9999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99999
        sDumminame = sDumminame & CStr(Wert1)

        sOriName = File1.list(i)

        'Jetzt erste Datei umbenennen und ins Transout kopieren
        oldpath = sLOPfa & File1.list(i)
        newpath = sLoTransPfa & sDumminame
        lRet = CopyFile(oldpath, newpath, lfail)
        
        Pause (1)
        
        UebertrageEinzelDat sDumminame, sLoTransPfa

        For j = 1 To lvRemote.ListItems.Count
            
            If lvRemote.ListItems.Item(j) = sDumminame Then
                bFound = True
                
                lvRemote.ListItems.Item(j).Selected = True

                Pause (3)
                EinzeldatUmbenennenR sOriName
                Pause (3)
                
                For l = 1 To lvRemote.ListItems.Count
                
                    If UCase(lvRemote.ListItems.Item(l)) = UCase(sOriName) Then
                        
                        Kill sLoTransPfa & sDumminame
                        Kill sLOPfa & sOriName
                        bFound = True
                        Exit For
                        
                    Else
                        bFound = False
                    End If
                Next l
            Else
                lvRemote.ListItems.Item(j).Selected = False
                bFound = False
            End If
        Next j
        
        If bFound = False Then
            UebertrageAllesEinzelnvonLokal_forDronovaCouponEinL = False
            Exit For

        End If

    Next i
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 35600 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "UebertrageAllesEinzelnvonLokal_forDronovaCouponEinL"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Function
Private Function UebertrageAllesEinzelnvonLokal_forBiogarten(sLOPfa As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim l               As Integer
    Dim Wert1           As Long
    Dim iMax            As Integer
    Dim sOriName        As String
    Dim sDumminame      As String
    Dim oldpath         As String
    Dim newpath         As String
    Dim lfail           As Long
    Dim lRet            As Long
    Dim sLoTransPfa     As String
    Dim bFound          As Boolean
    
    UebertrageAllesEinzelnvonLokal_forBiogarten = True
    bFound = False
    sLoTransPfa = gcDBPfad    'Datenbankpfad
    If Right$(sLoTransPfa, 1) <> "\" Then
        sLoTransPfa = sLoTransPfa & "\"
    End If
    sLoTransPfa = sLoTransPfa & "TRANSOUT\"
    Kill sLoTransPfa & "*.*"
    
    File1.Path = sLOPfa
    File1.Refresh

    For i = 0 To File1.ListCount - 1
    
        Randomize
        Wert1 = 0
        sDumminame = "D"
        sOriName = ""
        Wert1 = Int((9999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99999
        sDumminame = sDumminame & CStr(Wert1)

        sOriName = File1.list(i)

        'Jetzt erste Datei umbenennen und ins Transout kopieren
        oldpath = sLOPfa & File1.list(i)
        newpath = sLoTransPfa & sDumminame
        lRet = CopyFile(oldpath, newpath, lfail)
        
        Pause (1)
        
        UebertrageEinzelDat sDumminame, sLoTransPfa

        For j = 1 To lvRemote.ListItems.Count
            
            If lvRemote.ListItems.Item(j) = sDumminame Then
                bFound = True
                
                lvRemote.ListItems.Item(j).Selected = True

                Pause (9)
                EinzeldatUmbenennenR sOriName
                Pause (9)
                
                For l = 1 To lvRemote.ListItems.Count
                
                    If UCase(lvRemote.ListItems.Item(l)) = UCase(sOriName) Then
                        
                        Kill sLoTransPfa & sDumminame
                        Kill sLOPfa & sOriName
                        bFound = True
                        Exit For
                        
                    Else
                        bFound = False
                    End If
                Next l
            Else
                lvRemote.ListItems.Item(j).Selected = False
                bFound = False
            End If
        Next j
        
        If bFound = False Then
            UebertrageAllesEinzelnvonLokal_forBiogarten = False
            Exit For

        End If

    Next i
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 35600 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "UebertrageAllesEinzelnvonLokal_forBiogarten"
        Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Function
Private Sub EndZipHol()
    On Error GoTo LOKAL_ERROR
    
    Dim sDabaProtoPfad As String
    
    sDabaProtoPfad = gcDBPfad               'Datenbankpfad
    If Right$(sDabaProtoPfad, 1) <> "\" Then
        sDabaProtoPfad = sDabaProtoPfad & "\"
    End If
    sDabaProtoPfad = sDabaProtoPfad & "ENDZIPIN\"
                
    Label16.Caption = "Neue Datenbankdatei wird gesucht"
    Label16.Refresh

    sHosti = gsZenFTPAdresse
    sUseri = gsZenFTPUSER
    sPassi = gsZenFTPPASS
    
    cmdConnect_Click 'Ftp connect

    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    Pause (3)
    WechsleInsUnterverzFTP "ENDZIP"
    Pause (3)
    WechsleInsUnterverzLOKAL sDabaProtoPfad
    Pause (3)
    UebertrageAllVonFTPEndzip
    
'    Pause (6)
'    DelAllFTP
    Pause (3)
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
        Pause (1)
    
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EndZipHol"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub EndZipweg()
On Error GoTo LOKAL_ERROR

    Dim sDabaProtoPfad As String
    
    sDabaProtoPfad = gcDBPfad               'Datenbankpfad
    If Right$(sDabaProtoPfad, 1) <> "\" Then
        sDabaProtoPfad = sDabaProtoPfad & "\"
    End If
    sDabaProtoPfad = sDabaProtoPfad & "ENDZIP\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = gsStammFTPAdresse
    sUseri = gsStammFTPUSER
    sPassi = gsStammFTPPASS
    
    cmdConnect_Click 'Ftp connect

    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    WechsleInsUnterverzFTP "ENDZIP"
    
    Pause (3)
    
    WechsleInsUnterverzLOKAL sDabaProtoPfad
    
    Pause (3)
    
    DelAllFTP
    
    Pause (3)
    
    UebertrageAllVonLOKAL
    
    Pause (3)
    
    DelAllLOKAL
    
    Pause (3)
    
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
        Pause (1)
    
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EndZipweg"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub

Private Function KassendateienanZentrale() As Boolean
    On Error GoTo LOKAL_ERROR

    KassendateienanZentrale = False

    Dim ierrz As Integer
    ierrz = 0

    Label16.Caption = "Kassendateien werden versendet..."
    Label16.Refresh

    Pause (1)

    sHosti = gsZenFTPAdresse
    sUseri = gsZenFTPUSER
    sPassi = gsZenFTPPASS

    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"

    Pause (1)
    WechsleInsUnterverzLOKAL gsZoutPfad & "\"
    Pause (1)
    WechsleInsUnterverzFTP "LIVEIN"
    Pause (1)
    UebertrageAllVonLOKAL
    
    cmdConnect_Click 'Ftp disconnect
    
    
    'danach wieder verbinden um auf höchster Ebene einzusteigen
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    
    Pause (1)
    WechsleInsUnterverzFTP "ZENIN"
    Pause (1)
    UebertrageAllVonLOKAL_Kassendateien
    Pause (1)
    DelAllLOKAL
    Pause (1)
    
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    

    If giKissFtpMode <> 10 Then

        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
        
        Command8_Click

    End If
    
    KassendateienanZentrale = True

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KassendateienanZentrale"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function

Private Sub WarenverteilungenanZentrale()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    Pause (1)
    cPfad = cPfad & "WVOUT\"
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh

    If CInt(gcFilNr) = 1 Then
        sHosti = gsStammFTPAdresse
        sUseri = gsStammFTPUSER
        sPassi = gsStammFTPPASS
    ElseIf CInt(gcFilNr) > 1 Then
        sHosti = gsZenFTPAdresse
        sUseri = gsZenFTPUSER
        sPassi = gsZenFTPPASS
    End If
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    Pause (1)
    
    If WechsleInsUnterverzFTP("WV") = False Then
        Exit Sub
    End If
    
    Pause (1)
    WechsleInsUnterverzLOKAL cPfad
    Pause (1)
    
    UebertrageAllVonLOKAL
    Pause (1)
    DelAllLOKAL
    Pause (1)
    
'    Label16.Caption = "FTP - Server wird getrennt..."
'    Label16.Refresh
    
    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"
    

    Verbindungtrennen
    
    If gbFTPautomatic = False Then
        Pause (1)
    
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WarenverteilungenanZentrale"
    
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub WarenverteilungenanZentraleX()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    Pause (1)
    cPfad = cPfad & "WVOUT\"
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh

'    If CInt(gcFilNr) = 1 Then
'        sHosti = gsStammFTPAdresse
'        sUseri = gsStammFTPUSER
'        sPassi = gsStammFTPPASS
'    ElseIf CInt(gcFilNr) > 1 Then

        sHosti = gsZenFTPAdresse
        sUseri = gsZenFTPUSER
        sPassi = gsZenFTPPASS
        
'    End If
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    Pause (2)
    
    If WechsleInsUnterverzFTP("WV") = False Then
        Exit Sub
    End If
    
    Pause (2)
    WechsleInsUnterverzLOKAL cPfad
    Pause (2)
    
    UebertrageAllVonLOKAL
    Pause (2)
    DelAllLOKAL
    Pause (2)
    
'    Label16.Caption = "FTP - Server wird getrennt..."
'    Label16.Refresh
    
    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"
    

    Verbindungtrennen
    
    If gbFTPautomatic = False Then
        Pause (1)
    
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WarenverteilungenanZentraleX"
    
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
    

Private Sub statleeren()
On Error GoTo LOKAL_ERROR

    Dim sDabaStatPfad As String
    
    sDabaStatPfad = gcDBPfad               'Datenbankpfad
    If Right(sDabaStatPfad, 1) <> "\" Then
        sDabaStatPfad = sDabaStatPfad & "\"
    End If
    sDabaStatPfad = sDabaStatPfad & "STAT\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    Pause (1)
    
    sHosti = gsStammFTPAdresse
    sUseri = gsStammFTPUSER
    sPassi = gsStammFTPPASS
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"

    Pause (2)
    
    If WechsleInsUnterverzFTP("IN") = False Then
        Exit Sub
    End If
    
    Pause (2)
    WechsleInsUnterverzLOKAL sDabaStatPfad
    Pause (2)
    
    UebertrageAllVonLOKAL
    Pause (2)
    DelAllLOKAL
    Pause (2)
    

'    Label16.Caption = "FTP - Server wird getrennt..."
'    Label16.Refresh
    
    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"
    
    If giKissFtpMode <> 7 Then
        Verbindungtrennen
        Pause (1)
        If gbFTPautomatic = False Then
            Pause (1)
        
            Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
            Label16.Refresh
            
            Command8.Visible = True
            Command8.Caption = "Beenden"
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "statleeren"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Protokolleeren()
On Error GoTo LOKAL_ERROR

    Label16.Caption = "Protokolle werden übertragen..."
    Label16.Refresh
    
    sHosti = gsStammFTPAdresse
    sUseri = gsStammFTPUSER
    sPassi = gsStammFTPPASS
    
    cmdConnect.Caption = "Connect"
    cmdConnect_Click 'Ftp connect
    cmdConnect.Caption = "Disconnect"
    
    Pause (1)

    Dim sDabaProtoPfad As String
    
    sDabaProtoPfad = gcDBPfad               'Datenbankpfad
    If Right(sDabaProtoPfad, 1) <> "\" Then
        sDabaProtoPfad = sDabaProtoPfad & "\"
    End If
    sDabaProtoPfad = sDabaProtoPfad & "Protok\"
    
  
    If WechsleInsUnterverzFTP("Protokol") = False Then
        Exit Sub
    End If
    
    Pause (1)
    WechsleInsUnterverzLOKAL sDabaProtoPfad
    Pause (1)
    UebertrageAllVonLOKAL
    Pause (1)
    DelAllLOKAL
    

    Pause (1)
    
    cmdConnect.Caption = "Disconnect"
    cmdConnect_Click 'Ftp disconnect
    cmdConnect.Caption = "Connect"
    
    If giKissFtpMode <> 10 And giKissFtpMode <> 1 And giKissFtpMode <> 3 And giKissFtpMode <> 4 Then
        Verbindungtrennen
        Pause (1)
        If gbFTPautomatic = False Then
            Pause (1)
        
            Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
            Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
            Label16.Refresh
            
            Command8.Visible = True
            Command8.Caption = "Beenden"
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Protokolleeren"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ReweZuUns()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
'    sHosti = "80.86.85.121"
'    sUseri = "rewe"
'    sPassi = "a1cqwv"
    
    sHosti = "212.117.66.101"
    sUseri = "kisswws"
    sPassi = "kisswws2016()"
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    WechsleInsUnterverzLOKAL sBestPfad
    Pause (2)
    
    
    WechsleInsUnterverzFTP "ORDER"
    Pause (2)
    
    
    
    UebertrageAllesEinzelnvonLokal_forRewe sBestPfad
'''    UebertrageAllVonLOKAL
    Pause (1)
'''    WechsleInsUnterverzFTP "ABRECH"
'''    Pause (1)
'''    UebertrageAllVonLOKAL
'''    Pause (1)
'''    WechsleInsUnterverzFTP "SIC"
'''    Pause (1)
'''    UebertrageAllVonLOKAL
'''    Pause (1)
'''
'''    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ReweZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub PuralZuUns()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = "80.86.85.121"
    sUseri = "pural"
    sPassi = "R5Ghgw3"
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    WechsleInsUnterverzLOKAL sBestPfad
    Pause (1)
    UebertrageAllesEinzelnvonLokal_forPural sBestPfad

    Pause (1)

    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PuralZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub CarnotZuUns()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = "80.86.85.121"
    sUseri = "carnot"
    sPassi = "Ui62Wes"
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    WechsleInsUnterverzLOKAL sBestPfad
    Pause (1)
    UebertrageAllesEinzelnvonLokal_forCarnot sBestPfad

    Pause (1)

    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "CarnotZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub DronovCouponEinlZuUns()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = "80.86.85.121"
    sUseri = "drocou2"
    sPassi = "W_6GtBxP"
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden......"
    Label16.Refresh
    
    Pause (1)
    
    WechsleInsUnterverzLOKAL sBestPfad
    Pause (2)
    

    Label16.Caption = "Daten werden übertragen..."
    Label16.Refresh
    UebertrageAllVonLOKAL
    
    Pause (2)

    Label16.Caption = "Daten werden übertragen......"
    Label16.Refresh
    DelAllLOKAL
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
''''    UebertrageAllesEinzelnvonLokal_forDronovaCouponEinL sBestPfad

    Pause (1)

    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DronovCouponEinlZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub BiogartenZuUns()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = "www.biooffice.de"
    sUseri = "Bi012954"
    sPassi = "mairuebe"
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    WechsleInsUnterverzLOKAL sBestPfad
    Pause (1)
    UebertrageAllesEinzelnvonLokal_forBiogarten sBestPfad

    Pause (1)

    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BiogartenZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub LueningZuUns()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = "80.86.85.121"
    sUseri = "lueningb"
    sPassi = "stada"
    
'    sUseri = "heinz"
'    sPassi = "stada"
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    

    Pause (3)
    WechsleInsUnterverzLOKAL sBestPfad
    
    Pause (3)
    UebertrageAllesEinzelnvonLokal_forLuening sBestPfad
    
    
    
''    UebertrageAllVonLOKAL
''    Pause (1)
''    WechsleInsUnterverzFTP "ABRECH"
''    Pause (1)
''    UebertrageAllVonLOKAL
''    Pause (1)
''    WechsleInsUnterverzFTP "SIC"
''    Pause (1)
''    UebertrageAllVonLOKAL
    Pause (1)
    
''    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Bestellung ist erfolgt." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LueningZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub RFSZuUns()
    Dim sIPPfad As String
    
    sIPPfad = gcDBPfad             'IPpfad
    If Right$(sIPPfad, 1) <> "\" Then
        sIPPfad = sIPPfad & "\"
    End If
    sIPPfad = sIPPfad & "IP\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = "80.86.85.121"
    sUseri = "ipausw"
    sPassi = "r6g1ll9"
   
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    Pause (1)

    WechsleInsUnterverzLOKAL sIPPfad
    Pause (2)
    UebertrageAllVonLOKAL
''    Pause (1)
''    WechsleInsUnterverzFTP "ABRECH"
''    Pause (1)
''    UebertrageAllVonLOKAL
''    Pause (1)
''    WechsleInsUnterverzFTP "SIC"
''    Pause (1)
''    UebertrageAllVonLOKAL
    Pause (2)
    
    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    Pause (2)
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "RFSZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub VEDESZuUns()
    Dim sVedesPfad As String
    
    sVedesPfad = gcDBPfad             'IPpfad
    If Right$(sVedesPfad, 1) <> "\" Then
        sVedesPfad = sVedesPfad & "\"
    End If
    sVedesPfad = sVedesPfad & "VEDES\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    LeseVedesFTP
    
    If gsVEDES_HOST <> "" And gsVEDES_USER <> "" And gsVEDES_PW <> "" Then
         sHosti = gsVEDES_HOST
         sUseri = gsVEDES_USER
         sPassi = gsVEDES_PW
    Else
         sHosti = "80.86.85.121"
         sUseri = "vedesau"
         sPassi = "ft783hj7"
    End If
    
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    Pause (4)

    WechsleInsUnterverzLOKAL sVedesPfad
    Pause (4)
    
    WechsleInsUnterverzFTP "Datenservice"
    
    Pause (4)
    WechsleInsUnterverzFTP "User"
    
    Pause (4)
    WechsleInsUnterverzFTP "M0" & gsVEDES_USER & "00"
    
    Pause (4)
    WechsleInsUnterverzFTP "POSDM"
    
    Pause (4)
    
    UebertrageAllVonLOKAL
''    Pause (1)
    
''    Pause (1)
''    UebertrageAllVonLOKAL
''    Pause (1)
''    WechsleInsUnterverzFTP "SIC"
''    Pause (1)
''    UebertrageAllVonLOKAL
    Pause (4)
    
    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    Pause (4)
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "VEDESZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub BelaBestellungen()
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
 
    sHosti = "217.70.204.69"
    sUseri = "9812283"
    sPassi = "98#12283"
    
'    sHosti = "80.86.85.121"
'    sUseri = "bbela"
'    sPassi = "stada"
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    Pause (1)
    
    WechsleInsUnterverzLOKAL sBestPfad
    
    Pause (1)
    UebertrageAllesEinzelnvonLokal_forBela sBestPfad
'    UebertrageAllVonLOKAL

    Pause (1)
    
'    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BelaBestellungen"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub RinklinBestellungen()
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    
    'hinterlegt in LISRT mit Format EDIRINKLIN
    
    
 
    sHosti = ermXWert_fromEDIRINKLIN("adress") '"80.86.85.121"
    sUseri = ermXWert_fromEDIRINKLIN("bUser") '"heinz2"
    sPassi = ermXWert_fromEDIRINKLIN("Pass") '"stada"
    
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    Pause (1)
    
    WechsleInsUnterverzLOKAL sBestPfad
    
    Pause (1)
    UebertrageAllesEinzelnvonLokal_forRinklin sBestPfad
'    UebertrageAllVonLOKAL

    Pause (1)
    
'    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "RinklinBestellungen"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub MensonBestellungen()
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    
    'hinterlegt in LISRT mit Format EDIRINKLIN
    
    
 
    sHosti = ermXWert_fromEDIMENSON("adress") '"80.86.85.121"
    sUseri = ermXWert_fromEDIMENSON("bUser") '"heinz2"
    sPassi = ermXWert_fromEDIMENSON("Pass") '"stada"
    
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    Pause (1)
    
    WechsleInsUnterverzLOKAL sBestPfad
    
    Pause (1)
    UebertrageAllesEinzelnvonLokal_forRinklin sBestPfad
'    UebertrageAllVonLOKAL

    Pause (1)
    
'    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MensonBestellungen"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub BoerBestellungen()
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    
    'hinterlegt in LISRT mit Format EDIBOER
    
    sHosti = ermXWert_fromEDIBOER("adress") '"80.86.85.121"
    sUseri = ermXWert_fromEDIBOER("bUser") '"heinz2"
    sPassi = ermXWert_fromEDIBOER("Pass") '"stada"
    
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    Pause (1)
    
    WechsleInsUnterverzLOKAL sBestPfad
    
    Pause (1)
    UebertrageAllesEinzelnvonLokal_forRinklin sBestPfad
'    UebertrageAllVonLOKAL

    Pause (1)
    
'    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BoerBestellungen"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub BudniBestellungen()
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
 
    sHosti = "ftp.budni.de"
    sUseri = "Dronova_Ernst"
    sPassi = "Budn!3rnst"
    
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    Pause (1)
    
    WechsleInsUnterverzLOKAL sBestPfad
    
    Pause (3)
    WechsleInsUnterverzFTP "PXI_BESTELL" 'scharf ist immer PXI 'Test ist EXI
    Pause (3)
    UebertrageAllesEinzelnvonLokal_forBudni sBestPfad


    Pause (1)
    
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BudniBestellungen"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub BudniLieferavis()
    On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim sBudniKundnr    As String
    Dim rsLi            As DAO.Recordset
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sBudniKundnr = ""
    
    sSQL = "select KUNDNR from LISRT where FORMAT = 'EDIBUDNI'"
    Set rsLi = gdBase.OpenRecordset(sSQL)
    If Not rsLi.EOF Then
        sBudniKundnr = Trim(rsLi!Kundnr)
    End If
    rsLi.Close: Set rsLi = Nothing
    
    If Val(sBudniKundnr) = 0 Then
        Exit Sub
    End If

'    sHosti = "ftp.budni.de"
'    sUseri = "Dronova_Ernst"
'    sPassi = "Budn!3rnst"
    
    sHosti = "ftp.budni.de"
    sUseri = "Dronova_Brnas"
    sPassi = "Budn!jger16"
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    Pause (1)
    
    WechsleInsUnterverzLOKAL gsKinPfad & "\"
    
    
    Pause (3)
    
    Label16.Caption = "nach Lieferscheinen suchen..."
    Label16.Refresh
    
    WechsleInsUnterverzFTP "PXI_LIEFERAVIS"
    Pause (3)
    UebertrageAllVonFTP_Budni sBudniKundnr
    
    
    
''    Label16.Caption = "Lieferscheine löschen..."
''    Label16.Refresh
''
''    Pause (3)
''
    'lassen wir auf dem Server, so kann KISSLIVE diese zur Rechnungserstellung nutzen
'    DeleinteilfromFTP_Budni sBudniKundnr
    
    Pause (3)
    


    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BudniLieferavis"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Lüning_Stada_holen()
    On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim sLüningKundnr   As String
    Dim rsLi            As DAO.Recordset
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sLüningKundnr = ""
    
    sSQL = "select KUNDNR from LISRT where FORMAT = 'EDILUENING'"
    Set rsLi = gdBase.OpenRecordset(sSQL)
    If Not rsLi.EOF Then
        sLüningKundnr = Trim(rsLi!Kundnr)
    End If
    rsLi.Close: Set rsLi = Nothing
    
    If Val(sLüningKundnr) = 0 Then
        Exit Sub
    End If

    sHosti = "80.86.85.121"
    sUseri = "luening"
    sPassi = "stada"
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    Pause (1)
    
    WechsleInsUnterverzLOKAL gsKinPfad & "\"
    
    
    Pause (3)
    
    Label16.Caption = "nach Stammdaten suchen..."
    Label16.Refresh
    
    WechsleInsUnterverzFTP "IN"
    Pause (3)
    
    UebertrageAllVonFTP_Lüning sLüningKundnr
    
    
    
    Label16.Caption = "Stammdaten löschen..."
    Label16.Refresh

    Pause (3)
    
    
    
    DeleinteilfromFTP_Lüning sLüningKundnr
    
    Pause (3)
    


    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Lüning_Stada_holen"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Bela_Stada_holen()
    On Error GoTo LOKAL_ERROR
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    

    sHosti = "80.86.85.121"
    sUseri = "bela"
    sPassi = "stada"
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    Pause (1)
    
    WechsleInsUnterverzLOKAL gsKinPfad & "\"
    
    
    Pause (3)
    
    Label16.Caption = "nach Stammdaten suchen..."
    Label16.Refresh
    
    WechsleInsUnterverzFTP "IN"
    Pause (3)
    
    UebertrageAllVonFTP_Bela
    
    
    
    
    
    
    Pause (3)
    


    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Bela_Stada_holen"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub CosparGfkZuUns()
    On Error GoTo LOKAL_ERROR
    
    Dim sGFKPfad As String
    
    sGFKPfad = gcDBPfad             'GFKpfad
    If Right$(sGFKPfad, 1) <> "\" Then
        sGFKPfad = sGFKPfad & "\"
    End If
    sGFKPfad = sGFKPfad & "GFK\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = "80.86.85.121"
    sUseri = "cosgfk"
    sPassi = "u6gw0j"
   
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    

    WechsleInsUnterverzLOKAL sGFKPfad
    Pause (1)
    UebertrageAllVonLOKAL
    Pause (1)
    WechsleInsUnterverzFTP "ABRECH"
    Pause (1)
    UebertrageAllVonLOKAL
    Pause (1)
    WechsleInsUnterverzFTP "SIC"
    Pause (1)
    UebertrageAllVonLOKAL
    Pause (1)
    
    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "CosparGfkZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub VEDES_DSL_ZuVedes()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = gcDBPfad              'Beautypfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "VEDESDSL\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    LeseVedesFTP_DSL
    
    If gsVEDES_HOST_DSL <> "" And gsVEDES_USER_DSL <> "" And gsVEDES_PW_DSL <> "" Then
         sHosti = gsVEDES_HOST_DSL
         sUseri = gsVEDES_USER_DSL
         sPassi = gsVEDES_PW_DSL
    Else
         sHosti = "80.86.85.121"
         sUseri = "vedesau"
         sPassi = "ft783hj7"
    End If
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    

    WechsleInsUnterverzLOKAL sBestPfad
    Pause (5)
    
    WechsleInsUnterverzFTP "in"
    Pause (5)
    
    UebertrageAllVonLOKAL
    Pause (1)
    DelAllLOKAL
    
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "VEDES_DSL_ZuVedes"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub JedeBestZuUns()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = gsZenFTPAdresse
    sUseri = gsZenFTPUSER
    sPassi = gsZenFTPPASS
    
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    WechsleInsUnterverzLOKAL sBestPfad
    Pause (1)
    WechsleInsUnterverzFTP "IN"
    Pause (1)
    UebertrageAllVonLOKAL
    Pause (1)
    
    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "JedeBestZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub CotyZuUns()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = "80.86.85.121"
    sUseri = "coty"
    sPassi = "khjd523ikavb"
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
'    WechsleInsUnterverzFTP "BESTELL"
    WechsleInsUnterverzLOKAL sBestPfad
    
    UebertrageAllVonLOKAL
    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "CotyZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub BiedroZuUns()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = "80.86.85.121"
    sUseri = "biedro"
    sPassi = "uo92gh49d"
    cmdConnect_Click 'Ftp connect
    
    Pause (3)
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
'    WechsleInsUnterverzFTP "BESTELL"
    WechsleInsUnterverzLOKAL sBestPfad
    
    Pause (3)
    
    UebertrageAllVonLOKAL
    
    Pause (3)
    
    DelAllLOKAL
    
    Pause (3)
    
    
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Bestellung ist erfolgt." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BiedroZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub LOREALZuUns()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = "80.86.85.121"
    sUseri = "kiss"
    sPassi = "34df8k"
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
'    WechsleInsUnterverzFTP "BESTELL"
    WechsleInsUnterverzLOKAL sBestPfad
    
    UebertrageAllVonLOKAL
    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Bestellung ist erfolgt." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LOREALZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub ErnstBestellungen()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = "80.86.85.121"
    sUseri = "er_best"
    sPassi = "7umv9p"
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
'    WechsleInsUnterverzFTP "BESTELL"
    WechsleInsUnterverzLOKAL sBestPfad
    
    UebertrageAllVonLOKAL
    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErnstBestellungen"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub BBIZuUns()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = App.Path               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "EDI\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = "80.86.85.121"
    sUseri = "bbi"
    sPassi = "5k9cMA"
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
'    WechsleInsUnterverzFTP "BESTELL"
    WechsleInsUnterverzLOKAL sBestPfad
    
    UebertrageAllVonLOKAL
    DelAllLOKAL
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BBIZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub FTZuUns()
    On Error GoTo LOKAL_ERROR
    
    Dim sBestPfad As String
    
    sBestPfad = gcDBPfad               'Bestellpfad
    If Right$(sBestPfad, 1) <> "\" Then
        sBestPfad = sBestPfad & "\"
    End If
    sBestPfad = sBestPfad & "XML\"
    
    Label16.Caption = "Übertragung wird vorbereitet..."
    Label16.Refresh
    
    sHosti = gsZenFTPAdresse
    sUseri = gsZenFTPUSER
    sPassi = gsZenFTPPASS
    
'    sHosti = gsStammFTPAdresse
'    sUseri = gsStammFTPUSER
'    sPassi = gsStammFTPPASS
    cmdConnect_Click 'Ftp connect
    
    Label16.Caption = "FTP - Server wird verbunden..."
    Label16.Refresh
    
    WechsleInsUnterverzFTP "EXPORT"
    
    WechsleInsUnterverzLOKAL sBestPfad
    
    UebertrageAllVonLOKAL
    
    DelAllLOKAL
    
    gbErfolg = True
    Label16.Caption = "FTP - Server wird getrennt..."
    Label16.Refresh
    
    cmdConnect_Click 'Ftp disconnect
    
    Verbindungtrennen
    
    If gbFTPautomatic = False Then
    
        Pause (1)
        
        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
        Label16.Refresh
        
        Command8.Visible = True
        Command8.Caption = "Beenden"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FTZuUns"
    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
''''Private Sub ReweZuUns()
''''    On Error GoTo LOKAL_ERROR
''''    Dim sDabaBestPfad As String
''''
''''    sDabaBestPfad = gcDBPfad               'Datenbankpfad
''''    If Right(sDabaBestPfad, 1) <> "\" Then
''''        sDabaBestPfad = sDabaBestPfad & "\"
''''    End If
''''    sDabaBestPfad = sDabaBestPfad & "Bestell\"
''''
''''    Label16.Caption = "Übertragung wird vorbereitet..."
''''    Label16.Refresh
''''
''''    sHosti = gsStammFTPAdresse
''''    sUseri = gsStammFTPUSER
''''    sPassi = gsStammFTPPASS
''''
''''    cmdConnect.Caption = "Connect"
''''    cmdConnect_Click 'Ftp connect
''''    cmdConnect.Caption = "Disconnect"
''''
''''    Pause 1
''''    If WechsleInsUnterverzFTP("BESTELL") = False Then
''''
''''        Exit Sub
''''    End If
''''
''''    WechsleInsUnterverzLOKAL sDabaBestPfad
''''
''''    UebertrageAllVonLOKAL
''''    DelAllLOKAL
''''    gbErfolg = True
'''''    Label16.Caption = "FTP - Server wird getrennt..."
'''''    Label16.Refresh
''''
''''    cmdConnect.Caption = "Disconnect"
''''    cmdConnect_Click 'Ftp disconnect
''''    cmdConnect.Caption = "Connect"
''''    Verbindungtrennen
''''
''''    If gbFTPautomatic = False Then
''''
''''        Pause (1)
''''
''''        Label16.Caption = "Die Übertragung war fehlerfrei." & vbCrLf & vbCrLf
''''        Label16.Caption = Label16.Caption & "Drücken Sie 'Beenden'!"
''''        Label16.Refresh
''''
''''        Command8.Visible = True
''''        Command8.Caption = "Beenden"
''''    End If
''''
''''Exit Sub
''''LOKAL_ERROR:
''''    Fehler.gsDescr = err.Description
''''    Fehler.gsNumber = err.Number
''''    Fehler.gsFormular = Me.name
''''    Fehler.gsFunktion = "ReweZuUns"
''''    Fehler.gsFehlertext = "Im Programmteil FTP - Datenübertragung ist ein Fehler aufgetreten."
''''
''''    Fehlermeldung1
''''End Sub

