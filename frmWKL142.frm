VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWKL142 
   Caption         =   "Artikel suchen"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL142.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C000&
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
      Height          =   8625
      Left            =   -480
      TabIndex        =   1
      Top             =   1440
      Width           =   11895
      Begin VB.CheckBox Check14 
         BackColor       =   &H00C0C000&
         Caption         =   "ohne Ex"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   98
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   4560
         MaxLength       =   13
         TabIndex        =   82
         Text            =   "Text3"
         Top             =   840
         Width           =   2295
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   2160
         TabIndex        =   20
         Top             =   3600
         Width           =   9375
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   58
            Left            =   8880
            TabIndex        =   79
            Top             =   3000
            Width           =   735
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
            Caption         =   ">>>"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   57
            Left            =   8160
            TabIndex        =   78
            Top             =   3000
            Width           =   735
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
            Caption         =   "<<<"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   29
            Left            =   7440
            TabIndex        =   77
            Top             =   3000
            Width           =   735
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
            Caption         =   " "
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   43
            Left            =   6720
            TabIndex        =   76
            Top             =   3000
            Width           =   735
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
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   44
            Left            =   6000
            TabIndex        =   75
            Top             =   3000
            Width           =   735
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
            Caption         =   ","
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   12
            Left            =   5280
            TabIndex        =   74
            Top             =   3000
            Width           =   735
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
            Caption         =   "M"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   13
            Left            =   4560
            TabIndex        =   73
            Top             =   3000
            Width           =   735
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
            Caption         =   "N"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   1
            Left            =   3840
            TabIndex        =   72
            Top             =   3000
            Width           =   735
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
            Caption         =   "B"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   21
            Left            =   3120
            TabIndex        =   71
            Top             =   3000
            Width           =   735
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
            Caption         =   "V"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   2
            Left            =   2400
            TabIndex        =   70
            Top             =   3000
            Width           =   735
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
            Height          =   735
            Index           =   23
            Left            =   1680
            TabIndex        =   69
            Top             =   3000
            Width           =   735
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
            Caption         =   "X"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   24
            Left            =   960
            TabIndex        =   68
            Top             =   3000
            Width           =   735
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
            Caption         =   "Y"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   42
            Left            =   8520
            TabIndex        =   67
            Top             =   2280
            Width           =   735
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
            Caption         =   "#"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   26
            Left            =   7800
            TabIndex        =   66
            Top             =   2280
            Width           =   735
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
            Caption         =   "Ä"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   27
            Left            =   7080
            TabIndex        =   65
            Top             =   2280
            Width           =   735
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
            Caption         =   "Ö"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   11
            Left            =   6360
            TabIndex        =   64
            Top             =   2280
            Width           =   735
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
            Caption         =   "L"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   10
            Left            =   5640
            TabIndex        =   63
            Top             =   2280
            Width           =   735
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
            Caption         =   "K"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   9
            Left            =   4920
            TabIndex        =   62
            Top             =   2280
            Width           =   735
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
            Caption         =   "J"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   7
            Left            =   4200
            TabIndex        =   61
            Top             =   2280
            Width           =   735
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
            Caption         =   "H"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   6
            Left            =   3480
            TabIndex        =   60
            Top             =   2280
            Width           =   735
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
            Caption         =   "G"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   5
            Left            =   2760
            TabIndex        =   59
            Top             =   2280
            Width           =   735
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
            Caption         =   "F"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   3
            Left            =   2040
            TabIndex        =   58
            Top             =   2280
            Width           =   735
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
            Caption         =   "D"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   18
            Left            =   1320
            TabIndex        =   57
            Top             =   2280
            Width           =   735
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
            Caption         =   "S"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   0
            Left            =   600
            TabIndex        =   56
            Top             =   2280
            Width           =   735
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
            Caption         =   "A"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   41
            Left            =   9000
            TabIndex        =   55
            Top             =   1560
            Width           =   735
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
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   40
            Left            =   8280
            TabIndex        =   54
            Top             =   1560
            Width           =   735
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
            Caption         =   "*"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   28
            Left            =   7560
            TabIndex        =   53
            Top             =   1560
            Width           =   735
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
            Caption         =   "Ü"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   15
            Left            =   6840
            TabIndex        =   52
            Top             =   1560
            Width           =   735
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
            Caption         =   "P"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   14
            Left            =   6120
            TabIndex        =   51
            Top             =   1560
            Width           =   735
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
            Caption         =   "O"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   8
            Left            =   5400
            TabIndex        =   50
            Top             =   1560
            Width           =   735
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
            Caption         =   "I"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   20
            Left            =   4680
            TabIndex        =   49
            Top             =   1560
            Width           =   735
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
            Caption         =   "U"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   25
            Left            =   3960
            TabIndex        =   48
            Top             =   1560
            Width           =   735
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
            Caption         =   "Z"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   19
            Left            =   3240
            TabIndex        =   47
            Top             =   1560
            Width           =   735
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
            Caption         =   "T"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   17
            Left            =   2520
            TabIndex        =   46
            Top             =   1560
            Width           =   735
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
            Caption         =   "R"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   4
            Left            =   1800
            TabIndex        =   45
            Top             =   1560
            Width           =   735
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
            Caption         =   "E"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   22
            Left            =   1080
            TabIndex        =   44
            Top             =   1560
            Width           =   735
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
            Caption         =   "W"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   16
            Left            =   360
            TabIndex        =   43
            Top             =   1560
            Width           =   735
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
            Caption         =   "Q"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   55
            Left            =   7200
            TabIndex        =   42
            Top             =   840
            Width           =   1815
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
            Caption         =   "LEEREN"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   54
            Left            =   6480
            TabIndex        =   41
            Top             =   840
            Width           =   735
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
            Height          =   735
            Index           =   53
            Left            =   5760
            TabIndex        =   40
            Top             =   840
            Width           =   735
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
            Height          =   735
            Index           =   52
            Left            =   5040
            TabIndex        =   39
            Top             =   840
            Width           =   735
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
            Height          =   735
            Index           =   51
            Left            =   4320
            TabIndex        =   38
            Top             =   840
            Width           =   735
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
            Height          =   735
            Index           =   50
            Left            =   3600
            TabIndex        =   37
            Top             =   840
            Width           =   735
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
            Height          =   735
            Index           =   49
            Left            =   2880
            TabIndex        =   36
            Top             =   840
            Width           =   735
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
            Height          =   735
            Index           =   48
            Left            =   2160
            TabIndex        =   35
            Top             =   840
            Width           =   735
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
            Height          =   735
            Index           =   47
            Left            =   1440
            TabIndex        =   34
            Top             =   840
            Width           =   735
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
            Height          =   735
            Index           =   46
            Left            =   720
            TabIndex        =   33
            Top             =   840
            Width           =   735
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
            Height          =   735
            Index           =   45
            Left            =   0
            TabIndex        =   32
            Top             =   840
            Width           =   735
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
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   56
            Left            =   7200
            TabIndex        =   31
            Top             =   120
            Width           =   1815
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
            Caption         =   "RÜCKG"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   39
            Left            =   6480
            TabIndex        =   30
            Top             =   120
            Width           =   735
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
            Caption         =   "?"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   38
            Left            =   5760
            TabIndex        =   29
            Top             =   120
            Width           =   735
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
            Caption         =   "="
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   37
            Left            =   5040
            TabIndex        =   28
            Top             =   120
            Width           =   735
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
            Caption         =   ")"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   36
            Left            =   4320
            TabIndex        =   27
            Top             =   120
            Width           =   735
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
            Caption         =   "("
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   35
            Left            =   3600
            TabIndex        =   26
            Top             =   120
            Width           =   735
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
            Caption         =   "/"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   34
            Left            =   2880
            TabIndex        =   25
            Top             =   120
            Width           =   735
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
            Caption         =   "&&"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   33
            Left            =   2160
            TabIndex        =   24
            Top             =   120
            Width           =   735
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
            Caption         =   "%"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   32
            Left            =   1440
            TabIndex        =   23
            Top             =   120
            Width           =   735
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
            Caption         =   "$"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   31
            Left            =   720
            TabIndex        =   22
            Top             =   120
            Width           =   735
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
            Caption         =   "§"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   735
            Index           =   30
            Left            =   0
            TabIndex        =   21
            Top             =   120
            Width           =   735
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
            Caption         =   "!"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "-1"
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
            Index           =   5
            Left            =   9480
            TabIndex        =   81
            Top             =   1080
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "Zielfeld:"
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
            Index           =   3
            Left            =   9480
            TabIndex        =   80
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   2
         Left            =   10080
         TabIndex        =   19
         Top             =   7920
         Width           =   1695
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   1
         Left            =   8280
         TabIndex        =   18
         Top             =   7920
         Visible         =   0   'False
         Width           =   1695
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
         Caption         =   "Wählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   0
         Left            =   10080
         TabIndex        =   17
         Top             =   480
         Width           =   1695
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
         Caption         =   "S&uchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   4560
         MaxLength       =   13
         TabIndex        =   16
         Text            =   "Text3"
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   120
         MaxLength       =   35
         TabIndex        =   0
         Text            =   "Text3"
         Top             =   840
         Width           =   4335
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   9
         Left            =   3720
         TabIndex        =   15
         Top             =   7920
         Visible         =   0   'False
         Width           =   1695
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
         Caption         =   "Bestand"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   10
         Left            =   120
         TabIndex        =   14
         Top             =   7920
         Visible         =   0   'False
         Width           =   1695
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
         Caption         =   "Info"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikelbezeichnung"
         Height          =   255
         Index           =   0
         Left            =   8760
         TabIndex        =   13
         Top             =   0
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferant, Linie, Artikelbezeichnung"
         Height          =   255
         Index           =   1
         Left            =   8760
         TabIndex        =   12
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   6960
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   1680
         Width           =   2895
      End
      Begin sevCommand3.Command Command66 
         Height          =   360
         Index           =   0
         Left            =   9360
         TabIndex        =   10
         Top             =   1320
         Width           =   495
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
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   120
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   1680
         Width           =   4335
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00C0C000&
         Caption         =   "auch ""Schwarze"""
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   8
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   6960
         MaxLength       =   35
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   840
         Width           =   2895
      End
      Begin sevCommand3.Command Command66 
         Height          =   360
         Index           =   1
         Left            =   9360
         TabIndex        =   6
         Top             =   480
         Width           =   495
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
         Height          =   360
         Index           =   2
         Left            =   3960
         TabIndex        =   5
         Top             =   1320
         Width           =   495
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
         Height          =   360
         Index           =   3
         Left            =   11280
         TabIndex        =   4
         Top             =   1320
         Width           =   495
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
      Begin sevCommand3.Command Command4 
         Height          =   330
         Index           =   15
         Left            =   10080
         TabIndex        =   2
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
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
         Caption         =   "weitere"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin MSComctlLib.TreeView List4 
         Height          =   5295
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   9340
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ListBox List2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   83
         Top             =   2160
         Visible         =   0   'False
         Width           =   11655
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "EAN-Code:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   94
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "LiefBestellNr:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   93
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikelname:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   92
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikel suchen"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   495
         Left            =   120
         TabIndex        =   91
         Top             =   120
         Width           =   6735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Sortierung nach:"
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
         Left            =   7080
         TabIndex        =   90
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "PGN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   6960
         TabIndex        =   89
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "alle Farben:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   9960
         TabIndex        =   88
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferant:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   87
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "Marke:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   6960
         TabIndex        =   86
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "kein Lieferant"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   85
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         Caption         =   "keine Auswahl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   7440
         TabIndex        =   84
         Top             =   1440
         Width           =   1935
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C000&
      Caption         =   "label8(3)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   97
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "label2(1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   96
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "label2(3)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   95
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmWKL142"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bFocusonList4 As Boolean
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    
    iZielIndex = Label3(5).Caption
    If iZielIndex < 0 Then
        Exit Sub
    End If
    
    If iZielIndex = 97 Then
        Select Case Index
            Case 0 To 54, 63 To 117
                Combo1.Text = Combo1.Text & Command0(Index).Caption
                Combo1.SetFocus
            Case 55, 62    'Löschen
                Combo1.Text = ""
                Combo1.SetFocus
                
            Case 56, 61
                If Len(Combo1.Text) > 0 Then
                    Combo1.Text = Left(Combo1.Text, Len(Combo1.Text) - 1)
                End If
                Combo1.SetFocus
                
            Case Is = 57
                Text3(0).SetFocus
            
            Case Is = 58
                Text3(1).SetFocus
                
            Case Is = 59
                Combo1.SetFocus
            
            Case Is = 60
                Combo1.SetFocus
        End Select
    ElseIf iZielIndex = 4 Then
        Select Case Index
            Case 45 To 54
                Text3(iZielIndex).Text = Text3(iZielIndex).Text & Command0(Index).Caption
                Text3(iZielIndex).SetFocus
            Case 55, 62    'Löschen
                Text3(iZielIndex).Text = ""
                Text3(iZielIndex).SetFocus
                
            Case 56, 61
                If Len(Text3(iZielIndex).Text) > 0 Then
                    Text3(iZielIndex).Text = Left(Text3(iZielIndex).Text, Len(Text3(iZielIndex).Text) - 1)
                End If
                Text3(iZielIndex).SetFocus
                
            Case Is = 57 'Zurück
                Text3(1).SetFocus
            
            Case Is = 58 'vor
                Text3(4).SetFocus
                
            Case Is = 59
                Text3(3).SetFocus
            
            Case Is = 60
                Text3(3).SetFocus
            
        End Select

    Else
    
        Select Case Index
            Case 0 To 54, 63 To 117
                Text3(iZielIndex).Text = Text3(iZielIndex).Text & Command0(Index).Caption
                Text3(iZielIndex).SetFocus
            Case 55, 62    'Löschen
                Text3(iZielIndex).Text = ""
                Text3(iZielIndex).SetFocus
                
            Case 56, 61
                If Len(Text3(iZielIndex).Text) > 0 Then
                    Text3(iZielIndex).Text = Left(Text3(iZielIndex).Text, Len(Text3(iZielIndex).Text) - 1)
                End If
                Text3(iZielIndex).SetFocus
                
            Case Is = 57
                Select Case iZielIndex
                    Case 4
                        Text3(1).SetFocus
                    Case 1
                        Text3(5).SetFocus
                    Case 5
                        Text3(6).SetFocus
                    Case 6
                        Text3(2).SetFocus
                    Case 2
                        Text3(0).SetFocus
                End Select
            
            Case Is = 58
                
                Select Case iZielIndex
                    Case 0
                        Text3(2).SetFocus
                    Case 2
                        Text3(6).SetFocus
                    Case 6
                        Text3(5).SetFocus
                    Case 5
                        Text3(1).SetFocus
                    Case 1
                        Text3(4).SetFocus
                End Select
            Case Is = 59
                Combo1.SetFocus
'                Text3(3).SetFocus
            
            Case Is = 60
                Combo1.SetFocus
'                Text3(3).SetFocus
        End Select
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Select Case Index
    Case Is = 0
        anzeige "normal", "Artikel suchen", Label9
        iRet = fnPruefeDialogSucheArtWKL20()
        If iRet = 0 Then
            List4.Visible = False
            List4.Refresh
            List2.Visible = False
            List2.Refresh
            
            Command4(10).Visible = False
            Command4(9).Visible = False
            Command4(1).Visible = False
            
            
            Me.Refresh

            SucheArtikelKasseWKL20
        Else
            MsgBox "Bitte die Suchkriterien erweitern!", vbInformation, "Winkiss Hinweis:"
            Text3(0).SetFocus
        End If
    Case 1 'Wählen
        If List4.SelectedItem Is Nothing Then
            MsgBox "Bitte einen Eintrag in der Liste auswählen!", vbInformation, "Winkiss Hinweis:"
        Else
            cLBSatz = List4.SelectedItem.Text
            gsARTNR = Left(cLBSatz, 6)
            
            Command4_Click 2
        End If
    
    Case 2
        voreinstellungspeichernE142C
        Unload frmWKL142
    Case 9
        If List4.SelectedItem Is Nothing Then
            MsgBox "Bitte einen Eintrag in der Liste auswählen!", vbCritical, "STOP!"
        Else
            cLBSatz = List4.SelectedItem.Text
            cLBSatz = Left(cLBSatz, 6)
            cLBSatz = Trim$(cLBSatz)
            gcArtNrFiliale = cLBSatz
            frmWKLae.Show 1
        End If
    Case 10
        If List4.SelectedItem Is Nothing Then
            MsgBox "Bitte einen Eintrag in der Liste auswählen!", vbCritical, "STOP!"
        Else
            cLBSatz = List4.SelectedItem.Text
            cLBSatz = Left(cLBSatz, 6)
            cLBSatz = Trim$(cLBSatz)
            gcArtNrFiliale = cLBSatz
            frmWKLam.Show 1
        End If
    Case 15
        List4.Visible = False
        List2.Visible = False
        List4.Nodes.Clear
        
        Me.Refresh
        
        If Command4(15).Caption = "weitere Artikel" Then
            ermweitereArtikel
        ElseIf Command4(15).Caption = "noch mehr Artikel" Then
            ermweitereArtikelT2
        End If
End Select
            
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel suchen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnPruefeDialogSucheArtWKL20() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp    As String
    
    ctmp = Text3(6).Text
    ctmp = Trim$(ctmp)
    If ctmp <> "" Then
        fnPruefeDialogSucheArtWKL20 = 0
        Exit Function
    End If
    
    ctmp = Label1(9).Tag
    ctmp = Trim$(ctmp)
    If ctmp <> "" Then
        fnPruefeDialogSucheArtWKL20 = 0
        Exit Function
    End If
    
    ctmp = Text3(5).Text
    ctmp = Trim$(ctmp)
    If ctmp <> "" Then
        If IsNumeric(Left(ctmp, 6)) Then
            fnPruefeDialogSucheArtWKL20 = 0
            Exit Function
        Else
            Text3(5).Text = ""
        End If
    End If
    
    ctmp = Text3(1).Text
    ctmp = Trim$(ctmp)
    If ctmp <> "" Then
        fnPruefeDialogSucheArtWKL20 = 0
        Exit Function
    End If
    
    ctmp = Text3(2).Text
    ctmp = Trim$(ctmp)
    If ctmp <> "" Then
        fnPruefeDialogSucheArtWKL20 = 0
        Exit Function
    End If
    
    ctmp = Text3(4).Text
    ctmp = Trim$(ctmp)
    If ctmp <> "" Then
        fnPruefeDialogSucheArtWKL20 = 0
        Exit Function
    End If
    
    Text3(0).Text = SwapStr(Text3(0).Text, "'", "")
    ctmp = Text3(0).Text
    ctmp = Trim$(ctmp)
    
    If ctmp <> "" Then
        If Len(ctmp) > 1 Then
            fnPruefeDialogSucheArtWKL20 = 0
            Exit Function
        End If
    End If
    
    fnPruefeDialogSucheArtWKL20 = 1
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeDialogSucheArtWKL20"
    Fehler.gsFehlertext = "Im Programmteil Artikel suchen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub ermweitereArtikel()
    On Error GoTo LOKAL_ERROR
    
    Dim lLinr       As Long
    Dim lLinie      As Long
    Dim lcount      As Long
    Dim cRabattier  As String
    Dim ctmp        As String
    Dim cArtNrExakt As String
    Dim cMWST       As String
    Dim cSQL        As String
    Dim cArtBez     As String
    Dim cLiBesNr    As String
    Dim cEAN        As String
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim cVKPR       As String
    Dim cBestand    As String
    Dim cRabatt     As String
    Dim cPGN        As String
    Dim cAWM        As String
    Dim cLinr       As String
    
    Dim dRabatt     As Double
    Dim dVkPr       As Double
    Dim dKVkPr1     As Double
    Dim dFaktorV    As Double
    Dim dFaktorE    As Double
    Dim bAnd        As Boolean
    Dim rsrs        As Recordset
    Dim rsrs1       As Recordset
    Dim rsRs2       As Recordset
    Dim cMerk       As String
    
    Dim iStufe As Integer
    
    dFaktorV = 1 + (gdMWStV / 100)
    dFaktorE = 1 + (gdMWStE / 100)
    iStufe = 1
    dRabatt = 1
    If Label2(3).Visible Then
        cRabatt = Label2(3).Caption
        cRabatt = fnMoveComma2Point$(cRabatt)
        dRabatt = Val(cRabatt)
        dRabatt = dRabatt / 100
        dRabatt = 1 - dRabatt
    End If
    iStufe = 2
    If Label2(1).Visible Then
        cRabatt = Label2(1).Caption
        cRabatt = fnMoveComma2Point$(cRabatt)
        dRabatt = Val(cRabatt)
        dRabatt = dRabatt / 100
        dRabatt = 1 - dRabatt
    End If
    iStufe = 3
    bAnd = False
    
    cArtBez = Text3(0).Text
    cArtBez = UCase$(cArtBez)
'    Text3(0).Text = cArtBez

    cArtBez = LTrim(cArtBez)
    
    cArtBez = SwapStr(cArtBez, "     ", "*")
    cArtBez = SwapStr(cArtBez, "    ", "*")
    cArtBez = SwapStr(cArtBez, "   ", "*")
    cArtBez = SwapStr(cArtBez, "  ", "*")
    cArtBez = SwapStr(cArtBez, " ", "*")
    
    cLiBesNr = Text3(1).Text
    cLiBesNr = Trim$(cLiBesNr)
    cLiBesNr = UCase$(cLiBesNr)
    Text3(1).Text = cLiBesNr
    iStufe = 4
    cEAN = Text3(2).Text
    cEAN = Trim$(cEAN)
    cEAN = UCase$(cEAN)
    cEAN = SwapStr(cEAN, ",", "")
    cEAN = SwapStr(cEAN, ".", "")
    
    If cEAN <> "" Then
        If IsNumeric(cEAN) Then
            Text3(2).Text = cEAN
        Else
            Text3(2).Text = cEAN
            Text3(2).SetFocus
            Exit Sub
        End If
    End If
    
    cPGN = Text3(4).Text
    cPGN = Trim$(cPGN)

    cLinr = Text3(5).Text
    cLinr = Left(Trim(cLinr), 6)
    If IsNumeric(Left(cLinr, 6)) Then
    
    Else
        cLinr = ""
    End If
    
    iStufe = 5
    cAWM = Label1(9).Tag
    cAWM = Trim$(cAWM)
    
    cSQL = "Insert into " & srechnertab & "ASEEK Select "
    cSQL = cSQL & " ARTNR  "
    cSQL = cSQL & ", BEZEICH  "
    cSQL = cSQL & ", LEKPR  "
    cSQL = cSQL & ", EKPR  "
    cSQL = cSQL & ", KVKPR1  "
    cSQL = cSQL & ", VKPR  "
    cSQL = cSQL & ", RKZ  "
    cSQL = cSQL & ", BESTAND  "
    cSQL = cSQL & ", MWST "
    cSQL = cSQL & ", 2 as SEEKMOD  "
    cSQL = cSQL & ", AWM "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LPZ "
    cSQL = cSQL & " from ARTIKEL A  where "
    cSQL = cSQL & " Artnr not in (select Artnr from " & srechnertab & "ASEEK) "
    cSQL = cSQL & " and "
    
    iStufe = 6
    
    If cArtBez <> "" Then
        cSQL = cSQL & " a.BEZEICH like '*" & cArtBez & "*' "
        bAnd = True
    End If
    iStufe = 7
    
    If Check9.Value = vbUnchecked Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        
        cSQL = cSQL & " a.AWM <> '92' "
        bAnd = True
    Else
    
    End If
    
    If Check14.Value = vbChecked Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        
        cSQL = cSQL & " a.RKZ = 'N' "
        bAnd = True
    Else
    
    End If
    
    
    If cLiBesNr <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        
        cSQL = cSQL & " a.LIBESNR like '" & cLiBesNr & "' "
        bAnd = True
    End If
    iStufe = 8
    If cPGN <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        
        cSQL = cSQL & " a.PGN = " & cPGN
        bAnd = True
    End If
    
    If cLinr <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        
        cSQL = cSQL & " a.LINR = " & cLinr
        bAnd = True
    End If
     
    'Marke
    cFeld = Text3(6).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If LoeseMarkenInArtnr1(cFeld) Then
            If bAnd Then
                cSQL = cSQL & " and "
            End If
            cSQL = cSQL & " a.artnr in (Select artnr from MA" & srechnertab & ") "
            bAnd = True
        Else
            anzeige "rot", "Keine Artikel!", Label9
            Exit Sub
        End If
    End If
    
    iStufe = 9
    If cAWM <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        
        cSQL = cSQL & " a.AWM = '" & cAWM & "' "
        bAnd = True
    End If
    
    iStufe = 10
    
    If cEAN <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        If Len(cEAN) > 6 Or (Len(cEAN) = 6 And Left(cEAN, 1) = "9") Then
        
            If Len(cEAN) = 8 Then
                If Left(cEAN, 1) = "2" Then
                    cEAN = Mid$(cEAN, 2, 6)
                    cSQL = cSQL & " a.ARTNR = " & cEAN
                    bAnd = True
                 Else
                    cSQL = cSQL & "(a.EAN = '" & cEAN & "' or a.EAN2 = '" & cEAN & "' or a.EAN3 = '" & cEAN & "' ) "
                    bAnd = True
                End If
            Else
                cSQL = cSQL & "(a.EAN = '" & cEAN & "' or a.EAN2 = '" & cEAN & "' or a.EAN3 = '" & cEAN & "' ) "
                bAnd = True
            
            End If
        Else
        '1.wunsch
            If IsNumeric(cEAN) Then
                cSQL = cSQL & " a.ARTNR = " & cEAN
                bAnd = True
            Else
                cSQL = cSQL & " a.ARTNR = -1 "
                bAnd = True
            End If
        End If
    End If
    
    If bAnd Then
        cSQL = cSQL & " and "
    End If
    cSQL = cSQL & "  ( A.SYNSTATUS is null or A.SYNSTATUS = 'E' or A.SYNSTATUS = 'A' ) "
    
    iStufe = 11
'    hier sortierung
    If Option2(0).Value = True Then
        cSQL = cSQL & " order by a.BEZEICH "
    Else
        cSQL = cSQL & " order by a.LINR, a.LPZ, a.BEZEICH "
    End If
    iStufe = 12
    
    
    Screen.MousePointer = 11
    anzeige "ROT2", "Artikel werden gesucht(2)...", Label9
    
    gdBase.Execute cSQL, dbFailOnError
 
    cSQL = "Select * from " & srechnertab & "ASEEK where seekmod = 2 "
    If Option2(0).Value = True Then
        cSQL = cSQL & " order by BEZEICH "
    Else
        cSQL = cSQL & " order by LINR, LPZ, BEZEICH "
    End If
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iStufe = 17
            If Not IsNull(rsrs!MWST) Then
                cMWST = rsrs!MWST
            Else
                cMWST = "V"
            End If
            iStufe = 18
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cFeld & " "
            iStufe = 19
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            iStufe = 20
            If Not IsNull(rsrs!vkpr) Then
                dVkPr = rsrs!vkpr
            Else
                dVkPr = 0
            End If
            cFeld = Format$(dVkPr, "###,##0.00")
            cFeld = Trim$(cFeld)
            If Len(cFeld) > 9 Then
                cFeld = Space$(9)
            Else
                cFeld = Space$(9 - Len(cFeld)) & cFeld
            End If
            cLBSatz = cLBSatz & cFeld & " "
            iStufe = 21
            Select Case Val(Label8(3).Caption)
                Case Is = 0
                    If Not IsNull(rsrs!KVKPR1) Then
                        dKVkPr1 = rsrs!KVKPR1
                    Else
                        dKVkPr1 = 0
                    End If
                Case Is = 1
                    If Not IsNull(rsrs!vkpr) Then
                        dKVkPr1 = rsrs!vkpr
                    Else
                        dKVkPr1 = 0
                    End If
                Case Is = 2
                    If Not IsNull(rsrs!lekpr) Then
                        dKVkPr1 = rsrs!lekpr
                    Else
                        dKVkPr1 = 0
                    End If
                    If cMWST = "V" Then
                        dKVkPr1 = dKVkPr1 * dFaktorV
                    End If
                    If cMWST = "E" Then
                        dKVkPr1 = dKVkPr1 * dFaktorE
                    End If
                    
                Case Is = 3
                    If Not IsNull(rsrs!ekpr) Then
                        dKVkPr1 = rsrs!ekpr
                    Else
                        dKVkPr1 = 0
                    End If
                    If cMWST = "V" Then
                        dKVkPr1 = dKVkPr1 * dFaktorV
                    End If
                    If cMWST = "E" Then
                        dKVkPr1 = dKVkPr1 * dFaktorE
                    End If
                Case Is = 4 'Spez kvk
                    dKVkPr1 = LeseSpezpreis(CLng(rsrs!artnr), 0)
                    If dKVkPr1 = 0 Then
                        If Not IsNull(rsrs!KVKPR1) Then
                            dKVkPr1 = rsrs!KVKPR1
                        Else
                            dKVkPr1 = 0
                        End If
                    End If
                Case Is = 5 'lvk m A
                    If rsrs!RABATT_OK = "N" Then
                        If Not IsNull(rsrs!KVKPR1) Then
                            dKVkPr1 = rsrs!KVKPR1
                        Else
                            dKVkPr1 = 0
                        End If
                    Else
                        If Not IsNull(rsrs!vkpr) Then
                            dKVkPr1 = rsrs!vkpr
                        Else
                            dKVkPr1 = 0
                        End If
                    End If
            End Select
            iStufe = 22
            dKVkPr1 = dKVkPr1 * dRabatt
            cFeld = Format$(dKVkPr1, "###,##0.00")
            
            cFeld = Trim$(cFeld)
            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            iStufe = 23
            If Not IsNull(rsrs!BESTAND) Then
                dVkPr = rsrs!BESTAND
            Else
                dVkPr = 0
            End If
            cFeld = Format$(dVkPr, "#,##0")
            cFeld = Trim$(cFeld)
            If Len(cFeld) > 5 Then cFeld = 0
            cFeld = Space$(5 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            iStufe = 24
            If Not IsNull(rsrs!RKZ) Then
                cFeld = rsrs!RKZ
            Else
                cFeld = "N"
            End If
            cFeld = Trim$(cFeld)
            cLBSatz = cLBSatz & cFeld & Space$(1)
            
            If Not IsNull(rsrs!AWM) Then
                cFeld = rsrs!AWM
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cLBSatz = cLBSatz & cFeld & Space(2 - Len(cFeld))
            
            
            cMerk = Left(ZeigeArtmerk(Trim(Left(cLBSatz, 6))), 1)
            cLBSatz = cLBSatz & Space(1 - Len(cMerk)) & cMerk
            
            iStufe = 25
            List4.Nodes.Add Text:=cLBSatz
            
            rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    iStufe = 26
    
    Screen.MousePointer = 11
    farbelist4 Me
    
    
    Dim cartT As String
    iStufe = 27
    If List4.Nodes.Count > 0 Then
        For lcount = 1 To List4.Nodes.Count
            ctmp = List4.Nodes(lcount)
            cartT = Trim$(Left(ctmp, 6))
            If cartT = Trim$(cArtNrExakt) Then
                List4.Nodes(lcount).Selected = True
                List4.Nodes(lcount).EnsureVisible
                Exit For
            End If
        Next lcount
        anzeige "normal", List4.Nodes.Count & " Artikel", Label9
'        anzeige "normal", lCount - 1 & " Artikel", Label9
    Else
        If cArtBez <> "" Then
            If InStr(cArtBez, "*") > 0 Then
                Command4(15).Caption = "noch mehr Artikel"
                Command4(15).Visible = True
                Command4_Click 15
                Exit Sub
            Else
                anzeige "rot", "Keine Artikel!", Label9
            
            End If
        
            
        Else
            anzeige "rot", "Keine Artikel!", Label9
        End If
    End If
    iStufe = 28
    Frame7.Visible = False
    List2.Visible = True
    List4.Visible = True
    
    Command4(10).Visible = True
    Command4(9).Visible = True
    Command4(1).Visible = True
    List4.SetFocus
    Screen.MousePointer = 0
    
    'noch weitere Ergebnisse
    
    cArtBez = Trim(cArtBez)
    
    If cArtBez <> "" Then
        If InStr(cArtBez, "*") > 0 Then
            Command4(15).Caption = "noch mehr Artikel"
            Command4(15).Visible = True
        Else
            Command4(15).Caption = "weitere Artikel"
            Command4(15).Visible = False
        End If
    End If
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermweitereArtikel"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten. " & iStufe
     
    Fehlermeldung1
End Sub
Private Sub ermweitereArtikelT2()
    On Error GoTo LOKAL_ERROR
    
    Dim lLinr       As Long
    Dim lLinie      As Long
    Dim lcount      As Long
    Dim cRabattier  As String
    Dim ctmp        As String
    Dim cArtNrExakt As String
    Dim cMWST       As String
    Dim cSQL        As String
    Dim cArtBez     As String
    Dim cLiBesNr    As String
    Dim cEAN        As String
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim cVKPR       As String
    Dim cBestand    As String
    Dim cRabatt     As String
    Dim cPGN        As String
    Dim cAWM        As String
    Dim cLinr       As String
    Dim cMerk       As String
    
    Dim dRabatt     As Double
    Dim dVkPr       As Double
    Dim dKVkPr1     As Double
    Dim dFaktorV    As Double
    Dim dFaktorE    As Double
    Dim bAnd        As Boolean
    Dim rsrs        As Recordset
    Dim rsrs1       As Recordset
    Dim rsRs2       As Recordset
    Dim i           As Integer
    Dim j           As Integer
    
    Dim iStufe As Integer
    
    dFaktorV = 1 + (gdMWStV / 100)
    dFaktorE = 1 + (gdMWStE / 100)
    iStufe = 1
    dRabatt = 1
    If Label2(3).Visible Then
        cRabatt = Label2(3).Caption
        cRabatt = fnMoveComma2Point$(cRabatt)
        dRabatt = Val(cRabatt)
        dRabatt = dRabatt / 100
        dRabatt = 1 - dRabatt
    End If
    iStufe = 2
    If Label2(1).Visible Then
        cRabatt = Label2(1).Caption
        cRabatt = fnMoveComma2Point$(cRabatt)
        dRabatt = Val(cRabatt)
        dRabatt = dRabatt / 100
        dRabatt = 1 - dRabatt
    End If
    iStufe = 3
    bAnd = False
    
    cArtBez = Text3(0).Text
    cArtBez = UCase$(cArtBez)
'    Text3(0).Text = cArtBez

    cArtBez = LTrim(cArtBez)
    
    cArtBez = SwapStr(cArtBez, "     ", "*")
    cArtBez = SwapStr(cArtBez, "    ", "*")
    cArtBez = SwapStr(cArtBez, "   ", "*")
    cArtBez = SwapStr(cArtBez, "  ", "*")
    cArtBez = SwapStr(cArtBez, " ", "*")
    'jetzt malrauskriegen wieviele Teilstrings sind drin
    
    Dim lZaehler    As Long
    Dim lStart      As Long
    Dim lPos        As Long
    Dim lPosT       As Long
    ReDim cTeilstring(1 To 1) As String
    
    lStart = 0
    lZaehler = 0
    lPos = 100
    
    Do While lPos > 0
        lPos = InStr(lStart + 1, cArtBez, "*")
        
        lZaehler = lZaehler + 1
        ReDim Preserve cTeilstring(1 To lZaehler) As String
        
        lPosT = 0
        lPosT = InStr(lStart + 1, cArtBez, "*")
        
        If lPosT = 0 Then
            cTeilstring(lZaehler) = Mid(cArtBez, lStart + 1, Len(cArtBez) - (lStart))
        Else
            cTeilstring(lZaehler) = Mid(cArtBez, lStart + 1, lPosT - (lStart + 1))
        End If
        
        
        lStart = lPos
    Loop
    
    'Anzahl Teilstrings stehen fest
'    For i = 1 To LZaehler
'
'    Next i
    
    
    cLiBesNr = Text3(1).Text
    cLiBesNr = Trim$(cLiBesNr)
    cLiBesNr = UCase$(cLiBesNr)
    Text3(1).Text = cLiBesNr
    iStufe = 4
    cEAN = Text3(2).Text
    cEAN = Trim$(cEAN)
    cEAN = UCase$(cEAN)
    cEAN = SwapStr(cEAN, ",", "")
    cEAN = SwapStr(cEAN, ".", "")
    
    If cEAN <> "" Then
        If IsNumeric(cEAN) Then
            Text3(2).Text = cEAN
        Else
            Text3(2).Text = cEAN
            Text3(2).SetFocus
            Exit Sub
        End If
    End If
    
    cPGN = Text3(4).Text
    cPGN = Trim$(cPGN)

    cLinr = Text3(5).Text
    cLinr = Left(Trim(cLinr), 6)
    If IsNumeric(Left(cLinr, 6)) Then
    
    Else
        cLinr = ""
    End If
    
    iStufe = 5
    cAWM = Label1(9).Tag
    cAWM = Trim$(cAWM)
    
    Dim cSQLAnfang As String
    
    cSQLAnfang = "Insert into " & srechnertab & "ASEEK Select "
    cSQLAnfang = cSQLAnfang & " ARTNR  "
    cSQLAnfang = cSQLAnfang & ", BEZEICH  "
    cSQLAnfang = cSQLAnfang & ", LEKPR  "
    cSQLAnfang = cSQLAnfang & ", EKPR  "
    cSQLAnfang = cSQLAnfang & ", KVKPR1  "
    cSQLAnfang = cSQLAnfang & ", VKPR  "
    cSQLAnfang = cSQLAnfang & ", RKZ  "
    cSQLAnfang = cSQLAnfang & ", BESTAND  "
    cSQLAnfang = cSQLAnfang & ", MWST "
    cSQLAnfang = cSQLAnfang & ", 3 as SEEKMOD  "
    cSQLAnfang = cSQLAnfang & ", AWM "
    cSQLAnfang = cSQLAnfang & ", LINR "
    cSQLAnfang = cSQLAnfang & ", LPZ "
    cSQLAnfang = cSQLAnfang & " from ARTIKEL A  where "
    cSQLAnfang = cSQLAnfang & " Artnr not in (select Artnr from " & srechnertab & "ASEEK) "
    cSQLAnfang = cSQLAnfang & " and "
    
'    Dim lKombis As Long
'
    Select Case lZaehler

        Case 2
'            lKombis = 1 * 2
        Case 3
'            lKombis = 1 * 2 * 3
        Case 4
        
        Case 5

        Case 6

        
        
        Case Else
            anzeige "rot", "Keine Artikel!", Label9
        
            Frame7.Visible = True
            List2.Visible = False
            List4.Visible = False
            
            Command4(10).Visible = False
            Command4(9).Visible = False
            Command4(1).Visible = False
            
            Command4(15).Caption = "weitere Artikel"
            Command4(15).Visible = False
            Text3(0).SetFocus
            Exit Sub
    End Select

    
    
    Select Case lZaehler
        Case 2
            
            anzeige "ROT2", "Artikel werden gesucht(3)...", Label9
            cSQL = ""
            cArtBez = "*" & cTeilstring(1) & "*" & cTeilstring(2) & "*"
            cSQL = cSQLAnfang & MachDenRest(cArtBez)
            gdBase.Execute cSQL, dbFailOnError
            
            anzeige "ROT2", "Artikel werden gesucht(4)...", Label9
            cSQL = ""
            cArtBez = "*" & cTeilstring(2) & "*" & cTeilstring(1) & "*"
            cSQL = cSQLAnfang & MachDenRest(cArtBez)
            gdBase.Execute cSQL, dbFailOnError

        Case 3, 4, 5, 6
        
            anzeige "ROT2", "Artikel werden gesucht(3)...", Label9
            cSQL = ""
            cArtBez = "*" & cTeilstring(1) & "*" & cTeilstring(2) & "*" & cTeilstring(3) & "*"
            cSQL = cSQLAnfang & MachDenRest(cArtBez)
            gdBase.Execute cSQL, dbFailOnError
            
            anzeige "ROT2", "Artikel werden gesucht(4)...", Label9
            cSQL = ""
            cArtBez = "*" & cTeilstring(1) & "*" & cTeilstring(3) & "*" & cTeilstring(2) & "*"
            cSQL = cSQLAnfang & MachDenRest(cArtBez)
            gdBase.Execute cSQL, dbFailOnError
            
            anzeige "ROT2", "Artikel werden gesucht(5)...", Label9
            cSQL = ""
            cArtBez = "*" & cTeilstring(2) & "*" & cTeilstring(1) & "*" & cTeilstring(3) & "*"
            cSQL = cSQLAnfang & MachDenRest(cArtBez)
            gdBase.Execute cSQL, dbFailOnError
            
            
            anzeige "ROT2", "Artikel werden gesucht(6)...", Label9
            cSQL = ""
            cArtBez = "*" & cTeilstring(2) & "*" & cTeilstring(3) & "*" & cTeilstring(1) & "*"
            cSQL = cSQLAnfang & MachDenRest(cArtBez)
            gdBase.Execute cSQL, dbFailOnError
            
            anzeige "ROT2", "Artikel werden gesucht(7)...", Label9
            cSQL = ""
            cArtBez = "*" & cTeilstring(3) & "*" & cTeilstring(1) & "*" & cTeilstring(2) & "*"
            cSQL = cSQLAnfang & MachDenRest(cArtBez)
            gdBase.Execute cSQL, dbFailOnError
            
            anzeige "ROT2", "Artikel werden gesucht(8)...", Label9
            cSQL = ""
            cArtBez = "*" & cTeilstring(3) & "*" & cTeilstring(2) & "*" & cTeilstring(1) & "*"
            cSQL = cSQLAnfang & MachDenRest(cArtBez)
            gdBase.Execute cSQL, dbFailOnError
        
        
    End Select

 
    cSQL = "Select * from " & srechnertab & "ASEEK where seekmod = 3 "
    If Option2(0).Value = True Then
        cSQL = cSQL & " order by BEZEICH "
    Else
        cSQL = cSQL & " order by LINR, LPZ, BEZEICH "
    End If
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iStufe = 17
            If Not IsNull(rsrs!MWST) Then
                cMWST = rsrs!MWST
            Else
                cMWST = "V"
            End If
            iStufe = 18
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cFeld & " "
            iStufe = 19
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            iStufe = 20
            If Not IsNull(rsrs!vkpr) Then
                dVkPr = rsrs!vkpr
            Else
                dVkPr = 0
            End If
            cFeld = Format$(dVkPr, "###,##0.00")
            cFeld = Trim$(cFeld)
            If Len(cFeld) > 9 Then
                cFeld = Space$(9)
            Else
                cFeld = Space$(9 - Len(cFeld)) & cFeld
            End If
            cLBSatz = cLBSatz & cFeld & " "
            iStufe = 21
            Select Case Val(Label8(3).Caption)
                Case Is = 0
                    If Not IsNull(rsrs!KVKPR1) Then
                        dKVkPr1 = rsrs!KVKPR1
                    Else
                        dKVkPr1 = 0
                    End If
                Case Is = 1
                    If Not IsNull(rsrs!vkpr) Then
                        dKVkPr1 = rsrs!vkpr
                    Else
                        dKVkPr1 = 0
                    End If
                Case Is = 2
                    If Not IsNull(rsrs!lekpr) Then
                        dKVkPr1 = rsrs!lekpr
                    Else
                        dKVkPr1 = 0
                    End If
                    If cMWST = "V" Then
                        dKVkPr1 = dKVkPr1 * dFaktorV
                    End If
                    If cMWST = "E" Then
                        dKVkPr1 = dKVkPr1 * dFaktorE
                    End If
                    
                Case Is = 3
                    If Not IsNull(rsrs!ekpr) Then
                        dKVkPr1 = rsrs!ekpr
                    Else
                        dKVkPr1 = 0
                    End If
                    If cMWST = "V" Then
                        dKVkPr1 = dKVkPr1 * dFaktorV
                    End If
                    If cMWST = "E" Then
                        dKVkPr1 = dKVkPr1 * dFaktorE
                    End If
                Case Is = 4 'Spez kvk
                    dKVkPr1 = LeseSpezpreis(CLng(rsrs!artnr), 0)
                    If dKVkPr1 = 0 Then
                        If Not IsNull(rsrs!KVKPR1) Then
                            dKVkPr1 = rsrs!KVKPR1
                        Else
                            dKVkPr1 = 0
                        End If
                    End If
                Case Is = 5 'lvk m A
                    If rsrs!RABATT_OK = "N" Then
                        If Not IsNull(rsrs!KVKPR1) Then
                            dKVkPr1 = rsrs!KVKPR1
                        Else
                            dKVkPr1 = 0
                        End If
                    Else
                        If Not IsNull(rsrs!vkpr) Then
                            dKVkPr1 = rsrs!vkpr
                        Else
                            dKVkPr1 = 0
                        End If
                    End If
                
            End Select
            iStufe = 22
            dKVkPr1 = dKVkPr1 * dRabatt
            cFeld = Format$(dKVkPr1, "###,##0.00")
            
            cFeld = Trim$(cFeld)
            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            iStufe = 23
            If Not IsNull(rsrs!BESTAND) Then
                dVkPr = rsrs!BESTAND
            Else
                dVkPr = 0
            End If
            cFeld = Format$(dVkPr, "#,##0")
            cFeld = Trim$(cFeld)
            If Len(cFeld) > 5 Then cFeld = 0
            cFeld = Space$(5 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            iStufe = 24
            If Not IsNull(rsrs!RKZ) Then
                cFeld = rsrs!RKZ
            Else
                cFeld = "N"
            End If
            cFeld = Trim$(cFeld)
            cLBSatz = cLBSatz & cFeld & Space$(1)
            
            If Not IsNull(rsrs!AWM) Then
                cFeld = rsrs!AWM
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cLBSatz = cLBSatz & cFeld & Space(2 - Len(cFeld))
            
            
            cMerk = Left(ZeigeArtmerk(Trim(Left(cLBSatz, 6))), 1)
            cLBSatz = cLBSatz & Space(1 - Len(cMerk)) & cMerk
            
            iStufe = 25
            List4.Nodes.Add Text:=cLBSatz
            
            rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    iStufe = 26
    
    Screen.MousePointer = 11
    farbelist4 Me
    
    
    Dim cartT As String
    iStufe = 27
    If List4.Nodes.Count > 0 Then
        For lcount = 1 To List4.Nodes.Count
            ctmp = List4.Nodes(lcount)
            cartT = Trim$(Left(ctmp, 6))
            If cartT = Trim$(cArtNrExakt) Then
                List4.Nodes(lcount).Selected = True
                List4.Nodes(lcount).EnsureVisible
                Exit For
            End If
        Next lcount
        anzeige "normal", List4.Nodes.Count & " Artikel", Label9
'        anzeige "normal", lCount - 1 & " Artikel", Label9
        
        
        Frame7.Visible = False
        List2.Visible = True
        List4.Visible = True
        
        Command4(10).Visible = True
        Command4(9).Visible = True
        Command4(1).Visible = True
        List4.SetFocus
        
    Else
        anzeige "rot", "Keine Artikel!", Label9
        
        Frame7.Visible = True
        List2.Visible = False
        List4.Visible = False
        
        Command4(10).Visible = False
        Command4(9).Visible = False
        Command4(1).Visible = False
        
        Command4(15).Caption = "weitere Artikel"
        Command4(15).Visible = False
        Text3(0).SetFocus
    End If
    iStufe = 28
    
    Screen.MousePointer = 0
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermweitereArtikelT2"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten. " & iStufe
     
    Fehlermeldung1
End Sub
Private Function MachDenRest(cArtBez As String) As String
On Error GoTo LOKAL_ERROR


    Dim cSQL        As String
    Dim bAnd   As Boolean
    Dim iStufe As Integer
    Dim cLiBesNr As String
    Dim cPGN As String
    Dim cLinr As String
    Dim cFeld As String
    Dim cAWM As String
    Dim cEAN As String
    
    
    cLiBesNr = Text3(1).Text
    cLiBesNr = Trim$(cLiBesNr)
    cLiBesNr = UCase$(cLiBesNr)
    Text3(1).Text = cLiBesNr
    iStufe = 4
    cEAN = Text3(2).Text
    cEAN = Trim$(cEAN)
    cEAN = UCase$(cEAN)
    cEAN = SwapStr(cEAN, ",", "")
    cEAN = SwapStr(cEAN, ".", "")
    
    If cEAN <> "" Then
        If IsNumeric(cEAN) Then
            Text3(2).Text = cEAN
        Else
            Text3(2).Text = cEAN
            Text3(2).SetFocus
            Exit Function
        End If
    End If
    
    cPGN = Text3(4).Text
    cPGN = Trim$(cPGN)

    cLinr = Text3(5).Text
    cLinr = Left(Trim(cLinr), 6)
    If IsNumeric(Left(cLinr, 6)) Then
    
    Else
        cLinr = ""
    End If
    
    iStufe = 5
    cAWM = Label1(9).Tag
    cAWM = Trim$(cAWM)

    cSQL = ""
    

        If cArtBez <> "" Then
            cSQL = cSQL & " a.BEZEICH like '" & cArtBez & "' "
            bAnd = True
        End If
        iStufe = 7
        If Check9.Value = vbUnchecked Then
            If bAnd Then
                cSQL = cSQL & " and "
            End If
            
            cSQL = cSQL & " a.AWM <> '92' "
            bAnd = True
        End If
        
        If Check14.Value = vbChecked Then
            If bAnd Then
                cSQL = cSQL & " and "
            End If
            
            cSQL = cSQL & " a.RKZ = 'N' "
            bAnd = True
        Else
        
        End If
        If cLiBesNr <> "" Then
            If bAnd Then
                cSQL = cSQL & " and "
            End If
            
            cSQL = cSQL & " a.LIBESNR like '" & cLiBesNr & "' "
            bAnd = True
        End If
        iStufe = 8
        If cPGN <> "" Then
            If bAnd Then
                cSQL = cSQL & " and "
            End If
            
            cSQL = cSQL & " a.PGN = " & cPGN
            bAnd = True
        End If
        If cLinr <> "" Then
            If bAnd Then
                cSQL = cSQL & " and "
            End If
            
            cSQL = cSQL & " a.LINR = " & cLinr
            bAnd = True
        End If
         
        'Marke
        cFeld = Text3(6).Text
        cFeld = Trim$(cFeld)
        If cFeld <> "" Then
            If LoeseMarkenInArtnr1(cFeld) Then
                If bAnd Then
                    cSQL = cSQL & " and "
                End If
                cSQL = cSQL & " a.artnr in (Select artnr from MA" & srechnertab & ") "
                bAnd = True
            Else
                anzeige "rot", "Keine Artikel!", Label9
                Exit Function
            End If
        End If
        
        iStufe = 9
        If cAWM <> "" Then
            If bAnd Then
                cSQL = cSQL & " and "
            End If
            
            cSQL = cSQL & " a.AWM = '" & cAWM & "' "
            bAnd = True
        End If
        
        iStufe = 10
        
        If cEAN <> "" Then
            If bAnd Then
                cSQL = cSQL & " and "
            End If
            If Len(cEAN) > 6 Or (Len(cEAN) = 6 And Left(cEAN, 1) = "9") Then
            
                If Len(cEAN) = 8 Then
                    If Left(cEAN, 1) = "2" Then
                        cEAN = Mid$(cEAN, 2, 6)
                        cSQL = cSQL & " a.ARTNR = " & cEAN
                        bAnd = True
                     Else
                        cSQL = cSQL & "(a.EAN = '" & cEAN & "' or a.EAN2 = '" & cEAN & "' or a.EAN3 = '" & cEAN & "' ) "
                        bAnd = True
                    End If
                Else
                    cSQL = cSQL & "(a.EAN = '" & cEAN & "' or a.EAN2 = '" & cEAN & "' or a.EAN3 = '" & cEAN & "' ) "
                    bAnd = True
                
                End If
            Else
            '1.wunsch
                If IsNumeric(cEAN) Then
                    cSQL = cSQL & " a.ARTNR = " & cEAN
                    bAnd = True
                Else
                    cSQL = cSQL & " a.ARTNR = -1 "
                    bAnd = True
                End If
            End If
        End If
        
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        cSQL = cSQL & "  ( A.SYNSTATUS is null or A.SYNSTATUS = 'E' or A.SYNSTATUS = 'A' ) "
        
        MachDenRest = cSQL

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MachDenRest"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten. " & iStufe
     
    Fehlermeldung1

End Function
Private Sub Command66_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0
            Text3_KeyUp 4, vbKeyF2, 0
        Case Is = 1
            Text3_KeyUp 6, vbKeyF2, 0
        Case Is = 2
            Text3_KeyUp 5, vbKeyF2, 0
        Case Is = 3    'F2 farbe
            Screen.MousePointer = 0
            gsBackcolor = Label1(9).BackColor
            gsForecolor = Label1(9).ForeColor
            gsArtikelFarbe = Label1(9).Tag
            
            frmWKL49.Show 1
            
            Label1(9).BackColor = gsBackcolor
            Label1(9).ForeColor = gsForecolor
            Label1(9).Tag = gsArtikelFarbe
            
            If gsArtikelFarbe <> "" Then
                Label1(9).Caption = "Farbauswahl"
            Else
                Label1(9).Caption = "alle Farben"
            End If
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command66_Click"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladenE142C()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    Dim bo0 As Integer
    Dim bo1 As Integer
    
    Set rs = gdApp.OpenRecordset("E142C")
    If Not rs.EOF Then
        If rs!bo0 = True Then
            Check9.Value = vbUnchecked
        Else
            Check9.Value = vbChecked
        End If
        
        If rs!bo1 = True Then
            Check14.Value = vbUnchecked
        Else
            Check14.Value = vbChecked
        End If
    End If
    rs.Close: Set rs = Nothing
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladenE142C"
    Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichernE142C()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim bo0 As Integer
    Dim bo1 As Integer
    
    loeschNEW "E142C", gdApp
    CreateTableT2 "E142C", gdApp
    
    If Check9.Value = vbChecked Then
        bo0 = 0
    Else
        bo0 = -1
    End If
    
    If Check14.Value = vbChecked Then
        bo1 = 0
    Else
        bo1 = -1
    End If
    
    sSQL = "Insert into E142C ( bo0,bo1) "
    sSQL = sSQL & " values (" & bo0 & "," & bo1 & ")"

    gdApp.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichernE142C"
    Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    WKL142Positionieren
    Modul6.Skalieren Me, True, True:
    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    If NewTableSuchenDBKombi("E142C", gdApp) Then
        If SpalteInTabellegefundenNEW("E142C", "BO1", gdApp) Then
            voreinstellungladenE142C
        End If
    End If

    Text3(0).Text = ""
    Text3(1).Text = ""
    Text3(2).Text = ""
    Text3(4).Text = ""
    Text3(5).Text = ""
    Text3(6).Text = ""
    bFocusonList4 = False
    List4.Visible = False
    List2.Visible = False
    List2.Clear
    List4.Nodes.Clear
    List2.AddItem "ArtNr. Artikelbezeichnung                  Listen-Vk Kunden-Vk Best. RKZ"
    Me.Refresh
    If gbFilNr And gcFilNr <> 0 Then
        Command4(9).Visible = True
    Else
        Command4(9).Visible = False
    End If
    
    Command4(15).Caption = "weitere Artikel"
    Command4(15).Visible = False
    
    Frame6.Visible = True
    anzeige "normal", "Artikel suchen", Label9
    Me.Refresh

    
    If NewTableSuchenDBKombi("EKASS", gdApp) = True Then
        voreinstellungladen
    End If
    
    Label2(3).Caption = 0
    If frmWKL141.Label1(0).Caption <> "" Then 'Kundenrabatt = gesrab 2(3)
        If IsNumeric(frmWKL141.Label1(0).Caption) Then
            Label2(3).Caption = frmWKL141.Label1(0).Caption
        End If
    End If
    
    Label8(3).Caption = 0
    If frmWKL141.Label1(1).Caption <> "" Then 'preiskz
        If IsNumeric(frmWKL141.Label1(1).Caption) Then
            Label8(3).Caption = frmWKL141.Label1(1).Caption
        End If
    End If
    
    Label3(5).Caption = "0"
'    Text3(0).SetFocus
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikel suchen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub voreinstellungladen()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    Dim cSQL As String
    
    cSQL = "select * from EKASS"
    Set rs = gdApp.OpenRecordset(cSQL)
    
    If Not rs.EOF Then
        rs.MoveFirst
        If Not IsNull(rs!bo0) Then
            Option2(0).Value = rs!bo0
        End If
        If Not IsNull(rs!bo1) Then
            Option2(1).Value = rs!bo1
        End If
    End If
    rs.Close: Set rs = Nothing

     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
    Resume Next
End Sub
Private Sub SucheArtikelKasseWKL20()
    On Error GoTo LOKAL_ERROR
    
    Dim lLinr       As Long
    Dim lLinie      As Long
    Dim lcount      As Long
    Dim cRabattier  As String
    Dim ctmp        As String
    Dim cArtNrExakt As String
    Dim cMWST       As String
    Dim cSQL        As String
    Dim cArtBez     As String
    Dim cLiBesNr    As String
    Dim cEAN        As String
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim cVKPR       As String
    Dim cBestand    As String
    Dim cRabatt     As String
    Dim cPGN        As String
    Dim cAWM        As String
    Dim cLinr       As String
    Dim cMerk       As String
    
    Dim dRabatt     As Double
    Dim dVkPr       As Double
    Dim dKVkPr1     As Double
    Dim dFaktorV    As Double
    Dim dFaktorE    As Double
    Dim bAnd        As Boolean
    Dim rsrs        As Recordset
    Dim rsrs1       As Recordset
    Dim rsRs2       As Recordset
    
    Dim iStufe As Integer
    
    dFaktorV = 1 + (gdMWStV / 100)
    dFaktorE = 1 + (gdMWStE / 100)
    iStufe = 1
    dRabatt = 1
    If Label2(3).Visible Then
        cRabatt = Label2(3).Caption
        cRabatt = fnMoveComma2Point$(cRabatt)
        dRabatt = Val(cRabatt)
        dRabatt = dRabatt / 100
        dRabatt = 1 - dRabatt
    End If
    iStufe = 2
    If Label2(1).Visible Then
        cRabatt = Label2(1).Caption
        cRabatt = fnMoveComma2Point$(cRabatt)
        dRabatt = Val(cRabatt)
        dRabatt = dRabatt / 100
        dRabatt = 1 - dRabatt
    End If
    iStufe = 3
    bAnd = False
    
    cArtBez = Text3(0).Text
    cArtBez = UCase$(cArtBez)

    cArtBez = LTrim(cArtBez)
    
    cArtBez = SwapStr(cArtBez, "     ", "*")
    cArtBez = SwapStr(cArtBez, "    ", "*")
    cArtBez = SwapStr(cArtBez, "   ", "*")
    cArtBez = SwapStr(cArtBez, "  ", "*")
    cArtBez = SwapStr(cArtBez, " ", "*")
    
    cLiBesNr = Text3(1).Text
    cLiBesNr = Trim$(cLiBesNr)
    cLiBesNr = UCase$(cLiBesNr)
    Text3(1).Text = cLiBesNr
    iStufe = 4
    cEAN = Text3(2).Text
    cEAN = Trim$(cEAN)
    cEAN = UCase$(cEAN)
    cEAN = SwapStr(cEAN, ",", "")
    cEAN = SwapStr(cEAN, ".", "")
    
    If cEAN <> "" Then
        If IsNumeric(cEAN) Then
            Text3(2).Text = cEAN
        Else
            Text3(2).Text = cEAN
            Text3(2).SetFocus
            Exit Sub
        End If
    End If
    
    cPGN = Text3(4).Text
    cPGN = Trim$(cPGN)

    cLinr = Text3(5).Text
    cLinr = Left(Trim(cLinr), 6)
    If IsNumeric(Left(cLinr, 6)) Then
    
    Else
        cLinr = ""
    End If
    
    iStufe = 5
    cAWM = Label1(9).Tag
    cAWM = Trim$(cAWM)
    
    loeschNEW srechnertab & "ASEEK", gdBase
    CreateTable srechnertab & "ASEEK", gdBase
    
    cSQL = "Insert into " & srechnertab & "ASEEK Select "
    cSQL = cSQL & " ARTNR  "
    cSQL = cSQL & ", BEZEICH  "
    cSQL = cSQL & ", LEKPR  "
    cSQL = cSQL & ", EKPR  "
    cSQL = cSQL & ", KVKPR1  "
    cSQL = cSQL & ", VKPR  "
    cSQL = cSQL & ", RKZ  "
    cSQL = cSQL & ", BESTAND  "
    cSQL = cSQL & ", MWST "
    cSQL = cSQL & ", 1 as SEEKMOD  "
    cSQL = cSQL & ", AWM "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LPZ "
    cSQL = cSQL & ", RABATT_OK "
    cSQL = cSQL & " from ARTIKEL A  where "

    iStufe = 6
    
    If cArtBez <> "" Then
        cSQL = cSQL & " a.BEZEICH like '" & cArtBez & "*' "
        bAnd = True
    End If
    iStufe = 7
    
    If Check9.Value = vbUnchecked Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        
        cSQL = cSQL & " a.AWM <> '92' "
        bAnd = True
    Else
    
    End If
    
    If Check14.Value = vbChecked Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        
        cSQL = cSQL & " a.RKZ = 'N' "
        bAnd = True
    Else
    
    End If
    
    
    If cLiBesNr <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        
        cSQL = cSQL & " a.LIBESNR like '" & cLiBesNr & "' "
        bAnd = True
    End If
    iStufe = 8
    If cPGN <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        
        cSQL = cSQL & " a.PGN = " & cPGN
        bAnd = True
    End If
    
    If cLinr <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        
        cSQL = cSQL & " a.LINR = " & cLinr
        bAnd = True
    End If
     
    'Marke
    cFeld = Text3(6).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If LoeseMarkenInArtnr1(cFeld) Then
            If bAnd Then
                cSQL = cSQL & " and "
            End If
            cSQL = cSQL & " a.artnr in (Select artnr from MA" & srechnertab & ") "
            bAnd = True
        Else
            anzeige "rot", "Keine Artikel!", Label9
            Exit Sub
        End If
    End If
    
    iStufe = 9
    If cAWM <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        
        cSQL = cSQL & " a.AWM = '" & cAWM & "' "
        bAnd = True
    End If
    
    iStufe = 10
    
    If cEAN <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        If Len(cEAN) > 6 Or (Len(cEAN) = 6 And Left(cEAN, 1) = "9") Then
        
            If Len(cEAN) = 8 Then
                If Left(cEAN, 1) = "2" Then
                    cEAN = Mid$(cEAN, 2, 6)
                    cSQL = cSQL & " a.ARTNR = " & cEAN
                    bAnd = True
                 Else
                    cSQL = cSQL & "(a.EAN = '" & cEAN & "' or a.EAN2 = '" & cEAN & "' or a.EAN3 = '" & cEAN & "' ) "
                    bAnd = True
                End If
            Else
                cSQL = cSQL & "(a.EAN = '" & cEAN & "' or a.EAN2 = '" & cEAN & "' or a.EAN3 = '" & cEAN & "' ) "
                bAnd = True
            
            End If
        Else
        '1.wunsch
            If IsNumeric(cEAN) Then
                cSQL = cSQL & " a.ARTNR = " & cEAN
                bAnd = True
            Else
                cSQL = cSQL & " a.ARTNR = -1 "
                bAnd = True
            End If
        End If
    End If
    
    If bAnd Then
        cSQL = cSQL & " and "
    End If
    cSQL = cSQL & "  ( A.SYNSTATUS is null or A.SYNSTATUS = 'E' or A.SYNSTATUS = 'A' ) "
    
    iStufe = 11

    
    List4.Nodes.Clear
    Screen.MousePointer = 11
    anzeige "ROT2", "Artikel werden gesucht...", Label9
    
    gdBase.Execute cSQL, dbFailOnError
 
    cSQL = "Select * from " & srechnertab & "ASEEK"
    
    If Option2(0).Value = True Then
        cSQL = cSQL & " order by BEZEICH "
    Else
        cSQL = cSQL & " order by LINR, LPZ, BEZEICH "
    End If
    
    
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    If Not rsrs.EOF Then
        If cEAN <> "" Then
            rsrs.MoveFirst
            iStufe = 13
            If Not IsNull(rsrs!linr) Then
                lLinr = rsrs!linr
            Else
                lLinr = 0
            End If
            iStufe = 14
            If Not IsNull(rsrs!LPZ) Then
                lLinie = rsrs!LPZ
            Else
                lLinie = 0
            End If
            iStufe = 15
            If Not IsNull(rsrs!artnr) Then
                cArtNrExakt = rsrs!artnr
            Else
                cArtNrExakt = ""
            End If
            cArtNrExakt = Space$(6 - Len(cArtNrExakt)) & cArtNrExakt
            rsrs.Close: Set rsrs = Nothing
            
            loeschNEW srechnertab & "ASEEK", gdBase
            CreateTable srechnertab & "ASEEK", gdBase
            
            cSQL = "Insert into " & srechnertab & "ASEEK Select "
            cSQL = cSQL & " ARTNR  "
            cSQL = cSQL & ", BEZEICH  "
            cSQL = cSQL & ", LEKPR  "
            cSQL = cSQL & ", EKPR  "
            cSQL = cSQL & ", KVKPR1  "
            cSQL = cSQL & ", VKPR  "
            cSQL = cSQL & ", RKZ  "
            cSQL = cSQL & ", BESTAND  "
            cSQL = cSQL & ", MWST "
            cSQL = cSQL & ", 1 as SEEKMOD  "
            cSQL = cSQL & ", AWM "
            cSQL = cSQL & ", LINR "
            cSQL = cSQL & ", LPZ "
            cSQL = cSQL & ", RABATT_OK "
            cSQL = cSQL & " from ARTIKEL where LINR = " & Trim$(Str$(lLinr)) & " and LPZ = " & Trim$(Str$(lLinie))
            cSQL = cSQL & " order by LINR, LPZ, BEZEICH "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Select * from " & srechnertab & "ASEEK"
            If Option2(0).Value = True Then
                cSQL = cSQL & " order by BEZEICH "
            Else
                cSQL = cSQL & " order by LINR, LPZ, BEZEICH "
            End If
            FnOpenrecordset rsrs, cSQL, 1, gdBase
            
            iStufe = 16
        End If
    End If
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iStufe = 17
            If Not IsNull(rsrs!MWST) Then
                cMWST = rsrs!MWST
            Else
                cMWST = "V"
            End If
            iStufe = 18
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cFeld & " "
            iStufe = 19
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            iStufe = 20
            If Not IsNull(rsrs!vkpr) Then
                dVkPr = rsrs!vkpr
            Else
                dVkPr = 0
            End If
            cFeld = Format$(dVkPr, "###,##0.00")
            cFeld = Trim$(cFeld)
            If Len(cFeld) > 9 Then
                cFeld = Space$(9)
            Else
                cFeld = Space$(9 - Len(cFeld)) & cFeld
            End If
            cLBSatz = cLBSatz & cFeld & " "
            iStufe = 21
            Select Case Val(Label8(3).Caption)
                Case Is = 0
                    If Not IsNull(rsrs!KVKPR1) Then
                        dKVkPr1 = rsrs!KVKPR1
                    Else
                        dKVkPr1 = 0
                    End If
                Case Is = 1
                    If Not IsNull(rsrs!vkpr) Then
                        dKVkPr1 = rsrs!vkpr
                    Else
                        dKVkPr1 = 0
                    End If
                Case Is = 2
                    If Not IsNull(rsrs!lekpr) Then
                        dKVkPr1 = rsrs!lekpr
                    Else
                        dKVkPr1 = 0
                    End If
                    If cMWST = "V" Then
                        dKVkPr1 = dKVkPr1 * dFaktorV
                    End If
                    If cMWST = "E" Then
                        dKVkPr1 = dKVkPr1 * dFaktorE
                    End If
                    
                Case Is = 3
                    If Not IsNull(rsrs!ekpr) Then
                        dKVkPr1 = rsrs!ekpr
                    Else
                        dKVkPr1 = 0
                    End If
                    If cMWST = "V" Then
                        dKVkPr1 = dKVkPr1 * dFaktorV
                    End If
                    If cMWST = "E" Then
                        dKVkPr1 = dKVkPr1 * dFaktorE
                    End If
                Case Is = 4 'Spez kvk
                    dKVkPr1 = LeseSpezpreis(CLng(rsrs!artnr), 0)
                    If dKVkPr1 = 0 Then
                        If Not IsNull(rsrs!KVKPR1) Then
                            dKVkPr1 = rsrs!KVKPR1
                        Else
                            dKVkPr1 = 0
                        End If
                    End If
                Case Is = 5 'lvk m A
                    If rsrs!RABATT_OK = "N" Then
                        If Not IsNull(rsrs!KVKPR1) Then
                            dKVkPr1 = rsrs!KVKPR1
                        Else
                            dKVkPr1 = 0
                        End If
                    Else
                        If Not IsNull(rsrs!vkpr) Then
                            dKVkPr1 = rsrs!vkpr
                        Else
                            dKVkPr1 = 0
                        End If
                    End If
                
            End Select
            iStufe = 22
            dKVkPr1 = dKVkPr1 * dRabatt
            cFeld = Format$(dKVkPr1, "###,##0.00")
            
            cFeld = Trim$(cFeld)
            cFeld = Space$(9 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            iStufe = 23
            If Not IsNull(rsrs!BESTAND) Then
                dVkPr = rsrs!BESTAND
            Else
                dVkPr = 0
            End If
            cFeld = Format$(dVkPr, "#,##0")
            cFeld = Trim$(cFeld)
            If Len(cFeld) > 5 Then cFeld = 0
            cFeld = Space$(5 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            iStufe = 24
            If Not IsNull(rsrs!RKZ) Then
                cFeld = rsrs!RKZ
            Else
                cFeld = "N"
            End If
            cFeld = Trim$(cFeld)
            cLBSatz = cLBSatz & cFeld & Space$(1)
            
            If Not IsNull(rsrs!AWM) Then
                cFeld = rsrs!AWM
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cLBSatz = cLBSatz & cFeld & Space(2 - Len(cFeld))
            
            
            cMerk = Left(ZeigeArtmerk(Trim(Left(cLBSatz, 6))), 1)
            cLBSatz = cLBSatz & Space(1 - Len(cMerk)) & cMerk
            
            iStufe = 25
            List4.Nodes.Add Text:=cLBSatz
            
            rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    iStufe = 26
    
    Screen.MousePointer = 11
    farbelist4 Me
    
    Dim cartT As String
    iStufe = 27
    If List4.Nodes.Count > 0 Then
        For lcount = 1 To List4.Nodes.Count
            ctmp = List4.Nodes(lcount)
            cartT = Trim$(Left(ctmp, 6))
            If cartT = Trim$(cArtNrExakt) Then
                List4.Nodes(lcount).Selected = True
                List4.Nodes(lcount).EnsureVisible
                Exit For
            End If
        Next lcount
        anzeige "normal", List4.Nodes.Count & " Artikel", Label9
    Else
        If cArtBez <> "" Then
            Command4(15).Visible = True
            Command4_Click 15
            Exit Sub
        Else
            anzeige "rot", "Keine Artikel!", Label9
        End If

    End If
    iStufe = 28
    Frame7.Visible = False
    List2.Visible = True
    List4.Visible = True
    
    Command4(10).Visible = True
    Command4(9).Visible = True
    Command4(1).Visible = True
    List4.SetFocus
    
    'noch weitere Ergebnisse
    If cArtBez <> "" Then
        Command4(15).Visible = True
        
    End If
    
    Screen.MousePointer = 0
    
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheArtikelKasseWKL20"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten. "
     
    Fehlermeldung1

    Resume Next
End Sub
Private Sub WKL142Positionieren()
On Error GoTo LOKAL_ERROR
    
    Frame6.Top = 0
    Frame6.Left = 0
    Frame6.Height = 9000
    Frame6.Width = 12000
    
    Frame7.Top = 4440
    Frame7.Left = 0
    Frame7.Height = 4455
    Frame7.Width = 11895
    
    List4.Height = 5295
    List4.Width = 11655
    List4.Top = 2400
    List4.Left = 120
    
    List2.Height = 870
    List2.Width = 11655
    List2.Top = 2160
    List2.Left = 120
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL142Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Kunde suchen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR
    
    loeschNEW srechnertab & "ASEEK", gdBase
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
Private Sub Label1_DblClick(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Index = 9 Then
        Label1(Index).Caption = "alle Farben"
        Label1(Index).Tag = ""
        Label1(Index).BackColor = Label5(5).BackColor
        Label1(Index).ForeColor = Label5(5).ForeColor
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_dblClick"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List4_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    
    If KeyCode = 13 Then
        Command4_Click 1
    End If
    If KeyCode = vbKeyEscape Then
        Command4_Click 2
    End If
    
    If KeyCode = vbKeyF2 Then
    
        If gbBEDKARTE = True Then
            
            If Identi(87) Then
            'ja man muss sich hier identifizieren
            
            
                'Ja
                Dim ctemp As String
                Dim bErlaubt As Boolean
                bErlaubt = False
                
                gcIdentUserName = ""
                gcIdentPass = ""
                gcIdentBedienerNr = ""
                glIdentLevel = -1
                    
                gsMeldestatus = "Identifikation"
                
                frmWK12a.Show 1
                
                If glIdentLevel = -1 Then
                    bErlaubt = False
                Else
                    
                    If gbZugriffNew Then
                        If glIdentLevel >= ermittlezugriff(87) Then
                            schreibeBEDIdentProtokoll "erfolgreich '" & gsProteil & "' geöffnet"
                            bErlaubt = True
                        Else
                            
                            ctemp = gcIdentUserName & " wurde identifiziert." & vbCrLf
                            ctemp = ctemp & gcIdentUserName & " hat nicht das Recht '" & gsProteil & "' zu nutzen."
                            schreibeBEDIdentProtokoll "versuchte erfolglos '" & gsProteil & "' durchzuführen"
                            MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
                            bErlaubt = False
                        End If
                    Else
                        bErlaubt = True
                    End If
                    
                End If
                
                If Not bErlaubt Then
                
                    If glIdentLevel = -1 Then
'                            MsgBox "Keine Zulassung für Stornos!", vbCritical, "STOP!"
                        End If
                    Exit Sub
                Else
                    
                End If
            End If
        Else
            If glLevel >= 5 Then
                
            Else
                MsgBox "Keine Zugangsberechtigung!", vbCritical, "KEIN ZUGRIFF"
                Exit Sub
            End If
        End If
    
        If List4.Nodes.Count > 0 Then
            For lcount = 1 To List4.Nodes.Count
                If List4.Nodes(lcount).Selected = True Then
                    gsARTNR = List4.Nodes(lcount)
                    gsARTNR = Trim$(Left(gsARTNR, 6))
                    Exit For
                End If
            Next lcount
        End If
        
        If gsARTNR <> "" Then
            frmWKL10.Show 1
            Me.Refresh
            Screen.MousePointer = 0
        End If
        gsARTNR = ""
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List4_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String
    
    If Index = 4 Then
        cValid = gcNUM & Chr$(8)
        
        
        cZeichen = Chr$(KeyAscii)
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        End If
    
    ElseIf Index = 5 Then
    
        cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
        cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
        cValid = cValid & "+äÄÜüÖöß"
        
        cZeichen = Chr$(KeyAscii)
        
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        End If
    

    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text3(Index).BackColor = glSelBack1
    Text3(Index).SelStart = 0
    Text3(Index).SelLength = Len(Text3(Index).Text)
    Label3(5).Caption = Trim$(Str$(Index))
    
    If bFocusonList4 = False Then
        Frame7.Visible = True
        List2.Visible = False
        List4.Visible = False

        Command4(10).Visible = False
        Command4(9).Visible = False
        Command4(1).Visible = False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Text3_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        If Index < 3 Or Index = 4 Or Index = 5 Then
            Command4_Click 0
        End If
        If Index = 3 Then
            Command4_Click 5
        End If
    End If
    
    If KeyCode = vbKeyEscape Then
        If Index < 3 Or Index = 4 Or Index = 5 Then
            Command4_Click 2
        End If
        If Index = 3 Then
            Command4_Click 3
        End If
    End If
    
    gF2Prompt.cFeld = ""
    gF2Prompt.cWert = ""
    gF2Prompt.cWert2 = ""
    gF2Prompt.cWahl = ""
    gF2Prompt.bMultiple = False
        
    If KeyCode = vbKeyF2 Then
        Select Case Index
            Case 4 'Pgn
                gF2Prompt.cFeld = "PGN"
            Case 5 'linr
                gF2Prompt.cFeld = "LINR"
            Case 6 'Marke
                gF2Prompt.cFeld = "MARKE"
        End Select
        
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
            If gF2Prompt.cWahl <> "" Then
                Text3(Index).SetFocus
                Text3(Index).Text = gF2Prompt.cWahl
                
            End If
            
            If Index = 5 Then
                Label5(3).Caption = gF2Prompt.cWert
                Label5(3).Refresh
                
            End If
            
            If Index = 4 Then
                Label5(8).Caption = gF2Prompt.cWert
                Label5(8).Refresh
                
            End If
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text3(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_Change(Index As Integer)
On Error GoTo LOKAL_ERROR
    Dim sNeuLinr As String
    Dim sNeuPGNNr As String
    Dim searchstr As String

    If Index = 5 Then
        LiefKuerzelAufloesung Label5(3), Text3(5)
    End If
    
    If Index = 4 Then
        If Len(Text3(4).Text) = 0 Then
            Label5(8).Caption = "keine Auswahl"
            Label5(8).Refresh
        End If
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_Change"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
