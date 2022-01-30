VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWKL68 
   Caption         =   "Bezahlen"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "frmWKL68.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame5 
      Appearance      =   0  '2D
      BackColor       =   &H000000FF&
      Caption         =   "Wichtige Gutscheininformationen"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   -720
      TabIndex        =   170
      Top             =   8040
      Visible         =   0   'False
      Width           =   1575
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1770
         Index           =   9
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   171
         Top             =   840
         Width           =   4815
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   495
         Index           =   26
         Left            =   4440
         TabIndex        =   172
         Top             =   240
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "x"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'Kein
      Caption         =   "Frame4"
      Height          =   12255
      Left            =   0
      TabIndex        =   153
      Top             =   8040
      Width           =   6135
      Begin VB.Frame Frame33 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   7095
         Left            =   1560
         TabIndex        =   154
         Top             =   120
         Width           =   9495
         Begin sevCommand3.Command Command33 
            Height          =   735
            Index           =   1
            Left            =   3720
            TabIndex        =   156
            Top             =   3360
            Width           =   3255
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin VB.TextBox Text333 
            Alignment       =   1  'Rechts
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   555
            Left            =   2880
            MaxLength       =   9
            TabIndex        =   155
            Top             =   2280
            Width           =   3375
         End
         Begin VB.Label Label77 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00C0E0FF&
            Caption         =   "Zahlung in Euro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   240
            TabIndex        =   168
            Top             =   120
            Width           =   6615
         End
         Begin VB.Label Label33 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "€"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   25
            Left            =   6600
            TabIndex        =   167
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label333 
            Alignment       =   1  'Rechts
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   5
            Left            =   3600
            TabIndex        =   166
            Top             =   4680
            Width           =   2655
         End
         Begin VB.Label Label333 
            Alignment       =   1  'Rechts
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   6
            Left            =   2280
            TabIndex        =   165
            Top             =   840
            Width           =   3975
         End
         Begin VB.Label Label33 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "Zurück:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   26
            Left            =   1200
            TabIndex        =   164
            Top             =   4680
            Width           =   1935
         End
         Begin VB.Label Label33 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Gegeben..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Index           =   27
            Left            =   240
            TabIndex        =   163
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label Label33 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "Summe:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Index           =   28
            Left            =   240
            TabIndex        =   162
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label33 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "€"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   29
            Left            =   6600
            TabIndex        =   161
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label33 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "€"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   30
            Left            =   6600
            TabIndex        =   160
            Top             =   4680
            Width           =   375
         End
         Begin VB.Label Label33 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "noch offen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   31
            Left            =   240
            TabIndex        =   159
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label333 
            Alignment       =   1  'Rechts
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   7
            Left            =   2880
            TabIndex        =   158
            Top             =   1440
            Width           =   3375
         End
         Begin VB.Label Label33 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "€"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Index           =   32
            Left            =   6600
            TabIndex        =   157
            Top             =   1440
            Width           =   375
         End
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  '2D
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   5160
      TabIndex        =   145
      Top             =   8040
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox Text1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         Index           =   8
         Left            =   1200
         MaxLength       =   13
         TabIndex        =   147
         Top             =   1800
         Width           =   2055
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   495
         Index           =   24
         Left            =   3360
         TabIndex        =   148
         Top             =   1800
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Ok"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nummer"
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
         Left            =   1200
         TabIndex        =   152
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nr des Gutscheins:"
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
         Index           =   35
         Left            =   1200
         TabIndex        =   151
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Wert des Gutscheins"
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
         Index           =   34
         Left            =   1200
         TabIndex        =   149
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "alter Gutschein"
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
         Index           =   57
         Left            =   120
         TabIndex        =   146
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.Frame Frame18 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      Height          =   1455
      Left            =   10800
      TabIndex        =   86
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
      Begin VB.Frame Frame20 
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
         Height          =   2655
         Left            =   0
         TabIndex        =   87
         Top             =   4320
         Visible         =   0   'False
         Width           =   7935
         Begin sevCommand3.Command Command13 
            Height          =   735
            Index           =   11
            Left            =   3960
            TabIndex        =   123
            Top             =   3480
            Width           =   720
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin sevCommand3.Command Command13 
            Height          =   735
            Index           =   10
            Left            =   3960
            TabIndex        =   122
            Top             =   4200
            Width           =   720
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin sevCommand3.Command Command13 
            Height          =   735
            Index           =   9
            Left            =   3240
            TabIndex        =   121
            Top             =   4200
            Width           =   720
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin sevCommand3.Command Command13 
            Height          =   735
            Index           =   8
            Left            =   2520
            TabIndex        =   120
            Top             =   4200
            Width           =   720
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin sevCommand3.Command Command13 
            Height          =   735
            Index           =   7
            Left            =   1800
            TabIndex        =   119
            Top             =   4200
            Width           =   720
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin sevCommand3.Command Command13 
            Height          =   735
            Index           =   6
            Left            =   1080
            TabIndex        =   118
            Top             =   4200
            Width           =   720
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin sevCommand3.Command Command13 
            Height          =   735
            Index           =   5
            Left            =   360
            TabIndex        =   117
            Top             =   4200
            Width           =   720
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin sevCommand3.Command Command13 
            Height          =   735
            Index           =   4
            Left            =   3240
            TabIndex        =   116
            Top             =   3480
            Width           =   720
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin sevCommand3.Command Command13 
            Height          =   735
            Index           =   3
            Left            =   2520
            TabIndex        =   115
            Top             =   3480
            Width           =   720
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin sevCommand3.Command Command13 
            Height          =   735
            Index           =   2
            Left            =   1800
            TabIndex        =   114
            Top             =   3480
            Width           =   720
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin sevCommand3.Command Command13 
            Height          =   735
            Index           =   1
            Left            =   1080
            TabIndex        =   113
            Top             =   3480
            Width           =   720
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin sevCommand3.Command Command13 
            Height          =   735
            Index           =   0
            Left            =   360
            TabIndex        =   112
            Top             =   3480
            Width           =   720
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   18
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
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   93
            Text            =   "Text6"
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   2280
            MaxLength       =   8
            TabIndex        =   92
            Text            =   "Text6"
            Top             =   1440
            Width           =   2415
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   91
            Text            =   "Text6"
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   3
            Left            =   3840
            MaxLength       =   2
            TabIndex        =   90
            Text            =   "Text6"
            Top             =   2040
            Width           =   855
         End
         Begin sevCommand3.Command Command12 
            Height          =   735
            Index           =   0
            Left            =   360
            TabIndex        =   89
            Top             =   2640
            Width           =   2055
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
            Caption         =   "OK"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command12 
            Height          =   735
            Index           =   1
            Left            =   2640
            TabIndex        =   88
            Top             =   2640
            Width           =   2055
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
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
         Begin VB.Label Label12 
            BackColor       =   &H00FFFF00&
            Caption         =   "-1"
            Height          =   255
            Left            =   9000
            TabIndex        =   124
            Top             =   3720
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C000&
            BackStyle       =   0  'Transparent
            Caption         =   "Kontonummer:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   360
            TabIndex        =   98
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C000&
            BackStyle       =   0  'Transparent
            Caption         =   "Bankleitzahl:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   360
            TabIndex        =   97
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C000&
            BackStyle       =   0  'Transparent
            Caption         =   "gültig bis:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   360
            TabIndex        =   96
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   3240
            TabIndex        =   95
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Zentriert
            BackStyle       =   0  'Transparent
            Caption         =   "manuelle Eingabe EC-Karte"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   375
            Left            =   120
            TabIndex        =   94
            Top             =   120
            Width           =   4815
         End
      End
      Begin VB.Frame Frame19 
         BorderStyle     =   0  'Kein
         Height          =   2295
         Left            =   600
         TabIndex        =   104
         Top             =   3240
         Width           =   3735
         Begin VB.ListBox List11 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2040
            Left            =   120
            TabIndex        =   105
            Top             =   120
            Width           =   3495
         End
      End
      Begin VB.Frame Frame22 
         BorderStyle     =   0  'Kein
         Height          =   2655
         Left            =   600
         TabIndex        =   99
         Top             =   5640
         Width           =   3735
         Begin sevCommand3.Command Command10 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   103
            Top             =   840
            Width           =   3495
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
            Caption         =   "manuelle Eingabe"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command10 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   102
            Top             =   2040
            Width           =   3495
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
            Caption         =   "Abbrechen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command10 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   101
            Top             =   1440
            Width           =   3495
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
            Caption         =   "Eingabe löschen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command10 
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   3495
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
            Caption         =   "Beleg drucken"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   120
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         CommPort        =   2
         DTREnable       =   -1  'True
         Handshaking     =   1
         RThreshold      =   1
         ParitySetting   =   2
         SThreshold      =   1
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte die EC-Karte durch den Kartenleser ziehen!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   600
         TabIndex        =   110
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "zu zahlender Betrag:"
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
         Index           =   1
         Left            =   600
         TabIndex        =   109
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Zahlung über EC-Lastschrift"
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
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   108
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   107
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Kunde / Kontoinhaber:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   106
         Top             =   2880
         Visible         =   0   'False
         Width           =   3735
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   255
      Index           =   1
      Left            =   10800
      TabIndex        =   78
      Top             =   8040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
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
      Caption         =   "Zurück"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'Kein
      Caption         =   "Kreditkartendetails"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   5040
      TabIndex        =   76
      Top             =   6240
      Visible         =   0   'False
      Width           =   1695
      Begin VB.CheckBox Check15 
         BackColor       =   &H00808000&
         Caption         =   "manuelle Eingabe der Kreditkartendaten"
         Height          =   975
         Left            =   4080
         TabIndex        =   180
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4080
         TabIndex        =   173
         Top             =   1920
         Visible         =   0   'False
         Width           =   2175
         Begin VB.ComboBox cboVerfall 
            Height          =   315
            Left            =   3480
            TabIndex        =   176
            Text            =   "Combo2"
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   9
            Left            =   1560
            MaxLength       =   35
            TabIndex        =   175
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   8
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   174
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "gültig bis:"
            Height          =   255
            Index           =   10
            Left            =   2760
            TabIndex        =   179
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Kartennummer:"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   178
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Sicherheitscode:"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   177
            Top             =   600
            Width           =   1335
         End
      End
      Begin sevCommand3.Command SSCommand8 
         Height          =   1095
         Index           =   6
         Left            =   2040
         TabIndex        =   141
         ToolTipText     =   "Sonstige"
         Top             =   3360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1931
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
         ToolTipTitle    =   "Sonstige"
         ButtonStyle     =   2
         Caption         =   "Sonstige"
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand8 
         Height          =   1095
         Index           =   5
         Left            =   120
         TabIndex        =   140
         ToolTipText     =   "EC-Karte"
         Top             =   3360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1931
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
         ToolTipTitle    =   "EC-Karte"
         ButtonStyle     =   2
         Caption         =   "EC-Karte"
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand8 
         Height          =   615
         Index           =   4
         Left            =   2880
         TabIndex        =   139
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
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
         Caption         =   "Barclay Card"
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand8 
         Height          =   1095
         Index           =   3
         Left            =   2040
         TabIndex        =   138
         ToolTipText     =   "Diners Club"
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1931
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
         ToolTipTitle    =   "Diners Club"
         ButtonStyle     =   2
         Caption         =   "Diners Club"
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand8 
         Height          =   1095
         Index           =   2
         Left            =   120
         TabIndex        =   137
         ToolTipText     =   "American Express"
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1931
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
         ToolTipTitle    =   "American Express"
         ButtonStyle     =   2
         Caption         =   "American Express"
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand8 
         Height          =   1095
         Index           =   1
         Left            =   2040
         TabIndex        =   136
         ToolTipText     =   "Eurocard / Mastercard"
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1931
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
         ToolTipTitle    =   "Eurocard / Mastercard"
         ButtonStyle     =   2
         Caption         =   "Eurocard / Mastercard"
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand8 
         Height          =   1095
         Index           =   0
         Left            =   120
         TabIndex        =   135
         ToolTipText     =   "Visa"
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1931
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
         ToolTipTitle    =   "Visa"
         ButtonStyle     =   2
         Caption         =   "Visa"
         Picture         =   "frmWKL68.frx":0442
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand8 
         Height          =   1095
         Index           =   7
         Left            =   960
         TabIndex        =   181
         ToolTipText     =   "Automatisch"
         Top             =   2040
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1931
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
         ToolTipTitle    =   "Automatisch"
         ButtonStyle     =   2
         Caption         =   "Automatisch"
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Kreditkarte auswählen!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   77
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '2D
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   6720
      TabIndex        =   50
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   120
         TabIndex        =   82
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "GutscheinNr:"
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
         Index           =   25
         Left            =   2520
         TabIndex        =   74
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bedienername:"
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
         Index           =   24
         Left            =   2520
         TabIndex        =   73
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "GutscheinNr:"
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
         Index           =   23
         Left            =   2520
         TabIndex        =   72
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Wert:"
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
         Index           =   22
         Left            =   2520
         TabIndex        =   71
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "GutscheinNr:"
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
         Index           =   21
         Left            =   120
         TabIndex        =   70
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BedNr(Ausgabe):"
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
         Index           =   20
         Left            =   120
         TabIndex        =   69
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "GutscheinNr:"
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
         Index           =   19
         Left            =   120
         TabIndex        =   68
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sraße:"
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
         Index           =   18
         Left            =   120
         TabIndex        =   67
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "GutscheinNr:"
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
         Index           =   17
         Left            =   2520
         TabIndex        =   66
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ort:"
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
         Index           =   16
         Left            =   2520
         TabIndex        =   65
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "GutscheinNr:"
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
         Left            =   2520
         TabIndex        =   64
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Index           =   14
         Left            =   2520
         TabIndex        =   63
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "GutscheinNr:"
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
         Left            =   2520
         TabIndex        =   62
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Titel:"
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
         Index           =   12
         Left            =   2520
         TabIndex        =   61
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "GutscheinNr:"
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
         Index           =   11
         Left            =   2520
         TabIndex        =   60
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ausgabe am:"
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
         Index           =   10
         Left            =   2520
         TabIndex        =   59
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "GutscheinNr:"
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
         Index           =   9
         Left            =   120
         TabIndex        =   58
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Plz:"
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
         Index           =   8
         Left            =   120
         TabIndex        =   57
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "GutscheinNr:"
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
         Left            =   120
         TabIndex        =   56
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Vorname:"
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
         Index           =   6
         Left            =   120
         TabIndex        =   55
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "GutscheinNr:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   54
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "KundNr:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   53
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "GutscheinNr:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GutscheinNr:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   51
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   6
      Left            =   3000
      MaxLength       =   9
      TabIndex        =   5
      Top             =   5280
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   5
      Left            =   3000
      MaxLength       =   9
      TabIndex        =   6
      Top             =   5745
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   4
      Left            =   7320
      MaxLength       =   13
      TabIndex        =   3
      Top             =   4305
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   3
      Left            =   3000
      MaxLength       =   9
      TabIndex        =   4
      Top             =   4800
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   2
      Left            =   3000
      MaxLength       =   9
      TabIndex        =   1
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Index           =   0
      Left            =   3000
      MaxLength       =   9
      TabIndex        =   0
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   1
      Left            =   3000
      MaxLength       =   9
      TabIndex        =   7
      Top             =   4320
      Width           =   3135
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   8
      Top             =   7200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
      Caption         =   "Abschließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   800
      Index           =   11
      Left            =   5040
      TabIndex        =   24
      Top             =   6240
      Width           =   920
      _Version        =   65536
      _ExtentX        =   1623
      _ExtentY        =   1411
      _StockProps     =   78
      Caption         =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   800
      Index           =   10
      Left            =   5040
      TabIndex        =   25
      Top             =   7080
      Width           =   920
      _Version        =   65536
      _ExtentX        =   1623
      _ExtentY        =   1411
      _StockProps     =   78
      Caption         =   ","
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   800
      Index           =   9
      Left            =   4080
      TabIndex        =   26
      Top             =   7080
      Width           =   920
      _Version        =   65536
      _ExtentX        =   1623
      _ExtentY        =   1411
      _StockProps     =   78
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   800
      Index           =   8
      Left            =   3120
      TabIndex        =   27
      Top             =   7080
      Width           =   920
      _Version        =   65536
      _ExtentX        =   1623
      _ExtentY        =   1411
      _StockProps     =   78
      Caption         =   "9"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   800
      Index           =   7
      Left            =   2160
      TabIndex        =   28
      Top             =   7080
      Width           =   920
      _Version        =   65536
      _ExtentX        =   1623
      _ExtentY        =   1411
      _StockProps     =   78
      Caption         =   "8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   800
      Index           =   6
      Left            =   1200
      TabIndex        =   29
      Top             =   7080
      Width           =   920
      _Version        =   65536
      _ExtentX        =   1623
      _ExtentY        =   1411
      _StockProps     =   78
      Caption         =   "7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   800
      Index           =   5
      Left            =   240
      TabIndex        =   30
      Top             =   7080
      Width           =   920
      _Version        =   65536
      _ExtentX        =   1623
      _ExtentY        =   1411
      _StockProps     =   78
      Caption         =   "6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   800
      Index           =   4
      Left            =   4080
      TabIndex        =   31
      Top             =   6240
      Width           =   920
      _Version        =   65536
      _ExtentX        =   1623
      _ExtentY        =   1411
      _StockProps     =   78
      Caption         =   "5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   800
      Index           =   3
      Left            =   3120
      TabIndex        =   32
      Top             =   6240
      Width           =   920
      _Version        =   65536
      _ExtentX        =   1623
      _ExtentY        =   1411
      _StockProps     =   78
      Caption         =   "4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   800
      Index           =   2
      Left            =   2160
      TabIndex        =   33
      Top             =   6240
      Width           =   920
      _Version        =   65536
      _ExtentX        =   1623
      _ExtentY        =   1411
      _StockProps     =   78
      Caption         =   "3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   800
      Index           =   1
      Left            =   1200
      TabIndex        =   34
      Top             =   6240
      Width           =   920
      _Version        =   65536
      _ExtentX        =   1623
      _ExtentY        =   1411
      _StockProps     =   78
      Caption         =   "2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   800
      Index           =   0
      Left            =   240
      TabIndex        =   35
      Top             =   6240
      Width           =   920
      _Version        =   65536
      _ExtentX        =   1623
      _ExtentY        =   1411
      _StockProps     =   78
      Caption         =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   375
      Index           =   12
      Left            =   2160
      TabIndex        =   38
      Top             =   2880
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   375
      Index           =   13
      Left            =   2160
      TabIndex        =   39
      Top             =   4320
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   375
      Index           =   14
      Left            =   2160
      TabIndex        =   40
      Top             =   3360
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   375
      Index           =   15
      Left            =   2160
      TabIndex        =   41
      Top             =   4800
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   495
      Index           =   16
      Left            =   6720
      TabIndex        =   42
      Top             =   4305
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   375
      Index           =   17
      Left            =   2160
      TabIndex        =   45
      Top             =   5760
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   495
      Index           =   18
      Left            =   9480
      TabIndex        =   46
      ToolTipText     =   "Auswählen - Eingetragene Gutscheinnummer wird ausgewählt, selbe Funktion wie ""Enter""-Taste"
      Top             =   4305
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "A"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   375
      Index           =   19
      Left            =   2160
      TabIndex        =   47
      Top             =   5280
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   375
      Index           =   20
      Left            =   2160
      TabIndex        =   79
      Top             =   3840
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   495
      Index           =   22
      Left            =   11160
      TabIndex        =   134
      ToolTipText     =   "Suche - Zeigt alle offenen Gutscheine an"
      Top             =   4305
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "S"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   495
      Index           =   21
      Left            =   6960
      TabIndex        =   142
      Top             =   5740
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "EC Last abrechnen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   495
      Index           =   23
      Left            =   10200
      TabIndex        =   143
      ToolTipText     =   "Gutschein ohne Nummer - Eingetragener Wert wird als Gutscheinwert übernommen, ohne Gutscheinnummer"
      Top             =   4320
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "G"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   375
      Index           =   25
      Left            =   2160
      TabIndex        =   150
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Bar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   7
      Left            =   3000
      MaxLength       =   9
      TabIndex        =   2
      Top             =   3840
      Width           =   3135
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000080&
      Caption         =   "Geldbeträge mit Komma eingeben!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   23
      Left            =   120
      TabIndex        =   169
      Top             =   2520
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "G ohne Nr."
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
      Index           =   33
      Left            =   10200
      TabIndex        =   144
      Top             =   4065
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Index           =   32
      Left            =   6120
      TabIndex        =   133
      Top             =   7560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Index           =   31
      Left            =   6120
      TabIndex        =   132
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Index           =   30
      Left            =   6120
      TabIndex        =   131
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   22
      Left            =   6240
      TabIndex        =   130
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   21
      Left            =   6240
      TabIndex        =   129
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label333 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000080&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   128
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "gegeben insg."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   20
      Left            =   120
      TabIndex        =   127
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Restgutschein"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   29
      Left            =   120
      TabIndex        =   126
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000080&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   28
      Left            =   3000
      TabIndex        =   125
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Index           =   27
      Left            =   6120
      TabIndex        =   111
      Top             =   7920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "zurück in Bar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   85
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label333 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000080&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   84
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   18
      Left            =   6240
      TabIndex        =   83
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "2. Karte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   17
      Left            =   120
      TabIndex        =   81
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   16
      Left            =   6240
      TabIndex        =   80
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Suche"
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
      Left            =   11160
      TabIndex        =   75
      Top             =   4065
      Width           =   615
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "Scheck"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   15
      Left            =   120
      TabIndex        =   49
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   14
      Left            =   6240
      TabIndex        =   48
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "EC Last"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   13
      Left            =   120
      TabIndex        =   44
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   12
      Left            =   6240
      TabIndex        =   43
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Focus"
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   37
      Top             =   6480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Gutscheinnummer:"
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
      Left            =   7320
      TabIndex        =   36
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   11
      Left            =   6240
      TabIndex        =   23
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   10
      Left            =   6240
      TabIndex        =   22
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   9
      Left            =   6240
      TabIndex        =   21
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "Dukaten"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   20
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "1. Karte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "Gutschein"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      Index           =   1
      X1              =   120
      X2              =   6600
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   6240
      TabIndex        =   17
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label333 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000080&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   16
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "noch offen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   7
      Left            =   6240
      TabIndex        =   14
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "Summe:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "Bargeld"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label333 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000080&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label33 
      BackColor       =   &H00000080&
      Caption         =   "€"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   6240
      TabIndex        =   10
      Top             =   120
      Width           =   375
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
      TabIndex        =   9
      Top             =   7920
      Width           =   9375
   End
End
Attribute VB_Name = "frmWKL68"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Gutschl(20) As GUTSCHEIN
Dim dGutschwert As Double
Dim dGutscheinauszahlung As Double
Dim glnewGutschnr As Long
Dim bAlterGutscheinImSpiel As Boolean
Private Function fnSchreibeDTADatenWKL20() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz        As Long
    Dim lDatum          As Long
    Dim cSQL            As String
    Dim cEmpfaenger     As String
    Dim cText           As String
    Dim cUhrZeit        As String
    Dim dBetrag         As Double
    Dim rsrs            As Recordset
    
    fnSchreibeDTADatenWKL20 = 0
    
    lDatum = Fix(Now)
    cUhrZeit = Format$(Now, "HH:MM:SS")
    
    cText = Label11(3).Caption
    cText = fnMoveComma2Point$(cText)
    dBetrag = Val(cText)
    cText = ""
    
    cSQL = "Select count(*) as ANZ_SATZ from DTA"
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!ANZ_SATZ) Then
            lAnzSatz = rsrs!ANZ_SATZ
        Else
            lAnzSatz = 0
        End If
    Else
        lAnzSatz = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
    lAnzSatz = lAnzSatz + 1
    
    cSQL = "Select * from DTA where BELEGNR = -1"
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    rsrs.AddNew
    rsrs!Empfaenger = "EC-KARTE"
    rsrs!Betrag = dBetrag
    rsrs!TextA = "5"
    rsrs!BLZ = gECKarte.BLZ
    rsrs!EKONTO = gECKarte.Konto1
    rsrs!BELEGNR = lAnzSatz
    rsrs!Datum = lDatum
    rsrs!Uhrzeit = cUhrZeit
    rsrs!SENDOK = False
    rsrs!FILIALE = CInt(gcFilNr)
    rsrs!zweck1 = gcKasNum & "9999"
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing

    cSQL = "Select * from DTA where BELEGNR = " & Trim$(Str$(lAnzSatz)) 'Lesen, ob Satz wirklich angekommen
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    If rsrs.EOF Then
        fnSchreibeDTADatenWKL20 = 1
    End If
    rsrs.Close: Set rsrs = Nothing
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDTADatenWKL20"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub KorrigiereDTADaten68(cBel As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim sWas            As String
    
    sWas = gcKasNum & "9999"
    cSQL = "Select * from DTA where zweck1 = '" & sWas & "'"
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    If Not rsrs.EOF Then
    
    
    
        Do While Not rsrs.EOF
            rsrs.Edit
        
            rsrs!zweck1 = cBel
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
    Fehler.gsFunktion = "KorrigiereDTADaten68"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub KorrigiereAlterG68(cBel As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim sWas            As String
    
    sWas = gcKasNum & "9999"
    cSQL = "Select * from ALTERG where BELEGNR = " & sWas
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    If Not rsrs.EOF Then
    
    
        Do While Not rsrs.EOF
            rsrs.Edit
        
            rsrs!BELEGNR = cBel
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
    Fehler.gsFunktion = "KorrigiereAlterG68"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DElausAlterG68()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim sWas            As String
    
    sWas = gcKasNum & "9999"
    cSQL = "Delete from ALTERG where BELEGNR = " & sWas
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DElausAlterG68"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check15_Click()
On Error GoTo LOKAL_ERROR
    
    If Check15.value = vbChecked Then
        Frame16.Visible = True
        
        fuellecbo_Verfall
    Else
        Frame16.Visible = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check15_Click"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecbo_Verfall()
    On Error GoTo LOKAL_ERROR
    
    Dim j As Integer
    Dim m As Integer
    
    For m = Month(DateValue(Now)) To 12
        cboVerfall.AddItem m & "/" & Year(DateValue(Now))
    Next m
    
    For j = Year(DateValue(Now)) + 1 To Year(DateValue(Now)) + 10
        For m = 1 To 12
            cboVerfall.AddItem m & "/" & j
        Next m
    Next j
    
    cboVerfall.Text = Month(DateValue(Now)) & "/" & Year(DateValue(Now))
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecbo_Verfall"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command10_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    
    Select Case index
        Case Is = 0     'Speichern
            If gECKarte.Datenstrom <> "" Then
                iRet = fnPruefeDatenstromECKarteWKL20()
                If iRet = 0 Then                'jetzt Bon drucken und Daten schreiben
                    iRet = DruckeKassenBonECLastWKL68()
                    If iRet = 0 Then
                        iRet = fnSchreibeDTADatenWKL20()
                        If iRet = 0 Then
                            List11.Clear
                            
                            SSCommand6(21).Visible = False
                            Command5(1).Visible = False
                            
                            Text1(5).Enabled = False

                            Command10_Click 3
                        Else
                            MsgBox "Daten konnten nicht gespeichert werden!", vbCritical, "STOP!"
                            List11.SetFocus
                        End If
                    End If
                End If
            Else
                List11.SetFocus
            End If
 
            
        Case Is = 2     'Löschen
            LeereDatenECKarteWKL20
            List11.Clear
            Label11(0).Visible = True
            Label11(4).Visible = False
            
'            Label13(0).Visible = False
'            Label13(1).Visible = False
'            Label13(2).Visible = False
            
            List11.SetFocus

        Case Is = 3     'Schließen
'            back2 3
            Frame18.Visible = False
            MSComm1.PortOpen = False
            anzeige "normal", "", Label9

        Case Is = 4     'Manuelle Eingabe
            LeereDatenECKarteWKL20
            List11.Clear

            Text6(0).Text = ""
            Text6(1).Text = ""
            Text6(2).Text = ""
            Text6(3).Text = ""
            Frame20.Visible = True
            Frame19.Visible = False
            Frame22.Visible = False
            
            Text6(0).SetFocus
    End Select
Exit Sub
LOKAL_ERROR:
    If err.Number = 8005 Or err.Number = 8012 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command10_Click"
        Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub BildeDatenStromWKL20()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatenStrom As String
    Dim ctmp        As String
    Dim iRet        As Integer
    
    cDatenStrom = "67" & String$(6, "0")
    ctmp = Text6(0).Text
    ctmp = Trim$(ctmp)
    ctmp = String$(10 - Len(ctmp), "0") & ctmp
    cDatenStrom = cDatenStrom & ctmp & vbCr
    
    ctmp = Text6(3).Text
    ctmp = Trim$(ctmp)
    ctmp = String$(2 - Len(ctmp), "0") & ctmp
    cDatenStrom = cDatenStrom & ctmp
    
    ctmp = Text6(2).Text
    ctmp = Trim$(ctmp)
    ctmp = String$(2 - Len(ctmp), "0") & ctmp
    cDatenStrom = cDatenStrom & ctmp
    
    ctmp = Text6(1).Text
    ctmp = Trim$(ctmp) & String$(1, "0")
    
    cDatenStrom = cDatenStrom & ctmp
    
    ctmp = Text6(0).Text
    ctmp = Trim$(ctmp)
    ctmp = String$(10 - Len(ctmp), "0") & ctmp
    
    cDatenStrom = cDatenStrom & ctmp & vbCrLf
    
    gECKarte.Datenstrom = cDatenStrom
    
    iRet = fnPruefeDatenstromECKarteWKL20()
    If iRet <> 0 Then
        LeereDatenECKarteWKL20
    Else
        'jetzt optional nach Hakensetzung
        'Lucks
        If gbNachKBbeiEC Then
            If Val(gckundnr) > 0 Then
                Label1(27).Caption = gckundnr
                gckundnr = ""
            End If
        End If
        LeseDatenECLastschriftWKL20
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BildeDatenStrom"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command12_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    
    Select Case index
        Case Is = 0
            iRet = fnPruefeEingabeECKarte()
            Select Case iRet
                Case Is = 0
                    BildeDatenStromWKL20
                    If List11.ListCount > 0 Then
                        Command12_Click 1
                    End If
                Case Is = 1
                    MsgBox "Bitte die Kontonummer eingeben!", vbCritical, "STOP!"
                    Text6(0).SetFocus
                    Exit Sub
                Case Is = 2
                    MsgBox "Bitte die Bankleitzahl eingeben!", vbCritical, "STOP!"
                    Text6(1).SetFocus
                    Exit Sub
                Case Is = 3
                    MsgBox "Bitte die Gültigkeit (Format MM / JJ) eingeben!", vbCritical, "STOP!"
                    Text6(2).SetFocus
                    Exit Sub
                Case Is = 4
                    MsgBox "Bitte die Gültigkeit (Format MM / JJ) eingeben!", vbCritical, "STOP!"
                    Text6(3).SetFocus
                    Exit Sub
                Case Is = 5
                    MsgBox "Der eingegebene Monat ist falsch!", vbCritical, "STOP!"
                    Text6(2).SetFocus
                    Exit Sub
            End Select
        Case Is = 1
            Frame19.Visible = True
            Frame22.Visible = True
            Frame20.Visible = False
            
            List11.SetFocus
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command12_Click"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Sub
Private Sub Command13_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If index < 10 Then
        Text6(Label12.Caption).Text = Text6(Label12.Caption).Text & Command13(index).Caption
    ElseIf index = 10 Then
        Text6(Label12.Caption).Text = ""
    ElseIf index = 11 Then
        If Len(Text6(Label12.Caption).Text) > 0 Then
            Text6(Label12.Caption).Text = Left(Text6(Label12.Caption).Text, Len(Text6(Label12.Caption).Text) - 1)
        End If
    End If
    
    Text6(Label12.Caption).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command13_Click"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command33_Click(index As Integer)
On Error GoTo LOKAL_ERROR
      
    Unload frmWKL68
         
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command33_Click"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(index As Integer)
On Error GoTo LOKAL_ERROR

anzeige "normal", "", Label9

Select Case index
    Case 0
        Command5(1).Visible = False
        Command5(0).Visible = False
        If allesPrüfen Then
        
            gdRückgeldaus68 = CDbl(Label333(3).Caption)
            If abschliessen Then
            
                zeigabschlussframe
                
            Else
                Command5(0).Visible = True
            End If
        Else
            Command5(1).Visible = True
            Command5(0).Visible = True
        End If
        gbBackaus68 = False
    Case 1 'zurück
    
        gcKreditKarte = ""
        gcKreditKarte2 = ""
    
        gbBackaus68 = True
        
    
        DElausAlterG68
        Unload frmWKL68
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigabschlussframe()
On Error GoTo LOKAL_ERROR

Dim i As Integer

Frame4.Visible = True
Frame4.BackColor = &HC0C0C0    '&H8000000F
Frame33.BackColor = &HC0E0FF

SSCommand6(25).Visible = False

For i = 25 To 32
    Label33(i).ForeColor = &H0&
    Label33(i).BackColor = &H80&

Next i

For i = 5 To 7
    Label333(i).ForeColor = &H0&
    Label333(i).BackColor = &HC0E0FF
Next i



Label77.ForeColor = &H0&
Label77.BackColor = &HC0E0FF

Label333(5).Caption = Label333(3).Caption
Label333(6).Caption = Label333(0).Caption
Text333.Text = Label333(1).Caption

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigabschlussframe"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub AktualisiereZahlung68(cOperation As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lHeute          As Long
    Dim cSQL            As String
    Dim cGutschnr       As String
    Dim i               As Integer
    
    lHeute = Fix(Now)
    
    For i = 0 To 19
        If Gutschl(i).gutschnr <> 0 Then
            cGutschnr = Gutschl(i).gutschnr
            
            'neuen Wert für bereits durch Gutscheine gezahlten Betrag ermitteln
            If cOperation = "+" Then
            
                Gutschein_Einloesen cGutschnr, lHeute, Format$(Now, "HH:MM:SS"), CStr(gdBonNr), gcKasNum, gcBedienerNr
                
                insertGUTschHIS Fix(Now), Format$(Now, "HH:MM:SS"), CStr(gdBonNr), gcKasNum, "EI", 0, cGutschnr, gcBedienerNr
            End If
            
            If cOperation = "-" Then
            
                Gutschein_Einloesen_Aufheben cGutschnr
            
                cSQL = "Delete from GUHIS where GUTSCHNRO = " & cGutschnr
                gdBase.Execute cSQL, dbFailOnError
            End If
        
        End If
    Next i
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AktualisiereZahlung68"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Function abschliessen() As Boolean
On Error GoTo LOKAL_ERROR

    abschliessen = False
    
    Dim cPfad As String
    Dim iZaehler As Integer
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    iZaehler = 1
    
    'die Kasse Stop Funktion für eine bestimmte Kasse
    
    
    Do While FileExists(cPfad & "KASSSTOP_ALLE.TXT")
        Pause 1
        gsAnzeigeText = "Die Kasse wird für ein paar Sekunden unterbrochen," & vbCrLf
        gsAnzeigeText = gsAnzeigeText & "da an einem anderen Rechner ein KASSENABSCHLUSS durchgeführt wird." & vbCrLf
        
        MsgBox gsAnzeigeText, vbInformation, iZaehler & ". Meldung"
        Pause 1
        iZaehler = iZaehler + 1
    Loop
    
    
    'HoleNeueBonNrWKL20
    HoleNeueBonNrWKL20_NEU 'bonnr wird gleich eingetragen
    
    AktualisiereZahlung68 "+"
    
    RestGutscheinVerarbeitung
    
    Label1(30).Caption = gdBonNr
    
    TSSBerechnung
    DruckeKassenBonWKL68
    
    
    
    If Text1(5).Text <> "" Then
        KorrigiereDTADaten68 Label1(30).Caption
    End If
    
    KorrigiereAlterG68 Label1(30).Caption
    
    
    InsertAFCBuchModul20For68
    
    If CheckofP = True Then
        InsertProvision
    End If
    
    If CheckofX = True Then
        InsertXMarkierung
    End If
    
    'mach wie immer
'    UpdateAFCStat68Test3
'    UpdateAFCStat68Test4
'    UpdateAFCStat68Test5
    UpdateAFCStat68Test6
    
    'bei GutschModus
    
    If gbGutscheinBeiVKversteuern = True Then
        AddOn_AFCSTAT_GutschModus
    End If
    
    
    
    ReInitDialog20WK68
    
    abschliessen = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "abschliessen"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub TSSBerechnung()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim cSQL        As String
    Dim ctmp        As String
    Dim cLBSatz     As String
    Dim cExtend     As String
    Dim cWasSuchteMan     As String
    Dim dLiNr       As Double
    Dim dEkpr       As Double
    Dim dWert       As Double
    Dim iFeld       As Integer
    Dim iDbNr       As Integer
    Dim rsrs        As Recordset
    Dim rsKJ        As Recordset
    Dim rsArt       As Recordset
    Dim cArtMWSt    As String
    
    '*** Kunden-Umsatz ****
    Dim dKdUmsatz   As Double
    Dim dKdBonus    As Double
    
    '*** KASSJOUR-Felder ****
    Dim cKJArtNr    As String
    Dim cKJBezeich  As String
    Dim cKJMenge    As String
    Dim cKJAZeit    As String
    Dim cKJKundNr   As String
    Dim cKJFiliale  As String
    Dim cKJKasNum   As String
    Dim cKJLiNr     As String
    Dim cKJLPZ      As String
    Dim cKJAGN      As String
    Dim cKJEAN      As String
    Dim cKJMwst     As String
    Dim cKJBelegNr  As String
    Dim cUmsOK      As String
    Dim cBonusOk    As String
    Dim cKJMopreis  As String
    
    Dim dKJEkpr     As Double
    Dim dKJVkpr     As Double
    Dim dKJPreis    As Double
    Dim dKJBest1    As Double
    Dim dVkPr       As Double
    Dim dKJPreis2   As Double
    Dim dSpanne     As Double
    Dim lKJADate    As Long
    Dim lKJBediener As Long
    Dim sArtnr      As String
    Dim IAbschluss  As Long
    Dim ierrz       As Integer
    Dim dGeldwert   As Double
    Dim sRechner    As String
    Dim sPreisKz    As String
    
    Dim lPos As Long
    
    Dim cpfaddb As String
        
    cpfaddb = gcDBPfad
    If Right$(cpfaddb, 1) <> "\" Then
        cpfaddb = cpfaddb & "\"
    End If
    
    
    ctmp = Trim$(frmWKL20!Label2(7).Caption)
    If Val(ctmp) < 0 Then
        ctmp = "0"
    End If
    cKJKundNr = ctmp
    
    If Val(cKJKundNr) > 0 Then
        'dann nach Preiskz fragen
        sPreisKz = ermPREISKZ(cKJKundNr)
    End If

    
    sRechner = rechnername
    ierrz = 0
    
    lAnzSatz = frmWKL20!List1.ListCount

    cSQL = "Delete from AFCB" & sRechner & " "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    For lAktSatz = 0 To lAnzSatz - 1

        iFeld = 1
        cKJArtNr = ""
        cKJBezeich = ""
        cKJMenge = ""
        dKJPreis = 0
        lKJADate = 0
        cKJAZeit = ""
        lKJBediener = 0
        cKJKundNr = ""
        cKJFiliale = ""
        cKJKasNum = ""
        cKJLiNr = ""
        cKJLPZ = ""
        cKJAGN = ""
        cKJEAN = ""
        cKJMwst = ""
        dKJEkpr = 0
        dKJVkpr = 0
        cKJBelegNr = ""
        dKJBest1 = 0
        
        
        cLBSatz = frmWKL20!List1.list(lAktSatz)
        
'        MsgBox cLBSatz

        'Besonderheiten am Satzende

        'hier Besonders Merkmal - wird in Mopreis kassjour gespeichert
        
        If Len(cLBSatz) > 175 Then
            cKJMopreis = Mid(cLBSatz, 177, 8)
        Else
            cKJMopreis = "0"
        End If

        If Len(cLBSatz) > 157 Then
            cExtend = Mid(cLBSatz, 158, 18)
        Else
            cExtend = ""
        End If
        
        
        
        
        ctmp = Mid(cLBSatz, 7, 6)
        ctmp = Trim$(ctmp)
        sArtnr = ctmp
        
        '***************************************************
        '* Zeile ZWISCHENSUMME darf nicht übernommen werden!
        '***************************************************
        
        If ctmp <> "000000" Then
            
            cSQL = "Select LPZ, AGN, EAN, EKPR, linr, MWST, UMS_OK, BONUS_OK, Spanne from Artikel where Artnr = " & ctmp
            Set rsArt = gdBase.OpenRecordset(cSQL)
            
'            FnOpenrecordset rsArt, cSQL, 1, gdBase

            If Not rsArt.EOF Then
                If Not IsNull(rsArt!LPZ) Then
                    cKJLPZ = rsArt!LPZ
                Else
                    cKJLPZ = ""
                End If
        
                If Not IsNull(rsArt!AGN) Then
                    cKJAGN = rsArt!AGN
                Else
                    cKJAGN = ""
                End If
                
                If Not IsNull(rsArt!EAN) Then
                    cKJEAN = rsArt!EAN
                Else
                    cKJEAN = ""
                End If
                
                If Not IsNull(rsArt!ekpr) Then
                    dEkpr = rsArt!ekpr
                Else
                    dEkpr = 0
                End If
                
                If Not IsNull(rsArt!linr) Then
                    dLiNr = rsArt!linr
                Else
                    dLiNr = 0
                End If
                
                If Not IsNull(rsArt!MWST) Then
                    cArtMWSt = rsArt!MWST
                Else
                    cArtMWSt = "V"
                End If
                
                
                'ist Preiskz = 6 also Netto dann mwst = O
                If Val(sPreisKz) = 6 Then
                    cArtMWSt = "O"
                End If
                
                
                If Not IsNull(rsArt!UMS_OK) Then
                    cUmsOK = rsArt!UMS_OK
                Else
                    cUmsOK = "J"
                End If
                
                If Not IsNull(rsArt!BONUS_OK) Then
                    cBonusOk = rsArt!BONUS_OK
                Else
                    cBonusOk = "J"
                End If
                
                If Not IsNull(rsArt!SPANNE) Then
                    dSpanne = rsArt!SPANNE
                Else
                    dSpanne = 0
                End If
                             
            Else
                dEkpr = 0
                dLiNr = 0
            End If
            
            rsArt.Close: Set rsArt = Nothing
            
            Set rsrs = gdBase.OpenRecordset("AFCB" & sRechner, dbOpenTable)
                        
            rsrs.AddNew
            rsrs!SYNStatus = "A"

            If ctmp = "666666" Then
                If gbGutscheinBeiVKversteuern = True Then
                    cBonusOk = "N"
                    cUmsOK = "J"
                    cArtMWSt = "V"
                Else
                    cBonusOk = "N"
                    cUmsOK = "N"
                    cArtMWSt = "O"
                End If
            End If

            ctmp = Mid(cLBSatz, 148, 3)
            ctmp = Trim$(ctmp)

            rsrs!abednu = Val(ctmp)
            lKJBediener = Val(ctmp)

            rsrs!AFLAG = 0
            
            If Left(cLBSatz, 1) = "x" Then
                ctmp = Mid(cLBSatz, 2, 4)
            Else
                ctmp = Mid(cLBSatz, 1, 5)
            End If

           
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!aMenge = Val(ctmp)
            cKJMenge = ctmp

            ctmp = Mid(cLBSatz, 60, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!APREIS = Val(ctmp)
            dKJPreis = rsrs!APREIS

            ctmp = Mid(cLBSatz, 7, 6)
            ctmp = Trim$(ctmp)
            rsrs!aartnr = Val(ctmp)
            cKJArtNr = ctmp

            ctmp = Mid(cLBSatz, 14, 35)
            ctmp = Trim$(ctmp)

            rsrs!ABEZEICH = ctmp
            cKJBezeich = ctmp

            rsrs!ADATE = Fix(Now)
            rsrs!AZEIT = Format$(Now, "HH:MM:SS")
            lKJADate = rsrs!ADATE
            cKJAZeit = rsrs!AZEIT

            rsrs!AMWSK = cArtMWSt
            cKJMwst = cArtMWSt

            If ctmp = "V" Then
                ctmp = Mid(cLBSatz, 104, 9)
            ElseIf ctmp = "E" Then
                ctmp = Mid(cLBSatz, 114, 9)
            Else
                ctmp = "0"
            End If
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!AMWST = Val(ctmp)

            ctmp = frmWKL20!Label2(7).Caption
            ctmp = Trim$(ctmp)
            If Val(ctmp) < 0 Then
                ctmp = "0"
            End If
            rsrs!AKUNUM = Val(ctmp)
            cKJKundNr = ctmp

            rsrs!BELEGNR = gdBonNr
            cKJBelegNr = rsrs!BELEGNR

            rsrs!kasnum = Val(gcKasNum)
            cKJFiliale = gcFilNr

            rsrs!BUCHFLAG = 0

            ctmp = Mid(cLBSatz, 50, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            
            rsrs!AALTPREIS = Format(Val(ctmp), "#####0.00")

            ctmp = Mid(cLBSatz, 128, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)

            If Val(ctmp) = 0 Then
                ctmp = Mid(cLBSatz, 50, 9)
                ctmp = Trim$(ctmp)
                ctmp = fnMoveComma2Point$(ctmp)
                
                rsrs!AVKPR = Format(Val(ctmp), "#####0.00")

                dKJVkpr = rsrs!AVKPR
            Else
                rsrs!AVKPR = Val(ctmp)
                dKJVkpr = rsrs!AVKPR
            End If

            If dEkpr = 0 Then
            
                If sArtnr = "666668" Or sArtnr = "666669" Then
                    If gdZeitungsSpanne <> 0 Then
                        dEkpr = EKausNettospanneerrechnen(gdZeitungsSpanne, Val(ctmp), cArtMWSt)
                    End If
                Else
                    If dSpanne <> 0 Then
                        dEkpr = EKausNettospanneerrechnen(dSpanne, Val(ctmp), cArtMWSt)
                    End If
                End If
            
            End If

            rsrs!ALEKPR = dEkpr
            dKJEkpr = dEkpr

            rsrs!linr = dLiNr
            cKJLiNr = Trim$(Str$(dLiNr))

            If gcKreditKarte <> "" Then
                rsrs!kk_art = gcKreditKarte
            Else
                rsrs!kk_art = gcZahlMittel
            End If

            ctmp = Mid(cLBSatz, 138, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)

            dKJBest1 = Val(ctmp)
            rsrs!BESTAND = Val(ctmp)
            rsrs!ZHLGGUTSCH = 0
            

            
            rsrs!UMS_OK = cUmsOK
            rsrs!BONUS_OK = cBonusOk
            rsrs!FILIALNR = Val(gcFilNr)
            rsrs.Update

            rsrs.Close: Set rsrs = Nothing

        End If
    Next lAktSatz
    
    
'    If gcZahlMittel = "BA" Then
'        If gbGutscheinBeiVKversteuern = True Then
'            cSQL = "Select sum(preis) as NICHTUMS from KJ" & sRechner & " where ums_OK = 'N' and kk_art = 'BA'  "
'        Else
'            cSQL = "Select sum(preis) as NICHTUMS from KJ" & sRechner & " where ums_OK = 'N' and kk_art = 'BA' and ARTNR <> 666666 "
'        End If
'
'        Set rsRS = gdBase.OpenRecordset(cSQL)
'        If Not rsRS.EOF Then
'            If Not IsNull(rsRS!NICHTUMS) Then
'                insertNichtUmsBar lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, CDbl(rsRS!NICHTUMS)
'            End If
'        End If
'    End If
'
    
    
    '''''''' Oliver Alte TSE START
'
'    'TODO TSE FINISH
'    Dim dUmsatzVolleMwst As Double
'    Dim dUmsatzErmMwst As Double
'    Dim dUmsatzOhneMwst As Double
'
'    Dim dUmsatzGesamt As Double
'
'
'
'    dUmsatzVolleMwst = ermWertforTSS("V", "J", "AFCB" & sRechner)
'    dUmsatzErmMwst = ermWertforTSS("E", "J", "AFCB" & sRechner)
'    dUmsatzOhneMwst = ermWertforTSS("O", "J", "AFCB" & sRechner)
'
'    dUmsatzGesamt = dUmsatzVolleMwst + dUmsatzErmMwst + dUmsatzOhneMwst
'
'
'    Dim dBargegeben As Double: dBargegeben = 0
'    Dim dUnBargegeben As Double: dUnBargegeben = 0
'
'
'
'
'    If Text1(0).Text <> "" Then
'        If IsNumeric(Text1(0).Text) = True Then
'            dBargegeben = CDbl(Text1(0).Text)
'        End If
'    End If
'
'
'
'
'
'
'
'
'    If gcZahlMittel = "BA" Then
'        TSS.Finish frmWKL20!WinsockTSE, Beleg, dUmsatzVolleMwst, dUmsatzErmMwst, dUmsatzOhneMwst, "EUR", dUmsatzGesamt, 0
'    Else
'        TSS.Finish frmWKL20!WinsockTSE, Beleg, dUmsatzVolleMwst, dUmsatzErmMwst, dUmsatzOhneMwst, "EUR", 0, dUmsatzGesamt
'    End If
'
    '''''''' Oliver Alte TSE ENDE

    
   

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TSSBerechnung"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ErfasseGutscheinWKL68() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim lbednu      As Long
    Dim lDatum      As Long
    Dim lKUNDNR     As Long
    
    Dim ctmp        As String
    Dim cArtNr      As String
    Dim cBezeich    As String
    Dim cMin        As String
    Dim cMax        As String
    Dim cSQL        As String
    Dim sWert       As String
    Dim dAnz        As Double
    Dim dCount      As Double
    Dim dWert       As Double
    Dim rsrs        As Recordset
    
    ErfasseGutscheinWKL68 = 0
    sWert = "A"
    
    dAnz = 1
    cArtNr = "666666"

    frmWKLak.Show 1
    If glGutschNr > -1 Then
        ErfasseGutscheinWKL68 = glGutschNr
        
        lbednu = Val(Text1(0).Text)
        lDatum = Fix(Now)
        ctmp = Text1(1).Text
        
        If Len(ctmp) = 1 Then
        
            ctmp = "01"
        
        End If
        If InStr(ctmp, ",") = 0 Then
            ctmp = Left(ctmp, Len(ctmp) - 2) & "," & Right(ctmp, 2)
        End If
        
        ctmp = fnMoveComma2Point$(ctmp)
        
        dWert = Val(ctmp)
        lKUNDNR = Val(frmWKL20.Label2(7).Caption)
        
        Insert_Gutschein ErfasseGutscheinWKL68, lKUNDNR, dWert, lbednu, gcFilNr, gcKasNum, "RESTGUTSCHEIN", lDatum, sWert
        
'        cSQL = "Select * from GUTSCH where GUTSCHNR = 0"
'        FnOpenrecordset rsRS, cSQL, 1, gdBase
'        rsRS.AddNew
'        rsRS!gutschnr = ErfasseGutscheinWKL68
'        rsRS!BEDNU = lbednu
'        rsRS!DAT_AUSG = lDatum
'        rsRS!Wert = dWert
'        rsRS!Kundnr = lKundnr
'        rsRS!SYNStatus = "A"
'        rsRS!Status = swert
'        rsRS!FILIALE = gcFilNr
'        rsRS.Update
'        rsRS.Close: Set rsRS = Nothing

    End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErfasseGutscheinWKL68"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Private Sub RestGutscheinVerarbeitung()
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Dim cSQL As String
    Dim lbednu As Long
    Dim lDatum As Long
    Dim lGutschnr As Long
    Dim dSumRückGUTSCH As Double
    Dim i As Integer
    Dim dMaxGutschWert As Double
    Dim dGutschwert As Double
    Dim lGutschNrfrom_MaxWert As Long
    
    dSumRückGUTSCH = 0
    glnewGutschnr = 0
    Label1(31).Caption = ""
    
    If Label1(28).Caption <> "" Then
        dSumRückGUTSCH = CDbl(Label1(28).Caption)
    End If
    
    If dSumRückGUTSCH > 0 Then
    
        lbednu = Val(frmWKL20!Text1(0).Text)
        lDatum = Fix(Now)
        
        If gbRGO = True Then 'Original nummer behalten
        
            dGutschwert = 0
            dMaxGutschWert = 0
            lGutschNrfrom_MaxWert = 0
            
            For i = 0 To 19
                If Gutschl(i).gutschnr <> 0 Then
                    lGutschnr = Gutschl(i).gutschnr
                    dGutschwert = Gutschl(i).gutschwert
                    
                    If dGutschwert > dMaxGutschWert Then
                        dMaxGutschWert = dGutschwert
                        lGutschNrfrom_MaxWert = lGutschnr
                    End If
                    
                    
                    
'                    Exit For
                End If
            Next i
            
            lGutschnr = lGutschNrfrom_MaxWert
            
            If Gutschein_vorhanden(lGutschnr) = False Then
                glnewGutschnr = lGutschnr
                Label1(31).Caption = lGutschnr
                If lGutschnr = 0 Then
                    Exit Sub
                End If

                Insert_Gutschein lGutschnr, 0, dSumRückGUTSCH, lbednu, gcFilNr, gcKasNum, "RESTGUTSCHEIN", lDatum, "A"
                ProtokolliereRueckGutscheinWK20g dSumRückGUTSCH, lGutschnr
                
            Else
                'update
                
                glnewGutschnr = lGutschnr
           
                Label1(31).Caption = lGutschnr
            
                Update_Gutschein lGutschnr, 0, dSumRückGUTSCH, lbednu, gcFilNr, gcKasNum, "", lDatum, "E"
                ProtokolliereRueckGutscheinWK20g dSumRückGUTSCH, lGutschnr
                
            End If
        

            Exit Sub
        End If
        
    
        If gbRGO = False Then
            lGutschnr = NewGutschein
            glnewGutschnr = lGutschnr
            Label1(31).Caption = lGutschnr
            If lGutschnr = 0 Then
                Exit Sub
            End If
            
            Insert_Gutschein lGutschnr, 0, dSumRückGUTSCH, lbednu, gcFilNr, gcKasNum, "RESTGUTSCHEIN", lDatum, "A"
            ProtokolliereRueckGutscheinWK20g dSumRückGUTSCH, lGutschnr
            
        Else
            'Restgutschein = Originalgutschein
            
            dGutschwert = 0
            dMaxGutschWert = 0
            lGutschNrfrom_MaxWert = 0
            
            For i = 0 To 19
                If Gutschl(i).gutschnr <> 0 Then
                    lGutschnr = Gutschl(i).gutschnr
                    dGutschwert = Gutschl(i).gutschwert
                    
                    If dGutschwert > dMaxGutschWert Then
                        dMaxGutschWert = dGutschwert
                        lGutschNrfrom_MaxWert = lGutschnr
                    End If
                    
                    
                    
'                    Exit For
                End If
            Next i
            
            lGutschnr = lGutschNrfrom_MaxWert
            glnewGutschnr = lGutschnr
           
            Label1(31).Caption = lGutschnr
            
            Update_Gutschein lGutschnr, 0, dSumRückGUTSCH, lbednu, gcFilNr, gcKasNum, "", lDatum, "E"
            ProtokolliereRueckGutscheinWK20g dSumRückGUTSCH, lGutschnr

        End If
        
        
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "RestGutscheinVerarbeitung"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
            
End Sub
Private Function Gutschein_vorhanden(lGutschnr As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp    As String
    Dim ctmp1   As String
    Dim lWert   As Long
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim iRet    As Integer
    
    Gutschein_vorhanden = False
    
    lWert = lGutschnr
    
    If gbKL_LIVEGUTSCHEIN Then
    
        If fTestLogin_SQLDABA_Error = 0 Then 'ist alles OK? Datenbank erreichbar?
            'alles okay
            
        Else
        
            schreibeProtokollVPNTXT "Unterbrechung"
        
            Dim sTemp As String
            sTemp = "Bitte starten Sie diesen Rechner neu" & vbCrLf
            sTemp = sTemp & "oder schließen Sie das Schloss und starten Sie WinKiss neu."
        
            MsgBox sTemp, vbCritical + vbOKOnly, "Gutschein-Datenbank nicht erreichbar"
            Exit Function
        End If
    

        Dim stConnect As String
        
        If gsKL_DSN <> "" Then
            stConnect = "ODBC;DSN=" & gsKL_DSN & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
        Else
            stConnect = "ODBC;DRIVER=SQL Server;SERVER=" & gsKL_ADRESSE & ";DATABASE=" & gsKL_DATENBANKNAME & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
        End If
    
        Dim dbEAN As DAO.Database
        Set dbEAN = OpenDatabase(gsKL_DATENBANKNAME, dbDriverNoPrompt, False, stConnect)
        
        
        
        cSQL = "Select * from GUTSCHEINE where GUTSCHNR = '" & CStr(lWert) & "'"
        Set rsrs = dbEAN.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            Gutschein_vorhanden = True
            Exit Function
        End If
        rsrs.Close: Set rsrs = Nothing
        
        dbEAN.Close
    Else

        cSQL = "Select * from GUTSCH where GUTSCHNR = " & lWert
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            Gutschein_vorhanden = True
            Exit Function
        End If
        rsrs.Close: Set rsrs = Nothing
        
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Gutschein_vorhanden"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub UpdateAFCStat68Test3()
    On Error GoTo LOKAL_ERROR

    Dim lStornoAnz      As Long
    Dim lDatum          As Long
    Dim lAktSatz        As Long
    Dim lAnzSatz        As Long

    Dim cArtNr          As String
    Dim cUmsOK          As String
    Dim cSQL            As String
    Dim cErzielterPreis As String
    Dim ctmp            As String
    Dim cNormal         As String
    Dim cPosSumme       As String
    Dim cArtRabatt      As String
    Dim cLBSatz         As String

    Dim dUmsatz         As Double     'Summe des Verkaufs ohne Gutscheine
    Dim dEchterUmsatz   As Double     'Summe des Verkaufs ohne Gutscheine
    Dim dFalscherUmsatz As Double     'Summe des Verkaufs ohne Gutscheine
    
    Dim dNichtUmsatz    As Double     'Summe nichtumsatzrelevanter VK zB Dukaten

    Dim dUmsatz2        As Double     'Summe des Verkaufs inkl. Gutscheine
    Dim dSPreisAnz      As Double     'Anzahl Positionen mit Sonderpreis
    Dim dSPreisGes      As Double     'Summe aller Positionen mit Sonderpreis
    Dim dKundenZahl     As Double     'Konstante 1
    Dim dArtRabAnz      As Double     'Anzahl Positionen mit Artikelrabatt
    Dim dArtRabGes      As Double     'Summe des gegebenen Artikelrabatts
    Dim dGesRabAnz      As Double     'Anzahl Positionen mit Gesamtrabatt
    Dim dGesRabGes      As Double     'Summe des gegebenen Gesamtrabatts
    Dim dWertGutschein  As Double

    Dim rsrs                As Recordset

    Dim dStornoWert         As Double
    Dim dZhlgGutsch         As Double
    Dim dSumDukaten         As Double

    Dim dSumKreditkarten    As Double
    Dim dSumScheck          As Double
    Dim dSumECLAST          As Double
    Dim dSumBar             As Double
    
    Dim dGegScheck          As Double

    Dim dSumGUTSCHKreditkarten    As Double
    Dim dSumGUTSCHScheck          As Double
    Dim dSumGUTSCHECLAST          As Double
    Dim dSumGUTSCHBar             As Double
    Dim dSumGutschDukate          As Double



    Dim dSumRückBar                 As Double
    Dim dSumRückGUTSCH              As Double
    Dim dUmsatzGutsch               As Double
    Dim dSumUmsBar                  As Double

    Dim dGutschGutschAnteil         As Double
    Dim dGutschBarAnteil            As Double
    Dim dGutschECLASTAnteil         As Double
    Dim dGutschKreditkartenAnteil   As Double
    Dim dGutschScheckAnteil         As Double
    Dim dGutschDukateAnteil         As Double
    Dim bUmsatz                     As Boolean
    
    bUmsatz = True
    
    dNichtUmsatz = 0

    dGutschGutschAnteil = 0
    dGutschBarAnteil = 0
    dGutschECLASTAnteil = 0
    dGutschKreditkartenAnteil = 0
    dGutschScheckAnteil = 0
    dGutschDukateAnteil = 0


    dSumGUTSCHKreditkarten = 0
    dSumGUTSCHScheck = 0
    dSumGUTSCHECLAST = 0
    dSumGUTSCHBar = 0
    dSumGutschDukate = 0

    lDatum = Fix(Now)
    dKundenZahl = 0
    dSumKreditkarten = 0
    dSumDukaten = 0
    dSumScheck = 0
    dGegScheck = 0
    dSumECLAST = 0
    dSumBar = 0
    dSumRückBar = 0
    dSumUmsBar = 0
    dSumRückGUTSCH = 0
    dZhlgGutsch = 0
    dUmsatzGutsch = 0
    dwertGutverkauf = 0

    '*******************************************
    '* Was hat der Kunde insgesamt zu zahlen?
    '*******************************************
    If Label333(0).Caption <> "" Then
        dUmsatz = CDbl(Label333(0).Caption) 'dZuZahlen
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Gutscheinen?
    '*******************************************
    If Text1(1).Text <> "" And IsNumeric(Text1(1).Text) Then
        dZhlgGutsch = CDbl(Text1(1).Text) 'dEinrGutsch
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Dukaten?
    '*******************************************
    
    If Text1(3).Text <> "" And IsNumeric(Text1(3).Text) Then
        dSumDukaten = CDbl(Text1(3).Text)
        dSumGutschDukate = dSumDukaten
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Kreditkarte?
    '*******************************************
    If Text1(2).Text <> "" And IsNumeric(Text1(2).Text) Then
        dSumKreditkarten = dSumKreditkarten + CDbl(Text1(2).Text)
        dSumGUTSCHKreditkarten = dSumGUTSCHKreditkarten + dSumKreditkarten
    End If

    If Text1(7).Text <> "" And IsNumeric(Text1(7).Text) Then
        dSumKreditkarten = dSumKreditkarten + CDbl(Text1(7).Text)
        dSumGUTSCHKreditkarten = dSumGUTSCHKreditkarten + dSumKreditkarten
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Scheck?
    '*******************************************
    If Text1(6).Text <> "" And IsNumeric(Text1(6).Text) Then
        dSumScheck = CDbl(Text1(6).Text)
        dGegScheck = dSumScheck
        dSumGUTSCHScheck = dSumScheck
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels ECLAST?
    '*******************************************
    If Text1(5).Text <> "" And IsNumeric(Text1(5).Text) Then
        dSumECLAST = CDbl(Text1(5).Text)
        dSumGUTSCHECLAST = dSumECLAST
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Bargeld?
    '*******************************************
    If Text1(0).Text <> "" And IsNumeric(Text1(0).Text) Then
        dSumBar = CDbl(Text1(0).Text)
        dSumGUTSCHBar = dSumBar
    End If

    '*******************************************
    '* Was kriegt der Kunde mittels Bargeld zurück?
    '*******************************************
    If Label333(3).Caption <> "" Then
        dSumRückBar = CDbl(Label333(3).Caption)
    End If

    '*******************************************
    '* Was kriegt der Kunde mittels RestGutschein zurück?
    '*******************************************
    If Label1(28).Caption <> "" Then
        dSumRückGUTSCH = CDbl(Label1(28).Caption)
    End If

    dUmsatzGutsch = dZhlgGutsch '- dSumRückGUTSCH

    dEchterUmsatz = 0
    dWertGutschein = 0

    '*******************************************
    '* Untersuche jeden einzelnen Artikel
    '*******************************************

    lAnzSatz = frmWKL20!List1.ListCount

    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = frmWKL20!List1.list(lAktSatz)

        cArtNr = Mid(cLBSatz, 7, 6)

        '*******************************************
        '* Lies Kennzeichen Umsatzrelevanz
        '*******************************************
        If Len(cLBSatz) > 155 Then
            cUmsOK = Mid(cLBSatz, 156, 1)
        Else
            cUmsOK = "J"
        End If
        If cUmsOK <> "J" And cUmsOK <> "N" Then
            cUmsOK = "J"
        End If

        '*******************************************
        '* Lies regulären Stückpreis
        '*******************************************
        cNormal = Mid(cLBSatz, 128, 9)
        cNormal = Trim$(cNormal)
        cNormal = fnMoveComma2Point$(cNormal)

        '*******************************************
        '* Lies Stückpreis, zu dem verkauft wurde
        '*******************************************
        ctmp = Mid(cLBSatz, 74, 9)
        ctmp = Trim$(ctmp)
        ctmp = fnMoveComma2Point$(ctmp)

        '*******************************************
        '* Lies Gesamtpreis der Position
        '*******************************************
        cPosSumme = Mid(cLBSatz, 94, 9)
        cPosSumme = Trim$(cPosSumme)
        cPosSumme = fnMoveComma2Point$(cPosSumme)

        '*******************************************
        '* Lies Artikelrabatt der Position
        '*******************************************
        cArtRabatt = Mid(cLBSatz, 124, 3)
        cArtRabatt = Trim$(cArtRabatt)
        cArtRabatt = fnMoveComma2Point$(cArtRabatt)

        '**********************************************
        '* Lies den echten Verkaufspreis der Position
        '**********************************************
        cErzielterPreis = Mid(cLBSatz, 60, 9)
        cErzielterPreis = Trim$(cErzielterPreis)
        cErzielterPreis = fnMoveComma2Point$(cErzielterPreis)
        
        
        
        If gbGutscheinBeiVKversteuern = True Then
             '**********************************************
            '* Ermittle die echte Umsatzsumme
            '* -
            '* - kein nicht umsatzrelevanten Artikel
            '**********************************************
            
            
            If cArtNr <> "666666" Then
                If cUmsOK <> "N" Then
                    dEchterUmsatz = dEchterUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                Else
                    dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
                    
                    dFalscherUmsatz = dFalscherUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                End If
            Else
                '******************************************************
                '* Wenn Gutschein, dann summiere verkaufte Gutscheine
                '******************************************************
                
                
                dEchterUmsatz = dEchterUmsatz + Val(cErzielterPreis)
                dKundenZahl = 1
    
                dWertGutschein = dWertGutschein + Val(cErzielterPreis)
                dwertGutverkauf = dwertGutverkauf + Val(cErzielterPreis)
            End If
            
        Else
        

            '**********************************************
            '* Ermittle die echte Umsatzsumme
            '* - keine Gutscheine
            '* - kein nicht umsatzrelevanten Artikel
            '**********************************************
            
            
            If cArtNr <> "666666" Then
                If cUmsOK <> "N" Then
                    dEchterUmsatz = dEchterUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                Else
                    dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
                    
                    dFalscherUmsatz = dFalscherUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                End If
            Else
                '******************************************************
                '* Wenn Gutschein, dann summiere verkaufte Gutscheine
                '******************************************************
                dFalscherUmsatz = dFalscherUmsatz + Val(cErzielterPreis)
                dKundenZahl = 1
    
                dWertGutschein = dWertGutschein + Val(cErzielterPreis)
                dwertGutverkauf = dwertGutverkauf + Val(cErzielterPreis)
            End If
        End If


        '**********************************************
        '* Wenn regulärer Stückpreis und Stückpreis
        '* des Verkaufes abweichen, Zähler für
        '* Sonderpreis um 1 heraufsetzen und die
        '* Sonderpreissumme erhöhen
        '**********************************************
        If Val(ctmp) <> Val(cNormal) Then
            If cNormal = 0 Then
                dSPreisAnz = 0
                dSPreisGes = 0
            Else
                dSPreisAnz = dSPreisAnz + 1
                dSPreisGes = dSPreisGes + Val(cPosSumme)
            End If
        End If

        '*******************************************
        '* Wenn Artikelrabatt gewährt wurde,
        '* Zähler für Artikelrabatt um 1 heraufsetzen
        '* und ArtRabattsumme erhöhen
        '*******************************************
        If Val(cArtRabatt) <> 0 Then
            ctmp = Mid(cLBSatz, 84, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)

            dArtRabAnz = dArtRabAnz + 1
            dArtRabGes = dArtRabGes + Val(ctmp)
        ElseIf frmWKL20!Label2(3).Visible Then
            dGesRabAnz = dGesRabAnz + 1
        End If

        '*******************************************
        '* Wenn erzielter Preis < 0, dann
        '* Zähler für Storno um 1 heraufsetzen
        '* und Stornosumme erhöhen
        '*******************************************
        If Val(cErzielterPreis) < 0 Then
            If IstArtikelnichtStornierfähig(cArtNr) = False Then
                dStornoWert = dStornoWert + Val(cErzielterPreis)
                lStornoAnz = lStornoAnz + 1
            End If
        End If
    Next lAktSatz

    If frmWKL20!Label2(3).Visible Then
        dGesRabGes = fnHoleGesamtRabattModul20#()
    End If

    If dwertGutverkauf > 0 Then
        If dZhlgGutsch <= dwertGutverkauf Then
            dwertGutverkauf = dZhlgGutsch
            updateafcstat "ZHLGGUTSCH", (-1 * dwertGutverkauf), gcKasNum
        End If
    End If

    cSQL = "Select * from AFCSTAT where ADATE = " & Trim$(Str$(lDatum)) & " and KASNUM = " & gcKasNum
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    
    
    

    If dNichtUmsatz > 0 Then
    
    
    
    
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    


    dSumUmsBar = dSumBar - dSumRückBar + dGutscheinauszahlung

    Dim dRestfalscherumsatz As Double
    dRestfalscherumsatz = 0
    
    Dim dZuVergebenerEchterumsatz As Double
    dZuVergebenerEchterumsatz = 0
    
    If dEchterUmsatz > 0 Then dZuVergebenerEchterumsatz = dEchterUmsatz

    If dFalscherUmsatz > 0 Then 'Wert der Gutscheinverkäufe
    
    
        dRestfalscherumsatz = dFalscherUmsatz
        
        
        
        
        
        
        
        
        
        
        
        
''''''        'dukaten
''''''
''''''        If dSumDukaten < dRestfalscherumsatz Then 'dann dukaten abfragen
''''''
''''''            If dZhlgGutsch > dRestfalscherumsatz Then
''''''                dSumDukaten = dSumDukaten
''''''                dRestfalscherumsatz = 0
''''''            Else
''''''                If dSumDukaten >= dZuVergebenerEchterumsatz Then
''''''                    dSumDukaten = dZuVergebenerEchterumsatz
''''''
''''''                    dSumGutschDukate = dsumdukate - dZuVergebenerEchterumsatz
''''''                    dZuVergebenerEchterumsatz = 0
''''''
''''''
''''''
''''''                Else
''''''
''''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumDukaten
''''''
''''''                    dSumGutschDukate = 0
''''''
''''''                End If
'''''''                dRestfalscherumsatz = dFalscherUmsatz
''''''            End If
''''''
''''''
''''''        Else
''''''            If dZhlgGutsch >= dRestfalscherumsatz Then
''''''                dSumDukaten = dSumDukaten
''''''                dRestfalscherumsatz = 0
''''''            Else
''''''                dSumDukaten = dSumDukaten - dRestfalscherumsatz + dZhlgGutsch
'''''''                dRestfalscherumsatz = dFalscherUmsatz
''''''            End If
''''''        End If
''''''
''''''
''''''
''''''
''''''
''''''
''''''
''''''
''''''
''''''        'Kreditkarten
''''''
''''''        If dSumKreditkarten < dRestfalscherumsatz Then 'dann KK abfragen
''''''
''''''            If dZhlgGutsch > dRestfalscherumsatz Then
''''''
''''''
''''''                dSumKreditkarten = dSumKreditkarten
''''''                dRestfalscherumsatz = 0
''''''            Else
''''''
''''''                If dSumKreditkarten >= dZuVergebenerEchterumsatz Then
''''''                    dSumKreditkarten = dZuVergebenerEchterumsatz
''''''                    dZuVergebenerEchterumsatz = 0
''''''                Else
''''''
''''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumKreditkarten
''''''
''''''                End If
''''''
''''''
'''''''                dRestfalscherumsatz = dFalscherUmsatz
''''''            End If
''''''
''''''        Else
''''''
''''''            If dZhlgGutsch >= dRestfalscherumsatz Then
''''''
''''''                dSumKreditkarten = dSumKreditkarten
''''''                dRestfalscherumsatz = 0
''''''            Else
''''''
''''''                If dSumKreditkarten >= dZuVergebenerEchterumsatz Then
''''''                    dSumKreditkarten = dZuVergebenerEchterumsatz
''''''                    dZuVergebenerEchterumsatz = 0
''''''                Else
''''''                    dSumKreditkarten = dSumKreditkarten
''''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumKreditkarten
''''''                End If
''''''
'''''''                dRestfalscherumsatz = dFalscherUmsatz
''''''
''''''            End If
''''''        End If
''''''
''''''
''''''
''''''
''''''
''''''        '*****************BAR
''''''        If dSumUmsBar < dRestfalscherumsatz Then 'erst bar abfragen
''''''
''''''            If dZhlgGutsch > dRestfalscherumsatz Then
''''''
''''''                dSumUmsBar = dSumUmsBar ' - dFalscherUmsatz + dZhlgGutsch
''''''                dRestfalscherumsatz = 0
''''''            Else
''''''
''''''                If dSumUmsBar >= dZuVergebenerEchterumsatz Then
''''''                    dSumUmsBar = dZuVergebenerEchterumsatz
''''''                    dZuVergebenerEchterumsatz = 0
''''''                Else
''''''
''''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumUmsBar
''''''
''''''                End If
''''''
'''''''                dRestfalscherumsatz = dFalscherUmsatz
''''''
''''''            End If
''''''
''''''        Else
''''''            If dZhlgGutsch > dRestfalscherumsatz Then
''''''
''''''                dSumUmsBar = dSumUmsBar ' - dFalscherUmsatz + dZhlgGutsch
''''''                dRestfalscherumsatz = 0
''''''            Else
''''''
''''''                If dSumUmsBar >= dZuVergebenerEchterumsatz Then
''''''                    dSumUmsBar = dZuVergebenerEchterumsatz
''''''                    dZuVergebenerEchterumsatz = 0
''''''                Else
''''''                    dSumUmsBar = dSumUmsBar
''''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumUmsBar
''''''                End If
''''''
'''''''                dRestfalscherumsatz = dFalscherUmsatz
''''''            End If
''''''        End If
''''''
''''''        '*****************BAR Ende
''''''
''''''
''''''
''''''        'scheck
''''''
''''''        If dSumScheck < dRestfalscherumsatz Then 'dann Scheck abfragen
''''''            If dZhlgGutsch > dRestfalscherumsatz Then
''''''                dSumScheck = dSumScheck
''''''                dRestfalscherumsatz = 0
''''''            Else
''''''                If dSumScheck >= dZuVergebenerEchterumsatz Then
''''''                    dSumScheck = dZuVergebenerEchterumsatz
''''''                    dZuVergebenerEchterumsatz = 0
''''''                Else
''''''
''''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumScheck
''''''
''''''                End If
'''''''                dRestfalscherumsatz = dFalscherUmsatz
''''''            End If
''''''        Else
''''''            If dZhlgGutsch >= dRestfalscherumsatz Then
''''''                dSumScheck = dSumScheck
''''''                dRestfalscherumsatz = 0
''''''            Else
''''''                dSumScheck = dSumScheck - dRestfalscherumsatz + dZhlgGutsch
'''''''                dRestfalscherumsatz = dFalscherUmsatz
''''''            End If
''''''        End If
''''''
''''''
''''''
''''''
''''''
''''''
''''''
''''''        'Ec Last
''''''
''''''        If dSumECLAST < dRestfalscherumsatz Then 'dann ec last abfragen
''''''            If dZhlgGutsch > dRestfalscherumsatz Then
''''''                dSumECLAST = dSumECLAST
''''''                dRestfalscherumsatz = 0
''''''            Else
''''''                If dSumECLAST >= dZuVergebenerEchterumsatz Then
''''''                    dSumECLAST = dZuVergebenerEchterumsatz
''''''                    dZuVergebenerEchterumsatz = 0
''''''                Else
''''''
''''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumECLAST
''''''
''''''                End If
''''''                dRestfalscherumsatz = 0
''''''            End If
''''''        Else
''''''            If dZhlgGutsch >= dRestfalscherumsatz Then
''''''                dSumECLAST = dSumECLAST
''''''                dRestfalscherumsatz = 0
''''''            Else
''''''                dSumECLAST = dSumECLAST - dRestfalscherumsatz + dZhlgGutsch
''''''                dRestfalscherumsatz = 0
''''''            End If
''''''        End If
''''''
''''''
''''''
''''''
''''''
''''''
''''''
        
        
        
        
        
        
        
        
        
        
        
        
        'Kreditkarten

        If dSumKreditkarten < dRestfalscherumsatz Then 'dann KK abfragen

            If dZhlgGutsch > dRestfalscherumsatz Then


                dSumKreditkarten = dSumKreditkarten
                dRestfalscherumsatz = 0
            Else

                If dSumKreditkarten >= dZuVergebenerEchterumsatz Then
                    dSumKreditkarten = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dSumKreditkarten = dSumKreditkarten ' - dRestfalscherumsatz + dZhlgGutsch
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumKreditkarten
                End If


                dRestfalscherumsatz = dFalscherUmsatz
            End If

        Else

            If dZhlgGutsch >= dRestfalscherumsatz Then

                dSumKreditkarten = dSumKreditkarten
                dRestfalscherumsatz = 0
            Else

                If dSumKreditkarten >= dZuVergebenerEchterumsatz Then
                    dSumKreditkarten = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dSumKreditkarten = dSumKreditkarten
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumKreditkarten
                End If

                dRestfalscherumsatz = dFalscherUmsatz

            End If
        End If





        '*****************BAR
        If dSumUmsBar < dRestfalscherumsatz Then 'erst bar abfragen

            If dZhlgGutsch > dRestfalscherumsatz Then

                dSumUmsBar = dSumUmsBar ' - dFalscherUmsatz + dZhlgGutsch
                dRestfalscherumsatz = 0
            Else

                If dSumUmsBar >= dZuVergebenerEchterumsatz Then
                    dSumUmsBar = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dSumUmsBar = dSumUmsBar
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumUmsBar
                End If

                dRestfalscherumsatz = dFalscherUmsatz

            End If

        Else
            If dZhlgGutsch > dRestfalscherumsatz Then

                dSumUmsBar = dSumUmsBar ' - dFalscherUmsatz + dZhlgGutsch
                dRestfalscherumsatz = 0
            Else

                If dSumUmsBar >= dZuVergebenerEchterumsatz Then
                    dSumUmsBar = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dSumUmsBar = dSumUmsBar
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumUmsBar
                End If

                dRestfalscherumsatz = dFalscherUmsatz
            End If
        End If

        '*****************BAR Ende



        'scheck

        If dSumScheck < dRestfalscherumsatz Then 'dann Scheck abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumScheck = dSumScheck
                dRestfalscherumsatz = 0
            Else
                If dSumScheck >= dZuVergebenerEchterumsatz Then
                    dSumScheck = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dSumScheck = dSumScheck ' - dRestfalscherumsatz + dZhlgGutsch
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumScheck
                End If
                dRestfalscherumsatz = dFalscherUmsatz
            End If
        Else
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumScheck = dSumScheck
                dRestfalscherumsatz = 0
            Else
                dSumScheck = dSumScheck - dRestfalscherumsatz + dZhlgGutsch
                dRestfalscherumsatz = dFalscherUmsatz
            End If
        End If

''         'dukaten
''
''        If dSumDukaten < dRestfalscherumsatz Then 'dann dukaten abfragen
''
''            If dZhlgGutsch > dRestfalscherumsatz Then
''                dSumDukaten = dSumDukaten
''                dRestfalscherumsatz = 0
''            Else
''                If dSumDukaten >= dZuVergebenerEchterumsatz Then
''                    dSumDukaten = dZuVergebenerEchterumsatz
''                    dZuVergebenerEchterumsatz = 0
''                Else
''                    dSumDukaten = dSumDukaten ' - dRestfalscherumsatz + dZhlgGutsch
''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumDukaten
''                End If
''                dRestfalscherumsatz = dFalscherUmsatz
''            End If
''
''
''        Else
''            If dZhlgGutsch >= dRestfalscherumsatz Then
''                dSumDukaten = dSumDukaten
''                dRestfalscherumsatz = 0
''            Else
''                dSumDukaten = dSumDukaten - dRestfalscherumsatz + dZhlgGutsch
''                dRestfalscherumsatz = dFalscherUmsatz
''            End If
''        End If





        'Ec Last

        If dSumECLAST < dRestfalscherumsatz Then 'dann ec last abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumECLAST = dSumECLAST
                dRestfalscherumsatz = 0
            Else
                If dSumECLAST >= dZuVergebenerEchterumsatz Then
                    dSumECLAST = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dSumECLAST = dSumECLAST ' - dRestfalscherumsatz + dZhlgGutsch
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumECLAST
                End If
                dRestfalscherumsatz = 0
            End If
        Else
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumECLAST = dSumECLAST
                dRestfalscherumsatz = 0
            Else
                dSumECLAST = dSumECLAST - dRestfalscherumsatz + dZhlgGutsch
                dRestfalscherumsatz = 0
            End If
        End If
        
        
        
       
        
        
        
        
        
        
        
        
        
        
        
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

    '**************************************************************
    '* Wieviel Umsatz hat der Kunde in bar gemacht?
    '**************************************************************
    
'    If bUmsatz = True Then
    If dEchterUmsatz > 0 Then
        'erzielter Umsatz über Bargeld
        
        
    
        If Not IsNull(rsrs!UMS_BAR) Then
            rsrs!UMS_BAR = rsrs!UMS_BAR + dSumUmsBar
        Else
            rsrs!UMS_BAR = dSumUmsBar
        End If
    
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde mit EC-Lastschrift gemacht?
        '**************************************************************
        If Not IsNull(rsrs!UMS_LAST) Then
            rsrs!UMS_LAST = rsrs!UMS_LAST + dSumECLAST
        Else
            rsrs!UMS_LAST = dSumECLAST
        End If
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde über Dukaten gemacht?
        '**************************************************************
        If Not IsNull(rsrs!DUKA) Then
            rsrs!DUKA = rsrs!DUKA + dSumDukaten
        Else
            rsrs!DUKA = dSumDukaten
        End If
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde mit Schecks gemacht?
        '**************************************************************
        If Not IsNull(rsrs!UMS_SCHECK) Then
            rsrs!UMS_SCHECK = rsrs!UMS_SCHECK + dSumScheck
        Else
            rsrs!UMS_SCHECK = dSumScheck
        End If
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde mit Kreditkarten gemacht?
        '**************************************************************
        If Not IsNull(rsrs!UMS_KARTE) Then
            rsrs!UMS_KARTE = rsrs!UMS_KARTE + dSumKreditkarten
            schreibeProtokollUNITXT CStr(dSumKreditkarten), "Kartenzahlung"
        Else
            rsrs!UMS_KARTE = dSumKreditkarten
            schreibeProtokollUNITXT CStr(dSumKreditkarten), "Kartenzahlung"
        End If
        
    End If

    '**************************************************************
    '* Datum und Kassennummer des Verbuchens schreiben
    '**************************************************************
    rsrs!ADATE = lDatum
    rsrs!kasnum = Val(gcKasNum)

    '**************************************************************
    '* Betrag der eingereichten Gutscheine verbuchen
    '**************************************************************
    If Not IsNull(rsrs!EINRGUTSCH) Then
        rsrs!EINRGUTSCH = rsrs!EINRGUTSCH + dZhlgGutsch
    Else
        rsrs!EINRGUTSCH = dZhlgGutsch
    End If

    '**************************************************************
    '* Betrag der generierten Rest-Gutscheine verbuchen
    '**************************************************************
    If Not IsNull(rsrs!RESTGUTSCH) Then
        rsrs!RESTGUTSCH = rsrs!RESTGUTSCH + dSumRückGUTSCH
    Else
        rsrs!RESTGUTSCH = dSumRückGUTSCH
    End If

    'Achtung
    '**************************************************************
    '* Betrag der Gutschein-AUszahlung verbuchen
    '**************************************************************
    If Not IsNull(rsrs!AUSZGUTSCH) Then
        rsrs!AUSZGUTSCH = rsrs!AUSZGUTSCH + Format(dGutscheinauszahlung, "######0.00")
    Else
        rsrs!AUSZGUTSCH = Format(dGutscheinauszahlung, "######0.00")
    End If


    'Achtung
    '**************************************************************
    '* Betrag des Umsatzes durch Gutschein-Einreichungen verbuchen
    '**************************************************************
    
    
    'Test echter Umsatz
    

    If dwertGutverkauf > 0 Then

        If dZhlgGutsch > dwertGutverkauf Then

            If Not IsNull(rsrs!ZHLGGUTSCH) Then
                rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
            Else
                rsrs!ZHLGGUTSCH = dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
                
                If rsrs!ZHLGGUTSCH < 0 Then rsrs!ZHLGGUTSCH = 0
            End If
        
        Else
        
            If Not IsNull(rsrs!ZHLGGUTSCH) Then
                rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH
            Else
                rsrs!ZHLGGUTSCH = dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH
            End If
        
        End If


    Else
        If Not IsNull(rsrs!ZHLGGUTSCH) Then
            rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
        Else
            rsrs!ZHLGGUTSCH = dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
        End If

    End If
    
    
    
    
    
    
    
    
    

    'Achtung
    '**************************************************************
    '* Betrag des Gutschein-Verkäufe verbuchen (Bar?)
    '**************************************************************
    If Not IsNull(rsrs!GUTSCHEIN) Then
        rsrs!GUTSCHEIN = rsrs!GUTSCHEIN + dWertGutschein
    Else
        rsrs!GUTSCHEIN = dWertGutschein
    End If

    '**************************************************************
    '* Betrag der Scheck - Verkäufe verbuchen
    '**************************************************************

    If Not IsNull(rsrs!SCHVERKAUF) Then
        rsrs!SCHVERKAUF = rsrs!SCHVERKAUF + dGegScheck
    Else
        rsrs!SCHVERKAUF = dGegScheck
    End If

    '**************************************************************
    '* Sonderpreise verbuchen
    '**************************************************************
    If Not IsNull(rsrs!SPREIS_ANZ) Then
        rsrs!SPREIS_ANZ = rsrs!SPREIS_ANZ + dSPreisAnz
    Else
        rsrs!SPREIS_ANZ = dSPreisAnz
    End If

    If Not IsNull(rsrs!SPREIS_GES) Then
        rsrs!SPREIS_GES = rsrs!SPREIS_GES + dSPreisGes
    Else
        rsrs!SPREIS_GES = dSPreisGes
    End If



    '**************************************************************
    '* Kundenzahl schreiben
    '**************************************************************
    If Not IsNull(rsrs!Kundenzahl) Then
        rsrs!Kundenzahl = rsrs!Kundenzahl + dKundenZahl
    Else
        rsrs!Kundenzahl = dKundenZahl
    End If
    '**************************************************************
    '* Artikelrabatte schreiben
    '**************************************************************
    If Not IsNull(rsrs!ARTRAB_ANZ) Then
        rsrs!ARTRAB_ANZ = rsrs!ARTRAB_ANZ + dArtRabAnz
    Else
        rsrs!ARTRAB_ANZ = dArtRabAnz
    End If

    If Not IsNull(rsrs!ARTRAB_GES) Then
        rsrs!ARTRAB_GES = rsrs!ARTRAB_GES + dArtRabGes
    Else
        rsrs!ARTRAB_GES = dArtRabGes
    End If
    '**************************************************************
    '* Gesamtrabatte schreiben
    '**************************************************************
    If Not IsNull(rsrs!GESRAB_ANZ) Then
        rsrs!GESRAB_ANZ = rsrs!GESRAB_ANZ + dGesRabAnz
    Else
        rsrs!GESRAB_ANZ = dGesRabAnz
    End If

    If Not IsNull(rsrs!GESRAB_GES) Then
        rsrs!GESRAB_GES = rsrs!GESRAB_GES + dGesRabGes
    Else
        rsrs!GESRAB_GES = dGesRabGes
    End If
    '**************************************************************
    '* Bonnummer schreiben
    '**************************************************************
    
    If Not IsNull(rsrs!BELEGNR) Then
        If gdBonNr < CLng(rsrs!BELEGNR) Then
            
        Else
            rsrs!BELEGNR = gdBonNr
        End If
    Else
        rsrs!BELEGNR = gdBonNr
    End If
    
    

    '**************************************************************
    '* Stornos schreiben
    '**************************************************************
    If Not IsNull(rsrs!STORNO_GES) Then
        rsrs!STORNO_GES = rsrs!STORNO_GES + dStornoWert
    Else
        rsrs!STORNO_GES = dStornoWert
    End If

    If Not IsNull(rsrs!STORNO_ANZ) Then
        rsrs!STORNO_ANZ = rsrs!STORNO_ANZ + lStornoAnz
    Else
        rsrs!STORNO_ANZ = lStornoAnz
    End If

    '******************************************
    '* Wie sind neue Gutscheine bezahlt worden
    '******************************************

    '******************************************
    '* Betrag Gutscheinverkäufe in Bar setzen
    '******************************************
    Dim dRestZahlungsmittelforGutsch As Double
    dRestZahlungsmittelforGutsch = 0

    If dWertGutschein > 0 Then
    
        dRestZahlungsmittelforGutsch = dWertGutschein
        
        If dZhlgGutsch < dRestZahlungsmittelforGutsch Then 'dann gutschein aus gutschein abfragen
            dGutschGutschAnteil = dZhlgGutsch
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dGutschGutschAnteil
        Else
            dGutschGutschAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If
        
'        If dSumGutschDukate < dRestZahlungsmittelforGutsch Then 'dann dukate abfragen
'            dGutschDukateAnteil = dSumGutschDukate
'            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGutschDukate
'        Else
'            dGutschDukateAnteil = dRestZahlungsmittelforGutsch
'            dRestZahlungsmittelforGutsch = 0
'        End If

        If (dSumGUTSCHBar - dSumRückBar + dGutscheinauszahlung) < dRestZahlungsmittelforGutsch Then  'erst bar abfragen
            dGutschBarAnteil = dSumGUTSCHBar - dSumRückBar + dGutscheinauszahlung '+ dGutschGutschAnteil
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - (dSumGUTSCHBar - dSumRückBar + dGutscheinauszahlung)
        Else
            dGutschBarAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If


        If dSumGUTSCHScheck < dRestZahlungsmittelforGutsch Then 'dann Scheck abfragen
            dGutschScheckAnteil = dSumGUTSCHScheck
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHScheck
        Else
            dGutschScheckAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If

        If dSumGUTSCHECLAST < dRestZahlungsmittelforGutsch Then 'dann last abfragen
            dGutschECLASTAnteil = dSumGUTSCHECLAST
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHECLAST
        Else
            dGutschECLASTAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If
        
        
        
        If dSumGUTSCHKreditkarten < dRestZahlungsmittelforGutsch Then 'dann KK abfragen
            dGutschKreditkartenAnteil = dSumGUTSCHKreditkarten
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHKreditkarten
        Else
            dGutschKreditkartenAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If

    End If

    If dGutschBarAnteil > 0 Then
        '******************************************
        '* Wert des Barverkaufes setzen
        '******************************************
        If Not IsNull(rsrs!BARVERKAUF) Then
            rsrs!BARVERKAUF = rsrs!BARVERKAUF + (dSumBar - dSumRückBar + dGutscheinauszahlung - dGutschBarAnteil)
        Else
            rsrs!BARVERKAUF = (dSumBar - dSumRückBar + dGutscheinauszahlung - dGutschBarAnteil)
        End If

    Else
        '******************************************
        '* Wert des Barverkaufes setzen
        '******************************************
        If Not IsNull(rsrs!BARVERKAUF) Then
            rsrs!BARVERKAUF = rsrs!BARVERKAUF + dSumBar - dSumRückBar + dGutscheinauszahlung
        Else
            rsrs!BARVERKAUF = dSumBar - dSumRückBar + dGutscheinauszahlung
        End If
    End If
    
    
    '**************************************************************
    '* Gutscheine gekauft mit Gutschein bezahlt!
    '**************************************************************
    If dGutschGutschAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHGUTSCH) Then
            rsrs!GUTSCHGUTSCH = rsrs!GUTSCHGUTSCH + dGutschGutschAnteil
        Else
            rsrs!GUTSCHGUTSCH = dGutschGutschAnteil
        End If
    End If
    
    If dGutschBarAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHBAR) Then
            rsrs!GUTSCHBAR = rsrs!GUTSCHBAR + Format(dGutschBarAnteil, "######0.00") 'Hier ist Format() neu 21.10.13
        Else
            rsrs!GUTSCHBAR = Format(dGutschBarAnteil, "######0.00") 'Hier ist Format() neu 21.10.13
        End If
    End If

    If dGutschKreditkartenAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHKAR) Then
            rsrs!GUTSCHKAR = rsrs!GUTSCHKAR + dGutschKreditkartenAnteil
            schreibeProtokollUNITXT CStr(dGutschKreditkartenAnteil), "Kartenzahlung"
        Else
            rsrs!GUTSCHKAR = dGutschKreditkartenAnteil
            schreibeProtokollUNITXT CStr(dGutschKreditkartenAnteil), "Kartenzahlung"
        End If
    End If

    If dGutschECLASTAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHLAST) Then
            rsrs!GUTSCHLAST = rsrs!GUTSCHLAST + dGutschECLASTAnteil
        Else
            rsrs!GUTSCHLAST = dGutschECLASTAnteil
        End If
    End If

    If dGutschScheckAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHSCH) Then
            rsrs!GUTSCHSCH = rsrs!GUTSCHSCH + dGutschScheckAnteil
        Else
            rsrs!GUTSCHSCH = dGutschScheckAnteil
        End If
    End If





    '**************************************************************
    '* Schreibvorgang durchführen
    '**************************************************************
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "UpdateAFCStat68Test3"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub UpdateAFCStat68Test4()
    On Error GoTo LOKAL_ERROR

    Dim lStornoAnz      As Long
    Dim lDatum          As Long
    Dim lAktSatz        As Long
    Dim lAnzSatz        As Long

    Dim cArtNr          As String
    Dim cUmsOK          As String
    Dim cSQL            As String
    Dim cErzielterPreis As String
    Dim ctmp            As String
    Dim cNormal         As String
    Dim cPosSumme       As String
    Dim cArtRabatt      As String
    Dim cLBSatz         As String

    Dim dUmsatz         As Double     'Summe des Verkaufs ohne Gutscheine
    Dim dEchterUmsatz   As Double     'Summe des Verkaufs ohne Gutscheine
    Dim dFalscherUmsatz As Double     'Summe des Verkaufs ohne Gutscheine
    
    Dim dNichtUmsatz    As Double     'Summe nichtumsatzrelevanter VK zB Dukaten

    Dim dUmsatz2        As Double     'Summe des Verkaufs inkl. Gutscheine
    Dim dSPreisAnz      As Double     'Anzahl Positionen mit Sonderpreis
    Dim dSPreisGes      As Double     'Summe aller Positionen mit Sonderpreis
    Dim dKundenZahl     As Double     'Konstante 1
    Dim dArtRabAnz      As Double     'Anzahl Positionen mit Artikelrabatt
    Dim dArtRabGes      As Double     'Summe des gegebenen Artikelrabatts
    Dim dGesRabAnz      As Double     'Anzahl Positionen mit Gesamtrabatt
    Dim dGesRabGes      As Double     'Summe des gegebenen Gesamtrabatts
    Dim dWertGutschein  As Double

    Dim rsrs                As Recordset

    Dim dStornoWert         As Double
    Dim dZhlgGutsch         As Double
    Dim dSumDukaten         As Double

    Dim dSumKreditkarten    As Double
    Dim dSumScheck          As Double
    Dim dSumECLAST          As Double
    Dim dSumBar             As Double
    
    Dim dGegScheck          As Double

    Dim dSumGUTSCHKreditkarten    As Double
    Dim dSumGUTSCHScheck          As Double
    Dim dSumGUTSCHECLAST          As Double
    Dim dSumGUTSCHBar             As Double
    Dim dSumGutschDukate          As Double



    Dim dSumRückBar                 As Double
    Dim dSumRückGUTSCH              As Double
    Dim dUmsatzGutsch               As Double
    Dim dSumUmsBar                  As Double

    Dim dGutschGutschAnteil         As Double
    Dim dGutschBarAnteil            As Double
    Dim dGutschECLASTAnteil         As Double
    Dim dGutschKreditkartenAnteil   As Double
    Dim dGutschScheckAnteil         As Double
    Dim dGutschDukateAnteil         As Double
    
    Dim bUmsatz                     As Boolean
    
    bUmsatz = True
    
    dNichtUmsatz = 0

    dGutschGutschAnteil = 0
    dGutschBarAnteil = 0
    dGutschECLASTAnteil = 0
    dGutschKreditkartenAnteil = 0
    dGutschScheckAnteil = 0
    dGutschDukateAnteil = 0


    dSumGUTSCHKreditkarten = 0
    dSumGUTSCHScheck = 0
    dSumGUTSCHECLAST = 0
    dSumGUTSCHBar = 0
    dSumGutschDukate = 0

    lDatum = Fix(Now)
    dKundenZahl = 0
    dSumKreditkarten = 0
    dSumDukaten = 0
    dSumScheck = 0
    dGegScheck = 0
    dSumECLAST = 0
    dSumBar = 0
    dSumRückBar = 0
    dSumUmsBar = 0
    dSumRückGUTSCH = 0
    dZhlgGutsch = 0
    dUmsatzGutsch = 0
    dwertGutverkauf = 0

    '*******************************************
    '* Was hat der Kunde insgesamt zu zahlen?
    '*******************************************
    If Label333(0).Caption <> "" Then
        dUmsatz = CDbl(Label333(0).Caption) 'dZuZahlen
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Gutscheinen?
    '*******************************************
    If Text1(1).Text <> "" And IsNumeric(Text1(1).Text) Then
        dZhlgGutsch = CDbl(Text1(1).Text) 'dEinrGutsch
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Dukaten?
    '*******************************************
    
    If Text1(3).Text <> "" And IsNumeric(Text1(3).Text) Then
        dSumDukaten = CDbl(Text1(3).Text)
        dSumGutschDukate = dSumDukaten
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Kreditkarte?
    '*******************************************
    If Text1(2).Text <> "" And IsNumeric(Text1(2).Text) Then
        dSumKreditkarten = dSumKreditkarten + CDbl(Text1(2).Text)
        dSumGUTSCHKreditkarten = dSumGUTSCHKreditkarten + dSumKreditkarten
    End If

    If Text1(7).Text <> "" And IsNumeric(Text1(7).Text) Then
        dSumKreditkarten = dSumKreditkarten + CDbl(Text1(7).Text)
        dSumGUTSCHKreditkarten = dSumGUTSCHKreditkarten + dSumKreditkarten
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Scheck?
    '*******************************************
    If Text1(6).Text <> "" And IsNumeric(Text1(6).Text) Then
        dSumScheck = CDbl(Text1(6).Text)
        dGegScheck = dSumScheck
        dSumGUTSCHScheck = dSumScheck
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels ECLAST?
    '*******************************************
    If Text1(5).Text <> "" And IsNumeric(Text1(5).Text) Then
        dSumECLAST = CDbl(Text1(5).Text)
        dSumGUTSCHECLAST = dSumECLAST
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Bargeld?
    '*******************************************
    If Text1(0).Text <> "" And IsNumeric(Text1(0).Text) Then
        dSumBar = CDbl(Text1(0).Text)
        dSumGUTSCHBar = dSumBar
    End If

    '*******************************************
    '* Was kriegt der Kunde mittels Bargeld zurück?
    '*******************************************
    If Label333(3).Caption <> "" Then
        dSumRückBar = CDbl(Label333(3).Caption)
    End If

    '*******************************************
    '* Was kriegt der Kunde mittels RestGutschein zurück?
    '*******************************************
    If Label1(28).Caption <> "" Then
        dSumRückGUTSCH = CDbl(Label1(28).Caption)
    End If
    
    'gegeben Ende*****************************************************************************************
    
    
    
    

    dUmsatzGutsch = dZhlgGutsch '- dSumRückGUTSCH

    dEchterUmsatz = 0
    dWertGutschein = 0

    '*******************************************
    '* Untersuche jeden einzelnen Artikel
    '*******************************************

    lAnzSatz = frmWKL20!List1.ListCount

    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = frmWKL20!List1.list(lAktSatz)

        cArtNr = Mid(cLBSatz, 7, 6)

        '*******************************************
        '* Lies Kennzeichen Umsatzrelevanz
        '*******************************************
        If Len(cLBSatz) > 155 Then
            cUmsOK = Mid(cLBSatz, 156, 1)
        Else
            cUmsOK = "J"
        End If
        If cUmsOK <> "J" And cUmsOK <> "N" Then
            cUmsOK = "J"
        End If

        '*******************************************
        '* Lies regulären Stückpreis
        '*******************************************
        cNormal = Mid(cLBSatz, 128, 9)
        cNormal = Trim$(cNormal)
        cNormal = fnMoveComma2Point$(cNormal)

        '*******************************************
        '* Lies Stückpreis, zu dem verkauft wurde
        '*******************************************
        ctmp = Mid(cLBSatz, 74, 9)
        ctmp = Trim$(ctmp)
        ctmp = fnMoveComma2Point$(ctmp)

        '*******************************************
        '* Lies Gesamtpreis der Position
        '*******************************************
        cPosSumme = Mid(cLBSatz, 94, 9)
        cPosSumme = Trim$(cPosSumme)
        cPosSumme = fnMoveComma2Point$(cPosSumme)

        '*******************************************
        '* Lies Artikelrabatt der Position
        '*******************************************
        cArtRabatt = Mid(cLBSatz, 124, 3)
        cArtRabatt = Trim$(cArtRabatt)
        cArtRabatt = fnMoveComma2Point$(cArtRabatt)

        '**********************************************
        '* Lies den echten Verkaufspreis der Position
        '**********************************************
        cErzielterPreis = Mid(cLBSatz, 60, 9)
        cErzielterPreis = Trim$(cErzielterPreis)
        cErzielterPreis = fnMoveComma2Point$(cErzielterPreis)
        
        
        
        If gbGutscheinBeiVKversteuern = True Then
             '**********************************************
            '* Ermittle die echte Umsatzsumme
            '* -
            '* - kein nicht umsatzrelevanten Artikel
            '**********************************************
            
            
            If cArtNr <> "666666" Then
                If cUmsOK <> "N" Then
                    dEchterUmsatz = dEchterUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                Else
                    dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
                    
                    dFalscherUmsatz = dFalscherUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                End If
            Else
                '******************************************************
                '* Wenn Gutschein, dann summiere verkaufte Gutscheine
                '******************************************************
                
                
                dEchterUmsatz = dEchterUmsatz + Val(cErzielterPreis)
                dKundenZahl = 1
    
                dWertGutschein = dWertGutschein + Val(cErzielterPreis)
                dwertGutverkauf = dwertGutverkauf + Val(cErzielterPreis)
            End If
            
        Else
        

            '**********************************************
            '* Ermittle die echte Umsatzsumme
            '* - keine Gutscheine
            '* - kein nicht umsatzrelevanten Artikel
            '**********************************************
            
            
            If cArtNr <> "666666" Then
                If cUmsOK <> "N" Then
                    dEchterUmsatz = dEchterUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                Else
                    dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
                    
                    dFalscherUmsatz = dFalscherUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                End If
            Else
                '******************************************************
                '* Wenn Gutschein, dann summiere verkaufte Gutscheine
                '******************************************************
                dFalscherUmsatz = dFalscherUmsatz + Val(cErzielterPreis)
                dKundenZahl = 1
    
                dWertGutschein = dWertGutschein + Val(cErzielterPreis)
                dwertGutverkauf = dwertGutverkauf + Val(cErzielterPreis)
            End If
        End If


        '**********************************************
        '* Wenn regulärer Stückpreis und Stückpreis
        '* des Verkaufes abweichen, Zähler für
        '* Sonderpreis um 1 heraufsetzen und die
        '* Sonderpreissumme erhöhen
        '**********************************************
        If Val(ctmp) <> Val(cNormal) Then
            If cNormal = 0 Then
                dSPreisAnz = 0
                dSPreisGes = 0
            Else
                dSPreisAnz = dSPreisAnz + 1
                dSPreisGes = dSPreisGes + Val(cPosSumme)
            End If
        End If

        '*******************************************
        '* Wenn Artikelrabatt gewährt wurde,
        '* Zähler für Artikelrabatt um 1 heraufsetzen
        '* und ArtRabattsumme erhöhen
        '*******************************************
        If Val(cArtRabatt) <> 0 Then
            ctmp = Mid(cLBSatz, 84, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)

            dArtRabAnz = dArtRabAnz + 1
            dArtRabGes = dArtRabGes + Val(ctmp)
        ElseIf frmWKL20!Label2(3).Visible Then
            dGesRabAnz = dGesRabAnz + 1
        End If

        '*******************************************
        '* Wenn erzielter Preis < 0, dann
        '* Zähler für Storno um 1 heraufsetzen
        '* und Stornosumme erhöhen
        '*******************************************
        If Val(cErzielterPreis) < 0 Then
            If IstArtikelnichtStornierfähig(cArtNr) = False Then
                dStornoWert = dStornoWert + Val(cErzielterPreis)
                lStornoAnz = lStornoAnz + 1
            End If
        End If
    Next lAktSatz

    If frmWKL20!Label2(3).Visible Then
        dGesRabGes = fnHoleGesamtRabattModul20#()
    End If

    If dwertGutverkauf > 0 Then
        If dZhlgGutsch <= dwertGutverkauf Then
            dwertGutverkauf = dZhlgGutsch
            updateafcstat "ZHLGGUTSCH", (-1 * dwertGutverkauf), gcKasNum
        End If
    End If

    cSQL = "Select * from AFCSTAT where ADATE = " & Trim$(Str$(lDatum)) & " and KASNUM = " & gcKasNum
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    
    
    

    If dNichtUmsatz > 0 Then
    
    
    
    
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    


    dSumUmsBar = dSumBar - dSumRückBar + dGutscheinauszahlung

    Dim dRestfalscherumsatz As Double
    dRestfalscherumsatz = 0
    
    Dim dZuVergebenerEchterumsatz As Double
    dZuVergebenerEchterumsatz = 0
    
    If dEchterUmsatz > 0 Then dZuVergebenerEchterumsatz = dEchterUmsatz

    If dFalscherUmsatz > 0 Then 'Wert der Gutscheinverkäufe
    
    
        dRestfalscherumsatz = dFalscherUmsatz
        
        
''
''
''        ' CW20190320 TEST Begin
''        If dZhlgGutsch > dRestfalscherumsatz Then
''            dRestfalscherumsatz = 0
''        Else
''            dRestfalscherumsatz = dRestfalscherumsatz - dZhlgGutsch
''        End If
''           ' CW20190320 TEST End
''
        
        
        
        
        'dukaten

        If dSumDukaten < dRestfalscherumsatz Then 'dann dukaten abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumDukaten = dSumDukaten
                dRestfalscherumsatz = 0
            Else
                If dSumDukaten >= dZuVergebenerEchterumsatz Then
                    dSumDukaten = dZuVergebenerEchterumsatz

                    dSumGutschDukate = dSumDukaten - dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumDukaten
                    dSumGutschDukate = 0
                End If
            End If
        Else 'falls dukaten mehr oder gleich als dRestfalscherumsatz
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumDukaten = dSumDukaten
            Else
                dSumDukaten = dSumDukaten - dRestfalscherumsatz + dZhlgGutsch
                dSumGutschDukate = dRestfalscherumsatz
            End If
            dRestfalscherumsatz = 0
        End If


        'Kreditkarten

        If dSumKreditkarten < dRestfalscherumsatz Then 'dann KK abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumKreditkarten = dSumKreditkarten
                dRestfalscherumsatz = 0
            Else
                If dSumKreditkarten >= dZuVergebenerEchterumsatz Then
                    dSumGUTSCHKreditkarten = dSumKreditkarten - dZuVergebenerEchterumsatz 'CW20190320
                    dSumKreditkarten = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumKreditkarten
                    dSumGUTSCHKreditkarten = 0 'CW20190320
                End If
            End If

        Else
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumKreditkarten = dSumKreditkarten
                dRestfalscherumsatz = 0
            Else
                If dSumKreditkarten >= dZuVergebenerEchterumsatz Then
                    dSumGUTSCHKreditkarten = dSumKreditkarten - dZuVergebenerEchterumsatz 'CW20190320
                    dSumKreditkarten = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dSumKreditkarten = dSumKreditkarten
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumKreditkarten
                    dSumGUTSCHBar = 0  'CW20190320
                End If
            End If
        End If

        '*****************BAR
        If dSumUmsBar < dRestfalscherumsatz Then 'erst bar abfragen

            If dZhlgGutsch > dRestfalscherumsatz Then

                dSumUmsBar = dSumUmsBar ' - dFalscherUmsatz + dZhlgGutsch
                dRestfalscherumsatz = 0
            Else
                If dSumUmsBar >= dZuVergebenerEchterumsatz Then
                    dSumGUTSCHBar = dSumUmsBar - dZuVergebenerEchterumsatz 'CW20190320
                    dSumUmsBar = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumUmsBar
                    dSumGUTSCHBar = 0 'CW20190320
                End If
            End If

        Else
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumUmsBar = dSumUmsBar ' - dFalscherUmsatz + dZhlgGutsch
                dRestfalscherumsatz = 0
            Else

                If dSumUmsBar >= dZuVergebenerEchterumsatz Then
                    dSumGUTSCHBar = dSumUmsBar - dZuVergebenerEchterumsatz 'CW20190320
                    dSumUmsBar = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dSumUmsBar = dSumUmsBar
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumUmsBar
                    dSumGUTSCHBar = 0  'CW20190320
                End If
            End If
        End If

        '*****************BAR Ende



        'scheck

        If dSumScheck < dRestfalscherumsatz Then 'dann Scheck abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumScheck = dSumScheck
                dRestfalscherumsatz = 0
            Else
                If dSumScheck >= dZuVergebenerEchterumsatz Then
                    dSumScheck = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else

                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumScheck

                End If
'                dRestfalscherumsatz = dFalscherUmsatz
            End If
        Else
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumScheck = dSumScheck
            Else
                dSumScheck = dSumScheck - dRestfalscherumsatz + dZhlgGutsch
                dSumGUTSCHScheck = dRestfalscherumsatz 'CW20190320
            End If
             dRestfalscherumsatz = 0 'CW20190320
        End If
             
        'Ec Last
        If dSumECLAST < dRestfalscherumsatz Then 'dann ec last abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumECLAST = dSumECLAST
                dRestfalscherumsatz = 0
            Else
                If dSumECLAST >= dZuVergebenerEchterumsatz Then
                    dSumECLAST = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumECLAST
                End If
                dRestfalscherumsatz = 0
            End If
        Else
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumECLAST = dSumECLAST
            Else
                dSumECLAST = dSumECLAST - dRestfalscherumsatz + dZhlgGutsch
                dSumGUTSCHECLAST = dRestfalscherumsatz 'CW20190320
            End If
             dRestfalscherumsatz = 0 'CW20190320
        End If







        
        
        
        
        
        
        
'''''
'''''
'''''
'''''
'''''
'''''        'Kreditkarten
'''''
'''''        If dSumKreditkarten < dRestfalscherumsatz Then 'dann KK abfragen
'''''
'''''            If dZhlgGutsch > dRestfalscherumsatz Then
'''''
'''''
'''''                dSumKreditkarten = dSumKreditkarten
'''''                dRestfalscherumsatz = 0
'''''            Else
'''''
'''''                If dSumKreditkarten >= dZuVergebenerEchterumsatz Then
'''''                    dSumKreditkarten = dZuVergebenerEchterumsatz
'''''                    dZuVergebenerEchterumsatz = 0
'''''                Else
'''''                    dSumKreditkarten = dSumKreditkarten ' - dRestfalscherumsatz + dZhlgGutsch
'''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumKreditkarten
'''''                End If
'''''
'''''
'''''                dRestfalscherumsatz = dFalscherUmsatz
'''''            End If
'''''
'''''        Else
'''''
'''''            If dZhlgGutsch >= dRestfalscherumsatz Then
'''''
'''''                dSumKreditkarten = dSumKreditkarten
'''''                dRestfalscherumsatz = 0
'''''            Else
'''''
'''''                If dSumKreditkarten >= dZuVergebenerEchterumsatz Then
'''''                    dSumKreditkarten = dZuVergebenerEchterumsatz
'''''                    dZuVergebenerEchterumsatz = 0
'''''                Else
'''''                    dSumKreditkarten = dSumKreditkarten
'''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumKreditkarten
'''''                End If
'''''
'''''                dRestfalscherumsatz = dFalscherUmsatz
'''''
'''''            End If
'''''        End If
'''''
'''''
'''''
'''''
'''''
'''''        '*****************BAR
'''''        If dSumUmsBar < dRestfalscherumsatz Then 'erst bar abfragen
'''''
'''''            If dZhlgGutsch > dRestfalscherumsatz Then
'''''
'''''                dSumUmsBar = dSumUmsBar ' - dFalscherUmsatz + dZhlgGutsch
'''''                dRestfalscherumsatz = 0
'''''            Else
'''''
'''''                If dSumUmsBar >= dZuVergebenerEchterumsatz Then
'''''                    dSumUmsBar = dZuVergebenerEchterumsatz
'''''                    dZuVergebenerEchterumsatz = 0
'''''                Else
'''''                    dSumUmsBar = dSumUmsBar
'''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumUmsBar
'''''                End If
'''''
'''''                dRestfalscherumsatz = dFalscherUmsatz
'''''
'''''            End If
'''''
'''''        Else
'''''            If dZhlgGutsch > dRestfalscherumsatz Then
'''''
'''''                dSumUmsBar = dSumUmsBar ' - dFalscherUmsatz + dZhlgGutsch
'''''                dRestfalscherumsatz = 0
'''''            Else
'''''
'''''                If dSumUmsBar >= dZuVergebenerEchterumsatz Then
'''''                    dSumUmsBar = dZuVergebenerEchterumsatz
'''''                    dZuVergebenerEchterumsatz = 0
'''''                Else
'''''                    dSumUmsBar = dSumUmsBar
'''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumUmsBar
'''''                End If
'''''
'''''                dRestfalscherumsatz = dFalscherUmsatz
'''''            End If
'''''        End If
'''''
'''''        '*****************BAR Ende
'''''
'''''
'''''
'''''        'scheck
'''''
'''''        If dSumScheck < dRestfalscherumsatz Then 'dann Scheck abfragen
'''''            If dZhlgGutsch > dRestfalscherumsatz Then
'''''                dSumScheck = dSumScheck
'''''                dRestfalscherumsatz = 0
'''''            Else
'''''                If dSumScheck >= dZuVergebenerEchterumsatz Then
'''''                    dSumScheck = dZuVergebenerEchterumsatz
'''''                    dZuVergebenerEchterumsatz = 0
'''''                Else
'''''                    dSumScheck = dSumScheck ' - dRestfalscherumsatz + dZhlgGutsch
'''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumScheck
'''''                End If
'''''                dRestfalscherumsatz = dFalscherUmsatz
'''''            End If
'''''        Else
'''''            If dZhlgGutsch >= dRestfalscherumsatz Then
'''''                dSumScheck = dSumScheck
'''''                dRestfalscherumsatz = 0
'''''            Else
'''''                dSumScheck = dSumScheck - dRestfalscherumsatz + dZhlgGutsch
'''''                dRestfalscherumsatz = dFalscherUmsatz
'''''            End If
'''''        End If
'''''
'''''''         'dukaten
'''''''
'''''''        If dSumDukaten < dRestfalscherumsatz Then 'dann dukaten abfragen
'''''''
'''''''            If dZhlgGutsch > dRestfalscherumsatz Then
'''''''                dSumDukaten = dSumDukaten
'''''''                dRestfalscherumsatz = 0
'''''''            Else
'''''''                If dSumDukaten >= dZuVergebenerEchterumsatz Then
'''''''                    dSumDukaten = dZuVergebenerEchterumsatz
'''''''                    dZuVergebenerEchterumsatz = 0
'''''''                Else
'''''''                    dSumDukaten = dSumDukaten ' - dRestfalscherumsatz + dZhlgGutsch
'''''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumDukaten
'''''''                End If
'''''''                dRestfalscherumsatz = dFalscherUmsatz
'''''''            End If
'''''''
'''''''
'''''''        Else
'''''''            If dZhlgGutsch >= dRestfalscherumsatz Then
'''''''                dSumDukaten = dSumDukaten
'''''''                dRestfalscherumsatz = 0
'''''''            Else
'''''''                dSumDukaten = dSumDukaten - dRestfalscherumsatz + dZhlgGutsch
'''''''                dRestfalscherumsatz = dFalscherUmsatz
'''''''            End If
'''''''        End If
'''''
'''''
'''''
'''''
'''''
'''''        'Ec Last
'''''
'''''        If dSumECLAST < dRestfalscherumsatz Then 'dann ec last abfragen
'''''            If dZhlgGutsch > dRestfalscherumsatz Then
'''''                dSumECLAST = dSumECLAST
'''''                dRestfalscherumsatz = 0
'''''            Else
'''''                If dSumECLAST >= dZuVergebenerEchterumsatz Then
'''''                    dSumECLAST = dZuVergebenerEchterumsatz
'''''                    dZuVergebenerEchterumsatz = 0
'''''                Else
'''''                    dSumECLAST = dSumECLAST ' - dRestfalscherumsatz + dZhlgGutsch
'''''                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumECLAST
'''''                End If
'''''                dRestfalscherumsatz = 0
'''''            End If
'''''        Else
'''''            If dZhlgGutsch >= dRestfalscherumsatz Then
'''''                dSumECLAST = dSumECLAST
'''''                dRestfalscherumsatz = 0
'''''            Else
'''''                dSumECLAST = dSumECLAST - dRestfalscherumsatz + dZhlgGutsch
'''''                dRestfalscherumsatz = 0
'''''            End If
'''''        End If
'''''
'''''
'''''
'''''
'''''
'''''
'''''
'''''
        
        
        
        
        
        
        
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    

    '**************************************************************
    '* Wieviel Umsatz hat der Kunde in bar gemacht?
    '**************************************************************
    
'    If bUmsatz = True Then
    If dEchterUmsatz > 0 Then
        'erzielter Umsatz über Bargeld
        
        
    
        If Not IsNull(rsrs!UMS_BAR) Then
            rsrs!UMS_BAR = rsrs!UMS_BAR + dSumUmsBar
        Else
            rsrs!UMS_BAR = dSumUmsBar
        End If
    
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde mit EC-Lastschrift gemacht?
        '**************************************************************
        If Not IsNull(rsrs!UMS_LAST) Then
            rsrs!UMS_LAST = rsrs!UMS_LAST + dSumECLAST
        Else
            rsrs!UMS_LAST = dSumECLAST
        End If
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde über Dukaten gemacht?
        '**************************************************************
        If Not IsNull(rsrs!DUKA) Then
            rsrs!DUKA = rsrs!DUKA + dSumDukaten
        Else
            rsrs!DUKA = dSumDukaten
        End If
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde mit Schecks gemacht?
        '**************************************************************
        If Not IsNull(rsrs!UMS_SCHECK) Then
            rsrs!UMS_SCHECK = rsrs!UMS_SCHECK + dSumScheck
        Else
            rsrs!UMS_SCHECK = dSumScheck
        End If
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde mit Kreditkarten gemacht?
        '**************************************************************
        If Not IsNull(rsrs!UMS_KARTE) Then
            rsrs!UMS_KARTE = rsrs!UMS_KARTE + dSumKreditkarten
            schreibeProtokollUNITXT CStr(dSumKreditkarten), "Kartenzahlung"
        Else
            rsrs!UMS_KARTE = dSumKreditkarten
            schreibeProtokollUNITXT CStr(dSumKreditkarten), "Kartenzahlung"
        End If
        
    End If

    '**************************************************************
    '* Datum und Kassennummer des Verbuchens schreiben
    '**************************************************************
    rsrs!ADATE = lDatum
    rsrs!kasnum = Val(gcKasNum)

    '**************************************************************
    '* Betrag der eingereichten Gutscheine verbuchen
    '**************************************************************
    If Not IsNull(rsrs!EINRGUTSCH) Then
        rsrs!EINRGUTSCH = rsrs!EINRGUTSCH + dZhlgGutsch
    Else
        rsrs!EINRGUTSCH = dZhlgGutsch
    End If

    '**************************************************************
    '* Betrag der generierten Rest-Gutscheine verbuchen
    '**************************************************************
    If Not IsNull(rsrs!RESTGUTSCH) Then
        rsrs!RESTGUTSCH = rsrs!RESTGUTSCH + dSumRückGUTSCH
    Else
        rsrs!RESTGUTSCH = dSumRückGUTSCH
    End If

    'Achtung
    '**************************************************************
    '* Betrag der Gutschein-AUszahlung verbuchen
    '**************************************************************
    If Not IsNull(rsrs!AUSZGUTSCH) Then
        rsrs!AUSZGUTSCH = rsrs!AUSZGUTSCH + Format(dGutscheinauszahlung, "######0.00")
    Else
        rsrs!AUSZGUTSCH = Format(dGutscheinauszahlung, "######0.00")
    End If


    'Achtung
    '**************************************************************
    '* Betrag des Umsatzes durch Gutschein-Einreichungen verbuchen
    '**************************************************************
    
    
    'Test echter Umsatz
    

    If dwertGutverkauf > 0 Then

        If dZhlgGutsch > dwertGutverkauf Then

            If Not IsNull(rsrs!ZHLGGUTSCH) Then
                rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
            Else
                rsrs!ZHLGGUTSCH = dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
                
                If rsrs!ZHLGGUTSCH < 0 Then rsrs!ZHLGGUTSCH = 0
            End If
        
        Else
        
            If Not IsNull(rsrs!ZHLGGUTSCH) Then
                rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH
            Else
                rsrs!ZHLGGUTSCH = dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH
            End If
        
        End If


    Else
        If Not IsNull(rsrs!ZHLGGUTSCH) Then
            rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
        Else
            rsrs!ZHLGGUTSCH = dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
        End If

    End If
    
    
    
    
    
    
    
    
    

    'Achtung
    '**************************************************************
    '* Betrag des Gutschein-Verkäufe verbuchen (Bar?)
    '**************************************************************
    If Not IsNull(rsrs!GUTSCHEIN) Then
        rsrs!GUTSCHEIN = rsrs!GUTSCHEIN + dWertGutschein
    Else
        rsrs!GUTSCHEIN = dWertGutschein
    End If

    '**************************************************************
    '* Betrag der Scheck - Verkäufe verbuchen
    '**************************************************************

    If Not IsNull(rsrs!SCHVERKAUF) Then
        rsrs!SCHVERKAUF = rsrs!SCHVERKAUF + dGegScheck
    Else
        rsrs!SCHVERKAUF = dGegScheck
    End If

    '**************************************************************
    '* Sonderpreise verbuchen
    '**************************************************************
    If Not IsNull(rsrs!SPREIS_ANZ) Then
        rsrs!SPREIS_ANZ = rsrs!SPREIS_ANZ + dSPreisAnz
    Else
        rsrs!SPREIS_ANZ = dSPreisAnz
    End If

    If Not IsNull(rsrs!SPREIS_GES) Then
        rsrs!SPREIS_GES = rsrs!SPREIS_GES + dSPreisGes
    Else
        rsrs!SPREIS_GES = dSPreisGes
    End If



    '**************************************************************
    '* Kundenzahl schreiben
    '**************************************************************
    If Not IsNull(rsrs!Kundenzahl) Then
        rsrs!Kundenzahl = rsrs!Kundenzahl + dKundenZahl
    Else
        rsrs!Kundenzahl = dKundenZahl
    End If
    '**************************************************************
    '* Artikelrabatte schreiben
    '**************************************************************
    If Not IsNull(rsrs!ARTRAB_ANZ) Then
        rsrs!ARTRAB_ANZ = rsrs!ARTRAB_ANZ + dArtRabAnz
    Else
        rsrs!ARTRAB_ANZ = dArtRabAnz
    End If

    If Not IsNull(rsrs!ARTRAB_GES) Then
        rsrs!ARTRAB_GES = rsrs!ARTRAB_GES + dArtRabGes
    Else
        rsrs!ARTRAB_GES = dArtRabGes
    End If
    '**************************************************************
    '* Gesamtrabatte schreiben
    '**************************************************************
    If Not IsNull(rsrs!GESRAB_ANZ) Then
        rsrs!GESRAB_ANZ = rsrs!GESRAB_ANZ + dGesRabAnz
    Else
        rsrs!GESRAB_ANZ = dGesRabAnz
    End If

    If Not IsNull(rsrs!GESRAB_GES) Then
        rsrs!GESRAB_GES = rsrs!GESRAB_GES + dGesRabGes
    Else
        rsrs!GESRAB_GES = dGesRabGes
    End If
    '**************************************************************
    '* Bonnummer schreiben
    '**************************************************************
    
    If Not IsNull(rsrs!BELEGNR) Then
        If gdBonNr < CLng(rsrs!BELEGNR) Then
            
        Else
            rsrs!BELEGNR = gdBonNr
        End If
    Else
        rsrs!BELEGNR = gdBonNr
    End If
    
    

    '**************************************************************
    '* Stornos schreiben
    '**************************************************************
    If Not IsNull(rsrs!STORNO_GES) Then
        rsrs!STORNO_GES = rsrs!STORNO_GES + dStornoWert
    Else
        rsrs!STORNO_GES = dStornoWert
    End If

    If Not IsNull(rsrs!STORNO_ANZ) Then
        rsrs!STORNO_ANZ = rsrs!STORNO_ANZ + lStornoAnz
    Else
        rsrs!STORNO_ANZ = lStornoAnz
    End If

    '******************************************
    '* Wie sind neue Gutscheine bezahlt worden
    '******************************************

    '******************************************
    '* Betrag Gutscheinverkäufe in Bar setzen
    '******************************************
    Dim dRestZahlungsmittelforGutsch As Double
    dRestZahlungsmittelforGutsch = 0

    If dWertGutschein > 0 Then
    
        dRestZahlungsmittelforGutsch = dWertGutschein
        
        If dZhlgGutsch < dRestZahlungsmittelforGutsch Then 'dann gutschein aus gutschein abfragen
            dGutschGutschAnteil = dZhlgGutsch
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dGutschGutschAnteil
        Else
            dGutschGutschAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If
        
        If dSumGutschDukate < dRestZahlungsmittelforGutsch Then 'dann dukate abfragen
            dGutschDukateAnteil = dSumGutschDukate
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGutschDukate
        Else
            dGutschDukateAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If

'      If (dSumGUTSCHBar - dSumRückBar + dGutscheinauszahlung) < dRestZahlungsmittelforGutsch Then  'erst bar abfragen
'             dGutschBarAnteil = dSumGUTSCHBar - dSumRückBar + dGutscheinauszahlung '+ dGutschGutschAnteil
'            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - (dSumGUTSCHBar - dSumRückBar + dGutscheinauszahlung)
'        Else
'            dGutschBarAnteil = dRestZahlungsmittelforGutsch
'            dRestZahlungsmittelforGutsch = 0
'        End If
       
        If dSumGUTSCHBar < dRestZahlungsmittelforGutsch Then  'erst bar abfragen
            dGutschBarAnteil = dSumGUTSCHBar
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHBar
        Else
            dGutschBarAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If


        If dSumGUTSCHScheck < dRestZahlungsmittelforGutsch Then 'dann Scheck abfragen
            dGutschScheckAnteil = dSumGUTSCHScheck
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHScheck
        Else
            dGutschScheckAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If

        If dSumGUTSCHECLAST < dRestZahlungsmittelforGutsch Then 'dann last abfragen
            dGutschECLASTAnteil = dSumGUTSCHECLAST
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHECLAST
        Else
            dGutschECLASTAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If
        
        
        
        If dSumGUTSCHKreditkarten < dRestZahlungsmittelforGutsch Then 'dann KK abfragen
            dGutschKreditkartenAnteil = dSumGUTSCHKreditkarten
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHKreditkarten
        Else
            dGutschKreditkartenAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If

    End If

    If dGutschBarAnteil > 0 Then
        '******************************************
        '* Wert des Barverkaufes setzen
        '******************************************
        If Not IsNull(rsrs!BARVERKAUF) Then
            rsrs!BARVERKAUF = rsrs!BARVERKAUF + (dSumBar - dSumRückBar + dGutscheinauszahlung - dGutschBarAnteil)
        Else
            rsrs!BARVERKAUF = (dSumBar - dSumRückBar + dGutscheinauszahlung - dGutschBarAnteil)
        End If

    Else
        '******************************************
        '* Wert des Barverkaufes setzen
        '******************************************
        If Not IsNull(rsrs!BARVERKAUF) Then
            rsrs!BARVERKAUF = rsrs!BARVERKAUF + dSumBar - dSumRückBar + dGutscheinauszahlung
        Else
            rsrs!BARVERKAUF = dSumBar - dSumRückBar + dGutscheinauszahlung
        End If
    End If
    
    
    '**************************************************************
    '* Gutscheine gekauft mit Gutschein bezahlt!
    '**************************************************************
    If dGutschGutschAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHGUTSCH) Then
            rsrs!GUTSCHGUTSCH = rsrs!GUTSCHGUTSCH + dGutschGutschAnteil
        Else
            rsrs!GUTSCHGUTSCH = dGutschGutschAnteil
        End If
    End If
    
    If dGutschBarAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHBAR) Then
            rsrs!GUTSCHBAR = rsrs!GUTSCHBAR + Format(dGutschBarAnteil, "######0.00") 'Hier ist Format() neu 21.10.13
        Else
            rsrs!GUTSCHBAR = Format(dGutschBarAnteil, "######0.00") 'Hier ist Format() neu 21.10.13
        End If
    End If

    If dGutschKreditkartenAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHKAR) Then
            rsrs!GUTSCHKAR = rsrs!GUTSCHKAR + dGutschKreditkartenAnteil
            schreibeProtokollUNITXT CStr(dGutschKreditkartenAnteil), "Kartenzahlung"
        Else
            rsrs!GUTSCHKAR = dGutschKreditkartenAnteil
            schreibeProtokollUNITXT CStr(dGutschKreditkartenAnteil), "Kartenzahlung"
        End If
    End If

    If dGutschECLASTAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHLAST) Then
            rsrs!GUTSCHLAST = rsrs!GUTSCHLAST + dGutschECLASTAnteil
        Else
            rsrs!GUTSCHLAST = dGutschECLASTAnteil
        End If
    End If

    If dGutschScheckAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHSCH) Then
            rsrs!GUTSCHSCH = rsrs!GUTSCHSCH + dGutschScheckAnteil
        Else
            rsrs!GUTSCHSCH = dGutschScheckAnteil
        End If
    End If





    '**************************************************************
    '* Schreibvorgang durchführen
    '**************************************************************
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "UpdateAFCStat68Test4"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub UpdateAFCStat68Test5()
    On Error GoTo LOKAL_ERROR

    Dim lStornoAnz      As Long
    Dim lDatum          As Long
    Dim lAktSatz        As Long
    Dim lAnzSatz        As Long

    Dim cArtNr          As String
    Dim cUmsOK          As String
    Dim cSQL            As String
    Dim cErzielterPreis As String
    Dim ctmp            As String
    Dim cNormal         As String
    Dim cPosSumme       As String
    Dim cArtRabatt      As String
    Dim cLBSatz         As String

    Dim dUmsatz         As Double     'Summe des Verkaufs ohne Gutscheine
    Dim dEchterUmsatz   As Double     'Summe des Verkaufs ohne Gutscheine
    Dim dFalscherUmsatz As Double     'Summe des Verkaufs ohne Gutscheine
    
    Dim dNichtUmsatz    As Double     'Summe nichtumsatzrelevanter VK zB Dukaten

    Dim dUmsatz2        As Double     'Summe des Verkaufs inkl. Gutscheine
    Dim dSPreisAnz      As Double     'Anzahl Positionen mit Sonderpreis
    Dim dSPreisGes      As Double     'Summe aller Positionen mit Sonderpreis
    Dim dKundenZahl     As Double     'Konstante 1
    Dim dArtRabAnz      As Double     'Anzahl Positionen mit Artikelrabatt
    Dim dArtRabGes      As Double     'Summe des gegebenen Artikelrabatts
    Dim dGesRabAnz      As Double     'Anzahl Positionen mit Gesamtrabatt
    Dim dGesRabGes      As Double     'Summe des gegebenen Gesamtrabatts
    Dim dWertGutschein  As Double

    Dim rsrs                As Recordset

    Dim dStornoWert         As Double
    Dim dZhlgGutsch         As Double
    Dim dSumDukaten         As Double

    Dim dSumKreditkarten    As Double
    Dim dSumScheck          As Double
    Dim dSumECLAST          As Double
    Dim dSumBar             As Double
    
    Dim dGegScheck          As Double

    Dim dSumGUTSCHKreditkarten    As Double
    Dim dSumGUTSCHScheck          As Double
    Dim dSumGUTSCHECLAST          As Double
    Dim dSumGUTSCHBar             As Double
    Dim dSumGutschDukate          As Double



    Dim dSumRückBar                 As Double
    Dim dSumRückGUTSCH              As Double
    Dim dUmsatzGutsch               As Double
    Dim dSumUmsBar                  As Double

    Dim dGutschGutschAnteil         As Double
    Dim dGutschBarAnteil            As Double
    Dim dGutschECLASTAnteil         As Double
    Dim dGutschKreditkartenAnteil   As Double
    Dim dGutschScheckAnteil         As Double
    Dim dGutschDukateAnteil         As Double
    
    Dim bUmsatz                     As Boolean
    
    bUmsatz = True
    
    dNichtUmsatz = 0

    dGutschGutschAnteil = 0
    dGutschBarAnteil = 0
    dGutschECLASTAnteil = 0
    dGutschKreditkartenAnteil = 0
    dGutschScheckAnteil = 0
    dGutschDukateAnteil = 0


    dSumGUTSCHKreditkarten = 0
    dSumGUTSCHScheck = 0
    dSumGUTSCHECLAST = 0
    dSumGUTSCHBar = 0
    dSumGutschDukate = 0

    lDatum = Fix(Now)
    dKundenZahl = 0
    dSumKreditkarten = 0
    dSumDukaten = 0
    dSumScheck = 0
    dGegScheck = 0
    dSumECLAST = 0
    dSumBar = 0
    dSumRückBar = 0
    dSumUmsBar = 0
    dSumRückGUTSCH = 0
    dZhlgGutsch = 0
    dUmsatzGutsch = 0
    dwertGutverkauf = 0

    '*******************************************
    '* Was hat der Kunde insgesamt zu zahlen?
    '*******************************************
    If Label333(0).Caption <> "" Then
        dUmsatz = CDbl(Label333(0).Caption) 'dZuZahlen
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Gutscheinen?
    '*******************************************
    If Text1(1).Text <> "" And IsNumeric(Text1(1).Text) Then
        dZhlgGutsch = CDbl(Text1(1).Text) 'dEinrGutsch
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Dukaten?
    '*******************************************
    
    If Text1(3).Text <> "" And IsNumeric(Text1(3).Text) Then
        dSumDukaten = CDbl(Text1(3).Text)
        dSumGutschDukate = dSumDukaten
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Kreditkarte?
    '*******************************************
    If Text1(2).Text <> "" And IsNumeric(Text1(2).Text) Then
        dSumKreditkarten = dSumKreditkarten + CDbl(Text1(2).Text)
        dSumGUTSCHKreditkarten = dSumGUTSCHKreditkarten + dSumKreditkarten
    End If

    If Text1(7).Text <> "" And IsNumeric(Text1(7).Text) Then
        dSumKreditkarten = dSumKreditkarten + CDbl(Text1(7).Text)
        dSumGUTSCHKreditkarten = dSumGUTSCHKreditkarten + dSumKreditkarten
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Scheck?
    '*******************************************
    If Text1(6).Text <> "" And IsNumeric(Text1(6).Text) Then
        dSumScheck = CDbl(Text1(6).Text)
        dGegScheck = dSumScheck
        dSumGUTSCHScheck = dSumScheck
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels ECLAST?
    '*******************************************
    If Text1(5).Text <> "" And IsNumeric(Text1(5).Text) Then
        dSumECLAST = CDbl(Text1(5).Text)
        dSumGUTSCHECLAST = dSumECLAST
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Bargeld?
    '*******************************************
    If Text1(0).Text <> "" And IsNumeric(Text1(0).Text) Then
        dSumBar = CDbl(Text1(0).Text)
        dSumGUTSCHBar = dSumBar
    End If

    '*******************************************
    '* Was kriegt der Kunde mittels Bargeld zurück?
    '*******************************************
    If Label333(3).Caption <> "" Then
        dSumRückBar = CDbl(Label333(3).Caption)
    End If

    '*******************************************
    '* Was kriegt der Kunde mittels RestGutschein zurück?
    '*******************************************
    If Label1(28).Caption <> "" Then
        dSumRückGUTSCH = CDbl(Label1(28).Caption)
    End If
    
    'gegeben Ende*****************************************************************************************
    
    
    
    

    dUmsatzGutsch = dZhlgGutsch '- dSumRückGUTSCH

    dEchterUmsatz = 0
    dWertGutschein = 0

    '*******************************************
    '* Untersuche jeden einzelnen Artikel
    '*******************************************

    lAnzSatz = frmWKL20!List1.ListCount

    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = frmWKL20!List1.list(lAktSatz)

        cArtNr = Mid(cLBSatz, 7, 6)

        '*******************************************
        '* Lies Kennzeichen Umsatzrelevanz
        '*******************************************
        If Len(cLBSatz) > 155 Then
            cUmsOK = Mid(cLBSatz, 156, 1)
        Else
            cUmsOK = "J"
        End If
        If cUmsOK <> "J" And cUmsOK <> "N" Then
            cUmsOK = "J"
        End If

        '*******************************************
        '* Lies regulären Stückpreis
        '*******************************************
        cNormal = Mid(cLBSatz, 128, 9)
        cNormal = Trim$(cNormal)
        cNormal = fnMoveComma2Point$(cNormal)

        '*******************************************
        '* Lies Stückpreis, zu dem verkauft wurde
        '*******************************************
        ctmp = Mid(cLBSatz, 74, 9)
        ctmp = Trim$(ctmp)
        ctmp = fnMoveComma2Point$(ctmp)

        '*******************************************
        '* Lies Gesamtpreis der Position
        '*******************************************
        cPosSumme = Mid(cLBSatz, 94, 9)
        cPosSumme = Trim$(cPosSumme)
        cPosSumme = fnMoveComma2Point$(cPosSumme)

        '*******************************************
        '* Lies Artikelrabatt der Position
        '*******************************************
        cArtRabatt = Mid(cLBSatz, 124, 3)
        cArtRabatt = Trim$(cArtRabatt)
        cArtRabatt = fnMoveComma2Point$(cArtRabatt)

        '**********************************************
        '* Lies den echten Verkaufspreis der Position
        '**********************************************
        cErzielterPreis = Mid(cLBSatz, 60, 9)
        cErzielterPreis = Trim$(cErzielterPreis)
        cErzielterPreis = fnMoveComma2Point$(cErzielterPreis)
        
        
        
        If gbGutscheinBeiVKversteuern = True Then
             '**********************************************
            '* Ermittle die echte Umsatzsumme
            '* -
            '* - kein nicht umsatzrelevanten Artikel
            '**********************************************
            
            
            If cArtNr <> "666666" Then
                If cUmsOK <> "N" Then
                    dEchterUmsatz = dEchterUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                Else
                    dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
                    
                    dFalscherUmsatz = dFalscherUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                End If
            Else
                '******************************************************
                '* Wenn Gutschein, dann summiere verkaufte Gutscheine
                '******************************************************
                
                
                dEchterUmsatz = dEchterUmsatz + Val(cErzielterPreis)
                dKundenZahl = 1
    
                dWertGutschein = dWertGutschein + Val(cErzielterPreis)
                dwertGutverkauf = dwertGutverkauf + Val(cErzielterPreis)
            End If
            
        Else
        

            '**********************************************
            '* Ermittle die echte Umsatzsumme
            '* - keine Gutscheine
            '* - kein nicht umsatzrelevanten Artikel
            '**********************************************
            
            
            If cArtNr <> "666666" Then
                If cUmsOK <> "N" Then
                    dEchterUmsatz = dEchterUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                Else
                    dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
                    
                    dFalscherUmsatz = dFalscherUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                End If
            Else
                '******************************************************
                '* Wenn Gutschein, dann summiere verkaufte Gutscheine
                '******************************************************
                dFalscherUmsatz = dFalscherUmsatz + Val(cErzielterPreis)
                dKundenZahl = 1
    
                dWertGutschein = dWertGutschein + Val(cErzielterPreis)
                dwertGutverkauf = dwertGutverkauf + Val(cErzielterPreis)
            End If
        End If


        '**********************************************
        '* Wenn regulärer Stückpreis und Stückpreis
        '* des Verkaufes abweichen, Zähler für
        '* Sonderpreis um 1 heraufsetzen und die
        '* Sonderpreissumme erhöhen
        '**********************************************
        If Val(ctmp) <> Val(cNormal) Then
            If cNormal = 0 Then
                dSPreisAnz = 0
                dSPreisGes = 0
            Else
                dSPreisAnz = dSPreisAnz + 1
                dSPreisGes = dSPreisGes + Val(cPosSumme)
            End If
        End If

        '*******************************************
        '* Wenn Artikelrabatt gewährt wurde,
        '* Zähler für Artikelrabatt um 1 heraufsetzen
        '* und ArtRabattsumme erhöhen
        '*******************************************
        If Val(cArtRabatt) <> 0 Then
            ctmp = Mid(cLBSatz, 84, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)

            dArtRabAnz = dArtRabAnz + 1
            dArtRabGes = dArtRabGes + Val(ctmp)
        ElseIf frmWKL20!Label2(3).Visible Then
            dGesRabAnz = dGesRabAnz + 1
        End If

        '*******************************************
        '* Wenn erzielter Preis < 0, dann
        '* Zähler für Storno um 1 heraufsetzen
        '* und Stornosumme erhöhen
        '*******************************************
        If Val(cErzielterPreis) < 0 Then
            If IstArtikelnichtStornierfähig(cArtNr) = False Then
                dStornoWert = dStornoWert + Val(cErzielterPreis)
                lStornoAnz = lStornoAnz + 1
            End If
        End If
    Next lAktSatz

    If frmWKL20!Label2(3).Visible Then
        dGesRabGes = fnHoleGesamtRabattModul20#()
    End If

    If dwertGutverkauf > 0 Then
        If dZhlgGutsch <= dwertGutverkauf Then
            dwertGutverkauf = dZhlgGutsch
            updateafcstat "ZHLGGUTSCH", (-1 * dwertGutverkauf), gcKasNum
        End If
    End If

    cSQL = "Select * from AFCSTAT where ADATE = " & Trim$(Str$(lDatum)) & " and KASNUM = " & gcKasNum
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    
    
    

    If dNichtUmsatz > 0 Then
    
    
    
    
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    


    dSumUmsBar = dSumBar - dSumRückBar + dGutscheinauszahlung

    Dim dRestfalscherumsatz As Double
    dRestfalscherumsatz = 0
    
    Dim dZuVergebenerEchterumsatz As Double
    dZuVergebenerEchterumsatz = 0
    
    If dEchterUmsatz > 0 Then dZuVergebenerEchterumsatz = dEchterUmsatz

    If dFalscherUmsatz > 0 Then 'Wert der Gutscheinverkäufe
    
    
        dRestfalscherumsatz = dFalscherUmsatz
        
        
''
''
''        ' CW20190320 TEST Begin
''        If dZhlgGutsch > dRestfalscherumsatz Then
''            dRestfalscherumsatz = 0
''        Else
''            dRestfalscherumsatz = dRestfalscherumsatz - dZhlgGutsch
''        End If
''           ' CW20190320 TEST End
''
        
        
        
        
        'dukaten

        If dSumDukaten < dRestfalscherumsatz Then 'dann dukaten abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumDukaten = dSumDukaten
                dRestfalscherumsatz = 0
            Else
                If dSumDukaten >= dZuVergebenerEchterumsatz Then
                    dSumDukaten = dZuVergebenerEchterumsatz

                    dSumGutschDukate = dSumDukaten - dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumDukaten
                    dSumGutschDukate = 0
                End If
            End If
        Else 'falls dukaten mehr oder gleich als dRestfalscherumsatz
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumDukaten = dSumDukaten
            Else
                dSumDukaten = dSumDukaten - dRestfalscherumsatz + dZhlgGutsch
                dSumGutschDukate = dRestfalscherumsatz
            End If
            dRestfalscherumsatz = 0
        End If


        'Kreditkarten

        If dSumKreditkarten < dRestfalscherumsatz Then 'dann KK abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumKreditkarten = dSumKreditkarten
                dRestfalscherumsatz = 0
            Else
                If dSumKreditkarten >= dZuVergebenerEchterumsatz Then
                    dSumGUTSCHKreditkarten = dSumKreditkarten - dZuVergebenerEchterumsatz 'CW20190320
                    dSumKreditkarten = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumKreditkarten
                    dSumGUTSCHKreditkarten = 0 'CW20190320
                End If
            End If

        Else
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumKreditkarten = dSumKreditkarten
                dRestfalscherumsatz = 0
            Else
                If dSumKreditkarten >= dZuVergebenerEchterumsatz Then
                    dSumGUTSCHKreditkarten = dSumKreditkarten - dZuVergebenerEchterumsatz 'CW20190320
                    dSumKreditkarten = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dSumKreditkarten = dSumKreditkarten
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumKreditkarten
                    dSumGUTSCHBar = 0  'CW20190320
                End If
            End If
        End If

        '*****************BAR
        If dSumUmsBar < dRestfalscherumsatz Then 'erst bar abfragen

            If dZhlgGutsch > dRestfalscherumsatz Then

                dSumUmsBar = dSumUmsBar ' - dFalscherUmsatz + dZhlgGutsch
                dRestfalscherumsatz = 0
            Else
                If dSumUmsBar >= dZuVergebenerEchterumsatz Then
                    dSumGUTSCHBar = dSumUmsBar - dZuVergebenerEchterumsatz 'CW20190320
                    dSumUmsBar = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumUmsBar
                    dSumGUTSCHBar = 0 'CW20190320
                End If
            End If

        Else
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumUmsBar = dSumUmsBar ' - dFalscherUmsatz + dZhlgGutsch
                dRestfalscherumsatz = 0
            Else

                If dSumUmsBar >= dZuVergebenerEchterumsatz Then
                    dSumGUTSCHBar = dSumUmsBar - dZuVergebenerEchterumsatz 'CW20190320
                    dSumUmsBar = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dSumUmsBar = dSumUmsBar
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumUmsBar
                    dSumGUTSCHBar = 0  'CW20190320
                End If
            End If
        End If

        '*****************BAR Ende



        'scheck

        If dSumScheck < dRestfalscherumsatz Then 'dann Scheck abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumScheck = dSumScheck
                dRestfalscherumsatz = 0
            Else
                If dSumScheck >= dZuVergebenerEchterumsatz Then
                    dSumScheck = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else

                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumScheck

                End If
'                dRestfalscherumsatz = dFalscherUmsatz
            End If
        Else
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumScheck = dSumScheck
            Else
                dSumScheck = dSumScheck - dRestfalscherumsatz + dZhlgGutsch
                dSumGUTSCHScheck = dRestfalscherumsatz 'CW20190320
            End If
             dRestfalscherumsatz = 0 'CW20190320
        End If
             
        'Ec Last
        If dSumECLAST < dRestfalscherumsatz Then 'dann ec last abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumECLAST = dSumECLAST
                dRestfalscherumsatz = 0
            Else
                If dSumECLAST >= dZuVergebenerEchterumsatz Then
                    dSumECLAST = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumECLAST
                End If
                dRestfalscherumsatz = 0
            End If
        Else
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumECLAST = dSumECLAST
            Else
                dSumECLAST = dSumECLAST - dRestfalscherumsatz + dZhlgGutsch
                dSumGUTSCHECLAST = dRestfalscherumsatz 'CW20190320
            End If
             dRestfalscherumsatz = 0 'CW20190320
        End If


    End If
    
    

    '**************************************************************
    '* Wieviel Umsatz hat der Kunde in bar gemacht?
    '**************************************************************
    
'    If bUmsatz = True Then
    If dEchterUmsatz > 0 Then
        'erzielter Umsatz über Bargeld
        
        
    
        If Not IsNull(rsrs!UMS_BAR) Then
            rsrs!UMS_BAR = rsrs!UMS_BAR + dSumUmsBar
        Else
            rsrs!UMS_BAR = dSumUmsBar
        End If
    
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde mit EC-Lastschrift gemacht?
        '**************************************************************
        If Not IsNull(rsrs!UMS_LAST) Then
            rsrs!UMS_LAST = rsrs!UMS_LAST + dSumECLAST
        Else
            rsrs!UMS_LAST = dSumECLAST
        End If
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde über Dukaten gemacht?
        '**************************************************************
        If Not IsNull(rsrs!DUKA) Then
            rsrs!DUKA = rsrs!DUKA + dSumDukaten
        Else
            rsrs!DUKA = dSumDukaten
        End If
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde mit Schecks gemacht?
        '**************************************************************
        If Not IsNull(rsrs!UMS_SCHECK) Then
            rsrs!UMS_SCHECK = rsrs!UMS_SCHECK + dSumScheck
        Else
            rsrs!UMS_SCHECK = dSumScheck
        End If
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde mit Kreditkarten gemacht?
        '**************************************************************
        If Not IsNull(rsrs!UMS_KARTE) Then
            rsrs!UMS_KARTE = rsrs!UMS_KARTE + dSumKreditkarten
            schreibeProtokollUNITXT CStr(dSumKreditkarten), "Kartenzahlung"
        Else
            rsrs!UMS_KARTE = dSumKreditkarten
            schreibeProtokollUNITXT CStr(dSumKreditkarten), "Kartenzahlung"
        End If
        
    End If

    '**************************************************************
    '* Datum und Kassennummer des Verbuchens schreiben
    '**************************************************************
    rsrs!ADATE = lDatum
    rsrs!kasnum = Val(gcKasNum)

    '**************************************************************
    '* Betrag der eingereichten Gutscheine verbuchen
    '**************************************************************
    If Not IsNull(rsrs!EINRGUTSCH) Then
        rsrs!EINRGUTSCH = rsrs!EINRGUTSCH + dZhlgGutsch
    Else
        rsrs!EINRGUTSCH = dZhlgGutsch
    End If

    '**************************************************************
    '* Betrag der generierten Rest-Gutscheine verbuchen
    '**************************************************************
    If Not IsNull(rsrs!RESTGUTSCH) Then
        rsrs!RESTGUTSCH = rsrs!RESTGUTSCH + dSumRückGUTSCH
    Else
        rsrs!RESTGUTSCH = dSumRückGUTSCH
    End If

    'Achtung
    '**************************************************************
    '* Betrag der Gutschein-AUszahlung verbuchen
    '**************************************************************
    If Not IsNull(rsrs!AUSZGUTSCH) Then
        rsrs!AUSZGUTSCH = rsrs!AUSZGUTSCH + Format(dGutscheinauszahlung, "######0.00")
    Else
        rsrs!AUSZGUTSCH = Format(dGutscheinauszahlung, "######0.00")
    End If


    'Achtung
    '**************************************************************
    '* Betrag des Umsatzes durch Gutschein-Einreichungen verbuchen
    '**************************************************************
    
    
    'Test echter Umsatz
    

    If dwertGutverkauf > 0 Then

        If dZhlgGutsch > dwertGutverkauf Then

            If Not IsNull(rsrs!ZHLGGUTSCH) Then
                rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
            Else
                rsrs!ZHLGGUTSCH = dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
                
                If rsrs!ZHLGGUTSCH < 0 Then rsrs!ZHLGGUTSCH = 0
            End If
        
        Else
        
            If Not IsNull(rsrs!ZHLGGUTSCH) Then
                rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH
            Else
                rsrs!ZHLGGUTSCH = dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH
            End If
        
        End If


    Else
        If Not IsNull(rsrs!ZHLGGUTSCH) Then
            rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
        Else
            rsrs!ZHLGGUTSCH = dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
        End If

    End If
    
    
    
    
    
    
    
    
    

    'Achtung
    '**************************************************************
    '* Betrag des Gutschein-Verkäufe verbuchen (Bar?)
    '**************************************************************
    If Not IsNull(rsrs!GUTSCHEIN) Then
        rsrs!GUTSCHEIN = rsrs!GUTSCHEIN + dWertGutschein
    Else
        rsrs!GUTSCHEIN = dWertGutschein
    End If

    '**************************************************************
    '* Betrag der Scheck - Verkäufe verbuchen
    '**************************************************************

    If Not IsNull(rsrs!SCHVERKAUF) Then
        rsrs!SCHVERKAUF = rsrs!SCHVERKAUF + dGegScheck
    Else
        rsrs!SCHVERKAUF = dGegScheck
    End If

    '**************************************************************
    '* Sonderpreise verbuchen
    '**************************************************************
    If Not IsNull(rsrs!SPREIS_ANZ) Then
        rsrs!SPREIS_ANZ = rsrs!SPREIS_ANZ + dSPreisAnz
    Else
        rsrs!SPREIS_ANZ = dSPreisAnz
    End If

    If Not IsNull(rsrs!SPREIS_GES) Then
        rsrs!SPREIS_GES = rsrs!SPREIS_GES + dSPreisGes
    Else
        rsrs!SPREIS_GES = dSPreisGes
    End If



    '**************************************************************
    '* Kundenzahl schreiben
    '**************************************************************
    If Not IsNull(rsrs!Kundenzahl) Then
        rsrs!Kundenzahl = rsrs!Kundenzahl + dKundenZahl
    Else
        rsrs!Kundenzahl = dKundenZahl
    End If
    '**************************************************************
    '* Artikelrabatte schreiben
    '**************************************************************
    If Not IsNull(rsrs!ARTRAB_ANZ) Then
        rsrs!ARTRAB_ANZ = rsrs!ARTRAB_ANZ + dArtRabAnz
    Else
        rsrs!ARTRAB_ANZ = dArtRabAnz
    End If

    If Not IsNull(rsrs!ARTRAB_GES) Then
        rsrs!ARTRAB_GES = rsrs!ARTRAB_GES + dArtRabGes
    Else
        rsrs!ARTRAB_GES = dArtRabGes
    End If
    '**************************************************************
    '* Gesamtrabatte schreiben
    '**************************************************************
    If Not IsNull(rsrs!GESRAB_ANZ) Then
        rsrs!GESRAB_ANZ = rsrs!GESRAB_ANZ + dGesRabAnz
    Else
        rsrs!GESRAB_ANZ = dGesRabAnz
    End If

    If Not IsNull(rsrs!GESRAB_GES) Then
        rsrs!GESRAB_GES = rsrs!GESRAB_GES + dGesRabGes
    Else
        rsrs!GESRAB_GES = dGesRabGes
    End If
    '**************************************************************
    '* Bonnummer schreiben
    '**************************************************************
    
    If Not IsNull(rsrs!BELEGNR) Then
        If gdBonNr < CLng(rsrs!BELEGNR) Then
            
        Else
            rsrs!BELEGNR = gdBonNr
        End If
    Else
        rsrs!BELEGNR = gdBonNr
    End If
    
    

    '**************************************************************
    '* Stornos schreiben
    '**************************************************************
    If Not IsNull(rsrs!STORNO_GES) Then
        rsrs!STORNO_GES = rsrs!STORNO_GES + dStornoWert
    Else
        rsrs!STORNO_GES = dStornoWert
    End If

    If Not IsNull(rsrs!STORNO_ANZ) Then
        rsrs!STORNO_ANZ = rsrs!STORNO_ANZ + lStornoAnz
    Else
        rsrs!STORNO_ANZ = lStornoAnz
    End If

    '******************************************
    '* Wie sind neue Gutscheine bezahlt worden
    '******************************************

    '******************************************
    '* Betrag Gutscheinverkäufe in Bar setzen
    '******************************************
    Dim dRestZahlungsmittelforGutsch As Double
    dRestZahlungsmittelforGutsch = 0

    If dWertGutschein > 0 Then
    
        dRestZahlungsmittelforGutsch = dWertGutschein
        
        If dZhlgGutsch < dRestZahlungsmittelforGutsch Then 'dann gutschein aus gutschein abfragen
            dGutschGutschAnteil = dZhlgGutsch
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dGutschGutschAnteil
        Else
            dGutschGutschAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If
        
        If dSumGutschDukate < dRestZahlungsmittelforGutsch Then 'dann dukate abfragen
            dGutschDukateAnteil = dSumGutschDukate
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGutschDukate
        Else
            dGutschDukateAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If

'      If (dSumGUTSCHBar - dSumRückBar + dGutscheinauszahlung) < dRestZahlungsmittelforGutsch Then  'erst bar abfragen
'             dGutschBarAnteil = dSumGUTSCHBar - dSumRückBar + dGutscheinauszahlung '+ dGutschGutschAnteil
'            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - (dSumGUTSCHBar - dSumRückBar + dGutscheinauszahlung)
'        Else
'            dGutschBarAnteil = dRestZahlungsmittelforGutsch
'            dRestZahlungsmittelforGutsch = 0
'        End If
       
        If dSumGUTSCHBar < dRestZahlungsmittelforGutsch Then  'erst bar abfragen
            dGutschBarAnteil = dSumGUTSCHBar
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHBar
        Else
            dGutschBarAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If


        If dSumGUTSCHScheck < dRestZahlungsmittelforGutsch Then 'dann Scheck abfragen
            dGutschScheckAnteil = dSumGUTSCHScheck
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHScheck
        Else
            dGutschScheckAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If

        If dSumGUTSCHECLAST < dRestZahlungsmittelforGutsch Then 'dann last abfragen
            dGutschECLASTAnteil = dSumGUTSCHECLAST
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHECLAST
        Else
            dGutschECLASTAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If
        
        
        
        If dSumGUTSCHKreditkarten < dRestZahlungsmittelforGutsch Then 'dann KK abfragen
            dGutschKreditkartenAnteil = dSumGUTSCHKreditkarten
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHKreditkarten
        Else
            dGutschKreditkartenAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If

    End If

    If dGutschBarAnteil > 0 Then
        '******************************************
        '* Wert des Barverkaufes setzen
        '******************************************
        If Not IsNull(rsrs!BARVERKAUF) Then
            rsrs!BARVERKAUF = rsrs!BARVERKAUF + (dSumBar - dSumRückBar + dGutscheinauszahlung - dGutschBarAnteil)
        Else
            rsrs!BARVERKAUF = (dSumBar - dSumRückBar + dGutscheinauszahlung - dGutschBarAnteil)
        End If

    Else
        '******************************************
        '* Wert des Barverkaufes setzen
        '******************************************
        If Not IsNull(rsrs!BARVERKAUF) Then
            rsrs!BARVERKAUF = rsrs!BARVERKAUF + dSumBar - dSumRückBar + dGutscheinauszahlung
        Else
            rsrs!BARVERKAUF = dSumBar - dSumRückBar + dGutscheinauszahlung
        End If
    End If
    
    
    '**************************************************************
    '* Gutscheine gekauft mit Gutschein bezahlt!
    '**************************************************************
    If dGutschGutschAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHGUTSCH) Then
            rsrs!GUTSCHGUTSCH = rsrs!GUTSCHGUTSCH + dGutschGutschAnteil
        Else
            rsrs!GUTSCHGUTSCH = dGutschGutschAnteil
        End If
    End If
    
    If dGutschBarAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHBAR) Then
            rsrs!GUTSCHBAR = rsrs!GUTSCHBAR + Format(dGutschBarAnteil, "######0.00") 'Hier ist Format() neu 21.10.13
        Else
            rsrs!GUTSCHBAR = Format(dGutschBarAnteil, "######0.00") 'Hier ist Format() neu 21.10.13
        End If
    End If

    If dGutschKreditkartenAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHKAR) Then
            rsrs!GUTSCHKAR = rsrs!GUTSCHKAR + dGutschKreditkartenAnteil
            schreibeProtokollUNITXT CStr(dGutschKreditkartenAnteil), "Kartenzahlung"
        Else
            rsrs!GUTSCHKAR = dGutschKreditkartenAnteil
            schreibeProtokollUNITXT CStr(dGutschKreditkartenAnteil), "Kartenzahlung"
        End If
    End If

    If dGutschECLASTAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHLAST) Then
            rsrs!GUTSCHLAST = rsrs!GUTSCHLAST + dGutschECLASTAnteil
        Else
            rsrs!GUTSCHLAST = dGutschECLASTAnteil
        End If
    End If

    If dGutschScheckAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHSCH) Then
            rsrs!GUTSCHSCH = rsrs!GUTSCHSCH + dGutschScheckAnteil
        Else
            rsrs!GUTSCHSCH = dGutschScheckAnteil
        End If
    End If





    '**************************************************************
    '* Schreibvorgang durchführen
    '**************************************************************
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "UpdateAFCStat68Test5"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub UpdateAFCStat68Test6()
    On Error GoTo LOKAL_ERROR

    Dim lStornoAnz      As Long
    Dim lDatum          As Long
    Dim lAktSatz        As Long
    Dim lAnzSatz        As Long

    Dim cArtNr          As String
    Dim cUmsOK          As String
    Dim cSQL            As String
    Dim cErzielterPreis As String
    Dim ctmp            As String
    Dim cNormal         As String
    Dim cPosSumme       As String
    Dim cArtRabatt      As String
    Dim cLBSatz         As String

    Dim dUmsatz         As Double     'Summe des Verkaufs ohne Gutscheine
    Dim dEchterUmsatz   As Double     'Summe des Verkaufs ohne Gutscheine
    Dim dFalscherUmsatz As Double     'Summe des Verkaufs ohne Gutscheine
    
    Dim dNichtUmsatz    As Double     'Summe nichtumsatzrelevanter VK zB Dukaten

    Dim dUmsatz2        As Double     'Summe des Verkaufs inkl. Gutscheine
    Dim dSPreisAnz      As Double     'Anzahl Positionen mit Sonderpreis
    Dim dSPreisGes      As Double     'Summe aller Positionen mit Sonderpreis
    Dim dKundenZahl     As Double     'Konstante 1
    Dim dArtRabAnz      As Double     'Anzahl Positionen mit Artikelrabatt
    Dim dArtRabGes      As Double     'Summe des gegebenen Artikelrabatts
    Dim dGesRabAnz      As Double     'Anzahl Positionen mit Gesamtrabatt
    Dim dGesRabGes      As Double     'Summe des gegebenen Gesamtrabatts
    Dim dWertGutschein  As Double

    Dim rsrs                As Recordset

    Dim dStornoWert         As Double
    Dim dZhlgGutsch         As Double
    Dim dSumDukaten         As Double

    Dim dSumKreditkarten    As Double
    Dim dSumScheck          As Double
    Dim dSumECLAST          As Double
    Dim dSumBar             As Double
    
    Dim dGegScheck          As Double

    Dim dSumGUTSCHKreditkarten    As Double
    Dim dSumGUTSCHScheck          As Double
    Dim dSumGUTSCHECLAST          As Double
    Dim dSumGUTSCHBar             As Double
    Dim dSumGutschDukate          As Double
    Dim dSumGUTSCHKredit          As Double 'CW20190627


    Dim dSumRückBar                 As Double
    Dim dSumRückGUTSCH              As Double
    Dim dUmsatzGutsch               As Double
    Dim dSumUmsBar                  As Double

    Dim dGutschGutschAnteil         As Double
    Dim dGutschBarAnteil            As Double
    Dim dGutschECLASTAnteil         As Double
    Dim dGutschKreditkartenAnteil   As Double
    Dim dGutschScheckAnteil         As Double
    Dim dGutschDukateAnteil         As Double
    Dim dGutschKreditAnteil         As Double 'CW20190627
    
    Dim bUmsatz                     As Boolean
    
    bUmsatz = True
    
    dNichtUmsatz = 0

    dGutschGutschAnteil = 0
    dGutschBarAnteil = 0
    dGutschECLASTAnteil = 0
    dGutschKreditkartenAnteil = 0
    dGutschScheckAnteil = 0
    dGutschDukateAnteil = 0
    dGutschKreditAnteil = 0 'CW20190627


    dSumGUTSCHKreditkarten = 0
    dSumGUTSCHScheck = 0
    dSumGUTSCHECLAST = 0
    dSumGUTSCHBar = 0
    dSumGutschDukate = 0
    dSumGUTSCHKredit = 0 'CW20190627

    lDatum = Fix(Now)
    dKundenZahl = 0
    dSumKreditkarten = 0
    dSumDukaten = 0
    dSumScheck = 0
    dGegScheck = 0
    dSumECLAST = 0
    dSumBar = 0
    dSumRückBar = 0
    dSumUmsBar = 0
    dSumRückGUTSCH = 0
    dZhlgGutsch = 0
    dUmsatzGutsch = 0
    dwertGutverkauf = 0

    '*******************************************
    '* Was hat der Kunde insgesamt zu zahlen?
    '*******************************************
    If Label333(0).Caption <> "" Then
        dUmsatz = CDbl(Label333(0).Caption) 'dZuZahlen
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Gutscheinen?
    '*******************************************
    If Text1(1).Text <> "" And IsNumeric(Text1(1).Text) Then
        dZhlgGutsch = CDbl(Text1(1).Text) 'dEinrGutsch
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Dukaten?
    '*******************************************
    
    If Text1(3).Text <> "" And IsNumeric(Text1(3).Text) Then
        dSumDukaten = CDbl(Text1(3).Text)
        dSumGutschDukate = dSumDukaten
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Kreditkarte?
    '*******************************************
    If Text1(2).Text <> "" And IsNumeric(Text1(2).Text) Then
        dSumKreditkarten = dSumKreditkarten + CDbl(Text1(2).Text)
        dSumGUTSCHKreditkarten = dSumGUTSCHKreditkarten + dSumKreditkarten
    End If

    If Text1(7).Text <> "" And IsNumeric(Text1(7).Text) Then
        dSumKreditkarten = dSumKreditkarten + CDbl(Text1(7).Text)
        dSumGUTSCHKreditkarten = dSumGUTSCHKreditkarten + dSumKreditkarten
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Scheck?
    '*******************************************
    If Text1(6).Text <> "" And IsNumeric(Text1(6).Text) Then
        dSumScheck = CDbl(Text1(6).Text)
        dGegScheck = dSumScheck
        dSumGUTSCHScheck = dSumScheck
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels ECLAST?
    '*******************************************
    If Text1(5).Text <> "" And IsNumeric(Text1(5).Text) Then
        dSumECLAST = CDbl(Text1(5).Text)
        dSumGUTSCHECLAST = dSumECLAST
    End If

    '*******************************************
    '* Was zahlt der Kunde mittels Bargeld?
    '*******************************************
    If Text1(0).Text <> "" And IsNumeric(Text1(0).Text) Then
        dSumBar = CDbl(Text1(0).Text)
        dSumGUTSCHBar = dSumBar
    End If

    '*******************************************
    '* Was kriegt der Kunde mittels Bargeld zurück?
    '*******************************************
    If Label333(3).Caption <> "" Then
        dSumRückBar = CDbl(Label333(3).Caption)
    End If

    '*******************************************
    '* Was kriegt der Kunde mittels RestGutschein zurück?
    '*******************************************
    If Label1(28).Caption <> "" Then
        dSumRückGUTSCH = CDbl(Label1(28).Caption)
    End If
    
    'gegeben Ende*****************************************************************************************
    
    
    
    

    dUmsatzGutsch = dZhlgGutsch '- dSumRückGUTSCH

    dEchterUmsatz = 0
    dWertGutschein = 0

    '*******************************************
    '* Untersuche jeden einzelnen Artikel
    '*******************************************

    lAnzSatz = frmWKL20!List1.ListCount

    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = frmWKL20!List1.list(lAktSatz)

        cArtNr = Mid(cLBSatz, 7, 6)

        '*******************************************
        '* Lies Kennzeichen Umsatzrelevanz
        '*******************************************
        If Len(cLBSatz) > 155 Then
            cUmsOK = Mid(cLBSatz, 156, 1)
        Else
            cUmsOK = "J"
        End If
        If cUmsOK <> "J" And cUmsOK <> "N" Then
            cUmsOK = "J"
        End If

        '*******************************************
        '* Lies regulären Stückpreis
        '*******************************************
        cNormal = Mid(cLBSatz, 128, 9)
        cNormal = Trim$(cNormal)
        cNormal = fnMoveComma2Point$(cNormal)

        '*******************************************
        '* Lies Stückpreis, zu dem verkauft wurde
        '*******************************************
        ctmp = Mid(cLBSatz, 74, 9)
        ctmp = Trim$(ctmp)
        ctmp = fnMoveComma2Point$(ctmp)

        '*******************************************
        '* Lies Gesamtpreis der Position
        '*******************************************
        cPosSumme = Mid(cLBSatz, 94, 9)
        cPosSumme = Trim$(cPosSumme)
        cPosSumme = fnMoveComma2Point$(cPosSumme)

        '*******************************************
        '* Lies Artikelrabatt der Position
        '*******************************************
        cArtRabatt = Mid(cLBSatz, 124, 3)
        cArtRabatt = Trim$(cArtRabatt)
        cArtRabatt = fnMoveComma2Point$(cArtRabatt)

        '**********************************************
        '* Lies den echten Verkaufspreis der Position
        '**********************************************
        cErzielterPreis = Mid(cLBSatz, 60, 9)
        cErzielterPreis = Trim$(cErzielterPreis)
        cErzielterPreis = fnMoveComma2Point$(cErzielterPreis)
        
        
        
        If gbGutscheinBeiVKversteuern = True Then
             '**********************************************
            '* Ermittle die echte Umsatzsumme
            '* -
            '* - kein nicht umsatzrelevanten Artikel
            '**********************************************
            
            
            If cArtNr <> "666666" Then
                If cUmsOK <> "N" Then
                    dEchterUmsatz = dEchterUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                Else
                    dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
                    
                    dFalscherUmsatz = dFalscherUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                End If
            Else
                '******************************************************
                '* Wenn Gutschein, dann summiere verkaufte Gutscheine
                '******************************************************
                
                
                dEchterUmsatz = dEchterUmsatz + Val(cErzielterPreis)
                dKundenZahl = 1
    
                dWertGutschein = dWertGutschein + Val(cErzielterPreis)
                dwertGutverkauf = dwertGutverkauf + Val(cErzielterPreis)
            End If
            
        Else
        

            '**********************************************
            '* Ermittle die echte Umsatzsumme
            '* - keine Gutscheine
            '* - kein nicht umsatzrelevanten Artikel
            '**********************************************
            
            
            If cArtNr <> "666666" Then
                If cUmsOK <> "N" Then
                    dEchterUmsatz = dEchterUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                Else
                    dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
                    
                    dFalscherUmsatz = dFalscherUmsatz + Val(cErzielterPreis)
                    dKundenZahl = 1
                End If
            Else
                '******************************************************
                '* Wenn Gutschein, dann summiere verkaufte Gutscheine
                '******************************************************
                dFalscherUmsatz = dFalscherUmsatz + Val(cErzielterPreis)
                dKundenZahl = 1
    
                dWertGutschein = dWertGutschein + Val(cErzielterPreis)
                dwertGutverkauf = dwertGutverkauf + Val(cErzielterPreis)
            End If
        End If


        '**********************************************
        '* Wenn regulärer Stückpreis und Stückpreis
        '* des Verkaufes abweichen, Zähler für
        '* Sonderpreis um 1 heraufsetzen und die
        '* Sonderpreissumme erhöhen
        '**********************************************
        If Val(ctmp) <> Val(cNormal) Then
            If cNormal = 0 Then
                dSPreisAnz = 0
                dSPreisGes = 0
            Else
                dSPreisAnz = dSPreisAnz + 1
                dSPreisGes = dSPreisGes + Val(cPosSumme)
            End If
        End If

        '*******************************************
        '* Wenn Artikelrabatt gewährt wurde,
        '* Zähler für Artikelrabatt um 1 heraufsetzen
        '* und ArtRabattsumme erhöhen
        '*******************************************
        If Val(cArtRabatt) <> 0 Then
            ctmp = Mid(cLBSatz, 84, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)

            dArtRabAnz = dArtRabAnz + 1
            dArtRabGes = dArtRabGes + Val(ctmp)
        ElseIf frmWKL20!Label2(3).Visible Then
            dGesRabAnz = dGesRabAnz + 1
        End If

        '*******************************************
        '* Wenn erzielter Preis < 0, dann
        '* Zähler für Storno um 1 heraufsetzen
        '* und Stornosumme erhöhen
        '*******************************************
        If Val(cErzielterPreis) < 0 Then
            If IstArtikelnichtStornierfähig(cArtNr) = False Then
                dStornoWert = dStornoWert + Val(cErzielterPreis)
                lStornoAnz = lStornoAnz + 1
            End If
        End If
    Next lAktSatz

    If frmWKL20!Label2(3).Visible Then
        dGesRabGes = fnHoleGesamtRabattModul20#()
    End If


' CW20190627 Begin
'''''    If dwertGutverkauf > 0 Then
'''''        If dZhlgGutsch <= dwertGutverkauf Then
'''''            dwertGutverkauf = dZhlgGutsch
'''''            updateafcstat "ZHLGGUTSCH", (-1 * dwertGutverkauf), gcKasNum
'''''        End If
'''''    End If
' CW20190627 END

    cSQL = "Select * from AFCSTAT where ADATE = " & Trim$(Str$(lDatum)) & " and KASNUM = " & gcKasNum
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    
    
    

    If dNichtUmsatz > 0 Then
    
    
    
    
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    


    dSumUmsBar = dSumBar - dSumRückBar + dGutscheinauszahlung

    Dim dRestfalscherumsatz As Double
    dRestfalscherumsatz = 0
    
    Dim dZuVergebenerEchterumsatz As Double
    dZuVergebenerEchterumsatz = 0
    
    If dEchterUmsatz > 0 Then
        dZuVergebenerEchterumsatz = dEchterUmsatz
    End If
    
    If dFalscherUmsatz > 0 Then 'Wert der Gutscheinverkäufe
        dRestfalscherumsatz = dFalscherUmsatz

        'Reihenfolge: Kreditkarte, Dukate, Bar, Scheck, ECLast, Gutsch
        If dRestfalscherumsatz > 0 Then 'CW20190626 begin
            If dSumKreditkarten > 0 Then
                If dRestfalscherumsatz > dSumKreditkarten Then
                    dRestfalscherumsatz = dRestfalscherumsatz - dSumKreditkarten
                    dSumKreditkarten = 0
                Else
                    dSumKreditkarten = dSumKreditkarten - dRestfalscherumsatz
                    dRestfalscherumsatz = 0
                End If
            End If
                       
           If dSumDukaten > 0 Then
                If dRestfalscherumsatz > dSumDukaten Then
                    dRestfalscherumsatz = dRestfalscherumsatz - dSumDukaten
                    dSumDukaten = 0
                Else
                    dSumDukaten = dSumDukaten - dRestfalscherumsatz
                    dRestfalscherumsatz = 0
                End If
           End If
           
            If dSumUmsBar > 0 Then
                If dRestfalscherumsatz > dSumUmsBar Then
                    dRestfalscherumsatz = dRestfalscherumsatz - dSumUmsBar
                    dSumUmsBar = 0
                Else
                    dSumUmsBar = dSumUmsBar - dRestfalscherumsatz
                    dRestfalscherumsatz = 0
                End If
            End If
            
            If dSumScheck > 0 Then
                If dRestfalscherumsatz > dSumScheck Then
                    dRestfalscherumsatz = dRestfalscherumsatz - dSumScheck
                    dSumScheck = 0
                Else
                    dSumScheck = dSumScheck - dRestfalscherumsatz
                    dRestfalscherumsatz = 0
                End If
            End If
            
            If dSumECLAST > 0 Then
                If dRestfalscherumsatz > dSumECLAST Then
                    dRestfalscherumsatz = dRestfalscherumsatz - dSumECLAST
                    dSumECLAST = 0
                Else
                    dSumECLAST = dSumECLAST - dRestfalscherumsatz
                    dRestfalscherumsatz = 0
                End If
            End If

            If dUmsatzGutsch > 0 Then
                If dRestfalscherumsatz > dUmsatzGutsch Then
                    dRestfalscherumsatz = dRestfalscherumsatz - dUmsatzGutsch
                    dUmsatzGutsch = 0
                Else
                    dUmsatzGutsch = dUmsatzGutsch - dRestfalscherumsatz
                    dRestfalscherumsatz = 0
                End If 'CW20190626 end
            End If
        End If
      
        
        
        
        'Reihenfolge: Kreditkarte, Dukate, Bar, Scheck, ECLast, Gutsch
 'Kreditkarten
       
If dSumKreditkarten > 0 Then 'CW20190626
        If dSumKreditkarten < dRestfalscherumsatz Then 'dann KK abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumKreditkarten = dSumKreditkarten
                dRestfalscherumsatz = 0
            Else
                If dSumKreditkarten >= dZuVergebenerEchterumsatz Then
                    dSumGUTSCHKreditkarten = dSumKreditkarten - dZuVergebenerEchterumsatz 'CW20190320
                    dSumKreditkarten = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumKreditkarten
                    dSumGUTSCHKreditkarten = 0 'CW20190320
                End If
            End If

        Else
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumKreditkarten = dSumKreditkarten
                dRestfalscherumsatz = 0
            Else
                If dSumKreditkarten >= dZuVergebenerEchterumsatz Then
                    dSumGUTSCHKreditkarten = dSumKreditkarten - dZuVergebenerEchterumsatz 'CW20190320
                    dSumKreditkarten = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dSumKreditkarten = dSumKreditkarten
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumKreditkarten
                    dSumGUTSCHBar = 0  'CW20190320
                End If
            End If
        End If
End If 'CW20190626
        
        
        'dukaten
    If dSumDukaten > 0 Then 'CW20190626
   
        If dSumDukaten < dRestfalscherumsatz Then 'dann dukaten abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumDukaten = dSumDukaten
                dRestfalscherumsatz = 0
            Else
                If dSumDukaten >= dZuVergebenerEchterumsatz Then
                    dSumDukaten = dZuVergebenerEchterumsatz

                    dSumGutschDukate = dSumDukaten - dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumDukaten
                    dSumGutschDukate = 0
                End If
            End If
        Else 'falls dukaten mehr oder gleich als dRestfalscherumsatz
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumDukaten = dSumDukaten
            Else
                dSumDukaten = dSumDukaten - dRestfalscherumsatz + dZhlgGutsch
                dSumGutschDukate = dRestfalscherumsatz
            End If
            dRestfalscherumsatz = 0
        End If
    End If


        '*****************BAR
        If dSumUmsBar > 0 Then 'CW20190626
        
        If dSumUmsBar < dRestfalscherumsatz Then 'erst bar abfragen

            If dZhlgGutsch > dRestfalscherumsatz Then

                dSumUmsBar = dSumUmsBar ' - dFalscherUmsatz + dZhlgGutsch
                dRestfalscherumsatz = 0
            Else
                If dSumUmsBar >= dZuVergebenerEchterumsatz Then
                    dSumGUTSCHBar = dSumUmsBar - dZuVergebenerEchterumsatz 'CW20190320
                    dSumUmsBar = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumUmsBar
                    dSumGUTSCHBar = 0 'CW20190320
                End If
            End If

        Else
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumUmsBar = dSumUmsBar ' - dFalscherUmsatz + dZhlgGutsch
                dRestfalscherumsatz = 0
            Else

                If dSumUmsBar >= dZuVergebenerEchterumsatz Then
                    dSumGUTSCHBar = dSumUmsBar - dZuVergebenerEchterumsatz 'CW20190320
                    dSumUmsBar = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dSumUmsBar = dSumUmsBar
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumUmsBar
                    dSumGUTSCHBar = 0  'CW20190320
                End If
            End If
        End If
    End If 'CW20190626
        '*****************BAR Ende



        'scheck
If dSumScheck > 0 Then 'CW20190626
        If dSumScheck < dRestfalscherumsatz Then 'dann Scheck abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumScheck = dSumScheck
                dRestfalscherumsatz = 0
            Else
                If dSumScheck >= dZuVergebenerEchterumsatz Then
                    dSumScheck = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else

                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumScheck

                End If
'                dRestfalscherumsatz = dFalscherUmsatz
            End If
        Else
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumScheck = dSumScheck
            Else
                dSumScheck = dSumScheck - dRestfalscherumsatz + dZhlgGutsch
                dSumGUTSCHScheck = dRestfalscherumsatz 'CW20190320
            End If
             dRestfalscherumsatz = 0 'CW20190320
        End If
  End If 'CW20190626
  
        'Ec Last
        If dSumECLAST > 0 Then 'CW20190626
        If dSumECLAST < dRestfalscherumsatz Then 'dann ec last abfragen
            If dZhlgGutsch > dRestfalscherumsatz Then
                dSumECLAST = dSumECLAST
                dRestfalscherumsatz = 0
            Else
                If dSumECLAST >= dZuVergebenerEchterumsatz Then
                    dSumECLAST = dZuVergebenerEchterumsatz
                    dZuVergebenerEchterumsatz = 0
                Else
                    dZuVergebenerEchterumsatz = dZuVergebenerEchterumsatz - dSumECLAST
                End If
                dRestfalscherumsatz = 0
            End If
        Else
            If dZhlgGutsch >= dRestfalscherumsatz Then
                dSumECLAST = dSumECLAST
            Else
                dSumECLAST = dSumECLAST - dRestfalscherumsatz + dZhlgGutsch
                dSumGUTSCHECLAST = dRestfalscherumsatz 'CW20190320
            End If
             dRestfalscherumsatz = 0 'CW20190320
        End If
        End If 'CW20190626

    End If
    
    

    '**************************************************************
    '* Wieviel Umsatz hat der Kunde in bar gemacht?
    '**************************************************************
    
'    If bUmsatz = True Then
    If dEchterUmsatz > 0 Then
    
            '**************************************************************
        '* Wieviel Umsatz hat der Kunde mit Kreditkarten gemacht?
        '**************************************************************
        If Not IsNull(rsrs!UMS_KARTE) Then
            rsrs!UMS_KARTE = rsrs!UMS_KARTE + dSumKreditkarten
            schreibeProtokollUNITXT CStr(dSumKreditkarten), "Kartenzahlung"
        Else
            rsrs!UMS_KARTE = dSumKreditkarten
            schreibeProtokollUNITXT CStr(dSumKreditkarten), "Kartenzahlung"
        End If
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde über Dukaten gemacht?
        '**************************************************************
        If Not IsNull(rsrs!DUKA) Then
            rsrs!DUKA = rsrs!DUKA + dSumDukaten
        Else
            rsrs!DUKA = dSumDukaten
        End If
        
        
        'erzielter Umsatz über Bargeld
                  
        If Not IsNull(rsrs!UMS_BAR) Then
            rsrs!UMS_BAR = rsrs!UMS_BAR + dSumUmsBar
        Else
            rsrs!UMS_BAR = dSumUmsBar
        End If
    
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde mit Schecks gemacht?
        '**************************************************************
        If Not IsNull(rsrs!UMS_SCHECK) Then
            rsrs!UMS_SCHECK = rsrs!UMS_SCHECK + dSumScheck
        Else
            rsrs!UMS_SCHECK = dSumScheck
        End If
    
        '**************************************************************
        '* Wieviel Umsatz hat der Kunde mit EC-Lastschrift gemacht?
        '**************************************************************
        If Not IsNull(rsrs!UMS_LAST) Then
            rsrs!UMS_LAST = rsrs!UMS_LAST + dSumECLAST
        Else
            rsrs!UMS_LAST = dSumECLAST
        End If
        
    End If

    '**************************************************************
    '* Datum und Kassennummer des Verbuchens schreiben
    '**************************************************************
    rsrs!ADATE = lDatum
    rsrs!kasnum = Val(gcKasNum)

    '**************************************************************
    '* Betrag der eingereichten Gutscheine verbuchen
    '**************************************************************
    If Not IsNull(rsrs!EINRGUTSCH) Then
        rsrs!EINRGUTSCH = rsrs!EINRGUTSCH + dZhlgGutsch
    Else
        rsrs!EINRGUTSCH = dZhlgGutsch
    End If

    '**************************************************************
    '* Betrag der generierten Rest-Gutscheine verbuchen
    '**************************************************************
    If Not IsNull(rsrs!RESTGUTSCH) Then
        rsrs!RESTGUTSCH = rsrs!RESTGUTSCH + dSumRückGUTSCH
    Else
        rsrs!RESTGUTSCH = dSumRückGUTSCH
    End If

    'Achtung
    '**************************************************************
    '* Betrag der Gutschein-AUszahlung verbuchen
    '**************************************************************
    If Not IsNull(rsrs!AUSZGUTSCH) Then
        rsrs!AUSZGUTSCH = rsrs!AUSZGUTSCH + Format(dGutscheinauszahlung, "######0.00")
    Else
        rsrs!AUSZGUTSCH = Format(dGutscheinauszahlung, "######0.00")
    End If


    'Achtung
    '**************************************************************
    '* Betrag des Umsatzes durch Gutschein-Einreichungen verbuchen
    '**************************************************************
    
    
    'Test echter Umsatz
    

''''    If dwertGutverkauf > 0 Then
''''
''''        If dZhlgGutsch > dwertGutverkauf Then
''''
''''            If Not IsNull(rsrs!ZHLGGUTSCH) Then
''''                rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
''''            Else
''''                rsrs!ZHLGGUTSCH = dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
''''
''''                If rsrs!ZHLGGUTSCH < 0 Then rsrs!ZHLGGUTSCH = 0
''''            End If
''''
''''        Else
''''
''''            If Not IsNull(rsrs!ZHLGGUTSCH) Then
''''                rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH
''''            Else
''''                rsrs!ZHLGGUTSCH = dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH
''''            End If
''''
''''        End If
''''
''''
''''    Else
''''        If Not IsNull(rsrs!ZHLGGUTSCH) Then
''''            rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
''''        Else
''''            rsrs!ZHLGGUTSCH = dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH - dwertGutverkauf
''''        End If
''''
''''    End If
''''
    
    
    'neu 'CW20190627 ' Echter Umsatz aus Gutsch
        
        If dZhlgGutsch > 0 Then

            If Not IsNull(rsrs!ZHLGGUTSCH) Then
                rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH
            Else
                rsrs!ZHLGGUTSCH = dUmsatzGutsch - dGutscheinauszahlung - dSumRückGUTSCH
                
                If rsrs!ZHLGGUTSCH < 0 Then rsrs!ZHLGGUTSCH = 0
            End If
        End If

    
    'ende neu
    
    
    
    
    
    
    

    'Achtung
    '**************************************************************
    '* Betrag des Gutschein-Verkäufe verbuchen (Bar?)
    '**************************************************************
    If Not IsNull(rsrs!GUTSCHEIN) Then
        rsrs!GUTSCHEIN = rsrs!GUTSCHEIN + dWertGutschein
    Else
        rsrs!GUTSCHEIN = dWertGutschein
    End If

    '**************************************************************
    '* Betrag der Scheck - Verkäufe verbuchen
    '**************************************************************

    If Not IsNull(rsrs!SCHVERKAUF) Then
        rsrs!SCHVERKAUF = rsrs!SCHVERKAUF + dGegScheck
    Else
        rsrs!SCHVERKAUF = dGegScheck
    End If

    '**************************************************************
    '* Sonderpreise verbuchen
    '**************************************************************
    If Not IsNull(rsrs!SPREIS_ANZ) Then
        rsrs!SPREIS_ANZ = rsrs!SPREIS_ANZ + dSPreisAnz
    Else
        rsrs!SPREIS_ANZ = dSPreisAnz
    End If

    If Not IsNull(rsrs!SPREIS_GES) Then
        rsrs!SPREIS_GES = rsrs!SPREIS_GES + dSPreisGes
    Else
        rsrs!SPREIS_GES = dSPreisGes
    End If



    '**************************************************************
    '* Kundenzahl schreiben
    '**************************************************************
    If Not IsNull(rsrs!Kundenzahl) Then
        rsrs!Kundenzahl = rsrs!Kundenzahl + dKundenZahl
    Else
        rsrs!Kundenzahl = dKundenZahl
    End If
    '**************************************************************
    '* Artikelrabatte schreiben
    '**************************************************************
    If Not IsNull(rsrs!ARTRAB_ANZ) Then
        rsrs!ARTRAB_ANZ = rsrs!ARTRAB_ANZ + dArtRabAnz
    Else
        rsrs!ARTRAB_ANZ = dArtRabAnz
    End If

    If Not IsNull(rsrs!ARTRAB_GES) Then
        rsrs!ARTRAB_GES = rsrs!ARTRAB_GES + dArtRabGes
    Else
        rsrs!ARTRAB_GES = dArtRabGes
    End If
    '**************************************************************
    '* Gesamtrabatte schreiben
    '**************************************************************
    If Not IsNull(rsrs!GESRAB_ANZ) Then
        rsrs!GESRAB_ANZ = rsrs!GESRAB_ANZ + dGesRabAnz
    Else
        rsrs!GESRAB_ANZ = dGesRabAnz
    End If

    If Not IsNull(rsrs!GESRAB_GES) Then
        rsrs!GESRAB_GES = rsrs!GESRAB_GES + dGesRabGes
    Else
        rsrs!GESRAB_GES = dGesRabGes
    End If
    '**************************************************************
    '* Bonnummer schreiben
    '**************************************************************
    
    If Not IsNull(rsrs!BELEGNR) Then
        If gdBonNr < CLng(rsrs!BELEGNR) Then
            
        Else
            rsrs!BELEGNR = gdBonNr
        End If
    Else
        rsrs!BELEGNR = gdBonNr
    End If
    
    

    '**************************************************************
    '* Stornos schreiben
    '**************************************************************
    If Not IsNull(rsrs!STORNO_GES) Then
        rsrs!STORNO_GES = rsrs!STORNO_GES + dStornoWert
    Else
        rsrs!STORNO_GES = dStornoWert
    End If

    If Not IsNull(rsrs!STORNO_ANZ) Then
        rsrs!STORNO_ANZ = rsrs!STORNO_ANZ + lStornoAnz
    Else
        rsrs!STORNO_ANZ = lStornoAnz
    End If

    '******************************************
    '* Wie sind neue Gutscheine bezahlt worden
    '******************************************

    '******************************************
    '* Betrag Gutscheinverkäufe in Bar setzen
    '******************************************
    Dim dRestZahlungsmittelforGutsch As Double
    dRestZahlungsmittelforGutsch = 0

    If dWertGutschein > 0 Then ' Reihenfolge sehr wichtig, siehe wie oben bei Nichtumsatz
           'Reihenfolge: Kreditkarte, Dukate, Bar, Scheck, ECLast, Gutsch
    
        dRestZahlungsmittelforGutsch = dWertGutschein
        
        If dSumGUTSCHKreditkarten < dRestZahlungsmittelforGutsch Then 'dann KK abfragen
            dGutschKreditkartenAnteil = dSumGUTSCHKreditkarten
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHKreditkarten
        Else
            dGutschKreditkartenAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If
        
        If dSumGutschDukate < dRestZahlungsmittelforGutsch Then 'dann dukate abfragen
            dGutschDukateAnteil = dSumGutschDukate
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGutschDukate
        Else
            dGutschDukateAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If
        
        If dSumGUTSCHBar < dRestZahlungsmittelforGutsch Then  'erst bar abfragen
            dGutschBarAnteil = dSumGUTSCHBar
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHBar
        Else
            dGutschBarAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If
        
        If dSumGUTSCHScheck < dRestZahlungsmittelforGutsch Then 'dann Scheck abfragen
            dGutschScheckAnteil = dSumGUTSCHScheck
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHScheck
        Else
            dGutschScheckAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If
        
        If dSumGUTSCHECLAST < dRestZahlungsmittelforGutsch Then 'dann last abfragen
            dGutschECLASTAnteil = dSumGUTSCHECLAST
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHECLAST
        Else
            dGutschECLASTAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If
        
        If dZhlgGutsch < dRestZahlungsmittelforGutsch Then 'dann gutschein aus gutschein abfragen
            dGutschGutschAnteil = dZhlgGutsch
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dGutschGutschAnteil
        Else
            dGutschGutschAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If
    

        
        If dSumGUTSCHKredit < dRestZahlungsmittelforGutsch Then 'dann Kredit abfragen
            dGutschKreditAnteil = dSumGUTSCHKredit
            dRestZahlungsmittelforGutsch = dRestZahlungsmittelforGutsch - dSumGUTSCHKredit
        Else
            dGutschKreditAnteil = dRestZahlungsmittelforGutsch
            dRestZahlungsmittelforGutsch = 0
        End If

    End If

    If dGutschBarAnteil > 0 Then
        '******************************************
        '* Wert des Barverkaufes setzen
        '******************************************
        If Not IsNull(rsrs!BARVERKAUF) Then
            rsrs!BARVERKAUF = rsrs!BARVERKAUF + (dSumBar - dSumRückBar + dGutscheinauszahlung - dGutschBarAnteil)
        Else
            rsrs!BARVERKAUF = (dSumBar - dSumRückBar + dGutscheinauszahlung - dGutschBarAnteil)
        End If

    Else
        '******************************************
        '* Wert des Barverkaufes setzen
        '******************************************
        If Not IsNull(rsrs!BARVERKAUF) Then
            rsrs!BARVERKAUF = rsrs!BARVERKAUF + dSumBar - dSumRückBar + dGutscheinauszahlung
        Else
            rsrs!BARVERKAUF = dSumBar - dSumRückBar + dGutscheinauszahlung
        End If
    End If
    
    
    '**************************************************************
    '* Gutscheine gekauft mit Gutschein bezahlt!
    '**************************************************************
    If dGutschGutschAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHGUTSCH) Then
            rsrs!GUTSCHGUTSCH = rsrs!GUTSCHGUTSCH + dGutschGutschAnteil + dGutschDukateAnteil
        Else
            rsrs!GUTSCHGUTSCH = dGutschGutschAnteil + dGutschDukateAnteil
        End If
    End If
    
    If dGutschBarAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHBAR) Then
            rsrs!GUTSCHBAR = rsrs!GUTSCHBAR + Format(dGutschBarAnteil, "######0.00") 'Hier ist Format() neu 21.10.13
        Else
            rsrs!GUTSCHBAR = Format(dGutschBarAnteil, "######0.00") 'Hier ist Format() neu 21.10.13
        End If
    End If

    If dGutschKreditkartenAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHKAR) Then
            rsrs!GUTSCHKAR = rsrs!GUTSCHKAR + dGutschKreditkartenAnteil
            schreibeProtokollUNITXT CStr(dGutschKreditkartenAnteil), "Kartenzahlung"
        Else
            rsrs!GUTSCHKAR = dGutschKreditkartenAnteil
            schreibeProtokollUNITXT CStr(dGutschKreditkartenAnteil), "Kartenzahlung"
        End If
    End If

    If dGutschECLASTAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHLAST) Then
            rsrs!GUTSCHLAST = rsrs!GUTSCHLAST + dGutschECLASTAnteil
        Else
            rsrs!GUTSCHLAST = dGutschECLASTAnteil
        End If
    End If

    If dGutschScheckAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHSCH) Then
            rsrs!GUTSCHSCH = rsrs!GUTSCHSCH + dGutschScheckAnteil
        Else
            rsrs!GUTSCHSCH = dGutschScheckAnteil
        End If
    End If

' Kredite fehlte 'CW20190627
    If dGutschKreditAnteil > 0 Then
        If Not IsNull(rsrs!GUTSCHKRE) Then
            rsrs!GUTSCHKRE = rsrs!GUTSCHKRE + dGutschKreditAnteil
        Else
            rsrs!GUTSCHKRE = dGutschKreditAnteil
        End If
    End If


    '**************************************************************
    '* Schreibvorgang durchführen
    '**************************************************************
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "UpdateAFCStat68Test6"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub AddOn_AFCSTAT_GutschModus()
    On Error GoTo LOKAL_ERROR
    
    Dim cErzielterPreis As String
    Dim cArtNr As String
    Dim dNichtUmsatz As Double
    dNichtUmsatz = 0
    
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim cLBSatz As String
    Dim cUmsOK As String
    
    lAnzSatz = frmWKL20!List1.ListCount
    For lAktSatz = 0 To lAnzSatz - 1
         cLBSatz = frmWKL20!List1.list(lAktSatz)
        
        If Len(cLBSatz) > 156 Then
            cUmsOK = Mid(cLBSatz, 156, 1)
        Else
            cUmsOK = "J"
        End If

        cArtNr = Mid(cLBSatz, 7, 6)
        
        cErzielterPreis = Mid(cLBSatz, 60, 9)
        cErzielterPreis = Trim$(cErzielterPreis)
        cErzielterPreis = fnMoveComma2Point$(cErzielterPreis)
        
        If gbGutscheinBeiVKversteuern = True Then
            If cUmsOK <> "N" Then
    
            Else
                dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
            End If
        Else
            If cArtNr <> "666666" Then
                If cUmsOK <> "N" Then
                Else
                    dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
                End If
            End If
        End If
    Next lAktSatz
    
    Dim dUmsatz                 As Double
    dUmsatz = 0
    
    Dim dSummegesamt            As Double
    
    dSummegesamt = CDbl(Label333(0).Caption)
    dUmsatz = dSummegesamt - dNichtUmsatz
    
    Dim cGutschnr               As String
    Dim i                       As Integer
    Dim dGegebenGutschein       As Double
    Dim dateStichtag            As Date
    dateStichtag = ermStichtag
   
    Dim dateGutschAusgabeTAG    As Date
    
    For i = 0 To 19
        If Gutschl(i).gutschnr <> 0 Then
            cGutschnr = Gutschl(i).gutschnr
           
            'Ausgabedatum prüfen
            dateGutschAusgabeTAG = ermGutscheinAusgabeTag(cGutschnr)
            
            If dateGutschAusgabeTAG >= dateStichtag Then
                'werte summieren
                dGegebenGutschein = dGegebenGutschein + Gutschl(i).gutschwert
                
            End If
        End If
    Next i
    
    'gegeben Wert Gutschein
    
    Dim lDate As Long
    Dim sTime As String
    
    lDate = DateValue(Now)
    sTime = TimeValue(Now)
    
    If dGegebenGutschein > dUmsatz Then dGegebenGutschein = dUmsatz
    
    If dGegebenGutschein > 0 Then
        insert_Gemischte_Zahlung lDate, sTime, gdBonNr, gcKasNum, "nicht ums GUTSCHBETRAG", dGegebenGutschein
        updateafcstat "ZHLGGUTSCH", (-1 * dGegebenGutschein), gcKasNum
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "AddOn_AFCSTAT_GutschModus"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Function ermNichtUmsatz() As Double
    On Error GoTo LOKAL_ERROR
    
    
    
    ermNichtUmsatz = 0
    
    Dim cErzielterPreis As String
    Dim cArtNr As String
    Dim dNichtUmsatz As Double
    dNichtUmsatz = 0
    
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim cLBSatz As String
    Dim cUmsOK As String
    
    lAnzSatz = frmWKL20!List1.ListCount
    For lAktSatz = 0 To lAnzSatz - 1
         cLBSatz = frmWKL20!List1.list(lAktSatz)
        
        If Len(cLBSatz) > 156 Then
            cUmsOK = Mid(cLBSatz, 156, 1)
        Else
            cUmsOK = "J"
        End If

        cArtNr = Mid(cLBSatz, 7, 6)
        
        cErzielterPreis = Mid(cLBSatz, 60, 9)
        cErzielterPreis = Trim$(cErzielterPreis)
        cErzielterPreis = fnMoveComma2Point$(cErzielterPreis)
        
        If gbGutscheinBeiVKversteuern = True Then
            If cUmsOK <> "N" Then
    
            Else
                dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
            End If
        Else
        
            If cArtNr <> "666666" Then
                If cUmsOK <> "N" Then
                Else
                    dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
                End If
            End If
        End If
    Next lAktSatz
    
    Dim dUmsatz                 As Double
    dUmsatz = 0
    
    Dim dSummegesamt            As Double
    
    dSummegesamt = CDbl(Label333(0).Caption)
    dUmsatz = dSummegesamt - dNichtUmsatz
    
    Dim cGutschnr               As String
    Dim i                       As Integer
    Dim dGegebenGutschein       As Double
    Dim dateStichtag            As Date
    dateStichtag = ermStichtag
   
    Dim dateGutschAusgabeTAG    As Date
    
    For i = 0 To 19
        If Gutschl(i).gutschnr <> 0 Then
            cGutschnr = Gutschl(i).gutschnr
           
            'Ausgabedatum prüfen
            dateGutschAusgabeTAG = ermGutscheinAusgabeTag(cGutschnr)
            
            If dateGutschAusgabeTAG >= dateStichtag Then
                'werte summieren
                dGegebenGutschein = dGegebenGutschein + Gutschl(i).gutschwert
                
            End If
        End If
    Next i
    
    'gegeben Wert Gutschein
    
    Dim lDate As Long
    Dim sTime As String
    
    lDate = DateValue(Now)
    sTime = TimeValue(Now)
    
    If dGegebenGutschein > dUmsatz Then dGegebenGutschein = dUmsatz
    
    If dGegebenGutschein > 0 Then
    
        ermNichtUmsatz = dGegebenGutschein
'''        insert_Gemischte_Zahlung lDate, sTime, gdBonNr, gcKasNum, "nicht ums GUTSCHBETRAG", dGegebenGutschein
'''        updateafcstat "ZHLGGUTSCH", (-1 * dGegebenGutschein), gcKasNum
    End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermNichtUmsatz"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function




Private Sub ReInitDialog20WK68()
    On Error GoTo LOKAL_ERROR
    
    gbNumTaste = True

    frmWKL20!List1.Clear
    frmWKL20!List3.Nodes.Clear
    frmWKL20.Label41(1).Caption = 0

    frmWKL20!Label2(6).Caption = "0,00"
    
    LeereDialogModul20

    frmWKL20!Text1(0).Text = gcBedienerNr
    gcKreditKarte = ""
    gcZahlMittel = ""

    If gbBEDLEER = True Then
        frmWKL20!Label1(8).Caption = ""
        frmWKL20!Text1(0).Text = ""
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ReInitDialog20WK68"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub DruckeKassenBonWKL68()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz        As Long
    Dim lAktSatz        As Long
    Dim cLBSatz         As String
    Dim ctmp            As String
    
    Dim dWert           As Double
    Dim dFaktor         As Double
    Dim dGPreis         As Double
    Dim dVPreis         As Double
    Dim dMWSt           As Double
    Dim dMWStVoll       As Double
    Dim dMWStErm        As Double
    Dim bDebug          As Boolean
    
    bDebug = False
    
    dMWStVoll = gdMWStV
    dMWStErm = gdMWStE
    
    dFaktor = 1
    If frmWKL20.Label2(3).Visible Then
        ctmp = frmWKL20.Label2(3).Caption
        ctmp = fnMoveComma2Point$(ctmp)
        dWert = Val(ctmp)
        dFaktor = (100 - dWert) / 100
    End If
    
    lAnzSatz = frmWKL20.List1.ListCount
    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = frmWKL20.List1.list(lAktSatz)
        If bDebug Then
            MsgBox Len(cLBSatz) & vbCrLf & cLBSatz
            MsgBox "Menge = " & Mid(cLBSatz, 1, 5)
            MsgBox "ArtNr = " & Mid(cLBSatz, 7, 6)
            MsgBox "Bezeich = " & Mid(cLBSatz, 14, 35)
            MsgBox "EPreis nach Rabatt = " & Mid(cLBSatz, 50, 9)
            MsgBox "GPreis = " & Mid(cLBSatz, 60, 9)
            MsgBox "MWST-Kz = " & Mid(cLBSatz, 72, 1)
            MsgBox "Listenpreis/Sonderpreis = " & Mid(cLBSatz, 74, 9)
            MsgBox "Betrag ArtRabatt = " & Mid(cLBSatz, 84, 9)
            MsgBox "erzielter VK-Preis = " & Mid(cLBSatz, 94, 9)
            MsgBox "Betrag volle MWST = " & Mid(cLBSatz, 104, 9)
            MsgBox "Betrag erm. MWST = " & Mid(cLBSatz, 114, 9)
            MsgBox "ArtRabatt % = " & Mid(cLBSatz, 124, 3)
            MsgBox "Listen-VK = " & Mid(cLBSatz, 128, 9)
            MsgBox "Restmenge = " & Mid(cLBSatz, 138, 9)
        End If
                
        ctmp = Mid(cLBSatz, 60, 9)
        ctmp = fnMoveComma2Point$(ctmp)
        dWert = Val(ctmp)
                
        dVPreis = dWert * dFaktor
        ctmp = Format$(dVPreis, "#####0.00")
        ctmp = Space$(9 - Len(ctmp)) & ctmp
        Mid(cLBSatz, 94, 9) = ctmp
        
        ctmp = Mid(cLBSatz, 72, 1)
        If ctmp = "V" Then
            dMWSt = dVPreis * (dMWStVoll / 100)
            ctmp = Format$(dMWSt, "#####0.00")
            ctmp = Space$(9 - Len(ctmp)) & ctmp
            Mid(cLBSatz, 104, 9) = ctmp
        ElseIf ctmp = "E" Then
            dMWSt = dGPreis * (dMWStErm / 100)
            ctmp = Format$(dMWSt, "#####0.00")
            ctmp = Space$(9 - Len(ctmp)) & ctmp
            Mid(cLBSatz, 114, 9) = ctmp
        Else
            dMWSt = 0
        End If
        
        frmWKL20.List1.RemoveItem lAktSatz
        frmWKL20.List3.Nodes.Remove lAktSatz + 1
        frmWKL20.List1.AddItem cLBSatz, lAktSatz
        frmWKL20.List3.Nodes.Add Text:=Left(cLBSatz, 68)
        
        If bDebug Then
            cLBSatz = frmWKL20.List1.list(lAktSatz)
            MsgBox Len(cLBSatz) & vbCrLf & cLBSatz
            MsgBox Mid(cLBSatz, 1, 5)
            MsgBox Mid(cLBSatz, 7, 6)
            MsgBox Mid(cLBSatz, 14, 35)
            MsgBox Mid(cLBSatz, 50, 9)
            MsgBox Mid(cLBSatz, 60, 9)
            MsgBox Mid(cLBSatz, 72, 1)
            MsgBox Mid(cLBSatz, 74, 9)
            MsgBox Mid(cLBSatz, 84, 9)
            MsgBox Mid(cLBSatz, 94, 9)
            MsgBox Mid(cLBSatz, 104, 9)
            MsgBox Mid(cLBSatz, 114, 9)
            MsgBox Mid(cLBSatz, 124, 3)
            MsgBox Mid(cLBSatz, 128, 9)
            MsgBox "Restmenge = " & Mid(cLBSatz, 138, 9)
        End If
    Next lAktSatz
    
    Dim dNichtUmsatz As Double
    dNichtUmsatz = 0
    
    'bei GutschModus
    
    If gbGutscheinBeiVKversteuern = True Then
        
        dNichtUmsatz = ermNichtUmsatz
    
        SendeDaten2DruckerNeuWKL68_Bonus dNichtUmsatz
    Else
        SendeDaten2DruckerNeuWKL68_Bonus
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeKassenBonWKL68"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SendeDaten2DruckerECCASH()
    On Error GoTo LOKAL_ERROR
     
    Dim lcount                  As Long
    Dim lAnzZeile               As Long
    Dim cDaten                  As String
    Dim aDeviceName             As String
    Dim cEscapeSequenz          As String
    ReDim cDruckZeile(1 To 1) As String
    Dim iLenZeile               As Integer
    Dim sTempBeleg              As String
    Dim lPos                    As Long
    
    iLenZeile = 32

    setzedrucker gcBonDrucker
    
    'Drucker an, Display aus, Init Drucker
    aDeviceName = Printer.DeviceName
    cEscapeSequenz = gcInit
    OpenDrawer aDeviceName, cEscapeSequenz
    
    '***********************************************
    'Zeile Karteninfo
    '***********************************************
    
    If gsAdtBeleg <> "" Then
        sTempBeleg = gsAdtBeleg
        sTempBeleg = SwapStr(sTempBeleg, Space(35), "")
        lPos = InStr(sTempBeleg, "**")
        
        sTempBeleg = Left(sTempBeleg, lPos + 30)

        Do While Len(sTempBeleg) > 0
            cDaten = Left(sTempBeleg, 32)
            
            If Len(sTempBeleg) >= 33 Then
                sTempBeleg = Right(sTempBeleg, Len(sTempBeleg) - 33)
            Else
                sTempBeleg = ""
            End If
            
            If cDaten <> "" Then
                cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
        Loop
    End If
    
    '***********************************************
    'ein paar Leerzeilen drucken
    '***********************************************
    For lcount = 1 To gbLeereZeil
        If lcount = gbLeereZeil Then
            cEscapeSequenz = vbCrLf
        Else
            cEscapeSequenz = " " & vbCrLf
        End If
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    Next lcount
    
BON_DRUCKEN:
    If gbAPI = True Then
        OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
    Else
        OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
    End If

BON_SCHNEIDEN:
    'Kassenbon abschneiden
    If gbAPI = True Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcSchneiden
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
    gsAdtBeleg = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SendeDaten2DruckerECCASH"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SendeDaten2DruckerECCASH_Kundenbeleg()
    On Error GoTo LOKAL_ERROR

    Dim lcount                  As Long
    Dim lAnzZeile               As Long
    Dim cDaten                  As String
    Dim aDeviceName             As String
    Dim cEscapeSequenz          As String
    ReDim cDruckZeile(1 To 1) As String
    Dim iLenZeile               As Integer
    Dim sTempBeleg              As String
    Dim lPos                    As Long
    Dim i                       As Integer

    iLenZeile = 32

    setzedrucker gcBonDrucker

    'Drucker an, Display aus, Init Drucker
    aDeviceName = Printer.DeviceName
    cEscapeSequenz = gcInit
    OpenDrawer aDeviceName, cEscapeSequenz

    '***********************************************
    'Zeile Karteninfo
    '***********************************************
'    MsgBox gsAdtBeleg

    If gsAdtBeleg <> "" Then
        If InStr(gsAdtBeleg, "Unterschrift umseitig") > 0 Then
    
            sTempBeleg = gsAdtBeleg
            sTempBeleg = SwapStr(sTempBeleg, Space(35), "")
            lPos = InStr(sTempBeleg, "Unterschrift umseitig")
    
            sTempBeleg = Mid(sTempBeleg, 166, lPos - 200)
            
'            MsgBox sTempBeleg
            
            cDaten = "Kundenbeleg"
            If cDaten <> "" Then
                cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
    
            Do While Len(sTempBeleg) > 0
                cDaten = Left(sTempBeleg, 32)
                
                If Left(cDaten, 3) = "BLZ" Then
                    cDaten = Left(cDaten, 21) & Space(1) & "xxxxxxx" & Right(cDaten, 3)
                End If
                
                If Len(sTempBeleg) >= 33 Then
                    sTempBeleg = Right(sTempBeleg, Len(sTempBeleg) - 33)
                Else
                    sTempBeleg = ""
                End If
    
                If cDaten <> "" Then
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
            Loop
            
            '***********************************************
            'ein paar Leerzeilen drucken
            '***********************************************
            For lcount = 1 To gbLeereZeil
                If lcount = gbLeereZeil Then
                    cEscapeSequenz = vbCrLf
                Else
                    cEscapeSequenz = " " & vbCrLf
                End If
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            Next lcount
            
        ElseIf InStr(gsAdtBeleg, "Autorisierungsnummer") > 0 Or InStr(gsAdtBeleg, "EMV-Daten") > 0 Then

    
            sTempBeleg = gsAdtBeleg
            sTempBeleg = SwapStr(sTempBeleg, Space(35), "")
            
            If InStr(gsAdtBeleg, "Autorisierungsnummer") > 0 Then
                lPos = InStr(sTempBeleg, "Autorisierungsnummer")
            ElseIf InStr(gsAdtBeleg, "EMV-Daten") > 0 Then
                lPos = InStr(sTempBeleg, "EMV-Daten")
            End If
           
            sTempBeleg = Mid(sTempBeleg, 166, lPos + 30 - 164)
            
            cDaten = "Kundenbeleg"
            If cDaten <> "" Then
                cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
    
            Do While Len(sTempBeleg) > 0
                cDaten = Left(sTempBeleg, 32)
                
                If Left(cDaten, 9) = "Kartennr." Then
                    cDaten = "Kartennr." & Space(4) & "xxxxxxxxxxxxxxx" & Right(cDaten, 4)
                End If
                
                If Len(sTempBeleg) >= 33 Then
                    sTempBeleg = Right(sTempBeleg, Len(sTempBeleg) - 33)
                Else
                    sTempBeleg = ""
                End If
    
                If cDaten <> "" Then
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
            Loop
            
            '***********************************************
            'ein paar Leerzeilen drucken
            '***********************************************
            For lcount = 1 To gbLeereZeil
                If lcount = gbLeereZeil Then
                    cEscapeSequenz = vbCrLf
                Else
                    cEscapeSequenz = " " & vbCrLf
                End If
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            Next lcount
        End If
    End If

    

BON_DRUCKEN:
    If gbAPI = True Then
        OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
    Else
        OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
    End If

BON_SCHNEIDEN:
    'Kassenbon abschneiden
    If gbAPI = True Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcSchneiden
        OpenDrawer aDeviceName, cEscapeSequenz
    End If

''''    gsAdtBeleg = ""

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SendeDaten2DruckerECCASH_Kundenbeleg"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Private Sub SendeDaten2DruckerECCASH_Haendlerbeleg()
    On Error GoTo LOKAL_ERROR

    Dim lcount                  As Long
    Dim lAnzZeile               As Long
    Dim cDaten                  As String
    Dim aDeviceName             As String
    Dim cEscapeSequenz          As String
    ReDim cDruckZeile(1 To 1) As String
    Dim iLenZeile               As Integer
    Dim sTempBeleg              As String
    Dim lPos                    As Long
    Dim i                       As Integer

    iLenZeile = 32

''    setzedrucker gcBonDrucker

    'Drucker an, Display aus, Init Drucker
    aDeviceName = Printer.DeviceName
    cEscapeSequenz = gcInit
    OpenDrawer aDeviceName, cEscapeSequenz

    '***********************************************
    'Zeile Karteninfo
    '***********************************************
    
'    MsgBox gsAdtBeleg

    If gsAdtBeleg <> "" Then
    
        If InStr(gsAdtBeleg, "Unterschrift umseitig") > 0 Then
        
            sTempBeleg = gsAdtBeleg
            sTempBeleg = SwapStr(sTempBeleg, Space(35), "")
            lPos = InStr(sTempBeleg, "**")
    
            sTempBeleg = Left(sTempBeleg, lPos + 30)
    
            Do While Len(sTempBeleg) > 0
                cDaten = Left(sTempBeleg, 32)
    
                If Len(sTempBeleg) >= 33 Then
                    sTempBeleg = Right(sTempBeleg, Len(sTempBeleg) - 33)
                Else
                    sTempBeleg = ""
                End If
    
                If cDaten <> "" Then
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
            Loop
            
            cDaten = ""
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        
            cDaten = "Ermächtigung Lastschrifteinzug:"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz

            cDaten = ""
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            cDaten = "Ich ermächtige hiermit das"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz

            cDaten = "oben genannte Unternehmen,"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz

            cDaten = "den als Endsumme ausgewiesenen"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz

            cDaten = "Betrag von meinem durch Bank- "
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz

            cDaten = "leitzahl und Kontonummer be-"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz

            cDaten = "zeichneten Konto durch Last-"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            cDaten = "schrift einzuziehen."
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz

            cDaten = ""
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz

            cDaten = "Ermächtigung Adressweitergabe:"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz

            cDaten = ""
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz

            cDaten = "Ich weise mein Kreditinstitut,"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz

            cDaten = "das durch die Bankleitzahl "
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz

            cDaten = "bezeichnet ist, unwiderruflich"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            cDaten = "an, bei Nichteinlösung der Last-"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            cDaten = "schrift oder bei Widerspruch"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            cDaten = "gegen die Lastschrift dem"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            cDaten = "Unternehmen oder einem von"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            cDaten = "ihm beauftragten Dritten auf "
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            cDaten = "Anforderung meinen Namen und"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz

            cDaten = "meine Adresse mitzuteilen, damit"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            cDaten = "das Unternehmen seinen Anspruch"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            cDaten = "gegen mich geltend machen kann."
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            
    
            '+Unterschrift
            
            cDaten = ""
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            
            cDaten = ""
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            cDaten = ""
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            cDaten = "________________________________"
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            
            
            
            
            
            
            
            
            cDaten = "Unterschrift:"
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            
            KonvertAnsiAscii cDaten
            
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            
            
'            cDaten = "Kundenbeleg"
'            If cDaten <> "" Then
'                cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
'                KonvertAnsiAscii cDaten
'                cEscapeSequenz = cDaten & vbCrLf
'                lAnzZeile = lAnzZeile + 1
'                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'                cDruckZeile(lAnzZeile) = cEscapeSequenz
'            End If
        Else
            sTempBeleg = gsAdtBeleg
            sTempBeleg = SwapStr(sTempBeleg, Space(35), "")
            lPos = InStr(sTempBeleg, "**")
    
            sTempBeleg = Left(sTempBeleg, lPos + 30)
    
            Do While Len(sTempBeleg) > 0
                cDaten = Left(sTempBeleg, 32)
    
                If Len(sTempBeleg) >= 33 Then
                    sTempBeleg = Right(sTempBeleg, Len(sTempBeleg) - 33)
                Else
                    sTempBeleg = ""
                End If
    
                If cDaten <> "" Then
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
            Loop
        End If
    End If

    '***********************************************
    'ein paar Leerzeilen drucken
    '***********************************************
    For lcount = 1 To 6
        If lcount = 6 Then
            cEscapeSequenz = vbCrLf
        Else
            cEscapeSequenz = " " & vbCrLf
        End If
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    Next lcount

BON_DRUCKEN:
    If gbAPI = True Then
        OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
    Else
        OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
    End If

BON_SCHNEIDEN:
    'Kassenbon abschneiden
    If gbAPI = True Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcSchneiden
        OpenDrawer aDeviceName, cEscapeSequenz
    End If

    gsAdtBeleg = ""

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SendeDaten2DruckerECCASH_Haendlerbeleg"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub SendeDaten2DruckerNeuWKL68_Bonus(Optional dNichtUmsatz As Double = 0)
    On Error GoTo LOKAL_ERROR
        
    Dim lartnr                  As Long
    Dim lAnzSatz                As Long
    Dim lAktSatz                As Long
    Dim lcount                  As Long
    Dim lAnzZeile               As Long
    Dim lAnzLbSatz              As Long
    Dim lRet                    As Long

    Dim cLBSatz                 As String
    Dim cFeld                   As String
    Dim cDaten                  As String
    Dim ctmp                    As String
    Dim cMWST                   As String
    Dim cMWSTzzgl               As String
    Dim aDeviceName             As String
    Dim cEscapeSequenz          As String
    Dim cArtNr                  As String

    Dim iLevel                  As Integer
    Dim iAktCopy                As Integer
    Dim iStufe                  As Integer
    Dim iLenZeile               As Integer
    Dim iGesAnzahl              As Integer
    Dim i                       As Integer

    Dim dSumme                  As Double
    Dim dWert                   As Double
    Dim dMWStVoll               As Double
    Dim dMWStErm                As Double
    Dim dMWSTVzzgl              As Double
    Dim dMWSTEzzgl              As Double
    Dim dAktZeit                As Double
    Dim dNeuZeit                As Double
    Dim dMWSt                   As Double

    Dim gegebenInsgesamt        As Double
    Dim gegebenBAR              As Double
    Dim gegebenDUKATE           As Double
    Dim gegebenECLAST           As Double
    Dim gegebenKK1              As Double
    Dim cKK2art                 As String
    Dim cKK1art                 As String
    Dim gegebenKK2              As Double
    Dim gegebenSCHECK           As Double
    Dim gegebenGUTSCHEIN        As Double
    Dim dZurueckBAR             As Double
    Dim dZurueckGUTSCH          As Double
    Dim dEndbetrag              As Double
    Dim gutschwert              As Double
    
    Dim bBonZwang               As Boolean
    Dim gbStorni                As Boolean
    Dim dSparSatzsum            As Double
    
    dSparSatzsum = 0
    gbStorni = False
    bBonZwang = False
    iLevel = 0
    
    gbADTBON = False
    
    If Label333(3).Caption <> "" Then
        dZurueckBAR = CDbl(Label333(3).Caption)
    Else
        dZurueckBAR = 0
    End If
    
    If Label333(0).Caption <> "" Then
        dEndbetrag = CDbl(Label333(0).Caption)
    Else
        dEndbetrag = 0
    End If
    
    If Label1(28).Caption <> "" Then
        dZurueckGUTSCH = CDbl(Label1(28).Caption)
    Else
        dZurueckGUTSCH = 0
    End If
    
    If Label333(1).Caption <> "" Then
        gegebenInsgesamt = CDbl(Label333(1).Caption)
    Else
        gegebenInsgesamt = 0
    End If
    
    gegebenBAR = 0
    If Text1(0).Text <> "" Then
        If IsNumeric(Text1(0).Text) Then
            gegebenBAR = CDbl(Text1(0).Text)
        End If
    End If
    
    If Text1(3).Text <> "" Then 'And Val(Text1(3).Text) > 0
        gegebenDUKATE = CDbl(Text1(3).Text)
    Else
        gegebenDUKATE = 0
    End If
    
    If Text1(5).Text <> "" And Val(Text1(5).Text) > 0 Then
        gegebenECLAST = CDbl(Text1(5).Text)
    Else
        gegebenECLAST = 0
    End If
    
    
    
    If Text1(2).Text <> "" And Val(Text1(2).Text) > 0 Then
        gegebenKK1 = CDbl(Text1(2).Text)
        cKK1art = "" & gcKreditKarte & ""
    Else
        gegebenKK1 = 0
    End If
   
    If Text1(7).Text <> "" And Val(Text1(7).Text) > 0 Then
        gegebenKK2 = CDbl(Text1(7).Text)
        cKK2art = "" & gcKreditKarte2 & ""
    Else
        gegebenKK2 = 0
    End If
    
    If Text1(6).Text <> "" And Val(Text1(6).Text) > 0 Then
        gegebenSCHECK = CDbl(Text1(6).Text)
    Else
        gegebenSCHECK = 0
    End If
    
    If Text1(1).Text <> "" And Val(Text1(1).Text) > 0 Then
        gegebenGUTSCHEIN = CDbl(Text1(1).Text)
    Else
        gegebenGUTSCHEIN = 0
    End If
    
    'Drucker an, Display aus, Init Drucker
    aDeviceName = Printer.DeviceName
    cEscapeSequenz = gcInit
    OpenDrawer aDeviceName, cEscapeSequenz
    
SCHUBLADE:
    
    If gbLadeCom Then
        OpenDrawerViaComPortModul20
    Else
        If gbAPI = False Then
            dAktZeit = Time
            lRet = Shell("Command.com /C " & gcPfad & "LADE.EXE", 6)
            dNeuZeit = Time
            Do While dNeuZeit - dAktZeit < (2 / 86400)
                dNeuZeit = Time
            Loop
        Else
            aDeviceName = Printer.DeviceName
            cEscapeSequenz = gcLade
            OpenDrawer aDeviceName, cEscapeSequenz
        End If
    End If

StartPunkt:
    lAnzZeile = 0
    ReDim cDruckZeile(1 To 1) As String
    
    iAktCopy = iAktCopy + 1
    iLevel = 1
    cDaten = ""
    iLenZeile = 32
    dSumme = 0
    dMWStVoll = 0
    dMWStErm = 0
    dMWSTVzzgl = 0
    dMWSTEzzgl = 0
    
    '***********************************************
    'Hier geht's los
    '***********************************************
    
    lAnzSatz = frmWKL20.List1.ListCount
    iLevel = 2
    
    '***********************************************
    'Drucker wird auf BonDrucker geschaltet
    '***********************************************
    
    aDeviceName = gcBonDrucker
    iLevel = 3
    dMWStVoll = 0
    dMWStErm = 0
    
    '***********************************************
    'Drucker ein- und Kundendisplay ausschalten
    '***********************************************
    
    cEscapeSequenz = gcInit
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'ggf. Logo auf Kassenbon bringen
    '***********************************************
    If gbBonDruck Then
        If gcBild <> "" Then
            aDeviceName = Printer.DeviceName
            cEscapeSequenz = gcBild
            OpenDrawer aDeviceName, cEscapeSequenz
        End If
    End If
    
    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Kopfdaten 1.Zeile an Drucker senden
    '***********************************************
    
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
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Kopfdaten 2.Zeile an Drucker senden
    '***********************************************
    
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
    '***********************************************
    'Kopfdaten 3.Zeile an Drucker senden
    '***********************************************

    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION"
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
        iStufe = 3
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

    '***********************************************
    'Trennstrich drucken
    '***********************************************
    
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '//Zahlung mit Gutschein + EC-Lastschrift
    If gbGutschUNDlastschrift = True Then  'Achtung diese Variable muss auch zurückgesetzt werden
    
    
    
    Else
        '***********************************************
        'Artikelpositionen drucken
        '***********************************************
        
        iLevel = 4
        dSumme = 0
        iGesAnzahl = 0
        For lAktSatz = 0 To lAnzSatz - 1
            cLBSatz = frmWKL20.List1.list(lAktSatz)
            
            cFeld = Mid(cLBSatz, 7, 6)
            If cFeld = "666666" Then
                bBonZwang = True
            End If
            lartnr = CLng(cFeld)
            If cFeld <> "000000" Then
                '1.Zeile: ArtNr + MWSTKz + ArtBezeich
                cDaten = cFeld & " "
                
                cFeld = Mid(cLBSatz, 72, 1)
                cDaten = cDaten & cFeld & "  "
                cMWST = cFeld
                cMWSTzzgl = cMWST
                
                cFeld = Mid(cLBSatz, 14, 35)
                cFeld = Trim$(cFeld)
                If Len(cFeld) > 17 Then
                    cFeld = Left(cFeld, 17)
                End If
                
                If gbDivKosmetik = True Then
                    If lartnr = 666666 Then
                        cDaten = cDaten & cFeld
                    Else
                        cDaten = cDaten & gcDivKosmetik
                    End If
                Else
                    cDaten = cDaten & cFeld
                End If
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                '***********************************************
                'wenn Artikelermäßigung, dann drucken
                '***********************************************
                
                ctmp = Mid(cLBSatz, 124, 3)
                If Val(ctmp) > 0 And Left(gFirma.FirmaName, 5) <> "Stief" And gbRabatt Then
                    'Zeile nur bei Artikel-Ermäßigung drucken
                    
                    Dim dArtikelrabattinEuro As Double
                    dArtikelrabattinEuro = CDbl(Trim(Mid(cLBSatz, 84, 9)))
                    Dim dRabattierterGesamtPreisinEuro As Double
                    dRabattierterGesamtPreisinEuro = CDbl(Trim(Mid(cLBSatz, 60, 9)))
                    Dim dErgebnisinProz As Double
                    dErgebnisinProz = dArtikelrabattinEuro * 100 / (dRabattierterGesamtPreisinEuro + dArtikelrabattinEuro)
                    ctmp = Format$(dErgebnisinProz, "###,##0.00")
                        
                    
                    cDaten = "Rabatt:    " & ctmp & " %"
                    ctmp = Mid(cLBSatz, 84, 9)
                    dSparSatzsum = dSparSatzsum + CDbl(Trim(ctmp))
                    ctmp = fnMoveComma2Point$(ctmp)
                    ctmp = Space(9 - Len(ctmp)) & ctmp
                    cDaten = cDaten & ctmp
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
                
                '***********************************************
                'Anzahl, Einzelpreis, Positionspreis drucken
                '***********************************************
                iGesAnzahl = iGesAnzahl + Val(Mid(cLBSatz, 1, 5))
                If Val(Mid(cLBSatz, 1, 5)) < 0 Then
                    gbStorni = True
                End If
                ctmp = Mid(cLBSatz, 1, 5)
                ctmp = Trim$(ctmp)
                ctmp = ctmp & Space$(6 - Len(ctmp))
                cDaten = ctmp & " x"
                
                If gbRabatt Then
                    ctmp = Mid(cLBSatz, 74, 9)
                    ctmp = fnMoveComma2Point$(ctmp)
                    dWert = Val(ctmp)
                Else
                    ctmp = Mid(cLBSatz, 50, 9)
                    ctmp = fnMoveComma2Point$(ctmp)
                    dWert = Val(ctmp)
                End If
                
                If Left(gFirma.FirmaName, 5) = "Stief" Then
                    ctmp = Format$((dWert * 100), "########0")
                Else
                    ctmp = Format$(dWert, "#####0.00")
                End If
                ctmp = Space(11 - Len(ctmp)) & ctmp
                cDaten = cDaten & ctmp
                
                ctmp = Mid(cLBSatz, 60, 9)
                ctmp = fnMoveComma2Point$(ctmp)
                dWert = Val(ctmp)
                If giPreisKz > 1 And giPreisKz < 4 Then
                    If cMWSTzzgl = "V" Then
                        dMWSTVzzgl = dMWSTVzzgl + (dWert * (gdMWStV / 100))
                    End If
                    If cMWSTzzgl = "E" Then
                        dMWSTEzzgl = dMWSTEzzgl + (dWert * (gdMWStE / 100))
                    End If
                End If
                ctmp = Format$(dWert, "#####0.00")
                dSumme = dSumme + dWert
                ctmp = Space(13 - Len(ctmp)) & ctmp
                cDaten = cDaten & ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                '***********************************************
                'MWSt-Summe berechnen
                '***********************************************
                If cMWST = "V" Then
                    dMWSt = dWert / (100 + gdMWStV)
                    dMWSt = dMWSt * gdMWStV
                    dMWStVoll = dMWStVoll + dMWSt
                ElseIf cMWST = "E" Then
                   dMWSt = dWert / (100 + gdMWStE)
                    dMWSt = dMWSt * gdMWStE
                    dMWStErm = dMWStErm + dMWSt
                Else
                    dMWSt = 0
                End If
            Else
'                'Zeile mit Zwischensumme drucken
'                cDaten = "Zwischensumme:     "
'
'                ctmp = Mid(cLBSatz, 60, 9)
'                ctmp = fnMoveComma2Point$(ctmp)
'                dWert = Val(ctmp)
'                ctmp = Format$(dWert, "#####0.00")
'                ctmp = Space(13 - Len(ctmp)) & ctmp
'
'                cDaten = cDaten & ctmp

                'Zeile mit Zwischensumme drucken
                ctmp = Mid(cLBSatz, 13, Len(cLBSatz) - 13)
                ctmp = Left(Trim(ctmp), 32)
                cDaten = ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
        Next lAktSatz
        
        '***********************************************
        'Trennstrich drucken
        '***********************************************
        iLevel = 5
        cDaten = String$(iLenZeile, "-")
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        
        '***********************************************
        'Summe drucken
        '***********************************************
        
        ctmp = "Summe" & Space$(13) & gcWaehrung
'        ctmp = "Endbetrag" & Space$(9) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(dEndbetrag, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        iLevel = 5
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        '***********************************************
        'Zeile Trennstrich drucken
        '***********************************************
        cDaten = String$(iLenZeile, "_")
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        
        
        
        
        
        
        
        
        
        
        
    End If '//If gcZahlMittel = "LS" And giZahlArt = 1 Then
    
    
    
    '***********************************************
    'Gesamtrabatt drucken
    '***********************************************
        
    If frmWKL20.Label2(3).Visible And Left(gFirma.FirmaName, 5) <> "Stief" And gbRabatt Then
        'Zeile nur bei Gesamt-Ermäßigung drucken
        
        dWert = fnHoleGesamtRabattModul20#()
        ctmp = frmWKL20.Label2(3).Caption
        If Len(ctmp) > 6 Then
            ctmp = Left(ctmp, Len(ctmp) - 5)
        End If
        
        Dim cDruckGesRabattBezeichnung As String
            
        cDruckGesRabattBezeichnung = "GesRabatt: "
        cDaten = cDruckGesRabattBezeichnung & Space(5 - Len(ctmp)) & ctmp & "% " & gcWaehrung
        
        
        
        dSparSatzsum = dSparSatzsum + dWert
        
        ctmp = Format$(dWert, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Bei Preiskz = 2 oder 3 (EK-Preis) zzgl.MWST
    '***********************************************
    GoTo weiter1
    If giPreisKz > 1 And giPreisKz < 4 Then
        ctmp = "zzgl. MWSt.:  " & Format$(gdMWStV, "#0") & "%" & Space$(1) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(dMWSTVzzgl, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        gdSumme = gdSumme + dMWSTVzzgl
    
        ctmp = "zzgl. MWSt.:   " & Format$(gdMWStE, "#0") & "%" & Space$(1) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(dMWSTEzzgl, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        gdSumme = gdSumme + dMWSTEzzgl
        dSumme = gdSumme
    End If
    
weiter1:
    
    '***********************************************
    'Zahlungsart drucken gegeben insgesamt
    '***********************************************
    
    
    ctmp = "gegeben insg.: " & Space$(3) & gcWaehrung
    cDaten = ctmp
    ctmp = Format$(gegebenInsgesamt, "#####0.00")
    ctmp = Space$(11 - Len(ctmp)) & ctmp
    cDaten = cDaten & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    iLevel = 606

    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Zahlungsart drucken davon
    '***********************************************
    ctmp = "davon"
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    iLevel = 606

    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'davon in Bar
    '***********************************************
    If gegebenBAR > 0 Then
        ctmp = "als Bargeld: " & Space$(5) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(gegebenBAR, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        iLevel = 606
    
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'davon gutschein
    '***********************************************
    If gegebenGUTSCHEIN > 0 Then
    
        gbADTBON = True
        
        ctmp = "als Gutschein: " & Space$(3) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(gegebenGUTSCHEIN, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        iLevel = 606
    
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        
        
        For i = 0 To 19
            If Gutschl(i).gutschnr <> 0 Then
            
                gutschwert = Gutschl(i).gutschwert
            
                ctmp = "Gutschein:" & Space$(10 - Len(Gutschl(i).gutschnr)) & Gutschl(i).gutschnr
                cDaten = ctmp
                ctmp = Format$(gutschwert, "#####0.00")
                
                
                cDaten = cDaten & Space$((iLenZeile - Len(cDaten)) - Len(ctmp)) & ctmp
'                cDaten = cDaten & cTmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                
            Else
                Exit For
            End If
        Next i
    End If
    
    '***********************************************
    'davon scheck
    '***********************************************
    If gegebenSCHECK > 0 Then
        ctmp = "als Scheck: " & Space$(6) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(gegebenSCHECK, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        iLevel = 606
    
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'davon eclast
    '***********************************************
    If gegebenECLAST > 0 Then
        ctmp = "als EC Last: " & Space$(5) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(gegebenECLAST, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        iLevel = 606
    
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'davon kk1
    '***********************************************
    If gegebenKK1 > 0 Then
        ctmp = "Kreditkarte" & cKK1art & ":" & Space$(2) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(gegebenKK1, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        iLevel = 606
    
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'davon kk2
    '***********************************************
    If gegebenKK2 > 0 Then
        ctmp = "Kreditkarte" & cKK2art & ":" & Space$(2) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(gegebenKK2, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        iLevel = 606
    
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'davon dukate
    '***********************************************
    If gegebenDUKATE > 0 Then
        ctmp = "als Dukate(n): " & Space$(3) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(gegebenDUKATE, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        iLevel = 606
    
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If

    

    
    '***********************************************
    'Rückgeld drucken
    'Bei Barzahlung ohne Rückgeld werden die Zeilen
    'Zahlungsart und Rückgeld unterdrückt
    '***********************************************
    
    If dZurueckBAR > 0 Then
    
        iLevel = 607
        ctmp = "Zurück in Bar" & Space$(5) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(dZurueckBAR, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        iLevel = 608
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        
        
        
    End If
    
    If dZurueckGUTSCH > 0 Then
    
        iLevel = 607
        ctmp = "Restgutschein:" & Space$(4) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(dZurueckGUTSCH, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        iLevel = 608
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        If gbRGO = False Then
            If glnewGutschnr > 0 Then
                ctmp = "Gutschein:" & Space$(10 - Len(glnewGutschnr)) & glnewGutschnr
                cDaten = ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
        
        Else
            If glnewGutschnr > 0 Then
                ctmp = "Gutschein:" & Space$(10 - Len(glnewGutschnr)) & glnewGutschnr
                cDaten = ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            Else
                For i = 0 To 19
                    If Gutschl(i).gutschnr <> 0 Then
                        ctmp = "Gutschein:" & Space$(10 - Len(Gutschl(i).gutschnr)) & Gutschl(i).gutschnr
                        cDaten = ctmp
                        KonvertAnsiAscii cDaten
                        cEscapeSequenz = cDaten & vbCrLf
                        
                        lAnzZeile = lAnzZeile + 1
                        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                        cDruckZeile(lAnzZeile) = cEscapeSequenz
                        
                        Exit For
                    End If
                Next i
            End If
        End If
    End If
    
    
    '***********************************************
    'Zeile Leerzeile drucken
    '***********************************************
    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Zeile 'Endbeterag' drucken
    '***********************************************
    
    ctmp = "Endbetrag" & Space$(9) & gcWaehrung
    cDaten = ctmp
    ctmp = Format$(dEndbetrag, "#####0.00")
    ctmp = Space$(11 - Len(ctmp)) & ctmp
    cDaten = cDaten & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    iLevel = 6103
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Zeile Trennstrich drucken
    '***********************************************
    cDaten = String$(iLenZeile, "_")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = String$(iLenZeile, "_")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    '***********************************************
    'Zeile Leerzeile drucken
    '***********************************************
    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Zeile 'Storno Unterschrift' drucken
    '***********************************************
    
    
    If gbStorni Then
    
        If gb2BONST = True Then
            gbADTBON = True
        Else
            gbADTBON = False
        End If
        cDaten = "Betrag erhalten bzw. verrechnet "
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        
        cEscapeSequenz = vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        cEscapeSequenz = vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        
        cDaten = "      (Unterschrift Kunde)      "
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    End If
    
    If gbSparsatz And dSparSatzsum > 0 Then
        '***********************************************
        'Zeile 'Sie sparen' drucken
        '***********************************************
        
        ctmp = "Sie sparen " & Format$(dSparSatzsum, "#####0.00 ") & gcWaehrung
        dSparSatzsum = 0
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
    End If
    
    
    '***********************************************
    'Zeile 'Anzahl Artikel' drucken
    '***********************************************
    If iGesAnzahl > 1 Then
        ctmp = "Anzahl Artikel: " & iGesAnzahl
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Zeile 'Es bediente Sie' drucken
    '***********************************************
    
    ctmp = "Es bediente Sie"
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Zeile Bedienername drucken
    '***********************************************
    iLevel = 611
    
    ctmp = gcBediener
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    '***********************************************
    'Zeile 'Kassennummer' drucken
    '***********************************************
    
    ctmp = "Kasse: " & gcKasNum
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    iLevel = 612
    '***********************************************
    'Zeile Kundennummer drucken
    '***********************************************
    
    If Label1(27).Caption <> "0" Or Label1(27).Visible Then
        
        iLevel = 613
        ctmp = "Ihre KundenNr: " & Label1(27).Caption
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    
        If gbKUNDENA = True Then
        
            If gbKUIBONfirma Then
                ctmp = lookingForKundendaten(Trim(Label1(27).Caption)).firma
            
                If ctmp <> "" Then
                    If Len(ctmp) > 32 Then ctmp = Left(ctmp, 32)
                    cDaten = ctmp
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
            End If
            
            If gbKUIBONtitel Then
                ctmp = lookingForKundendaten(Trim(Label1(27).Caption)).titel
            
                If ctmp <> "" Then
                    If Len(ctmp) > 32 Then ctmp = Left(ctmp, 32)
                    cDaten = ctmp
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
            End If
            
            ctmp = ""
            If gbKUIBONvorname Then
                ctmp = lookingForKundendaten(Trim(Label1(27).Caption)).vorname
            End If
            
            If gbKUIBONname Then
                If ctmp = "" Then
                    ctmp = ctmp & ""
                Else
                    ctmp = ctmp & " "
                End If
                ctmp = ctmp & lookingForKundendaten(Trim(Label1(27).Caption)).nachname
            End If
    
        
            iLevel = 614
            If Len(ctmp) > 32 Then
                ctmp = Left(ctmp, 32)
            End If
            cDaten = ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            
            If gbKUIBONstrasse Then
                ctmp = lookingForKundendaten(Trim(Label1(27).Caption)).strasse
            
                iLevel = 615
                If Len(ctmp) > 32 Then
                    ctmp = Left(ctmp, 32)
                End If
                cDaten = ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            ctmp = ""
            If gbKUIBONplz Then
                ctmp = lookingForKundendaten(Trim(Label1(27).Caption)).Plz
            End If
            
            If gbKUIBONort Then
                If ctmp = "" Then
                    ctmp = ctmp & ""
                Else
                    ctmp = ctmp & " "
                End If
                ctmp = ctmp & lookingForKundendaten(Trim(Label1(27).Caption)).Ort
            End If
        
            iLevel = 616
            If Len(ctmp) > 32 Then
                ctmp = Left(ctmp, 32)
            End If
            cDaten = ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            iLevel = 617
            
            If gbKUIBONtel Then
                ctmp = lookingForKundendaten(Trim(Label1(27).Caption)).telefon
                If ctmp <> "" Then
                    cDaten = "Tel " & ctmp
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
            End If
            
            iLevel = 618
            
            If gbKUIBONmobil Then
                ctmp = lookingForKundendaten(Trim(Label1(27).Caption)).Mobiltel
                If ctmp <> "" Then
                    cDaten = "Mobil " & ctmp
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
            End If
            
        End If
    End If
    
    '***********************************************
    'Zeile Datum, BelegNr, Uhrzeit drucken
    '***********************************************
    iLevel = 615
    
    ctmp = Format$(Date, "DD.MM.YYYY")
    cDaten = ctmp
    
    iLevel = 6151
    ctmp = Format$(Now, "HH:MM")
    iLevel = 6152
    cDaten = cDaten & Space$(4) & ctmp
    
    If giZahlArt = giKOLLEGE Then
        ctmp = "0"
        gdBonNr = 0
    Else

        ctmp = Format$(gdBonNr, "#####0")
        If gbSPIEGEL Then
        
            Dim ctmp111 As String
            Dim N As Integer
            ctmp111 = ctmp
            ctmp = ""
            For N = Len(ctmp111) To 1 Step -1
            
                ctmp = ctmp & Mid(ctmp111, N, 1)
            
            Next N
            
            
        
        End If
    End If
    iLevel = 6153
    ctmp = gcKasNum & "/" & ctmp
    iLevel = 6154
    ctmp = Space$(8 - Len(ctmp)) & ctmp
    iLevel = 6155
    cDaten = cDaten & Space$(4) & ctmp
    
    iLevel = 6156
    KonvertAnsiAscii cDaten
    iLevel = 6157
    cEscapeSequenz = cDaten & vbCrLf
    
    iLevel = 6158
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    iLevel = 6159
    cDruckZeile(lAnzZeile) = cEscapeSequenz

    '***********************************************
    'Zeile Leerzeile drucken
    '***********************************************
    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Zeile Lieferschein bei Kreditverkäufen drucken
    '***********************************************
    
    iLevel = 7
    
    
    
    
    
    'bei eventueller Gutschein_Barauszahlung ein Unterschriftenfeld erzeugen
    If gbGUTSCHBARAUSZAHLUNGMITUNTER = True And dZurueckBAR > 0 And gegebenGUTSCHEIN > 0 Then
    
        cDaten = "Gutscheinauszahlung erhalten:"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '***********************************************
        'Zeile Leerzeile drucken
        '***********************************************
        cEscapeSequenz = vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        '***********************************************
        'Zeile Leerzeile drucken
        '***********************************************
        cEscapeSequenz = vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '***********************************************
        'Zeile Trennstrich drucken
        '***********************************************
        cDaten = String$(iLenZeile, "_")
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        cDaten = "(Unterschrift)"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz

    End If
    
    
    '****enventuell angefallene Garantiedaten einfügen 'Achtung 3 Mal im Programm
    
    If sind_Garatie_daten_zu_drucken(gdBonNr) = True Then
    
        
        cDaten = "*** Garantie - Informationen ***"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        ermGarantie_daten gdBonNr
        
        '3 Arrays auslesen
        Dim cWert As String
        Dim iCount As Integer
        For iCount = 1 To UBound(gcArrArtNr)
        
        
            'Artikelnummer
            cWert = gcArrArtNr(iCount)
            If cWert <> "" Then
                
            
                cDaten = "zu Artikel: " & cWert
                cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            End If
            
            'Seriennummer
            
            cWert = gcArrSerienNr(iCount)
            If cWert <> "" Then
                
                If Len(cWert) <= 22 Then
            
                    cDaten = "SerienNr: " & cWert
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                ElseIf Len(cWert) <= 32 Then
                
                    cDaten = "SerienNr: "
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                    cDaten = cWert
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                ElseIf Len(cWert) > 32 Then
                
                    Dim sNeuWert As String
                    
                    sNeuWert = Right(cWert, Len(cWert) - 22)
                
                    cDaten = "SerienNr: " & Left(cWert, 22)
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                    cDaten = Left(sNeuWert, 32)
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                    Do While Len(sNeuWert) >= 32
                        sNeuWert = Right(sNeuWert, Len(sNeuWert) - 32)
                        
                        If sNeuWert <> "" Then
                            cDaten = Left(sNeuWert, 32)
                            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                            KonvertAnsiAscii cDaten
                            cEscapeSequenz = cDaten & vbCrLf
                            lAnzZeile = lAnzZeile + 1
                            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                            cDruckZeile(lAnzZeile) = cEscapeSequenz
                        End If
                    Loop
                    
                    
                
                End If
            
            End If
            
            
            
            
            
        Next iCount
        
        
    
        cDaten = "*** Garantie - Informationen ***"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    End If
    
    
    
    '**** ENDE enventuell angefallene Garantiedate einfügen
    
    
    
    
    
    '***********************************************
    'Zeile Trennstrich drucken
    '***********************************************
    cDaten = String$(iLenZeile, gsSTERNZEICH)
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    
    'bei Gutschein sofort versteuern, dann den Gesamtumsatz und Nichtumsatz ermitteln
    'bei Bedarf auch die MWST-Anteile neu
    
    If gbGutscheinBeiVKversteuern = True Then
    
        Dim dBruttobetrag As Double
        Dim dNettobetrag As Double
        Dim dVolleMehrwertSteuer As Double
        
        dBruttobetrag = dEndbetrag - dNichtUmsatz
        
        dVolleMehrwertSteuer = dBruttobetrag * gdMWStV / (100 + gdMWStV)
        
        dNettobetrag = dBruttobetrag - dVolleMehrwertSteuer
        
        
        
        
        
        
        
        
        
        
        
        '***********************************************
        'Zeile Nettoumsatz
        '***********************************************
        iLevel = 6
        ctmp = "Nettoumsatz"
        cDaten = ctmp
        ctmp = Format$((dNettobetrag), "#####0.00")
        ctmp = Space$(15 - Len(ctmp)) & ctmp
        iLevel = 601
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
    
        iLevel = 602
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        
        
        
        
        
        
        '***********************************************
        'Zeile volle MWSt drucken
        '***********************************************
        If dVolleMehrwertSteuer <> 0 Then
            iLevel = 609
        
            ctmp = "MWSt.-Anteil: " & Format$(gdMWStV, "#0") & "%"
            cDaten = ctmp
            ctmp = Format$(dVolleMehrwertSteuer, "#####0.00")
            ctmp = Space$(9 - Len(ctmp)) & ctmp
            cDaten = cDaten & ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    
    
    
    
        '***********************************************
        'Zeile Bruttoumsatz
        '***********************************************
        
        iLevel = 6101
        ctmp = "Bruttoumsatz"
        cDaten = ctmp
        ctmp = Format$(dBruttobetrag, "#####0.00")
        ctmp = Space$(14 - Len(ctmp)) & ctmp
        iLevel = 6102
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        iLevel = 6103
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    Else
    
        If giPreisKz = 6 Then
    
            '***********************************************
            'Zeile Nettoumsatz
            '***********************************************
            iLevel = 6
            ctmp = "Nettoumsatz"
            cDaten = ctmp
            ctmp = Format$(dEndbetrag, "#####0.00")
            ctmp = Space$(15 - Len(ctmp)) & ctmp
            iLevel = 601
            cDaten = cDaten & ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
        
            iLevel = 602
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
        Else
    
            '***********************************************
            'Zeile Nettoumsatz
            '***********************************************
            iLevel = 6
            ctmp = "Nettoumsatz"
            cDaten = ctmp
            ctmp = Format$((dEndbetrag - (dMWStVoll + dMWStErm)), "#####0.00")
            ctmp = Space$(15 - Len(ctmp)) & ctmp
            iLevel = 601
            cDaten = cDaten & ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
        
            iLevel = 602
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            '***********************************************
            'Zeile volle MWSt drucken
            '***********************************************
            If dMWStVoll <> 0 Then
                iLevel = 609
            
                ctmp = "MWSt.-Anteil: " & Format$(gdMWStV, "#0") & "%"
                cDaten = ctmp
                ctmp = Format$(dMWStVoll, "#####0.00")
                ctmp = Space$(9 - Len(ctmp)) & ctmp
                cDaten = cDaten & ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
        
            '***********************************************
            'Zeile erm. MWSt drucken
            '***********************************************
            If dMWStErm <> 0 Then
                iLevel = 610
            
                ctmp = "MWSt.-Anteil: " & Format$(gdMWStE, "#0") & "%"
                cDaten = ctmp
                ctmp = Format$(dMWStErm, "#####0.00")
                ctmp = Space$(10 - Len(ctmp)) & ctmp
                cDaten = cDaten & ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
    
            '***********************************************
            'Zeile Bruttoumsatz
            '***********************************************
            
            iLevel = 6101
            ctmp = "Bruttoumsatz"
            cDaten = ctmp
            ctmp = Format$(dEndbetrag, "#####0.00")
            ctmp = Space$(14 - Len(ctmp)) & ctmp
            iLevel = 6102
            cDaten = cDaten & ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            iLevel = 6103
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
        End If
    
    End If
    
     

    '***********************************************
    'Zeile Trennstrich drucken
    '***********************************************
    cDaten = String$(iLenZeile, gsSTERNZEICH)
'    cDaten = String$(iLenZeile, "*")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    '***********************************************
    'Zeile Trennstrich drucken
    '***********************************************
    
    
    
    
 'TSE Footer START  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< START
       
       
  'wenn Endbetrag <= 0, dann TSE überspringen, weil ein
  'Kassenbon mit 0 Endbetrag nicht signiert wird
    If d68Summe > 0 Then
    
    Else
    
        R_StartTime = ""
        R_FinishTime = ""
        R_TransactionNr = ""
        R_QRCodeAlsText = ""
        R_QRCodeAlsImgPath = ""
        R_FinishSignatur = ""
        R_StartSignatur = ""
        GoTo NACH_TSE
        
    End If
    
    If E_TSE_Aktiv And TSE_OK Then
    
     
          TransactionSchreiben "", 1, 1, dMWStVoll, dMWStErm, 0, 0, 0, gegebenBAR, gegebenInsgesamt - gegebenBAR
     
        
            If TSE_OK Then

                    '''''''''''''''''''  TSE Start  ''''''''''''''''''''''''
                    cDaten = "TSE Start: " & R_StartTime

                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf

                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz

                    '''''''''''''''''''  TSE Ende  ''''''''''''''''''''''''''
                    cDaten = "TSE Ende: " & R_FinishTime

                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf

                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz

                    '''''''''''''''  TSE Transaction.Nr  '''''''''''''''''''
                    cDaten = "TSE Transaction.Nr: " & R_TransactionNr

                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf

                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz

                    '''''''''''''''''''  TSE Signatur  ''''''''''''''''''''''''
                    cDaten = "TSE Signatur: " & vbNewLine & SplitStringNachCharZahl(5, R_FinishSignatur, 32)

                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf

                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    

                    '''''''''''''''''''  TSE alle Info zusammen (Optional) ''''''''''''''''''''''''
'                    cDaten = "TSE Info: " & vbNewLine & SplitStringNachCharZahl(5, R_QRCodeAlsText, 32)
'
'                    KonvertAnsiAscii cDaten
'                    cEscapeSequenz = cDaten & vbCrLf
'
'                    lAnzZeile = lAnzZeile + 1
'                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
               Else

                    cDaten = "TSE nicht erreichbar !!!"

                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf

                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
               End If
        Else

            cDaten = "TSE ist deaktiviert/falsch" & vbNewLine & "     initialisiert !!!"

            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf

            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
        
'
'        cDaten = "TSE Start: " & TSS.TRStart 'Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
'
'        KonvertAnsiAscii cDaten
'        cEscapeSequenz = cDaten & vbCrLf
'
'        lAnzZeile = lAnzZeile + 1
'        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'        cDruckZeile(lAnzZeile) = cEscapeSequenz
'
'
'        cDaten = "TSE Ende: " & TSS.TRFinish 'Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
'
'        KonvertAnsiAscii cDaten
'        cEscapeSequenz = cDaten & vbCrLf
'
'        lAnzZeile = lAnzZeile + 1
'        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'        cDruckZeile(lAnzZeile) = cEscapeSequenz
'
'
'        cDaten = "TSE TaNr: " & TSS.TRNo 'Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
'
'        KonvertAnsiAscii cDaten
'        cEscapeSequenz = cDaten & vbCrLf
'
'        lAnzZeile = lAnzZeile + 1
'        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'        cDruckZeile(lAnzZeile) = cEscapeSequenz
'
'
'
'
'
'        cDaten = "TSE Serial: " & TSS.Serial 'Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
'
'        KonvertAnsiAscii cDaten
'        cEscapeSequenz = cDaten & vbCrLf
'
'        lAnzZeile = lAnzZeile + 1
'        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'        cDruckZeile(lAnzZeile) = cEscapeSequenz
'
'
        
 
    End If
    
  'TSE Footer ENDE <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ENDE
   
      
        '***********************************************
        'Zeile Trennstrich drucken
        '***********************************************
    
        cDaten = String$(iLenZeile, gsSTERNZEICH)
    '    cDaten = String$(iLenZeile, "*")
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz

NACH_TSE:

    '***********************************************
    'Fußzeile 1 drucken
    '***********************************************
    
    'Fußzeilen
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "KEIN GÜLTIGER KASSENBON!"
    Else
        cDaten = gcBonText(2)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    
    '***********************************************
    'Fußzeile 2 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(3)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        iLevel = 10
    End If
    
    '***********************************************
    'Fußzeile 3 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = ""
    Else
        cDaten = gcBonText(5)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        iLevel = 10
    End If
    
    '***********************************************
    'Fußzeile 4 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(6)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    '***********************************************
    'Fußzeile 5 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(7)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    
    '***********************************************
    'Fußzeile 6 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(8)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If

    '***********************************************
    'Fußzeile 7 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(9)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    
    '***********************************************
    'Fußzeile 8 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(10)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    
    '***********************************************
    'Fußzeile 9 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(11)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    
    
    
    
    
    
    
    'Am Ende eventuell einen Rabattgutschein für den nächsten Einkauf
    
    
    cLBSatz = ""
        
    Dim bGutscheinbedingung_erfuellt As Boolean
    bGutscheinbedingung_erfuellt = False
            
    If giBonusNr = 1 Then
        
        If gbWWKundBi Then 'nur mit Kundenbindung Gutschein anbieten
            If frmWKL20!Label2(7).Caption <> "0" And frmWKL20!Label2(7).Visible Then
                bGutscheinbedingung_erfuellt = True
            End If
        Else
            bGutscheinbedingung_erfuellt = True
        End If
        
        
    
        If bGutscheinbedingung_erfuellt = True Then
        
            
            'für alle bonusfähigen Artikel größer Schwellenwert gibt es einen Rabattgutschein
            Dim dRabattfWert As Double
            Dim dDruckRabattWert As Double
            
            dRabattfWert = ermWertrabattf_Artikel
            
            'ich drucke Gutschein - dafür nehme ich den Bonus beim Kunden zurück
                           
            
            BonusVeränderung "negativ", CLng(frmWKL20!Label2(7).Caption), dRabattfWert, 0
            
            If dRabattfWert >= CDbl(gsWWSchwellenwert) Then
                '***********************************************
                'Zeile Leerzeile drucken
                '***********************************************
                cDaten = String$(iLenZeile, " ")
                cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                '***********************************************
                'Zeile  drucken
                '***********************************************
                
                cDaten = "---------- Gutschein -----------"
                cDaten = Trim$(cDaten)
                If cDaten <> "" Then
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
                
                '***********************************************
                'Zeile Leerzeile drucken
                '***********************************************
                cDaten = String$(iLenZeile, " ")
                cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                
            
                If gsWWArt = "Prozent" Then
                    
                    dDruckRabattWert = dRabattfWert * CDbl(gsWWwert) / 100
                Else
                    
                    dDruckRabattWert = CDbl(gsWWwert)
                End If
                
                Dim sEinText As String
                Dim sWort As String
                
                sEinText = Trim(gsTextVor) & " " & Format(dDruckRabattWert, "###,##0.00") & " " & Trim(gsWWZeichen) & " " & Trim(gsTextNach) & " "
            
                If Len(sEinText) > iLenZeile Then
                    Do While Len(sEinText) > 0
                        sWort = Mid(sEinText, 1, InStr(1, sEinText, " ") - 1)
                        
                        If Len(cLBSatz & sWort & Space(1)) > iLenZeile Then
                            cDaten = cLBSatz
                            KonvertAnsiAscii cDaten
                            cEscapeSequenz = cDaten & vbCrLf
                            
                            lAnzZeile = lAnzZeile + 1
                            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                            cDruckZeile(lAnzZeile) = cEscapeSequenz
                            cLBSatz = ""
                        End If
                        
                        
                        
                        cLBSatz = cLBSatz & sWort & Space(1)
                        sEinText = Mid(sEinText, Len(sWort) + 2, Len(sEinText) - Len(sWort) + 1)
                    Loop
                    
                    
                    
                    
                    cDaten = cLBSatz
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                    If Val(gsWWBonusGDAUER) > 0 Then
                        cDaten = "Gültigkeit bis " & Format(DateValue(Now) + gsWWBonusGDAUER, "DD.MM.YYYY")
                        KonvertAnsiAscii cDaten
                        cEscapeSequenz = cDaten & vbCrLf
                        lAnzZeile = lAnzZeile + 1
                        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                        cDruckZeile(lAnzZeile) = cEscapeSequenz
                    End If
                    
                Else
                
                    
                    cLBSatz = sEinText
                    cDaten = cLBSatz
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
                


            End If
        End If
    End If
    
    
    
    
    '***********************************************
    'ein paar Leerzeilen drucken  <<<<<<<<<<<< START
    '***********************************************
    If Not MitQrCode Or Not E_TSE_Aktiv Or Not d68Summe > 0 Or Not TSE_OK Or altDruckModus Then
        'Barcode Bonus auf Bon
        If gsWWBonusArtnr <> "0" And dDruckRabattWert > 0 Then
    
        Else
    
    
            For lcount = 1 To gbLeereZeil
                If lcount = gbLeereZeil Then
                    cEscapeSequenz = vbCrLf
                Else
                    cEscapeSequenz = " " & vbCrLf
                End If
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            Next lcount
    
        End If
    End If
    '***********************************************
    'ein paar Leerzeilen drucken  <<<<<<<<<<<< ENDE
    '***********************************************
    
    
    
     
    'OpenDrawer3 benutzt die WindowsAPI
    'OpenDrawer4 geht über das PRINTER-Objekt
    
    'Schublade nur einmal öffnen
    
    If giAndersZahlung = giKREDIT And iAktCopy = 2 Then
        GoTo BON_DRUCKEN
    End If
    
    iLevel = 12
    'Kassenschublade öffnen (ich habe es auskommentiert, weil die Schublade oben unter (SCHUBLADE:)-Block schon geöffnet wurde)
    
'    If gbLadeCom Then
'        OpenDrawerViaComPortModul20
'    Else
'        If gbAPI = True Then
'            aDeviceName = Printer.DeviceName
'            cEscapeSequenz = gcLade
'            OpenDrawer aDeviceName, cEscapeSequenz
'        End If
'    End If
   
    
BON_DRUCKEN:

    gbBonDruck = True

    If gbBonDruck Then
        If gbAPI = True Then
            OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
        Else
            OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
        End If
    End If
    
    Dim bPlusLZ As Boolean
    bPlusLZ = False
    
    If gbBonDruck Then
        If gsWWBonusArtnr <> "0" Then
            If dDruckRabattWert > 0 Then
                Barcode_Bonus CStr(dDruckRabattWert), "7"
                bPlusLZ = True
            End If
        End If
    Else
        If gsWWBonusArtnr <> "0" And dDruckRabattWert > 0 Then
            bPlusLZ = True
        End If
    End If
        
    If iAktCopy = 1 Then
        'Bon-Daten sichern
        gdSumme = dEndbetrag
        SichernBonDaten cDruckZeile(), lAnzZeile, "GZ", Trim(Label1(27).Caption), False
    End If

'BON_SCHNEIDEN:
'    If gbBonDruck Then
'        'Kassenbon abschneiden
'        If gbAPI = True Then
'            aDeviceName = Printer.DeviceName
'            cEscapeSequenz = gcSchneiden
'            OpenDrawer aDeviceName, cEscapeSequenz
'        End If
'    End If
    
    iLevel = 11
ZWEITER_BON:
    
    If gbADTBON And iAktCopy < 2 Then
        GoTo StartPunkt
    End If
    Erase cDruckZeile
    
   'uncommit die folgende Zeile zum Drucken des QR-Codes auf dem Bon
    If MitQrCode And E_TSE_Aktiv Then
     'beim altDruckModus kann der Drucker beim Kunde kein QR-Code drucken (alte Drucker)
         If altDruckModus Then
         
         Else
            If gbBonDruck Then
              QRcodeDrucken
            End If
           
         End If
    End If
        
    If altDruckModus Then
       'Papier schneiden (alte Funktion)
        If gbBonDruck Then
            'Kassenbon abschneiden
            If gbAPI = True Then
                aDeviceName = Printer.DeviceName
                cEscapeSequenz = gcSchneiden
                OpenDrawer aDeviceName, cEscapeSequenz
            End If
        End If
    
    Else
        'Papier schneiden (neue Funktion)
        If gbBonDruck Then
         CutPapier
        End If
    End If
     
     
GUTSCHEIN:
    lAnzLbSatz = frmWKL20.List1.ListCount
    For lcount = 0 To lAnzLbSatz - 1
        cLBSatz = frmWKL20.List1.list(lcount)
        cArtNr = Mid(cLBSatz, 7, 6)
        If cArtNr = "666666" Then
            If gbNoBonGu = False Then
                
                 If Not altDruckModus Then
                    PaarLeereZeilenDrucken
                    Sleep 2000
                 End If
                
                DruckeGutscheinBonWKL68 cLBSatz
                
                If altDruckModus Then
                'Papier schneiden (alte Funktion)
                     If gbBonDruck Then
                         'Kassenbon abschneiden
                         If gbAPI = True Then
                             aDeviceName = Printer.DeviceName
                             cEscapeSequenz = gcSchneiden
                             OpenDrawer aDeviceName, cEscapeSequenz
                         End If
                     End If
                 
                 Else
                  'Papier schneiden (neue Funktion)
                  If gbBonDruck Then
                     CutPapier
                  End If
                 End If
        
                
            End If
            
        End If
    Next lcount
    
    If gbRGO = False Then
        If glnewGutschnr > 0 Then
            cLBSatz = Space$(100)
            'Rückgabe-Gutschein drucken
            ctmp = Trim$(Str$(glnewGutschnr))
            ctmp = Space$(8 - Len(ctmp)) & ctmp
            Mid(cLBSatz, 24, 8) = ctmp
    
            ctmp = Format$(dZurueckGUTSCH, "#####0.00")
            ctmp = Trim$(ctmp)
            ctmp = Space$(9 - Len(ctmp)) & ctmp
            Mid(cLBSatz, 60, 9) = ctmp
            DruckeGutscheinBonModul20 cLBSatz
        End If
    Else
        If glnewGutschnr > 0 Then
            cLBSatz = Space$(100)
            'Rückgabe-Gutschein drucken
            ctmp = Trim$(Str$(glnewGutschnr))
            ctmp = Space$(8 - Len(ctmp)) & ctmp
            Mid(cLBSatz, 24, 8) = ctmp
    
            ctmp = Format$(dZurueckGUTSCH, "#####0.00")
            ctmp = Trim$(ctmp)
            ctmp = Space$(9 - Len(ctmp)) & ctmp
            Mid(cLBSatz, 60, 9) = ctmp
            DruckeGutscheinBonModul20 cLBSatz
        End If
    End If
    
    
    
ENDE:
    
    gbADTBON = False
     
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SendeDaten2DruckerNeuWKL68_Bonus"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
'    Resume Next
    
End Sub
Private Function allesPrüfen() As Boolean
On Error GoTo LOKAL_ERROR

allesPrüfen = False

'1. Karte prüfen
If Text1(2).Text <> "" Then
    If IsNumeric(Text1(2).Text) = True Then
        If Label33(5).Caption = "1. Karte" Then
            anzeige "rot", "Geben Sie bitte die 1. Kreditkarte an!", Label9
            iWelchekarte = 1
            Label6(2).Caption = "1. Kreditkarte auswählen"
            Frame2.BackColor = glH2
            Frame2.Visible = True
            Frame18.Visible = False
            Frame1.Visible = False
            Exit Function
            
        Else
        
        End If
    
    End If
End If

If Text1(7).Text <> "" Then
    If IsNumeric(Text1(7).Text) = True Then
        If Label33(17).Caption = "2. Karte" Then
            anzeige "rot", "Geben Sie bitte die 2. Kreditkarte an!", Label9
            iWelchekarte = 2
            Label6(2).Caption = "2. Kreditkarte auswählen"
            Frame2.BackColor = glH2
            Frame2.Visible = True
            Frame18.Visible = False
            Frame1.Visible = False
            Exit Function
            
        Else
        
        End If
    
    End If
End If

If IsNumeric(Label333(2).Caption) = True Then
    If CDbl(Label333(2).Caption) > 0 Then
        anzeige "rot", "Lassen Sie sich mehr Zahlungsmittel geben!", Label9
        Exit Function
    End If
End If

If Text1(5).Text <> "" Then
    If Text1(5).Enabled = True Then
        anzeige "rot", "EC - Lastschrift ist noch nicht abgerechnet.", Label9
        Exit Function
    End If
End If


Dim dGegebenKarteSum As Double
dGegebenKarteSum = 0

'1. Karte prüfen
If Text1(2).Text <> "" Then
    If IsNumeric(Text1(2).Text) = True Then
        dGegebenKarteSum = dGegebenKarteSum + CDbl(Text1(2).Text)
    
    End If
End If

If Text1(7).Text <> "" Then
    If IsNumeric(Text1(7).Text) = True Then
        dGegebenKarteSum = dGegebenKarteSum + CDbl(Text1(7).Text)
    End If
End If



Dim dBackinBar As Double
dBackinBar = 0

If IsNumeric(Label333(3).Caption) = True Then
    dBackinBar = CDbl(Label333(3).Caption)
End If




'If dGegebenKarteSum > 0 And dBackinBar > 0 Then
'    anzeige "rot", "Bitte passende Kreditkartenzahlung vornehmen (kein Rückgeld!)", Label9
'    Exit Function
'End If




allesPrüfen = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "allesPrüfen"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
Dim i As Integer

For i = 0 To 19
    Gutschl(i).gutschnr = 0
    Gutschl(i).gutschwert = 0
Next i

dGutscheinauszahlung = 0
dGutschwert = 0
glnewGutschnr = 0

PositionierenWKL68
Modul6.Skalieren_Kasse Me, True, True: Modul6.Schrift Me: Modul6.Log Me
Modul6.Farbform Me, Nothing

vorbereitung

ausblend

bAlterGutscheinImSpiel = False

anzeige "normal", "", Label9
Label33(23).ForeColor = vbRed

Text1(4).TabIndex = 0

Me.Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnPruefeEingabeECKarte() As Integer
On Error GoTo LOKAL_ERROR
    
    Dim ctmp        As String
    
    fnPruefeEingabeECKarte = 0
    
    ctmp = Text6(0).Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        fnPruefeEingabeECKarte = 1
        Exit Function
    End If
    
    ctmp = Text6(1).Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        fnPruefeEingabeECKarte = 2
        Exit Function
    End If
    
    ctmp = Text6(2).Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        fnPruefeEingabeECKarte = 3
        Exit Function
    End If
    
    If Val(ctmp) < 1 Or Val(ctmp) > 12 Then
        fnPruefeEingabeECKarte = 5
        Exit Function
    End If
    
    ctmp = String$(2 - Len(ctmp), "0") & ctmp
    Text6(2).Text = ctmp
    
    ctmp = Text6(3).Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        fnPruefeEingabeECKarte = 4
        Exit Function
    End If
    
    ctmp = String$(2 - Len(ctmp), "0") & ctmp
    Text6(3).Text = ctmp
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeECKarte"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub PositionierenWKL68()
On Error GoTo LOKAL_ERROR
    
    With Frame1
        .Top = 120
        .Height = 3855
        .Width = 5055
        .Left = 6720
    End With
    
    With Frame2
        .Top = 120
        .Height = 4695
        .Width = 5055
        .Left = 6720
    End With
    
    With Frame3
        .Top = 120
        .Height = 3855
        .Width = 5055
        .Left = 6720
    End With
    
    With Frame18
        .Top = 120
        .Height = 8415
        .Width = 5055
        .Left = 6720
    End With
    
    With Frame20
        .Top = 720
        .Height = 7695
        .Width = 5055
        .Left = 0
    End With
    
    With Frame4
        .Top = 0
        .Height = 9000
        .Width = 12000
        .Left = 0
        .Visible = False
    End With
    
    With Frame33
        .Top = 1080
        .Height = 5895
        .Width = 7215
        .Left = 2280
    End With
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL68"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Zurückrechnen(iind As Integer) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cBetrag             As String
    Dim dBetrag             As Double
    Dim dSumme              As Double
    Dim cSumme              As String
    Dim cEUR                As String
    Dim dEUR                As Double
    Dim dNochOffen          As Double
    Dim dNochOffenjetzt     As Double
    Dim cNochOffen          As String
    Dim i                   As Integer
    Dim dOffen              As Double
    Dim dzurück             As Double
    
    Zurückrechnen = False
    
    '2.Betrag addieren
    cBetrag = Text1(iind).Text
    cBetrag = fnMoveComma2Point$(cBetrag)
    dBetrag = Val(cBetrag)
    
    
    Text1(iind).SetFocus
    
    cSumme = Label333(0).Caption
    cSumme = fnMoveComma2Point$(cSumme)
    dSumme = Val(cSumme)
    
    cNochOffen = Label333(2).Caption
    cNochOffen = fnMoveComma2Point$(cNochOffen)
    dNochOffen = Val(cNochOffen)
    
    '4.noch Offen" füllen
    dNochOffenjetzt = dNochOffen - dBetrag
    
    If dNochOffen > 0 Then

        Label333(2).Caption = Format$(dNochOffenjetzt, "#####0.00")
        Label333(3).Caption = "0,00" 'zurück
    ElseIf dNochOffen <= 0 Then
        Zurückrechnen = True
    
        Label333(2).Caption = "0,00"
        dzurück = dNochOffenjetzt * (-1)
        Label333(3).Caption = Format$(dzurück, "#####0.00") 'zurück
    End If



Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zurückrechnen"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Private Function Offenrechnen(iind As Integer) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cBetrag             As String
    Dim dBetrag             As Double
    Dim dNochOffenjetzt     As Double
    Dim dRestgutsch         As Double
    Dim i                   As Integer
    Dim iRet                As Integer
    Dim cBetrag1            As String
    Dim dBetrag1            As Double
    Dim ctmp                As String
    
    dBetrag = 0
    dBetrag1 = 0
    Offenrechnen = False
    
    Label333(3).Caption = "0,00"
    
    If iind <> 4 Then
        cBetrag1 = Text1(iind).Text
        cBetrag1 = fnMoveComma2Point$(cBetrag1)
        dBetrag1 = dBetrag1 + Val(cBetrag1)

    End If
    
    
    
    
    For i = 0 To 7
        If i = 4 Then
        
        Else
            cBetrag = Text1(i).Text
            cBetrag = fnMoveComma2Point$(cBetrag)
            dBetrag = dBetrag + Val(cBetrag)
        End If
    Next i
    
    
    
    If iind <> 1 And iind <> 5 Then
        Text1(iind).SetFocus
    End If
    
    If dBetrag = 0 Then
        Label333(2).Caption = Format$(d68Nochoffen, "#####0.00")
        Label333(3).Caption = "0,00" 'zurück
    End If
    
    
    'neu
    Dim dGutschAll As Double
    dGutschAll = 0
    
    If Text1(1).Text <> "" Then
        dGutschAll = CDbl(Text1(1).Text)
    End If
                
    If dGutschAll > 0 Then
        'Gutschein im Spiel - auch Dukaten? Text1(3)
        If Text1(3).Text <> "" Then
            dGutschAll = dGutschAll + CDbl(Text1(3).Text)
        End If
    End If

    
    
    
    dRestgutsch = dGutschAll - d68Summe
    '    neu ende
    
    
    
    
    
'    dRestgutsch = CDbl(Label1(28).Caption)
    
    dNochOffenjetzt = d68Nochoffen - dBetrag '- dRestgutsch
    
    If dNochOffenjetzt > 0 Then
        Label333(2).Caption = Format$(dNochOffenjetzt, "#####0.00")
        Label333(3).Caption = "0,00" 'zurück
    ElseIf dNochOffenjetzt <= 0 Then
        Offenrechnen = True
    
        Label333(2).Caption = "0,00"
        
        If dRestgutsch > 0 Then
        
        
            If dRestgutsch <= gdRESTGU Then
            
                Label333(3).Caption = Format$(dNochOffenjetzt * (-1), "#####0.00")  'zurück

                

            Else
                Label1(28).Caption = Format$(dNochOffenjetzt * (-1), "#####0.00")  'zurück
                'alles anschalten
                Label1(28).Visible = True
                Label1(29).Visible = True
                Label33(22).Visible = True
                If gbRESTinBAR = True Then
                    SSCommand6(25).Visible = True
                End If
            End If
        Else
            Label333(3).Caption = Format$(dNochOffenjetzt * (-1), "#####0.00") 'zurück
        End If
        
    End If
    
    cBetrag = ""
    dBetrag = 0
    
    For i = 0 To 7
        If i = 4 Then
        
        Else
            cBetrag = Text1(i).Text
            cBetrag = fnMoveComma2Point$(cBetrag)
            dBetrag = dBetrag + Val(cBetrag)
        End If
        
    Next i
    Label333(1).Caption = Format$(dBetrag, "#####0.00")  'gegeben insgesamt

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Offenrechnen"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'   Resume Next
End Function
Private Sub vorbereitung()
    On Error GoTo LOKAL_ERROR
    
    Label333(0).Caption = "0,00"
    Label333(0).Refresh
    
    Label333(2).Caption = "0,00"
    Label333(2).Refresh
    
    Label1(1).Caption = -1
    
    Label333(0).Caption = Format$(d68Summe, "#####0.00")
    Label333(0).Refresh
        
    Label333(2).Caption = Format$(d68Nochoffen, "#####0.00")
    Label333(2).Refresh
    
    Label1(27).Caption = c68Kdnr

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitung"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ausblend()
    On Error GoTo LOKAL_ERROR
    
    If gbAUSBLDU Then
        Label33(8).Visible = False
        SSCommand6(15).Visible = False
        Text1(3).Visible = False
        Label33(11).Visible = False
    End If
    
    If gbAUSBLSH Then
        Label33(15).Visible = False
        SSCommand6(19).Visible = False
        Text1(6).Visible = False
        Label33(14).Visible = False
    End If
    
    If gbAUSBLLS Then
        Label33(13).Visible = False
        SSCommand6(17).Visible = False
        SSCommand6(21).Visible = False
        Text1(5).Visible = False
        Label33(12).Visible = False
    End If

    SSCommand8(0).Enabled = False
    SSCommand8(1).Enabled = False
    SSCommand8(2).Enabled = False
    SSCommand8(3).Enabled = False
    SSCommand8(5).Enabled = False
    SSCommand8(6).Enabled = False
    
    If gbKK_Visa Then SSCommand8(0).Enabled = True
    If gbKK_EurocardMastercard Then SSCommand8(1).Enabled = True
    If gbKK_AmericanExpress Then SSCommand8(2).Enabled = True
    If gbKK_DinersClub Then SSCommand8(3).Enabled = True
    If gbKK_ECKarte Then SSCommand8(5).Enabled = True
    If gbKK_Sonstige Then SSCommand8(6).Enabled = True

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ausblend"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub List1_Click()
On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    
    cLBSatz = Mid(List1.list(List1.ListIndex), 1, InStr(1, List1.list(List1.ListIndex), " "))
    zeiggutschdetails cLBSatz

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List1_Click"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub MSComm1_OnComm()
    On Error GoTo LOKAL_ERROR
    Dim lPos As Long
    Dim iRet As Integer
    Dim ctmp As String
    Dim ctmp1 As String
    Dim cTmp2 As String
    Dim lStart As Long
    Dim lAktuell As Long

    Select Case MSComm1.CommEvent
        Case comEvReceive                   ' Anzahl empfangener Zeichen gleich RThreshold

            lStart = Timer
            Do
                lAktuell = Timer
            Loop While lAktuell < lStart + 1

            gECKarte.Datenstrom = gECKarte.Datenstrom & MSComm1.Input
            gECKarte.Original = gECKarte.Datenstrom
            ctmp = MSComm1.Input
            ctmp1 = MSComm1.Input
            cTmp2 = MSComm1.Input

            If ctmp <> "" Then
                gECKarte.Datenstrom = gECKarte.Datenstrom & ctmp
            End If
            If ctmp1 <> "" Then
                gECKarte.Datenstrom = gECKarte.Datenstrom & ctmp1
            End If
            If cTmp2 <> "" Then
                gECKarte.Datenstrom = gECKarte.Datenstrom & cTmp2
            End If

            gECKarte.Original = gECKarte.Datenstrom


            If Left(gECKarte.Datenstrom, 1) = Chr$(2) Then
                gECKarte.Datenstrom = Right(gECKarte.Datenstrom, Len(gECKarte.Datenstrom) - 4)
            End If
            If gbDebug Then
                MsgBox gECKarte.Datenstrom
            End If
            lPos = InStr(1, gECKarte.Datenstrom, "?")
            If lPos > 0 Then

                ctmp1 = Left(gECKarte.Datenstrom, lPos)

                ctmp = Mid(gECKarte.Datenstrom, lPos + 1, Len(gECKarte.Datenstrom) - lPos)

                If Len(ctmp) < 3 Then
                    MsgBox "Karte konnte 2.Spur nicht als EC-Karte identifizieren!", vbCritical, "STOP!"
                    LeereDatenECKarteWKL20
                    MSComm1.PortOpen = False
                    MSComm1.CommPort = gVerbindung.iComPort
                    MSComm1.InputLen = 0
                    MSComm1.Settings = gVerbindung.cSettings
                    MSComm1.RThreshold = 1
                    If Not MSComm1.PortOpen = True Then
                        MSComm1.PortOpen = True
                    End If

                    Exit Sub
                End If
                ctmp = Right(ctmp, Len(ctmp) - 3)

                gECKarte.Datenstrom = ctmp1 & ctmp

                lPos = InStr(lPos, gECKarte.Datenstrom, "?")
                If lPos > 0 Then
                    Mid(gECKarte.Datenstrom, lPos, 1) = Chr$(13)
                    'Kartenlesen (Datenstrom) abgeschlossen
                    iRet = fnPruefeDatenstromECKarteWKL20()
                    If iRet <> 0 Then
                        LeereDatenECKarteWKL20
                    Else
                        'jetzt optional nach Hakensetzung
                        'Lucks
                        If gbNachKBbeiEC Then
                            If Val(gckundnr) > 0 Then
                                Label1(27).Caption = gckundnr
                                gckundnr = ""
                            End If
                        End If
                        LeseDatenECLastschriftWKL20
                    End If
                Else
                    iRet = MsgBox("Karte konnte nicht oder nur unvollständig gelesen werden (Fehler 2.Spur)!" & vbCrLf & "Details?", vbQuestion + vbYesNo + vbDefaultButton2, "HINWEIS")
                    If iRet = vbYes Then
                        MsgBox gECKarte.Datenstrom, vbInformation, "gelesene Daten"
                    End If
                    LeereDatenECKarteWKL20
                End If
            Else
                iRet = MsgBox("Karte konnte nicht oder nur unvollständig gelesen werden (Fehler 1.Spur)!" & vbCrLf & "Details?", vbQuestion + vbYesNo + vbDefaultButton2, "HINWEIS")
                If iRet = vbYes Then
                    MsgBox gECKarte.Datenstrom, vbInformation, "gelesene Daten"
                End If
                LeereDatenECKarteWKL20
            End If
            MSComm1.PortOpen = False
            MSComm1.CommPort = gVerbindung.iComPort
            MSComm1.InputLen = 0
            MSComm1.Settings = gVerbindung.cSettings
            MSComm1.RThreshold = 1
            If Not MSComm1.PortOpen = True Then
                MSComm1.PortOpen = True
            End If

           ' Fehler

        Case comBreak     'gestoppt
        Case comCDTO    ' CD-Zeitüberschreitung
        Case comCTSTO   ' CTS-Zeitüberschreitung
        Case comDSRTO   ' DSR-Zeitüberschreitung
        Case comFrame   ' Fehler im Übertragungsraster (Framing Error)
        Case comOverrun ' Datenverlust
        Case comRxOver  ' Überlauf des Empfangspuffers
        Case comRxParity    ' Paritätsfehler

        Case comTxFull  ' Sendepuffer voll
    ' Ereignisse
        Case comEvCD    ' Pegeländerung auf DCD
        Case comEvCTS   ' Pegeländerung auf CTS
        Case comEvDSR   ' Pegeländerung auf DSR
        Case comEvRing  ' Pegeländerung auf RI (Ring Indicator)

        Case comEvSend  ' Im Sendepuffer befinden sich SThreshold Zeichen
    End Select

Exit Sub
LOKAL_ERROR:
    If err.Number = 8005 Or err.Number = 8012 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "MSComm1_OnComm"
        Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."

        Fehlermeldung1
    End If
End Sub
Private Sub List11_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim lPos As Long
    Dim iRet As Integer
    Dim ctmp1 As String
    Dim cTmp2 As String
    Dim cZiel As String

    cZeichen = Chr$(KeyAscii)

    gECKarte.Datenstrom = gECKarte.Datenstrom & cZeichen

    lPos = InStr(1, gECKarte.Datenstrom, vbCr)
    If lPos > 0 Then
        lPos = lPos + 1
        lPos = InStr(lPos, gECKarte.Datenstrom, vbCr)
        If lPos > 0 Then
            'Kartenlesen (Datenstrom) abgeschlossen
            iRet = fnPruefeDatenstromECKarteWKL20()
            If iRet <> 0 Then
                LeereDatenECKarteWKL20
            Else
                'jetzt optional nach Hakensetzung
                'Lucks
                If gbNachKBbeiEC Then
                    If Val(gckundnr) > 0 Then
                        Label1(27).Caption = gckundnr
                        gckundnr = ""
                    End If
                End If
                LeseDatenECLastschriftWKL20
            End If
        End If
    Else
        lPos = InStr(1, gECKarte.Datenstrom, Chr$(95))
        If lPos > 0 Then
            lPos = lPos + 1
            lPos = InStr(lPos, gECKarte.Datenstrom, Chr$(95))
            If lPos > 0 Then

                'jetzt Sonderzeichen rauswerfen!
                ctmp1 = gECKarte.Datenstrom
                cZiel = ""
                For iRet = 1 To Len(ctmp1)
                    cTmp2 = Mid(ctmp1, iRet, 1)
                    If InStr("1234567890" & Chr$(95) & Chr$(180), cTmp2) <> 0 Then
                        cZiel = cZiel & cTmp2
                    End If
                Next iRet

                'jetzt alle Chr$(95) durch Chr$(13) ersetzen
                ctmp1 = cZiel
                cZiel = ""
                For iRet = 1 To Len(ctmp1)
                    cTmp2 = Mid(ctmp1, iRet, 1)
                    If InStr("1234567890" & Chr$(180), cTmp2) <> 0 Then
                        cZiel = cZiel & cTmp2
                    Else
                        cZiel = cZiel & Chr$(13)
                    End If
                Next iRet

                gECKarte.Datenstrom = cZiel

                iRet = fnPruefeDatenstromECKarteWKL20()
                If iRet <> 0 Then
                    LeereDatenECKarteWKL20
                Else
                    'jetzt optional nach Hakensetzung
                    'Lucks
                    If gbNachKBbeiEC Then
                        If Val(gckundnr) > 0 Then
                            Label1(27).Caption = gckundnr
                            gckundnr = ""
                        End If
                    End If
                    LeseDatenECLastschriftWKL20
                End If
            End If
        End If
    End If


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List11_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function fnPruefeDatenstromECKarteWKL20() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim lPos As Long
    Dim lcount As Long
    Dim lJahr As Long
    Dim lMonat As Long
    Dim cZeichen As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim iFehlerstufe As Integer
    Dim lTrenner1 As Long
    Dim lTrenner2 As Long
    Dim cMerker As String
    fnPruefeDatenstromECKarteWKL20 = 0
    
    iFehlerstufe = 1
    
    If Len(gECKarte.Datenstrom) > 1 Then
        
        iFehlerstufe = 2
        cMerker = gECKarte.Datenstrom
        'MsgBox cMerker
        If Left(gECKarte.Datenstrom, 1) = "&" Then
            'Drei-Spur-Leser!!!
            lTrenner1 = InStr(1, gECKarte.Datenstrom, "_")
            gECKarte.Datenstrom = Mid(gECKarte.Datenstrom, lTrenner1 + 1, Len(gECKarte.Datenstrom) - lTrenner1)
            gECKarte.Datenstrom = gECKarte.Datenstrom & vbCrLf & Space$(2)
            gECKarte.Datenstrom = gECKarte.Datenstrom & Mid(cMerker, 6, 8)
            gECKarte.Datenstrom = gECKarte.Datenstrom & Space$(1) & Mid(cMerker, 15, 10)
            gECKarte.Datenstrom = gECKarte.Datenstrom & Space$(10)
            'MsgBox gECKarte.Datenstrom
        ElseIf Left(gECKarte.Datenstrom, 1) = "ö" Then
            'Aures Leser!!!
            gECKarte.Datenstrom = Right(gECKarte.Datenstrom, Len(gECKarte.Datenstrom) - 1)
            gECKarte.Datenstrom = SwapStr(gECKarte.Datenstrom, "_", "")
            gECKarte.Datenstrom = SwapStr(gECKarte.Datenstrom, "`", "")
            gECKarte.Datenstrom = SwapStr(gECKarte.Datenstrom, "'", "=")
            'MsgBox gECKarte.Datenstrom
        End If
        
        If Left(gECKarte.Datenstrom, 2) <> "67" Then
            iFehlerstufe = 3
            MsgBox "Karte konnte nicht als EC-Karte identifiziert werden!" & vbCrLf & "Startkennung ist " & Left(gECKarte.Datenstrom, 2), vbCritical, "STOP!"
            fnPruefeDatenstromECKarteWKL20 = 1
            Exit Function
        End If
        If Len(gECKarte.Datenstrom) >= 19 Then
            iFehlerstufe = 4
            gECKarte.Konto1 = Mid(gECKarte.Datenstrom, 9, 10)
        Else
            gECKarte.Konto1 = ""
        End If
        
        lPos = InStr(gECKarte.Datenstrom, vbCr)
        
        If Len(gECKarte.Datenstrom) > lPos + 24 Then
            iFehlerstufe = 5
            gECKarte.BLZ = Mid(gECKarte.Datenstrom, lPos + 5, 8)
            gECKarte.Konto2 = Mid(gECKarte.Datenstrom, lPos + 14, 10)
        Else
            gECKarte.Konto2 = ""
        End If
        
        If gECKarte.Konto1 <> "" And gECKarte.Konto2 <> "" Then
            iFehlerstufe = 6
            If gECKarte.Konto1 <> gECKarte.Konto2 Then
                iFehlerstufe = 7
                MsgBox "Karte konnte nicht als EC-Karte identifiziert werden!" & vbCrLf & "(ungleiche Kontonummern auf Spur 1 und Spur 2)", vbCritical, "STOP!"
                fnPruefeDatenstromECKarteWKL20 = 1
                Exit Function
            End If
        Else
            iFehlerstufe = 8
            MsgBox "Karte konnte nicht als EC-Karte identifiziert werden!" & vbCrLf & "(Kontonummer Spur 1 nicht überprüfbar gegen Spur 2)", vbCritical, "STOP!"
            fnPruefeDatenstromECKarteWKL20 = 1
            Exit Function
        End If
        iFehlerstufe = 9
        If Val(gECKarte.BLZ) < 10000000 Then
            MsgBox "Karte konnte nicht als EC-Karte identifiziert werden!" & vbCrLf & "(ungültige Bankleitzahl)", vbCritical, "STOP!"
            fnPruefeDatenstromECKarteWKL20 = 1
            Exit Function
        End If
        iFehlerstufe = 10
        If Val(gECKarte.Konto1) = 0 Then
            MsgBox "Karte konnte nicht als EC-Karte identifiziert werden!" & vbCrLf & "(ungültige Kontonummer)", vbCritical, "STOP!"
            fnPruefeDatenstromECKarteWKL20 = 1
            Exit Function
        End If
        iFehlerstufe = 11
        lPos = 0
        For lcount = 1 To Len(gECKarte.Datenstrom)
            cZeichen = Mid(gECKarte.Datenstrom, lcount, 1)
            If InStr("1234567890", cZeichen) = 0 Then
                lPos = lcount
                Exit For
            End If
        Next lcount
        iFehlerstufe = 12
        If lPos > 0 Then
            lPos = lPos + 1
            gECKarte.jahr = Mid(gECKarte.Datenstrom, lPos, 2)
            If Val(gECKarte.jahr) < 80 Then
                gECKarte.jahr = "20" & gECKarte.jahr
            Else
                gECKarte.jahr = "19" & gECKarte.jahr
            End If
            iFehlerstufe = 13
            lPos = lPos + 2
            gECKarte.Monat = Mid(gECKarte.Datenstrom, lPos, 2)
            gECKarte.Monat = String$(2 - Len(gECKarte.Monat), "0") & gECKarte.Monat
        Else
            gECKarte.jahr = ""
            gECKarte.Monat = ""
        End If
        iFehlerstufe = 14
        If gECKarte.jahr <> "" And gECKarte.Monat <> "" Then
            lJahr = Year(Now)
            lMonat = Month(Now)
            iFehlerstufe = 15
            If Val(gECKarte.jahr) < lJahr Then
                MsgBox "Die EC-Karte ist abgelaufen!", vbCritical, "STOP!"
                fnPruefeDatenstromECKarteWKL20 = 1
                Exit Function
            ElseIf Val(gECKarte.jahr) = lJahr And Val(gECKarte.Monat) < lMonat Then
                MsgBox "Die EC-Karte ist abgelaufen!", vbCritical, "STOP!"
                fnPruefeDatenstromECKarteWKL20 = 1
                Exit Function
            End If
        Else
            MsgBox "Karte konnte nicht als EC-Karte identifiziert werden!" & vbCrLf & "(ungültiges Verfalldatum)", vbCritical, "STOP!"
            fnPruefeDatenstromECKarteWKL20 = 1
            Exit Function
        End If
        iFehlerstufe = 16
        LeseBankLeitZahlWKL20 Label1(27)
        iFehlerstufe = 17
        
        If gECKarte.BankName = "" Then
            MsgBox "Bankleitzahl konnte nicht zugeordnet werden!", vbCritical, "STOP!"
            fnPruefeDatenstromECKarteWKL20 = 1
            Exit Function
        End If
        iFehlerstufe = 18

        cSQL = "Select max(BELEGNR) as MAXBELEG from DTA"
        FnOpenrecordset rsrs, cSQL, 1, gdBase
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!MAXBELEG) Then
                lcount = rsrs!MAXBELEG
            Else
                lcount = 0
            End If
            
        Else
            lcount = 0
        End If
        iFehlerstufe = 19

        gECKarte.LastSchriftNr = Trim$(Str$(lcount + 1))
        
    Else
        MsgBox "Lesefehler!", vbCritical, "STOP!"
        fnPruefeDatenstromECKarteWKL20 = 1
        Exit Function
    End If
    
    iFehlerstufe = 20

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PruefeDatenstromECKarteWKL20"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub LeseDatenECLastschriftWKL20()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum      As String
    Dim czeit       As String
    Dim cBedNr      As String
    Dim cZeile      As String
    
    cDatum = Format$(Now, "DD.MM.YYYY")
    czeit = Format$(Now, "HH:MM")
    cBedNr = Text1(0).Text
    
    cZeile = cDatum & " " & czeit & "   " & cBedNr
    
    List11.Clear
    
    List11.AddItem "Lastschriftbeleg Nr. " & gECKarte.LastSchriftNr
    List11.AddItem cZeile
    List11.AddItem " "
    List11.AddItem "Hiermit ermächtige ich"
    List11.AddItem " "
    List11.AddItem gFirma.FirmaName
    List11.AddItem gFirma.strasse
    List11.AddItem gFirma.Plz & " " & gFirma.Ort
    List11.AddItem " "
    List11.AddItem "zum Einzug von " & Label11(3).Caption
    List11.AddItem "von meinem Konto " & gECKarte.Konto1
    List11.AddItem "BANKLEITZAHL " & gECKarte.BLZ
    List11.AddItem "(" & gECKarte.BankName
    List11.AddItem " in " & gECKarte.BankOrt & ")"
    List11.AddItem " "
    List11.AddItem "Für den Fall der Nichteinlösung"
    List11.AddItem "weise ich meine Bank"
    List11.AddItem "unwiderruflich an, "
    List11.AddItem gFirma.FirmaName
    List11.AddItem "auf Anforderung meinen Namen"
    List11.AddItem "und meine Anschrift vollständig"
    List11.AddItem "mitzuteilen. Insofern soll"
    List11.AddItem gFirma.FirmaName
    List11.AddItem "ein eigener Anspruch zustehen."
    List11.AddItem " "
    
    '***neu
    If Label1(27).Caption <> "0" Then
    
        List11.AddItem "KundenNr: " & Label1(27).Caption
        List11.AddItem Left(lookingForKundendaten(Trim(Label1(27).Caption)).vorname, 32)
        List11.AddItem Left(lookingForKundendaten(Trim(Label1(27).Caption)).nachname, 32)
        List11.AddItem Left(lookingForKundendaten(Trim(Label1(27).Caption)).strasse, 32)
        List11.AddItem Left(lookingForKundendaten(Trim(Label1(27).Caption)).Plz & " " & lookingForKundendaten(Trim(Label1(27).Caption)).Ort, 32)
        List11.AddItem " "
    Else
    
        List11.AddItem "--------------------------------"
        List11.AddItem " "
        List11.AddItem "Name/Vorname:"
        List11.AddItem " "
        List11.AddItem " "
        List11.AddItem " "
        List11.AddItem "Straße:"
        List11.AddItem " "
        List11.AddItem " "
        List11.AddItem " "
        List11.AddItem "Plz/Ort:"
        List11.AddItem " "
        List11.AddItem " "
        List11.AddItem " "
        
    End If
    
    'neu Ende
    
    List11.AddItem "Unterschrift des Kontoinhabers"
'    If gECKarte.KontoInhaber <> "unbekannt" Then
'        List11.AddItem "(" & gECKarte.KontoInhaber & ")"
'    End If
    List11.AddItem " "
    List11.AddItem " "
    List11.AddItem " "
    List11.AddItem " "
    Label11(0).Visible = False
    Label11(4).Caption = "Kunde / Kontoinhaber: " & gECKarte.KontoInhaber
    Label11(4).Visible = True
    
'    Label13(0).Visible = True
'    Label13(1).Visible = True
'    Label13(2).Visible = True
    
'    Command10(0).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseDatenECLastschriftWKL20"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SSCommand6_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim i           As Integer
    Dim dretutsch   As Double
    Dim dGeldwert   As Double
    Dim lKUNDNR     As Long
    Dim dGutschAll  As Double
    Dim iRet        As Integer
    Dim ctmp        As String
    
    If Label1(1).Caption <> -1 Then
        
        Select Case index
            
            Case 0 To 9    '** Ziffern **
            
                'Extra
                If Label1(1).Caption = 2 Then
                
                    iWelchekarte = 1
                    Label33(5).Caption = "1. Karte"
                    Label6(2).Caption = "1. Kreditkarte auswählen"
                    Frame2.BackColor = glH2
                    Frame2.Visible = True
                    Frame18.Visible = False
                    Frame1.Visible = False
                    ShowDie6ButtonsOrJustOneButton
                    
                ElseIf Label1(1).Caption = 7 Then
                
                    iWelchekarte = 2
                    Label33(17).Caption = "2. Karte"
                    Label6(2).Caption = "2. Kreditkarte auswählen"
                    Frame2.BackColor = glH2
                    Frame2.Visible = True
                    Frame18.Visible = False
                    Frame1.Visible = False
                    ShowDie6ButtonsOrJustOneButton
                End If
                
                'Extra Ende
                If Text1(Label1(1).Caption).Enabled = True Then
                    Text1(Label1(1).Caption).Text = Text1(Label1(1).Caption).Text & SSCommand6(index).Caption
                    Text1(Label1(1).Caption).SetFocus
                End If
                
            Case Is = 10    '** Komma **
                If Text1(Label1(1).Caption).Enabled = True Then
                    If InStr(Text1(Label1(1).Caption).Text, ",") = 0 Then
                        Text1(Label1(1).Caption).Text = Text1(Label1(1).Caption).Text & SSCommand6(index).Caption
                    End If
                    Text1(Label1(1).Caption).SetFocus
                End If
                
            Case Is = 11    '** Clear **
                anzeige "normal", "", Label9
                'Extra
                If Label1(1).Caption = 2 Then
                    back2 1
                    
                    
                ElseIf Label1(1).Caption = 7 Then
                    back2 2
                    
                ElseIf Label1(1).Caption = 1 Then 'Gutscheinwert focus
                
                    bAlterGutscheinImSpiel = False
                    For i = 0 To 19
                        Gutschl(i).gutschnr = 0
                        Gutschl(i).gutschwert = 0
                    Next i
                    Frame1.Visible = False
                End If
                'Extra Ende
                
                
                If Text1(Label1(1).Caption).Enabled = True Then
                    Text1(Label1(1).Caption).Text = ""
                    Text1(Label1(1).Caption).SetFocus
                End If
                
            Case Is = 12    '** Clear in Bar **
                anzeige "normal", "", Label9
                Text1(0).Text = ""
                Text1(0).SetFocus
                
                Offenrechnen (0)
                
            Case Is = 13    '** Clear in Gutsch **
            
                anzeige "normal", "", Label9
                Text1(1).Text = ""
                Text1(4).SetFocus
                
                For i = 0 To 19
                    Gutschl(i).gutschnr = 0
                    Gutschl(i).gutschwert = 0
                Next i
                
                Frame1.Visible = False
                Label1(28).Visible = False
                Label1(28).Caption = "0"
                Label1(29).Visible = False
                Label33(22).Visible = False
                SSCommand6(25).Visible = False
                
                bAlterGutscheinImSpiel = False
                
                DElausAlterG68
                
                Offenrechnen (1)
                

            Case Is = 14    '** Clear in 1.Karte **
                back2 1
                
                Text1(2).Text = ""
                Text1(2).SetFocus
                
                Offenrechnen (2)
                
            Case Is = 15    '** Clear in Dukate **
                anzeige "normal", "", Label9
                Text1(3).Text = ""
                Text1(3).SetFocus
                
                
                'Neu
                Label1(28).Visible = False
                Label1(28).Caption = "0"
                Label1(29).Visible = False
                Label33(22).Visible = False
                SSCommand6(25).Visible = False
                'Ende Neu
                
                
                
                Offenrechnen (3)
                
            Case Is = 16    '** Clear in suchegutschnr **
                Text1(4).Text = ""
                Text1(4).SetFocus
                
            Case Is = 17    '** Clear in EC Last **
                anzeige "normal", "", Label9
                If Text1(5).Enabled = True Then
                    Text1(5).Text = ""
                    Text1(5).SetFocus
                    
                    Offenrechnen (5)
                End If
                
            Case 18
            
                Dim cFeld As String

                dGutschAll = 0
                dGutschwert = 0
                
                Text1(4).Text = SwapStr(Text1(4).Text, ",", "")
                Text1(4).Text = SwapStr(Text1(4).Text, "-", "")
                Text1(4).Text = SwapStr(Text1(4).Text, "+", "")
                
                cFeld = Text1(4).Text
                
                
                
                
                
                
                
                
                If gbGutschnrKomplett = True Then
                    'gradmann's neue Gutscheine
                            'so belassen
                            
                    If Len(Text1(4).Text) = 6 Then
                        'alter gutschein von Hand eingegeben
                    End If
                    
                    If Len(Text1(4).Text) = 8 Then
                        'neuer gutschein von Hand
                         
                    End If
                            
                            
                    If Len(Text1(4).Text) = 13 Then
                        'neuer Gutschein gescannt
                        If Left(Text1(4).Text, 1) = "2" Then
                            cFeld = Left(Text1(4).Text, 8)
                        End If
                    End If
                Else
                
                
                

                    'selbst erstellter
                    If Len(Text1(4).Text) = 8 Then
                    
                        If Left(Text1(4).Text, 1) = "2" Then
                            cFeld = Mid(Text1(4).Text, 2, 6)
                        End If
                        
                        If Left(Text1(4).Text, 1) = "0" Then
                            cFeld = Mid(Text1(4).Text, 2, 6)
                        End If
                        
                        If Left(Text1(4).Text, 1) = "9" Then
                            cFeld = Mid(Text1(4).Text, 2, 6)
                        End If
                    End If
                    
                    'Gottmann
                    If Left(Text1(4).Text, 2) = "00" Then
                        cFeld = Left(Text1(4).Text, Len(Text1(4).Text) - 1)
                    End If
                    
                    'BA Gutscheine 10stellig
                    'Achtung vielleicht kollidiert es mit Gottmann und Goedecke
                    'dann BA-Gutscheine als Programmvariable
                    'beispiel 0004221829
                    If Len(Text1(4).Text) = 10 Then
                        cFeld = Val(Text1(4).Text)
                    End If
                    
                    'selbst erstellter 13er an der Kasse
                    If Len(Text1(4).Text) = 13 Then
                    
                        If Left(Text1(4).Text, 2) = "22" Then
                            cFeld = Mid(Text1(4).Text, 3, 10)
    
                        End If
                        
                        If Left(Text1(4).Text, 2) = "21" Then
                            cFeld = Mid(Text1(4).Text, 7, 6)
    
                        End If
                    End If
                
                End If
                
                
                
                
                
                
                
                
                
                
                Text1(4).Text = Val(cFeld)
                
                
                
                If suchegutsch(Text1(4).Text) Then
                    Text1(4).Text = ""
                    Frame1.BackColor = glH2
                    
                    Frame2.Visible = False
                    Frame18.Visible = False
                    Frame1.Visible = True
                    
                    anzeige "normal", "", Label9
                Else
                    Screen.MousePointer = 0
                End If
                
                
                
                
                If Text1(1).Text <> "" Then
                    dGutschAll = CDbl(Text1(1).Text) + dGutschwert
                Else
                    dGutschAll = dGutschwert
                End If
                
                If dGutschAll > 0 Then
                    'Gutschein im Spiel - auch Dukaten? Text1(3)
                    If Text1(3).Text <> "" Then
                        dGutschAll = dGutschAll + CDbl(Text1(3).Text)
                    End If
                
                    If dGutschAll > d68Summe Then
                        dretutsch = dGutschAll - d68Summe
                        
                        dGutscheinauszahlung = 0
                        If dretutsch <= gdRESTGU Then
                            dGutscheinauszahlung = dretutsch

                        Else
                            Label1(28).Visible = True
                            Label1(28).Caption = Format$(dretutsch, "######0.00")
                            Label1(29).Visible = True
                            Label33(22).Visible = True
                            If gbRESTinBAR = True Then
                                SSCommand6(25).Visible = True
                            End If
                            
                           
                        
                        End If
                    Else
                    
                        Label1(28).Visible = False
                        Label1(28).Caption = "0"
                        Label1(29).Visible = False
                        Label33(22).Visible = False
                        SSCommand6(25).Visible = False
                    
                    End If
                End If
                
                If dGutschwert > 0 Then
                    If Text1(1).Text = "" Then
                        Text1(1).Text = Format$(dGutschwert, "######0.00")
                    Else
                        Text1(1).Text = Format$(CDbl(Text1(1).Text) + dGutschwert, "######0.00")
                    End If
                End If
                
                Dim cNotiz As String
                
                cNotiz = ermittleGutschNotizen(Trim(Label1(3).Caption))
                If cNotiz <> "" Then
                    Frame5.Visible = True
                    Text1(9).Text = cNotiz
                Else
                    Frame5.Visible = False
                    Text1(9).Text = ""
                End If
                
                
                
                 If bAlterGutscheinImSpiel = True And CDbl(Label1(28).Caption) > 0 Then
                    Dim sMessText As String
                    
                    sMessText = "Achtung: Sie lösen mindestens einen alten Gutschein ein, "
                    sMessText = sMessText & "der vor dem Stichtag verkauft wurde. Für diese "
                    sMessText = sMessText & " Gutscheine werden aus steuerrechtlichen Gründen "
                    sMessText = sMessText & " keine Restgutscheine erzeugt. Der offene Betrag kann nur ausbezahlt werden."
                    MsgBox sMessText, vbInformation + vbOKOnly, "Winkiss Hinweis:"
                 
                    SSCommand6_Click 25
                End If
                
                
                
                
            Case Is = 19    '** Clear in Scheck **
                anzeige "normal", "", Label9
                Text1(6).Text = ""
                Text1(6).SetFocus
                
                Offenrechnen (6)
            Case Is = 20    '** Clear in 2.Karte **
                back2 2
                
                Text1(7).Text = ""
                Text1(7).SetFocus
                
                Offenrechnen (7)
            Case 21
                If IsNumeric(Text1(5).Text) Then
                    If CDbl(Text1(5).Text) > 0 Then
                        lastanzeige
                        If Frame18.Visible Then
                            List11.SetFocus
                        End If
                    End If
                End If
                
            Case 22
                gLGutschnum = -1
                frmWKL100.Show 1
                If gLGutschnum > 0 Then
                    Text1(4).Text = gLGutschnum
                    SSCommand6_Click 18
                End If
                
            Case 23
                Frame3.BackColor = glH2
                Frame3.Visible = True
                Text1(8).Text = ""
                Text1(8).SetFocus
                Label1(36).Caption = ermnextAltegutschnr
                Label1(36).Refresh
                
            Case 24
                Dim lKJADate As Long
                Dim cKJAZeit As String
                
                lKJADate = Fix(Now)
                cKJAZeit = Format$(Now, "HH:MM:SS")
    
                If IsNumeric(Text1(8).Text) Then
                    If CDbl(Text1(8).Text) > 0 Then
                        dGeldwert = CDbl(Text1(8).Text)
                        
                        lKUNDNR = Val(Label1(27).Caption)
                        
                        If dGeldwert > 999 Then
                            iRet = MsgBox("Der gegebene Betrag von: " & Format$(dGeldwert, "#####0.00 " & gcWaehrung) & " ist ungewöhnlich hoch. Möchten Sie die Eingabe löschen?", vbQuestion + vbYesNo, "Winkiss Frage:")
                            If iRet = vbYes Then
                                Text1(8).Text = ""
                                Text1(8).SetFocus
                                SSCommand6_Click 13
                                Exit Sub
                            Else    'Trotz warnung weitergemacht
                                ctmp = "Der gegebene Betrag von: " & Format$(dGeldwert, "#####0.00 " & gcWaehrung)
                                ctmp = ctmp & " ist ungewöhnlich hoch. Möchten Sie die Eingabe löschen?   --> wurde von Bediener " & gcBediener & " mit 'Nein' beantwortet"
                                schreibeProtokollGZwarn ctmp
                            End If
                        Else
                            InsertAlterG lKJADate, cKJAZeit, gcKasNum, dGeldwert, lKUNDNR
                        End If
                        
                        
                        
                    End If
                End If
                Frame3.Visible = False
                
                'Teil2****
                dGutschAll = 0
                dGutschwert = 0
                
                If dGeldwert > 0 Then
                    If TrageAltgutschEin(dGeldwert, CLng(Label1(36).Caption)) Then
                        Text1(4).Text = ""
                        Frame1.BackColor = glH2
                        
                        Frame2.Visible = False
                        Frame18.Visible = False
                        Frame1.Visible = True
                        
                        anzeige "normal", "", Label9
                    End If
                End If
                
                If Text1(1).Text <> "" Then
                    dGutschAll = CDbl(Text1(1).Text) + dGutschwert
                Else
                    dGutschAll = dGutschwert
                End If
                
                If dGutschAll > 0 Then
                
                    'Gutschein im Spiel - auch Dukaten? Text1(3)
                    If Text1(3).Text <> "" Then
                        dGutschAll = dGutschAll + CDbl(Text1(3).Text)
                    End If
                
                    If dGutschAll > d68Summe Then
                        dretutsch = dGutschAll - d68Summe
                        
                        dGutscheinauszahlung = 0
                        If dretutsch <= gdRESTGU Then
                            dGutscheinauszahlung = dretutsch
                        Else
                            Label1(28).Visible = True
                            Label1(28).Caption = Format$(dretutsch, "######0.00")
                            Label1(29).Visible = True
                            Label33(22).Visible = True
                            If gbRESTinBAR = True Then
                                SSCommand6(25).Visible = True
                            End If
                        End If
                    Else
                    
                        Label1(28).Visible = False
                        Label1(28).Caption = "0"
                        Label1(29).Visible = False
                        Label33(22).Visible = False
                        SSCommand6(25).Visible = False
                    
                    End If
                End If
                
                If dGutschwert > 0 Then
                    If Text1(1).Text = "" Then
                        Text1(1).Text = Format$(dGutschwert, "######0.00")
                    Else
                        Text1(1).Text = Format$(CDbl(Text1(1).Text) + dGutschwert, "######0.00")
                    End If
                End If
                
                'ende Teil2 ****
                
            Case 25
            
                dGutscheinauszahlung = CDbl(Label1(28).Caption)
            
                Label333(3).Caption = Format$(CDbl(Label333(3).Caption) + CDbl(Label1(28).Caption), "######0.00")
                
                Label1(28).Visible = False
                Label1(28).Caption = "0"
                Label1(29).Visible = False
                Label33(22).Visible = False
                SSCommand6(25).Visible = False
            Case 26
                Frame5.Visible = False
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand6_Click"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermnextAltegutschnr() As Long
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset
    
    ermnextAltegutschnr = 0
    
    cSQL = "Select max(agnu) as maxi from ALTERG "
    Set rsGZ = gdBase.OpenRecordset(cSQL)
    If Not rsGZ.EOF Then
        If Not IsNull(rsGZ!maxi) Then
            ermnextAltegutschnr = rsGZ!maxi
        End If
    End If
    rsGZ.Close: Set rsGZ = Nothing
    
    ermnextAltegutschnr = ermnextAltegutschnr + 1

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermnextAltegutschnr"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub InsertAlterG(lDat As Long, czeit As String, cKass As String, dMoney As Double, lKUNDNR As Long)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset
    
    frmWKL20.DeleteGutscheinWKL20 "1"
    
'    cSQL = "Delete from Gutsch where gutschnr = 1"
'    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from ALTERG where belegnr = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!BELEGNR = gcKasNum & "9999"
    rsGZ!kasnum = cKass
    rsGZ!GELDWERT = dMoney
    rsGZ!FILIALE = gcFilNr
    rsGZ!BEDNU = Val(gcBedienerNr)
    rsGZ!Kundnr = lKUNDNR
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InsertAlterG"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lastanzeige()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim dWert As Double
    Dim cZeile1 As String
    Dim cZeile2 As String
    
'    gcZahlMittel = "LS"
'    giZahlArt = 1

    ctmp = Text1(5).Text
    ctmp = fnMoveComma2Point$(ctmp)
    dWert = Val(ctmp)
    
    If gFirma.BLZ = "" Or gFirma.Konto = "" Or gFirma.BankName = "" Then
    
        MsgBox "Für Ihr Unternehmen ist keine Kontoverbindung angegeben!" & vbCrLf & "(->Service ->Einstellungen ->Unternehmensdaten)" & vbCrLf & "EC-Lastschrift nicht möglich!", vbCritical, "STOP!"
'        Text1(1).SetFocus
    Else
        If dWert > 0 Then
            cZeile1 = "EC-Lastschrift:"
            cZeile2 = gcWaehrung & " " & Text1(5).Text
            cZeile2 = Space$(20 - Len(cZeile2)) & cZeile2
            ZeigeKundenDisplay cZeile1, cZeile2
            
            Label11(0).Visible = True
            Label11(4).Visible = False
            LeereDatenECKarteWKL20
            List11.Clear
            Label11(3).Caption = Format$(Text1(5).Text, "######0.00 ") & " " & gcWaehrung
            Frame18.BackColor = glH2
            Frame19.BackColor = glH2
            Frame20.BackColor = glH2
            Frame22.BackColor = glH2
            Frame18.Visible = True
            Frame1.Visible = False
            Frame2.Visible = False
            
            MSComm1.PortOpen = False
            MSComm1.CommPort = gVerbindung.iComPort
            MSComm1.InputLen = 0
            MSComm1.Settings = gVerbindung.cSettings
            MSComm1.RThreshold = 1
            If Not MSComm1.PortOpen = True Then
                MSComm1.PortOpen = True
            End If
            
        ElseIf dWert < 0 Then
            anzeige "ROT", "EC-Lastschrift ist bei negativen Beträgen nicht möglich!", Label9

        ElseIf dWert = 0 Then
            anzeige "ROT", "Der Endbetrag ist 0. EC-Lastschrift ist nicht möglich!", Label9
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 8005 Or err.Number = 8012 Then
        Resume Next
    ElseIf err.Number = 8002 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "lastanzeige"
        Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Function suchegutsch(cGutschnr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    
    Dim ibednu          As Integer
    Dim cbedname        As String
    Dim GutscheinKunde  As Kunde
    Dim cKundnr         As String
    Dim dateDatAusg     As Date
    Dim iFiliale        As Integer
    Dim dWert           As Double
    Dim i               As Integer
    
    suchegutsch = False
    
    If cGutschnr = "" Then
        Exit Function
    End If
    
    If IsNumeric(cGutschnr) = False Then
        Exit Function
    End If
    
    If Len(cGutschnr) = 13 Then
        cGutschnr = CStr(Val(Mid(cGutschnr, 2, 11)))
    End If
    
    If Len(cGutschnr) = 12 Then
        cGutschnr = CStr(Val(Mid(cGutschnr, 2, 10)))
    End If
    
    For i = 0 To 19
        If Gutschl(i).gutschnr = cGutschnr Then
            anzeige "ROT", "Diese Gutscheinnummer haben Sie schon ausgewählt.", Label9
            Frame1.BackColor = glH2
                    
            Frame2.Visible = False
            Frame18.Visible = False
            Frame1.Visible = True
            Exit Function
        End If
    Next i
    
    If gbKL_LIVEGUTSCHEIN Then
    
        If fTestLogin_SQLDABA_Error = 0 Then 'ist alles OK? Datenbank erreichbar?
            'alles okay
            
        Else
        
            schreibeProtokollVPNTXT "Unterbrechung"
        
            Dim sTemp As String
            sTemp = "Bitte starten Sie diesen Rechner neu" & vbCrLf
            sTemp = sTemp & "oder schließen Sie das Schloss und starten Sie WinKiss neu."
        
            MsgBox sTemp, vbCritical + vbOKOnly, "Gutschein-Datenbank nicht erreichbar"
            
            Exit Function
        End If
    
    
    
        Dim stConnect As String
        
        If gsKL_DSN <> "" Then
            stConnect = "ODBC;DSN=" & gsKL_DSN & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
        Else
            stConnect = "ODBC;DRIVER=SQL Server;SERVER=" & gsKL_ADRESSE & ";DATABASE=" & gsKL_DATENBANKNAME & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
        End If
        
        Dim dbEAN As DAO.Database
        Set dbEAN = OpenDatabase(gsKL_DATENBANKNAME, dbDriverNoPrompt, False, stConnect)
    
        '1. Suchen auf dem SQL-SERVER in der Gutscheintabelle
        
        sSQL = "Select WERT,AUSG_Datum,AUSG_Filiale,AUSG_Kundnr,AUSG_BEDIENER from GUTSCHEINE where GUTSCHNR = '" & cGutschnr & "'"
        sSQL = sSQL & " and (EINL_DATUM is null or EINL_DATUM = 0 ) "
        Set rsrs = dbEAN.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
        
        
            If Not IsNull(rsrs!Wert) Then
                dWert = CDbl(rsrs!Wert)
                dGutschwert = dWert
                suchegutsch = True
            End If
            
            If Not IsNull(rsrs!AUSG_DATUM) Then
                dateDatAusg = rsrs!AUSG_DATUM
            End If
            
            If Not IsNull(rsrs!AUSG_FILIALE) Then
                iFiliale = rsrs!AUSG_FILIALE
            End If
            
            If Not IsNull(rsrs!AUSG_Kundnr) Then
                cKundnr = rsrs!AUSG_Kundnr
            End If
            
            If Not IsNull(rsrs!AUSG_BEDIENER) Then
                ibednu = rsrs!AUSG_BEDIENER
            End If
        
        End If
        rsrs.Close
        dbEAN.Close
    

    Else
        
        sSQL = "select * from gutsch where gutschnr = " & cGutschnr
        sSQL = sSQL & " and (dat_einl is null or dat_einl = 0 ) "
        sSQL = sSQL & " and STATUS <> 'L' "
        
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
        
            If Not IsNull(rsrs!Wert) Then
                dWert = CDbl(rsrs!Wert)
                dGutschwert = dWert
                suchegutsch = True
            End If
            
            If Not IsNull(rsrs!DAT_AUSG) Then
                dateDatAusg = rsrs!DAT_AUSG
            End If
            
            If Not IsNull(rsrs!FILIALE) Then
                iFiliale = rsrs!FILIALE
            End If
            
            If Not IsNull(rsrs!Kundnr) Then
                cKundnr = rsrs!Kundnr
            End If
            
            If Not IsNull(rsrs!BEDNU) Then
                ibednu = rsrs!BEDNU
            End If
        
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
      
    
    If suchegutsch = True Then
        GutscheinKunde.Mobiltel = lookingForKundendaten(Trim(cKundnr)).Mobiltel
        GutscheinKunde.Plz = lookingForKundendaten(Trim(cKundnr)).Plz
        GutscheinKunde.Ort = lookingForKundendaten(Trim(cKundnr)).Ort
        GutscheinKunde.strasse = lookingForKundendaten(Trim(cKundnr)).strasse
        GutscheinKunde.telefon = lookingForKundendaten(Trim(cKundnr)).telefon
        GutscheinKunde.Email = lookingForKundendaten(Trim(cKundnr)).Email
        GutscheinKunde.titel = lookingForKundendaten(Trim(cKundnr)).titel
        GutscheinKunde.telefon = lookingForKundendaten(Trim(cKundnr)).telefon
        GutscheinKunde.firma = lookingForKundendaten(Trim(cKundnr)).firma
        GutscheinKunde.anrede = lookingForKundendaten(Trim(cKundnr)).anrede
        GutscheinKunde.vorname = lookingForKundendaten(Trim(cKundnr)).vorname
        GutscheinKunde.nachname = lookingForKundendaten(Trim(cKundnr)).nachname
        
        Label1(3).Caption = cGutschnr
        Label1(11).Caption = DateValue(dateDatAusg)
        
        If gbGutscheinBeiVKversteuern = True Then
        
            Dim dateStichtag As Date
            dateStichtag = ermStichtag
            
            If dateDatAusg < dateStichtag Then
                    
                bAlterGutscheinImSpiel = True
                
            End If
        End If
    
    
        
        
        Label1(5).Caption = cKundnr
        Label1(13).Caption = GutscheinKunde.titel
        
        Label1(7).Caption = GutscheinKunde.vorname
        Label1(15).Caption = GutscheinKunde.nachname
        
        Label1(9).Caption = GutscheinKunde.Plz
        Label1(17).Caption = GutscheinKunde.Ort
        
        Label1(19).Caption = GutscheinKunde.strasse
        Label1(23).Caption = Format$(dWert, "######0.00 €")
        
        Label1(21).Caption = ibednu
        Label1(25).Caption = ermfromBed("BEDNAME", CStr(ibednu))
        
        For i = 0 To 19
            If Gutschl(i).gutschnr = 0 Then
                Gutschl(i).gutschnr = cGutschnr
                Gutschl(i).gutschwert = dWert
                Exit For
            End If
        Next i
    Else
        anzeige "ROT", "Diese Gutscheinnummer ist nicht vorhanden.", Label9
    End If
    
    
    
    List1.Clear
    
    For i = 0 To 19
        If Gutschl(i).gutschnr <> 0 Then
            List1.AddItem Gutschl(i).gutschnr & " " & Format$(Gutschl(i).gutschwert, "######0.00 €")
        Else
            Exit For
        End If
    Next i
    

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "suchegutsch"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function TrageAltgutschEin(dGutwert As Double, lGutschNumber As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim GutscheinKunde  As Kunde
    Dim i As Integer
    TrageAltgutschEin = False
    
    GutscheinKunde.Mobiltel = ""
    GutscheinKunde.Plz = ""
    GutscheinKunde.Ort = ""
    GutscheinKunde.strasse = ""
    GutscheinKunde.telefon = ""
    GutscheinKunde.Email = ""
    GutscheinKunde.titel = ""
    GutscheinKunde.telefon = ""
    GutscheinKunde.firma = ""
    GutscheinKunde.anrede = ""
    GutscheinKunde.vorname = ""
    GutscheinKunde.nachname = ""
        
    Label1(3).Caption = "1"
    Label1(23).Caption = Format$(dGutwert, "######0.00 €")
        
    For i = 0 To 19
        If Gutschl(i).gutschnr = 0 Then
            Gutschl(i).gutschnr = lGutschNumber
            Gutschl(i).gutschwert = dGutwert
            Exit For
        End If
    Next i
    
    dGutschwert = dGutwert
    
    List1.Clear
    
    For i = 0 To 19
        If Gutschl(i).gutschnr <> 0 Then
            List1.AddItem Gutschl(i).gutschnr & " " & Format$(Gutschl(i).gutschwert, "######0.00 €")
        Else
            Exit For
        End If
    Next i
    
    TrageAltgutschEin = True
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TrageAltgutschEin"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub zeiggutschdetails(cGutschnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    
    Dim ibednu          As Integer
    Dim cbedname        As String
    Dim GutscheinKunde  As Kunde
    Dim cKundnr         As String
    Dim dateDatAusg     As Date
    Dim iFiliale        As Integer
    Dim dWert           As Double
    Dim i               As Integer
    
    If cGutschnr = "" Then
        Exit Sub
    End If
    
    If IsNumeric(cGutschnr) = False Then
        Exit Sub
    End If
    
    
    
    If gbKL_LIVEGUTSCHEIN Then
    
        If fTestLogin_SQLDABA_Error = 0 Then 'ist alles OK? Datenbank erreichbar?
            'alles okay
            
        Else
        
            schreibeProtokollVPNTXT "Unterbrechung"
        
            Dim sTemp As String
            sTemp = "Bitte starten Sie diesen Rechner neu" & vbCrLf
            sTemp = sTemp & "oder schließen Sie das Schloss und starten Sie WinKiss neu."
        
            MsgBox sTemp, vbCritical + vbOKOnly, "Gutschein-Datenbank nicht erreichbar"
            Exit Sub
        End If
    
    
    
        Dim stConnect As String
        
        If gsKL_DSN <> "" Then
            stConnect = "ODBC;DSN=" & gsKL_DSN & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
        Else
            stConnect = "ODBC;DRIVER=SQL Server;SERVER=" & gsKL_ADRESSE & ";DATABASE=" & gsKL_DATENBANKNAME & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
        End If
        
        
        
        
        Dim dbEAN As DAO.Database
        Set dbEAN = OpenDatabase(gsKL_DATENBANKNAME, dbDriverNoPrompt, False, stConnect)
    
        '1. Suchen auf dem SQL-SERVER in der Gutscheintabelle
        
        sSQL = "Select * from GUTSCHEINE where GUTSCHNR = '" & cGutschnr & "'"
        Set rsrs = dbEAN.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
        
        
            If Not IsNull(rsrs!Wert) Then
                dWert = CDbl(rsrs!Wert)
            End If
            
            If Not IsNull(rsrs!AUSG_DATUM) Then
                dateDatAusg = rsrs!AUSG_DATUM
            End If
            
            If Not IsNull(rsrs!AUSG_FILIALE) Then
                iFiliale = rsrs!AUSG_FILIALE
            End If
            
            If Not IsNull(rsrs!AUSG_Kundnr) Then
                cKundnr = rsrs!AUSG_Kundnr
            End If
            
            If Not IsNull(rsrs!AUSG_BEDIENER) Then
                ibednu = rsrs!AUSG_BEDIENER
            End If
        
        End If
        rsrs.Close
        dbEAN.Close

    Else
    
        sSQL = "select * from gutsch where gutschnr = " & cGutschnr
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
        
            If Not IsNull(rsrs!Wert) Then
                dWert = CDbl(rsrs!Wert)
            End If
            
            If Not IsNull(rsrs!DAT_AUSG) Then
                dateDatAusg = rsrs!DAT_AUSG
            End If
            
            If Not IsNull(rsrs!FILIALE) Then
                iFiliale = rsrs!FILIALE
            End If
            
            If Not IsNull(rsrs!Kundnr) Then
                cKundnr = rsrs!Kundnr
            End If
            
            If Not IsNull(rsrs!BEDNU) Then
                ibednu = rsrs!BEDNU
            End If
        
        End If
        rsrs.Close: Set rsrs = Nothing
        
    End If
    
    GutscheinKunde.Mobiltel = lookingForKundendaten(Trim(cKundnr)).Mobiltel
    GutscheinKunde.Plz = lookingForKundendaten(Trim(cKundnr)).Plz
    GutscheinKunde.Ort = lookingForKundendaten(Trim(cKundnr)).Ort
    GutscheinKunde.strasse = lookingForKundendaten(Trim(cKundnr)).strasse
    GutscheinKunde.telefon = lookingForKundendaten(Trim(cKundnr)).telefon
    GutscheinKunde.Email = lookingForKundendaten(Trim(cKundnr)).Email
    GutscheinKunde.titel = lookingForKundendaten(Trim(cKundnr)).titel
    GutscheinKunde.telefon = lookingForKundendaten(Trim(cKundnr)).telefon
    GutscheinKunde.firma = lookingForKundendaten(Trim(cKundnr)).firma
    GutscheinKunde.anrede = lookingForKundendaten(Trim(cKundnr)).anrede
    GutscheinKunde.vorname = lookingForKundendaten(Trim(cKundnr)).vorname
    GutscheinKunde.nachname = lookingForKundendaten(Trim(cKundnr)).nachname
    
    Label1(3).Caption = cGutschnr
    Label1(11).Caption = DateValue(dateDatAusg)
    
    Label1(5).Caption = cKundnr
    Label1(13).Caption = GutscheinKunde.titel
    
    Label1(7).Caption = GutscheinKunde.vorname
    Label1(15).Caption = GutscheinKunde.nachname
    
    Label1(9).Caption = GutscheinKunde.Plz
    Label1(17).Caption = GutscheinKunde.Ort
    
    Label1(19).Caption = GutscheinKunde.strasse
    Label1(23).Caption = Format$(dWert, "######0.00 €")
    
    Label1(21).Caption = ibednu
    Label1(25).Caption = ermfromBed("BEDNAME", CStr(ibednu))
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeiggutschdetails"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SSCommand8_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lCents  As Long
    Dim sCent   As String
    Dim lret1    As Long
    Dim sBLZ As String * 2000
    Dim lBuffer As Long

    Dim lerrCode As Long
    Dim serrMeldung As String * 8000
    Dim lForcePIN As Long
    
    Dim lRet        As Long
    Dim iRet        As Integer
    Dim sTraceNr    As String
    Dim hwnd&
    Dim Y As String
    Dim result&
    Dim Title$
    Dim lPos As Long
    Dim cNachkomma As String
    
    
    If iWelchekarte = 1 Then
        If InStr(Text1(2).Text, ",") = 0 Then
            anzeige "rot", "Geldbeträge mit Komma eingeben!", Label9
            Text1(2).SetFocus
            Exit Sub
        Else
            'nach dem Komma 2 stellen?
            lPos = InStr(1, Text1(2).Text, ",")
            cNachkomma = Right(Text1(2).Text, Len(Text1(2).Text) - lPos)
            If Len(cNachkomma) <> 2 Then
                anzeige "rot", "Geldbeträge mit 2 Nachkommastellen eingeben!", Label9
                Text1(2).SetFocus
                Exit Sub
            End If
        End If
        
    
        Select Case index
            Case 0
                gcKreditKarte = "VI"
            Case 1
                gcKreditKarte = "EU"
            Case 2
                gcKreditKarte = "AE"
            Case 3
                gcKreditKarte = "DC"
            Case 4
                gcKreditKarte = "BC"
            Case 5
                gcKreditKarte = "EC"
            Case 6
                gcKreditKarte = "SO"
            Case 7
                gcKreditKarte = "Automatisch"
        End Select
        
        
        Label33(5).Caption = "1. Karte (" & gcKreditKarte & ")"
        Label33(5).Refresh
        SSCommand6(14).Picture = SSCommand8(index).Picture
        SSCommand6(14).Caption = ""
        
        'Anfang Kartenterminal
        
        If gbEcash Then
            Select Case gsEPartner
'                Case Is = "ADT"
'                    If gsAdtVerfahren = "XML" Then
'                        anzeige "rot1", "Bedienen Sie jetzt das Kartenterminal!", Label9
'
'                        If CDbl(Text1(2).Text) < 0 Then
'                            If CInt(gADTclientId) = 0 Then
'                                'Storno
'                                sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "TA Nr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
'                                lRet = Shell("C:\Programme\EL-ME\SECpos\SECposPay\SECposPay.exe", vbHide) 'secpos
'                                AppActivate lRet ' True
'
'                                'erstmal zum storno navigieren
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'
'                                SendKeys "{Down}", True
'
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'
'                                SendKeys "{Down}", True
'
'                                SendKeys sTraceNr, True
'                                SendKeys "{enter}", True
'
'                                Call keybd_event(VK_LWIN, 0, 0, 0)
'                                Call keybd_event(77, 0, 0, 0)
'                                Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
'
'                                iRet = MsgBox("Storno ok? ", vbInformation + vbYesNo, "Winkiss Frage:")
'                                Pause 5
'                                 'secpospay suchen finden und schließen
'                                Y = "SECpos Pay"
'                                hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
'                                Do
'                                    result = GetWindowTextLength(hwnd) + 1
'                                    Title = Space(result)
'                                    result = GetWindowText(hwnd, Title, result)
'                                    Title = Left$(Title, Len(Title) - 1)
'
'                                    If InStr(1, Title, Y) Then
'                                        SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
'                                    End If
'                                    hwnd = GetWindow(hwnd, GW_HWNDNEXT)
'                                Loop Until hwnd = 0
'
'                                setzedrucker gcBonDrucker
'
'                                If iRet = vbNo Then
'                                    Screen.MousePointer = 0
'
'                                        Text1(2).Text = ""
''                                    Command5(0).Enabled = True
''                                    Command5(1).Enabled = True
'                                    Exit Sub
'                                End If
'                            Else
'                                'Storno neuer Weg
'
'                                If CInt(gADTclientId) > 0 Then
'                                    If gADTipAdress <> "" And gADTport <> "" Then
'                                        '192.168.1.14 '20002
'                                        lRet = ELMESettings(vbNullString, gADTclientId, vbNullString, gADTipAdress, gADTport, -1, -1, -1, vbNullString)
'                                    Else
'                                        lRet = ELMESettings(vbNullString, gADTclientId, vbNullString, vbNullString, -1, -1, -1, -1, vbNullString)
'                                    End If
'                                End If
'
'                                sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "TA Nr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
'                                lRet = ELMEReversal(CLng(sTraceNr))
'                                If lRet = 0 Then
'                                    lRet = ELMEGetPrint(sBLZ, 8000)
'                                    If lRet = 0 Then
'                                        gsAdtBeleg = sBLZ
'
'                                        If gADTipAdress <> "" And gADTport <> "" Then
'                                            SendeDaten2DruckerECCASH_Kundenbeleg
'                                            SendeDaten2DruckerECCASH_Haendlerbeleg
'                                        Else
'                                            SendeDaten2DruckerECCASH
'                                        End If
'                                    Else
'                                        MsgBox "Fehler ELMEGetPrint: " & lRet, vbCritical, "Winkiss Fehler:"
'                                        gsAdtBeleg = ""
'                                    End If
'                                Else
'                                    lRet = ELMEGetLastError(lerrCode, serrMeldung, 8000)
'                                    If lRet = 0 Then
'                                        MsgBox serrMeldung, vbCritical, "Winkiss Fehler:"
'                                    Else
'                                        MsgBox "Fehler ELMEGetLastError: " & lRet, vbCritical, "Winkiss Fehler:"
'                                    End If
'                                    'Abbruch
'                                    Screen.MousePointer = 0
'
'                                    Text1(2).Text = ""
'                                    Exit Sub
'                                    'Abbruch
'                                End If
'                            End If
'                        Else
'                            If CInt(gADTclientId) = 0 Then
'                                'Zahlung
'                                lRet = Shell("C:\Programme\EL-ME\SECpos\SECposPay\SECposPay.exe", vbHide) 'secpos
'                                AppActivate lRet
'                                SendKeys Text1(2).Text, True
'                                SendKeys "{enter}", True
'                                Call keybd_event(VK_LWIN, 0, 0, 0)
'                                Call keybd_event(77, 0, 0, 0)
'                                Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
'                                iRet = MsgBox("Zahlung ok? ", vbInformation + vbYesNo, "Winkiss Frage:")
'                                Pause 5
'                                'secpospay suchen finden und schließen
'                                Y = "SECpos Pay" '  (Terminal-ID: " & gsTerminalid & ")"
'                                hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
'                                Do
'                                    result = GetWindowTextLength(hwnd) + 1
'                                    Title = Space(result)
'                                    result = GetWindowText(hwnd, Title, result)
'                                    Title = Left$(Title, Len(Title) - 1)
'                                    If InStr(1, Title, Y) Then
'                                        SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
'                                    End If
'                                    hwnd = GetWindow(hwnd, GW_HWNDNEXT)
'                                Loop Until hwnd = 0
'                                setzedrucker gcBonDrucker
'                                If iRet = vbNo Then
'                                    'abbruch
'                                    Screen.MousePointer = 0
'                                        Text1(2).Text = ""
'                                    Exit Sub
'                                End If
'                            Else
'                                'Zahlung der neue Weg
'
'                                If CInt(gADTclientId) > 0 Then
'                                    If gADTipAdress <> "" And gADTport <> "" Then
'                                        '192.168.1.14 '20002
'                                        lRet = ELMESettings(vbNullString, gADTclientId, vbNullString, gADTipAdress, gADTport, -1, -1, -1, vbNullString)
'                                    Else
'                                        lRet = ELMESettings(vbNullString, gADTclientId, vbNullString, vbNullString, -1, -1, -1, -1, vbNullString)
'                                    End If
'                                End If
'
'                                lForcePIN = 0
'
'                                sCent = Text1(2).Text
'                                sCent = SwapStr(sCent, ",", "")
'                                lCents = CLng(sCent)
'                                lret1 = ELMEPay(lCents, lForcePIN)
'
'                                If lret1 = 0 Then
'                                    lret1 = ELMEGetPrint(sBLZ, 2000)
'                                    If lret1 = 0 Then
'                                        gsAdtBeleg = sBLZ
'                                        If gADTipAdress <> "" And gADTport <> "" Then
'                                            SendeDaten2DruckerECCASH_Kundenbeleg
'                                            SendeDaten2DruckerECCASH_Haendlerbeleg
'                                        Else
'                                            SendeDaten2DruckerECCASH
'                                        End If
'                                    Else
'                                        MsgBox "Fehler ELMEGetPrint: " & lret1, vbCritical, "Winkiss Fehler:"
'                                        gsAdtBeleg = ""
'                                    End If
'                                Else
'                                    lret1 = ELMEGetLastError(lerrCode, serrMeldung, 8000)
'                                    If lret1 = 0 Then
'                                        MsgBox serrMeldung, vbCritical, "Winkiss Fehler:"
'                                    Else
'                                        MsgBox "Fehler ELMEGetLastError: " & lret1, vbCritical, "Winkiss Fehler:"
'                                    End If
'                                    'Abbruch
'                                    Screen.MousePointer = 0
'                                        Text1(2).Text = ""
''                                    Command5(0).Enabled = True
''                                    Command5(1).Enabled = True
'                                    Exit Sub
'                                    'Abbruch
'                                End If
'                            End If
'                        End If
'                ElseIf gsAdtVerfahren = "INOUT" Then
'
'                End If
                
                Case "ELP"
                
                    anzeige "rot1", "Bedienen Sie jetzt das Kartenterminal!", Label9
                        
                    If CDbl(Text1(2).Text) < 0 Then
                        'Storno
                        sCent = Text1(2).Text
                        sCent = SwapStr(sCent, ",", "")
                        sCent = SwapStr(sCent, "-", "")
                        
                        sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "TA Nr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
                               
                        Storno_elPAY sTraceNr, sCent
                        
                        If giELPAY_Fehler > 0 Then
                                
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                            Text1(2).Text = ""
                            
                            anzeige "rot1", "Fehler am Kartenterminal!", Label9
                            
                            Exit Sub
                            'Abbruch
                        End If
                        
                    Else
        
                        'Zahlung
                        sCent = Text1(2).Text
                        sCent = SwapStr(sCent, ",", "")
                        
                        Zahlung_elPAY sCent
                        
                        If giELPAY_Fehler > 0 Then
                                
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                            Text1(2).Text = ""
                            
                            anzeige "rot1", "Fehler am Kartenterminal!", Label9

                            Exit Sub
                            'Abbruch
                        End If
                            
                    End If
                    
                Case "ZVT"
                
                    anzeige "rot1", "Bedienen Sie jetzt das Kartenterminal!", Label9
                        
                    If CDbl(Text1(2).Text) < 0 Then
                        'Storno
                        sCent = Text1(2).Text
                        sCent = SwapStr(sCent, ",", "")
                        sCent = SwapStr(sCent, "-", "")
                        
                        dlgTaNr.Show 1
                    
                        sTraceNr = dlgTaNr.Back
'                        sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "BNr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
                               
                        Storno_ZVT sTraceNr
                        
                        If giZVT_Fehler > 0 Then
                                
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                            Text1(2).Text = ""
                            
                            anzeige "rot1", "Fehler am Kartenterminal!", Label9
                            
                            Exit Sub
                            'Abbruch
                        End If
                        
                    Else
        
                        'Zahlung
                        sCent = Text1(2).Text
                        sCent = SwapStr(sCent, ",", "")
                        
                        Zahlung_ZVT sCent
                        
                        If giZVT_Fehler > 0 Then
                                
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                            Text1(2).Text = ""
                            
                            anzeige "rot1", "Fehler am Kartenterminal!", Label9

                            Exit Sub
                            'Abbruch
                        End If
                            
                    End If
                    
                Case "ZV2"
                
                    anzeige "rot1", "Bedienen Sie jetzt das Kartenterminal!", Label9
                        
                    If CDbl(Text1(2).Text) < 0 Then
                        'Storno
                        sCent = Text1(2).Text
                        sCent = SwapStr(sCent, ",", "")
                        sCent = SwapStr(sCent, "-", "")
                        
                        dlgTaNr.Show 1
                    
                        sTraceNr = dlgTaNr.Back
'                        sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "BNr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
                               
                        Storno_ZVT2 sTraceNr, sCent, False
                        
                        If giZVT2_Fehler > 0 Then
                                
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                            Text1(2).Text = ""
                            
                            anzeige "rot1", "Fehler am Kartenterminal!", Label9
                            
                            Exit Sub
                            'Abbruch
                        End If
                        
                    Else
        
                        'Zahlung
                        sCent = Text1(2).Text
                        sCent = SwapStr(sCent, ",", "")
                        
                        Zahlung_ZVT2 sCent, False
                        
                        If giZVT2_Fehler > 0 Then
                                
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                            Text1(2).Text = ""
                            
                            anzeige "rot1", "Fehler am Kartenterminal!", Label9

                            Exit Sub
                            'Abbruch
                        End If
                            
                    End If
            End Select
        End If
        anzeige "normal", "", Label9
        
        
        'Ende Kartenterminal
        
        
        
        
    ElseIf iWelchekarte = 2 Then
    
        If InStr(Text1(7).Text, ",") = 0 Then
            anzeige "rot", "Geldbeträge mit Komma eingeben!", Label9
            Text1(7).SetFocus
            Exit Sub
        Else
            'nach dem Komma 2 stellen?
            lPos = InStr(1, Text1(7).Text, ",")
            cNachkomma = Right(Text1(7).Text, Len(Text1(7).Text) - lPos)
            If Len(cNachkomma) <> 2 Then
                anzeige "rot", "Geldbeträge mit 2 Nachkommastellen eingeben!", Label9
                Text1(7).SetFocus
                Exit Sub
            End If
        End If
    
        Select Case index
            Case 0
                gcKreditKarte2 = "VI"
            Case 1
                gcKreditKarte2 = "EU"
            Case 2
                gcKreditKarte2 = "AE"
            Case 3
                gcKreditKarte2 = "DC"
            Case 4
                gcKreditKarte2 = "BC"
            Case 5
                gcKreditKarte2 = "EC"
            Case 6
                gcKreditKarte2 = "SO"
            Case 7
                gcKreditKarte2 = "Automatisch"
        End Select
        
        
        Label33(17).Caption = "2. Karte (" & gcKreditKarte2 & ")"
        Label33(17).Refresh
        
        SSCommand6(20).Picture = SSCommand8(index).Picture
        SSCommand6(20).Caption = ""
        
        'Anfang Kartenterminal
        
        If gbEcash Then
            Select Case gsEPartner
'                Case Is = "ADT"
'                    If gsAdtVerfahren = "XML" Then
'                        anzeige "rot1", "Bedienen Sie jetzt das Kartenterminal!", Label9
'
'                        If CDbl(Text1(7).Text) < 0 Then
'                            If CInt(gADTclientId) = 0 Then
'                                'Storno
'                                sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "TA Nr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
'                                lRet = Shell("C:\Programme\EL-ME\SECpos\SECposPay\SECposPay.exe", vbHide) 'secpos
'                                AppActivate lRet ' True
'
'                                'erstmal zum storno navigieren
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'
'                                SendKeys "{Down}", True
'
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'                                SendKeys "{TAB}", True
'
'                                SendKeys "{Down}", True
'
'                                SendKeys sTraceNr, True
'                                SendKeys "{enter}", True
'
'                                Call keybd_event(VK_LWIN, 0, 0, 0)
'                                Call keybd_event(77, 0, 0, 0)
'                                Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
'
'                                iRet = MsgBox("Storno ok? ", vbInformation + vbYesNo, "Winkiss Frage:")
'                                Pause 5
'                                 'secpospay suchen finden und schließen
'                                Y = "SECpos Pay"
'                                hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
'                                Do
'                                    result = GetWindowTextLength(hwnd) + 1
'                                    Title = Space(result)
'                                    result = GetWindowText(hwnd, Title, result)
'                                    Title = Left$(Title, Len(Title) - 1)
'
'                                    If InStr(1, Title, Y) Then
'                                        SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
'                                    End If
'                                    hwnd = GetWindow(hwnd, GW_HWNDNEXT)
'                                Loop Until hwnd = 0
'
'                                setzedrucker gcBonDrucker
'
'                                If iRet = vbNo Then
'                                    Screen.MousePointer = 0
'                                    Text1(7).Text = ""
'                                    Exit Sub
'                                End If
'                            Else
'                                'Storno neuer Weg
'
'                                If CInt(gADTclientId) > 0 Then
'                                    If gADTipAdress <> "" And gADTport <> "" Then
'                                        '192.168.1.14 '20002
'                                        lRet = ELMESettings(vbNullString, gADTclientId, vbNullString, gADTipAdress, gADTport, -1, -1, -1, vbNullString)
'                                    Else
'                                        lRet = ELMESettings(vbNullString, gADTclientId, vbNullString, vbNullString, -1, -1, -1, -1, vbNullString)
'                                    End If
'                                End If
'
'                                sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "TA Nr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
'                                lRet = ELMEReversal(CLng(sTraceNr))
'                                If lRet = 0 Then
'                                    lRet = ELMEGetPrint(sBLZ, 8000)
'                                    If lRet = 0 Then
'                                        gsAdtBeleg = sBLZ
'                                        If gADTipAdress <> "" And gADTport <> "" Then
'                                            SendeDaten2DruckerECCASH_Kundenbeleg
'                                            SendeDaten2DruckerECCASH_Haendlerbeleg
'                                        Else
'                                            SendeDaten2DruckerECCASH
'                                        End If
'                                    Else
'                                        MsgBox "Fehler ELMEGetPrint: " & lRet, vbCritical, "Winkiss Fehler:"
'                                        gsAdtBeleg = ""
'                                    End If
'                                Else
'                                    lRet = ELMEGetLastError(lerrCode, serrMeldung, 8000)
'                                    If lRet = 0 Then
'                                        MsgBox serrMeldung, vbCritical, "Winkiss Fehler:"
'                                    Else
'                                        MsgBox "Fehler ELMEGetLastError: " & lRet, vbCritical, "Winkiss Fehler:"
'                                    End If
'                                    'Abbruch
'                                    Screen.MousePointer = 0
'                                    Text1(7).Text = ""
'                                    Exit Sub
'                                    'Abbruch
'                                End If
'                            End If
'                        Else
'                            If CInt(gADTclientId) = 0 Then
'                                'Zahlung
'                                lRet = Shell("C:\Programme\EL-ME\SECpos\SECposPay\SECposPay.exe", vbHide) 'secpos
'                                AppActivate lRet
'                                SendKeys Text1(7).Text, True
'                                SendKeys "{enter}", True
'                                Call keybd_event(VK_LWIN, 0, 0, 0)
'                                Call keybd_event(77, 0, 0, 0)
'                                Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
'                                iRet = MsgBox("Zahlung ok? ", vbInformation + vbYesNo, "Winkiss Frage:")
'                                Pause 5
'                                'secpospay suchen finden und schließen
'                                Y = "SECpos Pay" '  (Terminal-ID: " & gsTerminalid & ")"
'                                hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
'                                Do
'                                    result = GetWindowTextLength(hwnd) + 1
'                                    Title = Space(result)
'                                    result = GetWindowText(hwnd, Title, result)
'                                    Title = Left$(Title, Len(Title) - 1)
'                                    If InStr(1, Title, Y) Then
'                                        SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
'                                    End If
'                                    hwnd = GetWindow(hwnd, GW_HWNDNEXT)
'                                Loop Until hwnd = 0
'                                setzedrucker gcBonDrucker
'                                If iRet = vbNo Then
'                                    'abbruch
'                                    Screen.MousePointer = 0
'                                        Text1(7).Text = ""
'
'                                    Exit Sub
'                                End If
'                            Else
'                                'Zahlung der neue Weg
'
'                                If CInt(gADTclientId) > 0 Then
'                                    If gADTipAdress <> "" And gADTport <> "" Then
'                                        '192.168.1.14 '20002
'                                        lRet = ELMESettings(vbNullString, gADTclientId, vbNullString, gADTipAdress, gADTport, -1, -1, -1, vbNullString)
'                                    Else
'                                        lRet = ELMESettings(vbNullString, gADTclientId, vbNullString, vbNullString, -1, -1, -1, -1, vbNullString)
'                                    End If
'                                End If
'
'                                lForcePIN = 0
'
'                                sCent = Text1(7).Text
'                                sCent = SwapStr(sCent, ",", "")
'                                lCents = CLng(sCent)
'                                lret1 = ELMEPay(lCents, lForcePIN)
'
'                                If lret1 = 0 Then
'                                    lret1 = ELMEGetPrint(sBLZ, 2000)
'                                    If lret1 = 0 Then
'                                        gsAdtBeleg = sBLZ
''                                        SendeDaten2DruckerECCASH
'                                        If gADTipAdress <> "" And gADTport <> "" Then
'                                            SendeDaten2DruckerECCASH_Kundenbeleg
'                                            SendeDaten2DruckerECCASH_Haendlerbeleg
'                                        Else
'                                            SendeDaten2DruckerECCASH
'                                        End If
'                                    Else
'                                        MsgBox "Fehler ELMEGetPrint: " & lret1, vbCritical, "Winkiss Fehler:"
'                                        gsAdtBeleg = ""
'                                    End If
'                                Else
'                                    lret1 = ELMEGetLastError(lerrCode, serrMeldung, 8000)
'                                    If lret1 = 0 Then
'                                        MsgBox serrMeldung, vbCritical, "Winkiss Fehler:"
'                                    Else
'                                        MsgBox "Fehler ELMEGetLastError: " & lret1, vbCritical, "Winkiss Fehler:"
'                                    End If
'                                    'Abbruch
'                                    Screen.MousePointer = 0
'                                    Text1(7).Text = ""
'                                    Exit Sub
'                                    'Abbruch
'                                End If
'                            End If
'                        End If
'                ElseIf gsAdtVerfahren = "INOUT" Then
'
'                End If
                
                Case "ELP"
                
                    anzeige "rot1", "Bedienen Sie jetzt das Kartenterminal!", Label9
                        
                    If CDbl(Text1(7).Text) < 0 Then
                        'Storno
                        sCent = Text1(7).Text
                        sCent = SwapStr(sCent, ",", "")
                        sCent = SwapStr(sCent, "-", "")
                        
                        sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "TA Nr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
                               
                        Storno_elPAY sTraceNr, sCent
                        
                        If giELPAY_Fehler > 0 Then
                                
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                            Text1(7).Text = ""
                            
                            anzeige "rot1", "Fehler am Kartenterminal!", Label9

                            Exit Sub
                            'Abbruch
                        End If
                    Else
        
                        'Zahlung
                        sCent = Text1(7).Text
                        sCent = SwapStr(sCent, ",", "")
                        
                        Zahlung_elPAY sCent
                        
                        If giELPAY_Fehler > 0 Then
                                
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                            Text1(7).Text = ""
                            
                            anzeige "rot1", "Fehler am Kartenterminal!", Label9

                            Exit Sub
                            'Abbruch
                        End If
                            
                    End If
                    
                    
                    
                Case "ZVT"
                
                    anzeige "rot1", "Bedienen Sie jetzt das Kartenterminal!", Label9
                        
                    If CDbl(Text1(7).Text) < 0 Then
                        'Storno
                        sCent = Text1(7).Text
                        sCent = SwapStr(sCent, ",", "")
                        sCent = SwapStr(sCent, "-", "")
                        
                        dlgTaNr.Show 1
                    
                        sTraceNr = dlgTaNr.Back
                        
'                        sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "BNr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
                               
                        Storno_ZVT sTraceNr
                        
                        If giZVT_Fehler > 0 Then
                                
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                            Text1(7).Text = ""
                            
                            anzeige "rot1", "Fehler am Kartenterminal!", Label9

                            Exit Sub
                            'Abbruch
                        End If
                    Else
        
                        'Zahlung
                        sCent = Text1(7).Text
                        sCent = SwapStr(sCent, ",", "")
                        
                        Zahlung_ZVT sCent
                        
                        If giZVT_Fehler > 0 Then
                                
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                            Text1(7).Text = ""
                            
                            anzeige "rot1", "Fehler am Kartenterminal!", Label9

                            Exit Sub
                            'Abbruch
                        End If
                            
                    End If
                    
                Case "ZV2"
                
                    anzeige "rot1", "Bedienen Sie jetzt das Kartenterminal!", Label9
                        
                    If CDbl(Text1(7).Text) < 0 Then
                        'Storno
                        sCent = Text1(7).Text
                        sCent = SwapStr(sCent, ",", "")
                        sCent = SwapStr(sCent, "-", "")
                        
                        dlgTaNr.Show 1
                    
                        sTraceNr = dlgTaNr.Back
                        
'                        sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "BNr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
                               
                        Storno_ZVT2 sTraceNr, sCent, False
                        
                        If giZVT2_Fehler > 0 Then
                                
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                            Text1(7).Text = ""
                            
                            anzeige "rot1", "Fehler am Kartenterminal!", Label9

                            Exit Sub
                            'Abbruch
                        End If
                    Else
        
                        'Zahlung
                        sCent = Text1(7).Text
                        sCent = SwapStr(sCent, ",", "")
                        
                        Zahlung_ZVT2 sCent, False
                        
                        If giZVT2_Fehler > 0 Then
                                
                            'Abbruch, so geht Abbruch
                            Screen.MousePointer = 0
                            Text1(7).Text = ""
                            
                            anzeige "rot1", "Fehler am Kartenterminal!", Label9

                            Exit Sub
                            'Abbruch
                        End If
                            
                    End If
            End Select
        End If
        anzeige "normal", "", Label9
        
        
        'Ende Kartenterminal
        
    
    End If
    
    Frame2.Visible = False

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand8_Click"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_Change(index As Integer)
On Error GoTo LOKAL_ERROR

    If index = 0 Or index = 1 Or index = 2 Or index = 3 Or index = 5 Or index = 6 Or index = 7 Then
    
        If index = 3 Then 'Dukaten
            'Neu
            Label1(28).Visible = False
            Label1(28).Caption = "0"
            Label1(29).Visible = False
            Label33(22).Visible = False
            SSCommand6(25).Visible = False
            'Ende Neu
            
        End If

        If Offenrechnen(index) Then

        End If
        
        
        
        

    End If
    
    
    
    If index = 2 Then
        If IsNumeric(Text1(2).Text) Then
            If CDbl(Text1(2).Text) <> 0 Then
                iWelchekarte = 1
                Label33(5).Caption = "1. Karte"
                Label6(2).Caption = "1. Kreditkarte auswählen"
                Frame2.BackColor = glH2
                Frame2.Visible = True
                Frame18.Visible = False
                Frame1.Visible = False
                
                ShowDie6ButtonsOrJustOneButton
                
                
'                If gbEcash Then
'                    Select Case gsEPartner
'                        Case "ELP"
'                            Check15.Visible = True
'                        Case "ZV2"
'                            Check15.Visible = True
'                    End Select
'                End If
                
                
                
            Else
                back2 1
            End If
        Else
            back2 1
        End If
    End If
    
    If index = 7 Then
        If IsNumeric(Text1(7).Text) Then
            If CDbl(Text1(7).Text) <> 0 Then
                iWelchekarte = 2
                Label33(17).Caption = "2. Karte"
                Label6(2).Caption = "2. Kreditkarte auswählen"
                Frame2.BackColor = glH2
                Frame2.Visible = True
                Frame18.Visible = False
                Frame1.Visible = False
                
                
                ShowDie6ButtonsOrJustOneButton
                
                
                
'                If gbEcash Then
'                    Select Case gsEPartner
'                        Case "ELP"
'                            Check15.Visible = True
'                        Case "ZV2"
'                            Check15.Visible = True
'                    End Select
'                End If
            Else
                back2 2
            End If
        Else
            back2 2
        End If
    End If
    

        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub checkgutschScan(Scanstring As TextBox)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim lWert As Long
    
    If Len(Scanstring.Text) = 8 Then
        ctmp = Mid(Scanstring.Text, 2, 6)
    Else
        ctmp = Scanstring.Text
    End If
    
    Scanstring.Text = ctmp
      
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "checkgutschScan"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub back2(ind As Integer)
    On Error GoTo LOKAL_ERROR
    
    If ind = 1 Then
        anzeige "normal", "", Label9
        SSCommand6(14).Caption = "C"
        SSCommand6(14).Picture = Nothing
        Label33(5).Caption = "1. Karte"
        Frame2.Visible = False
    ElseIf ind = 2 Then
        anzeige "normal", "", Label9
        Label33(17).Caption = "2. Karte"
        SSCommand6(20).Caption = "C"
        SSCommand6(20).Picture = Nothing
        Frame2.Visible = False
    ElseIf ind = 3 Then
'        Frame18.Visible = False
'        If MSComm1.PortOpen = True Then
'            MSComm1.PortOpen = False
'        End If
'        Text1(5).Text = ""
'        anzeige "normal", "", Label9
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "back2"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(index).BackColor = glSelBack1
    Label1(1).Caption = index
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    Select Case index
        Case 0 'bargeld
            cValid = "1234567890," & Chr$(8)
        Case 1 '
            cValid = "1234567890," & Chr$(8)
        Case 3 'dukaten
            cValid = "1234567890," & Chr$(8)
        Case 5 'ec last
            cValid = "1234567890," & Chr$(8)
        Case 6 'scheck
            cValid = "1234567890," & Chr$(8)
        Case 4 'Gutscheinnumer eingabe
            cValid = "1234567890" & Chr$(8)
        Case 2 'Kreditkarte
            cValid = "1234567890," & Chr$(8)
        Case 7  'Kreditkarte
            cValid = "1234567890," & Chr$(8)
        Case 8  'alter Gutschein Wert
            cValid = "1234567890," & Chr$(8)
    End Select
    
   
    cZeichen = Chr$(KeyAscii)
'    cZeichen = UCase$(cZeichen)
    
     If cZeichen = "," Then
        If InStr(Text1(index).Text, ",") > 0 Then
            KeyAscii = 0
            
        Else
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            Else
                KeyAscii = Asc(cZeichen)
            End If
        End If
    
    
    End If

    
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    
    If KeyCode = vbKeyReturn Then
        Select Case index
            Case 4
                SSCommand6_Click 18
       End Select
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Function DruckeKassenBonECLastWKL68()
    On Error GoTo LOKAL_ERROR

    Dim iRet            As Integer

    DruckeKassenBonECLastWKL68 = 0
    
    SendeDaten2DruckerLastSchriftWKL68

    Pause (1)
    
    iRet = MsgBox("Wurden die Kassenbons korrekt ausgedruckt?", vbQuestion + vbYesNo, "DRUCK OK?")
    If iRet <> vbYes Then
        DruckeKassenBonECLastWKL68 = 1
    Else

    End If
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeKassenBonECLastWKL68"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
'aufgeräumt
Private Sub SendeDaten2DruckerLastSchriftWKL68()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz        As Long
    Dim lAktSatz        As Long
    Dim lRet            As Long
    Dim lcount          As Long
    Dim lAnzZeile       As Long
    Dim lAnzLbSatz      As Long
    Dim lTask           As Long
    Dim cLBSatz         As String
    Dim cFeld           As String
    Dim cDaten          As String
    Dim ctmp            As String
    Dim cTmp2           As String
    Dim cMWST           As String
    Dim cText           As String
    Dim aDeviceName     As String
    Dim cEscapeSequenz  As String
    Dim cArtNr          As String
    
    Dim dGRabatt        As Double
    Dim dGRabattWert    As Double
    Dim dSumme          As Double
    Dim dWert           As Double
    Dim dEuro           As Double
    Dim dMWStVoll       As Double
    Dim dMWStErm        As Double
    Dim dAktZeit        As Double
    Dim dNeuZeit        As Double
    Dim iFileNr         As Integer
    Dim iLenZeile       As Integer
    Dim iLevel          As Integer
    Dim iAktCopy        As Integer
    ReDim cDruckZeile(1 To 1) As String
    
    iLevel = 0
    setzedrucker gcBonDrucker
    'Drucker an, Display aus, Init Drucker
    aDeviceName = Printer.DeviceName
    cEscapeSequenz = gcInit
    OpenDrawer aDeviceName, cEscapeSequenz
    DoEvents
    
StartPunkt:
    lAnzZeile = 0
    ReDim cDruckZeile(1 To 1) As String
    
    iAktCopy = iAktCopy + 1
    iLevel = 1
    cDaten = ""
    iLenZeile = 32
    
    '********************************************************
    '* Hier geht's los                                      *
    '********************************************************
    
    'Lastschrifttext an Drucker senden
    lAnzSatz = List11.ListCount
    
    iLevel = 2
    
    'Drucker wird auf BonDrucker geschaltet
    aDeviceName = gcBonDrucker
    
    iLevel = 3
    
    '***********************************************
    'Drucker ein- und Kundendisplay ausschalten
    '***********************************************
    
    cEscapeSequenz = gcInit
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    'Kopfdaten an Drucker senden
    
    If gcBild <> "" Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcBild
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '**********************************************************
    '* 1.Kopfzeile
    '**********************************************************
    
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
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '**********************************************************
    '* 2.Kopfzeile
    '**********************************************************
    
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
    
    '**********************************************************
    '* 3.Kopfzeile
    '**********************************************************
    
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
    
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    iLevel = 4
    
    
    
    dSumme = 0
    For lAktSatz = 0 To lAnzSatz - 1
    
        cLBSatz = List11.list(lAktSatz)
        
        cDaten = cLBSatz
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    Next lAktSatz
    
    iLevel = 5
    
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    iLevel = 6
    
    iLevel = 7
    
    '******************************************************
    '* 1.Fußzeile
    '******************************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "KEIN GÜLTIGER KASSENBON!"
    Else
        cDaten = gcBonText(2)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    
    '******************************************************
    '* 2.Fußzeile
    '******************************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(3)
    End If
    
    If Trim$(cDaten) <> "" Then
    
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    
    '******************************************************
    '* 3.Fußzeile
    '******************************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = ""
    Else
        cDaten = gcBonText(5)
    End If
    
    If Trim$(cDaten) <> "" Then
    
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    
    iLevel = 10
    
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
    
BON_DRUCKEN:
    
    If gbAPI = True Then
        OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
        DoEvents
    Else
        OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
        DoEvents
    End If
    
    'Bon-Daten sichern
    SichernBonDaten cDruckZeile(), lAnzZeile, "", "", False

    
BON_SCHNEIDEN:

    'Kassenbon abschneiden
    If gbAPI = True Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcSchneiden
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
    iLevel = 11
    
    'Druckbereich freigeben
    
    Erase cDruckZeile
    DoEvents
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SendeDaten2DruckerLastSchriftWKL68"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text6_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text6(index).BackColor = glSelBack1
    Text6(index).SelStart = 0
    Text6(index).SelLength = Len(Text6(index).Text)
    Label12.Caption = Trim$(Str$(index))
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text6_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text6_LostFocus(index As Integer)
    On Error GoTo LOKAL_ERROR

    Text6(index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text6_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text6_KeyPress(index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr$(8)
    
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
    If KeyAscii <> 0 And KeyAscii <> 8 Then
        If index < 3 Then
            If Len(Text6(index).Text) = Text6(index).MaxLength - 1 Then
                Text6(index + 1).SetFocus
            End If
        Else
            If Len(Text6(index).Text) = Text6(index).MaxLength - 1 Then
                Command12(0).SetFocus
            End If
        End If
    End If
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text6_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Barcode(cGutschnr As String)
On Error GoTo LOKAL_ERROR

    Dim lcount                  As Long
    Dim cZeichen                As String
    Dim cArtNr                  As String

    Dim aDeviceName             As String
    Dim cEscapeSequenz          As String
    Dim lPruefZiffer            As Long

    aDeviceName = Printer.DeviceName
    cEscapeSequenz = vbCrLf
    OpenDrawer aDeviceName, cEscapeSequenz

    cEscapeSequenz = Chr(27) & Chr(97) & Chr(1)
    OpenDrawer aDeviceName, cEscapeSequenz

    cEscapeSequenz = Chr(29) & Chr(72) & Chr(2)
    OpenDrawer aDeviceName, cEscapeSequenz
    
    'die Barcodehöhe
    
'    cEscapeSequenz = Chr(29) & Chr(104) & Chr(164)
'    OpenDrawer aDeviceName, cEscapeSequenz
    
    cEscapeSequenz = Chr(29) & Chr(104) & Chr(40)
    OpenDrawer aDeviceName, cEscapeSequenz
    
    cEscapeSequenz = Chr(29) & Chr(119) & Chr(3)
    OpenDrawer aDeviceName, cEscapeSequenz
    

    If Len(cGutschnr) = 1 Then
        cGutschnr = "0000000000" & cGutschnr
    ElseIf Len(cGutschnr) = 2 Then
        cGutschnr = "000000000" & cGutschnr
    ElseIf Len(cGutschnr) = 3 Then
        cGutschnr = "00000000" & cGutschnr
    ElseIf Len(cGutschnr) = 4 Then
        cGutschnr = "0000000" & cGutschnr
    ElseIf Len(cGutschnr) = 5 Then
        cGutschnr = "000000" & cGutschnr
    ElseIf Len(cGutschnr) = 6 Then
        cGutschnr = "00000" & cGutschnr
    ElseIf Len(cGutschnr) = 7 Then
        cGutschnr = "0000" & cGutschnr
    ElseIf Len(cGutschnr) = 8 Then
        cGutschnr = "000" & cGutschnr
    ElseIf Len(cGutschnr) = 9 Then
        cGutschnr = "00" & cGutschnr
    End If
    
    cArtNr = "2" & cGutschnr
'    cArtnr = "2" & "9" & cVorgang & "0" & cSumme
    
    Dim p1 As Integer
    Dim p2 As Integer
    Dim p3 As Integer
    Dim p4 As Integer
    Dim p5 As Integer
    Dim p6 As Integer
    Dim p7 As Integer
    Dim p8 As Integer
    Dim p9 As Integer
    Dim p10 As Integer
    Dim p11 As Integer
    Dim p12 As Integer
    Dim p13 As Integer
    
    Dim rest As Double
    Dim pz As Long
    
    
    p1 = Val(Mid(cArtNr, 1, 1)) * 1
    p2 = Val(Mid(cArtNr, 2, 1)) * 3
    p3 = Val(Mid(cArtNr, 3, 1)) * 1
    p4 = Val(Mid(cArtNr, 4, 1)) * 3
    p5 = Val(Mid(cArtNr, 5, 1)) * 1
    p6 = Val(Mid(cArtNr, 6, 1)) * 3
    p7 = Val(Mid(cArtNr, 7, 1)) * 1
    p8 = Val(Mid(cArtNr, 8, 1)) * 3
    p9 = Val(Mid(cArtNr, 9, 1)) * 1
    p10 = Val(Mid(cArtNr, 10, 1)) * 3
    p11 = Val(Mid(cArtNr, 11, 1)) * 1
    p12 = Val(Mid(cArtNr, 12, 1)) * 3
    p13 = p1 + p2 + p3 + p4 + p5 + p6 + p7 + p8 + p9 + p10 + p11 + p12
    
    rest = p13 Mod 10
    pz = 10 - rest
    If rest = 0 Then
        pz = 0
    End If
    
    cArtNr = cArtNr & Trim$(Str$(pz))
    
    cEscapeSequenz = gcBarCode & cArtNr
    OpenDrawer aDeviceName, cEscapeSequenz
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Barcode"
    Fehler.gsFehlertext = "Im Programmteil BEZAHLEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub ShowDie6ButtonsOrJustOneButton()

  If gbEcash Then
                
                  If gsEPartner = "ZV2" Then
                    
                            SSCommand8(0).Visible = False
                            SSCommand8(1).Visible = False
                            SSCommand8(2).Visible = False
                            SSCommand8(3).Visible = False
                            SSCommand8(5).Visible = False
                            SSCommand8(6).Visible = False
                            
                            SSCommand8(7).Visible = True
                        Else
                            SSCommand8(0).Visible = True
                            SSCommand8(1).Visible = True
                            SSCommand8(2).Visible = True
                            SSCommand8(3).Visible = True
                            SSCommand8(5).Visible = True
                            SSCommand8(6).Visible = True
                            
                            SSCommand8(7).Visible = False
                        
                  End If
   End If
End Sub

Private Sub DruckeGutscheinBonWKL68(cLBSatz As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzZeile       As Long
    Dim lcount          As Long
    Dim cDaten          As String
    Dim cEscapeSequenz  As String
    Dim aDeviceName     As String
    Dim ctmp            As String
    Dim cGPreis         As String
    
    Dim iStufe          As Integer
    Dim iLenZeile       As Integer
    ReDim cDruckZeile(1 To 1) As String
    
    If Not gbBonDruck Then
        GoTo ENDE
    End If
    
    iStufe = 0
    
    cGPreis = Mid(cLBSatz, 60, 9)
    ctmp = Mid(cLBSatz, 24, 8)
    ctmp = Trim(ctmp)
    
    iLenZeile = 32
    'Drucker ist bereits auf BonDrucker geschaltet
    aDeviceName = gcBonDrucker
    
    
    '***********************************************
    'ggf. Logo auf Kassenbon bringen
    '***********************************************

    If gcBild <> "" Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcBild
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
    cEscapeSequenz = gcInit
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    

    Barcode_Gutschein Trim(ctmp)

    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    

    iStufe = 1
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
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        iStufe = 2
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
        iStufe = 3
    End If
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION"
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
        iStufe = 3
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
    iStufe = 4
    
    '******************************************************************
    'hier
    cDaten = "G U T S C H E I N V E R K A U F"
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    iStufe = 5
    '******************************************************************
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = " " & vbCrLf
        
    iStufe = 6
    '******************************************************************
        
    cDaten = "Wert Gutscheins:    " & gcWaehrung & cGPreis
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    iStufe = 7
    '******************************************************************
    
    
    cDaten = "Nummer Gutschein:" & Space(15 - Len(ctmp)) & ctmp
    
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    iStufe = 8
    '******************************************************************
    gcFilNr = Trim(gcFilNr)
    cDaten = "Nummer der Filiale:" & Space(13 - Len(gcFilNr)) & gcFilNr
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    ctmp = "Kasse:                         " & gcKasNum
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    iStufe = 9
    '******************************************************************
    ctmp = Trim$(frmWKL20.Text1(0).Text)
    ctmp = Trim$(ctmp)
    ctmp = Space$(3 - Len(ctmp)) & ctmp
    cDaten = "Bedienernummer:              " & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    iStufe = 10
    '******************************************************************
    ctmp = Trim$(Str$(gdBonNr))
    ctmp = Trim$(ctmp)
    ctmp = gcKasNum & "/" & ctmp
    ctmp = Space$(10 - Len(ctmp)) & ctmp
    
    cDaten = "Belegnummer:          " & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    iStufe = 11
    '******************************************************************
    ctmp = Format$(Now, "DD.MM.YYYY")
    cDaten = "Datum:                " & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    iStufe = 12
    '******************************************************************
    
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    If gbBonGu2J Then
        cDaten = "Gültigkeitsdauer 4 Jahre"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        cDaten = "nach Ausstellungsdatum"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    
    iStufe = 13
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "KEIN GÜLTIGER KASSENBON!"
    Else
        cDaten = gcBonText(2)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        iStufe = 14
    End If
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(3)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        iStufe = 15
    End If
    '******************************************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = ""
    Else
        cDaten = gcBonText(5)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        iStufe = 15
    End If
    '******************************************************************
    
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
    
    iStufe = 16
    
    If gbAPI = True Then
        OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
    Else
        OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
    End If
    gdSumme = cGPreis
    SichernBonDaten cDruckZeile(), lAnzZeile, "", "", False, True
    
    Erase cDruckZeile
    iStufe = 17
    
'BON_SCHNEIDEN:
    'Kassenbon abschneiden
'    If gbAPI Then
'        aDeviceName = Printer.DeviceName
'        cEscapeSequenz = gcSchneiden
'        OpenDrawer aDeviceName, cEscapeSequenz
'    End If
    iStufe = 18
    
ENDE:

    '...und tschüß!
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeGutscheinBonWKL68"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
