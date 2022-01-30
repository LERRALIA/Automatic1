VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKLal 
   BackColor       =   &H00C0C000&
   Caption         =   "Gutscheinverwaltung"
   ClientHeight    =   8595
   ClientLeft      =   1875
   ClientTop       =   2115
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
   Icon            =   "frmWKLal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   3
      Left            =   7080
      TabIndex        =   53
      Top             =   7800
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Caption         =   "neue Strichcodes"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Height          =   4455
      Left            =   8160
      TabIndex        =   40
      Top             =   2640
      Width           =   3495
      Begin VB.Label Label6 
         Appearance      =   0  '2D
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   5
         Left            =   2805
         TabIndex        =   61
         Tag             =   "Shape"
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label6 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   2
         Left            =   2400
         TabIndex        =   60
         Tag             =   "Shape"
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label6 
         Appearance      =   0  '2D
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   4
         Left            =   1725
         TabIndex        =   59
         Tag             =   "Shape"
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label6 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   1
         Left            =   1320
         TabIndex        =   58
         Tag             =   "Shape"
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label6 
         Appearance      =   0  '2D
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   3
         Left            =   645
         TabIndex        =   57
         Tag             =   "Shape"
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label6 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   0
         Left            =   240
         TabIndex        =   56
         Tag             =   "Shape"
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label5 
         Appearance      =   0  '2D
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   55
         Tag             =   "Shape"
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label Label5 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   54
         Tag             =   "Shape"
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblUeberschrift 
         BackStyle       =   0  'Transparent
         Caption         =   "Stückzahlen im Vergleich"
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
         Index           =   1
         Left            =   240
         TabIndex        =   52
         Top             =   120
         Width           =   3135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   240
         X2              =   3240
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "eingelöste Gutscheine"
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
         Left            =   600
         TabIndex        =   51
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "verkaufte Gutscheine"
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
         TabIndex        =   50
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   49
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   48
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   47
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Label3"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   46
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Label3"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   45
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Label3"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   44
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   43
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   42
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   41
         Top             =   2400
         Width           =   315
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9480
      TabIndex        =   29
      Top             =   7800
      Width           =   2175
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame3"
      Height          =   5295
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   7815
      Begin VB.Frame Frame6 
         BackColor       =   &H0080FF80&
         Height          =   3135
         Left            =   -1320
         TabIndex        =   86
         Top             =   840
         Visible         =   0   'False
         Width           =   4815
         Begin VB.TextBox Text1 
            Height          =   2235
            Index           =   5
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   89
            Top             =   480
            Width           =   4575
         End
         Begin sevCommand3.Command Command1 
            Height          =   255
            Index           =   10
            Left            =   4320
            TabIndex        =   87
            Top             =   120
            Width           =   360
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
            Caption         =   "x"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "Notizen"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   88
            Top             =   120
            Width           =   1815
         End
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   9
         Left            =   4200
         TabIndex        =   85
         Top             =   4560
         Width           =   1185
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "Notizen"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   6
         Left            =   6375
         TabIndex        =   84
         Top             =   4560
         Width           =   1320
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "Details"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H008080FF&
         Height          =   3135
         Left            =   2520
         TabIndex        =   66
         Top             =   960
         Visible         =   0   'False
         Width           =   4815
         Begin sevCommand3.Command Command1 
            Height          =   255
            Index           =   8
            Left            =   4320
            TabIndex        =   82
            ToolTipText     =   "Bonansicht"
            Top             =   2040
            Width           =   360
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
            Caption         =   "F2"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command1 
            Height          =   255
            Index           =   7
            Left            =   4320
            TabIndex        =   67
            Top             =   2760
            Width           =   360
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
            Caption         =   "x"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "stehen nur dann zur Verfügung, wenn man die Gutscheine über die 'gemischte Zahlung' annimmt."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   495
            Index           =   19
            Left            =   2040
            TabIndex        =   83
            Top             =   120
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "Filialauswahl"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   11
            Left            =   2040
            TabIndex        =   81
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "Einlösedetails"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   80
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "Einlösefiliale:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   79
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "Filialauswahl"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   7
            Left            =   2040
            TabIndex        =   78
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "an Kasse:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   77
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "Filialauswahl"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   9
            Left            =   2040
            TabIndex        =   76
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "durch Bediener:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   75
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "Filialauswahl"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   615
            Index           =   12
            Left            =   2520
            TabIndex        =   74
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "diesen Bon bezahlt:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   73
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "Filialauswahl"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   14
            Left            =   2040
            TabIndex        =   72
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "Einlösetag:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   71
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "Filialauswahl"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   16
            Left            =   2040
            TabIndex        =   70
            Top             =   2400
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "Einlösezeitpunkt:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   69
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000001&
            Caption         =   "Filialauswahl"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   18
            Left            =   2040
            TabIndex        =   68
            Top             =   2760
            Width           =   2295
         End
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   0
         MultiSelect     =   2  'Erweitert
         TabIndex        =   13
         Top             =   360
         Width           =   7695
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   4560
         Width           =   1185
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   1
         Left            =   2760
         TabIndex        =   17
         Top             =   4560
         Width           =   1425
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "ausbuchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   2
         Left            =   1200
         TabIndex        =   16
         Top             =   4560
         Width           =   1545
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "reaktivieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   2
         Left            =   5400
         TabIndex        =   15
         Top             =   4560
         Width           =   960
         _ExtentX        =   0
         _ExtentY        =   0
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
         Caption         =   "Druck"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   0
         TabIndex        =   14
         Top             =   120
         Width           =   7695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   11415
      Begin VB.CheckBox Check38 
         Caption         =   "nur das aktuelle Jahr anzeigen"
         Height          =   240
         Left            =   360
         TabIndex        =   104
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame7"
         Height          =   615
         Left            =   5520
         TabIndex        =   93
         Top             =   960
         Width           =   1815
         Begin sevCommand3.Command Command4 
            Height          =   300
            Index           =   2
            Left            =   1440
            TabIndex        =   96
            ToolTipText     =   "Leeren"
            Top             =   240
            Width           =   300
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
            Version3        =   -1  'True
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Vormonat"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   95
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "aktueller Monat"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   94
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame7"
         Height          =   615
         Left            =   3600
         TabIndex        =   90
         Top             =   960
         Width           =   1815
         Begin sevCommand3.Command Command4 
            Height          =   300
            Index           =   3
            Left            =   1440
            TabIndex        =   97
            ToolTipText     =   "Leeren"
            Top             =   240
            Width           =   300
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
            Version3        =   -1  'True
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "aktueller Monat"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   92
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Vormonat"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   91
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   9480
         MaxLength       =   13
         TabIndex        =   64
         Top             =   240
         Width           =   1935
      End
      Begin sevCommand3.Command Command1 
         Height          =   255
         Index           =   5
         Left            =   9480
         TabIndex        =   63
         Top             =   1320
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "alte Gutscheine"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   255
         Index           =   4
         Left            =   9480
         TabIndex        =   62
         Top             =   1020
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "Restgutscheine"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   5880
         TabIndex        =   36
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   5880
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   3960
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   7440
         TabIndex        =   5
         Top             =   0
         Width           =   1935
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Wert"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   9
            Top             =   1320
            Width           =   3255
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Einlösedatum"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   8
            Top             =   960
            Width           =   3255
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Verkaufsdatum"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   600
            Width           =   3255
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Gutschein-Nummer"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   3255
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C000&
            Caption         =   "sortiert nach"
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
            Left            =   120
            TabIndex        =   11
            Top             =   0
            Width           =   1455
         End
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   0
         Left            =   9480
         TabIndex        =   4
         Top             =   600
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         Caption         =   "Suche Daten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "alle Gutscheine"
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
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "eingelöste Gutscheine"
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
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "offene Gutscheine"
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
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin sevCommand3.Command Command4 
         Height          =   360
         Index           =   20
         Left            =   5000
         TabIndex        =   98
         ToolTipText     =   "Kalender"
         Top             =   240
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
      Begin sevCommand3.Command Command4 
         Height          =   360
         Index           =   21
         Left            =   5000
         TabIndex        =   99
         ToolTipText     =   "Kalender"
         Top             =   600
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
      Begin sevCommand3.Command Command4 
         Height          =   360
         Index           =   0
         Left            =   6960
         TabIndex        =   100
         ToolTipText     =   "Kalender"
         Top             =   240
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
      Begin sevCommand3.Command Command4 
         Height          =   360
         Index           =   1
         Left            =   6960
         TabIndex        =   101
         ToolTipText     =   "Kalender"
         Top             =   600
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
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "GutscheinNr.:"
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
         Left            =   9480
         TabIndex        =   65
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "bis:"
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
         Left            =   5400
         TabIndex        =   39
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "von:"
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
         Left            =   5400
         TabIndex        =   38
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Einlösedatum"
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
         Left            =   5880
         TabIndex        =   37
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Wert"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   2760
         TabIndex        =   34
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Stück"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   2160
         TabIndex        =   33
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   135
         Index           =   8
         Left            =   2760
         TabIndex        =   32
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   135
         Index           =   7
         Left            =   2760
         TabIndex        =   31
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   135
         Index           =   6
         Left            =   2760
         TabIndex        =   30
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Ausgabedatum"
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
         Left            =   3960
         TabIndex        =   27
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "von:"
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
         Left            =   3480
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "bis:"
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
         Left            =   3480
         TabIndex        =   25
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   135
         Index           =   4
         Left            =   2280
         TabIndex        =   22
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   135
         Index           =   3
         Left            =   2280
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   135
         Index           =   2
         Left            =   2280
         TabIndex        =   20
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "suche"
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
         Index           =   1
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Hilfethemen:"
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
      Index           =   32
      Left            =   9000
      TabIndex        =   103
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Gutscheine mit Strichcode"
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
      Index           =   20
      Left            =   9000
      MouseIcon       =   "frmWKLal.frx":0442
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   102
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblAnzeige 
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
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   28
      Top             =   7800
      Width           =   6735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Gutscheinverwaltung"
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
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmWKLal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iPrueD          As Integer
Dim gitop           As Integer

Private Sub Check38_Click()
    On Error GoTo LOKAL_ERROR
    
    StueckErmittlung
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check38_Click"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Select Case Index
        Case Is = 3
            If FileExists(cPfad & "aWKL30j.rpt") Then 'Spezialreport a5
                lblanzeige.Caption = "Etikettengröße: Spezialdruck für " & gFirma.FirmaName
                lblanzeige.Refresh
            ElseIf FileExists(cPfad & "aWKL30is.rpt") Then 'Spezialreport a4
                lblanzeige.Caption = "Etikettengröße: Spezialdruck für " & gFirma.FirmaName
                lblanzeige.Refresh
            Else
                lblanzeige.Caption = "Etikettengröße: 35,6 mm x 16,9 mm"
                lblanzeige.Refresh
            End If
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_MouseMove"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    If Index = 20 Then
        Label1(20).ForeColor = glLink
    End If
    
    If Index = 14 Then
        Label1(14).ForeColor = glLink
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    Select Case Index
        Case Is = 2
            Option1(3).Value = False
            Option1(4).Value = False
            Text1(2).Text = ""
            Text1(3).Text = ""
            
        Case Is = 3
            Option1(5).Value = False
            Option1(6).Value = False
            Text1(0).Text = ""
            Text1(1).Text = ""
            
        Case Is = 0        ' Kalender
            Text1(2).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            Text1(3).SetFocus
        Case Is = 1        ' Kalender
            Text1(3).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
        Case Is = 20        ' Kalender
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            Text1(1).SetFocus
        Case Is = 21        ' Kalender
            Text1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            'fertig
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    PositionierenWKLal
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift(0)
    
    
    
    If NewTableSuchenDBKombi("EAL", gdBase) Then
        
        voreinstellungladen
    
    End If
    
    
    diagramm
    StueckErmittlung
    
    List1.Clear
    List1.AddItem "GutschNr       Wert   Ausgegeben    Eingelöst   Bed       Kunde (hat Gutschein gekauft)"
    
    
    If gbGutsch Then
        Command1(3).Visible = True
    End If
    
    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim bo0 As Integer
    
    loeschNEW "EAL", gdBase
    CreateTableT2 "EAL", gdBase
    
    
    If Check38.Value = vbChecked Then
        bo0 = 0
    Else
        bo0 = -1
    End If
    
    sSQL = "Insert into EAL ( bo0) values "
    sSQL = sSQL & "(" & bo0 & ")"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungladen()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    
    Set rs = gdBase.OpenRecordset("EAL")
    If Not rs.EOF Then
    
        If rs!bo0 = True Then
            Check38.Value = vbUnchecked
        Else
            Check38.Value = vbChecked
        End If
    End If
    rs.Close: Set rs = Nothing

     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub PositionierenWKLal()
    On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 960
    Frame1.Left = 240
    Frame1.Height = 1575
    Frame1.Width = 11415
    
    Frame3.Top = 2520
    Frame3.Left = 240
    Frame3.Height = 5175
    Frame3.Width = 7815
    
    Frame4.Top = 2640
    Frame4.Left = 8160
    Frame4.Height = 4335
    Frame4.Width = 3495
    
    Frame5.Top = 960
    Frame5.Left = 2520
    Frame5.Height = 3135
    Frame5.Width = 4815
    
    Frame6.Top = 960
    Frame6.Left = 2520
    Frame6.Height = 3135
    Frame6.Width = 4815
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKLal"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub diagramm()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim lDate   As Long
    Dim lVon    As Long
    Dim lBis    As Long
    Dim sDate   As String
    Dim rsrs    As Recordset
    Dim cFeld   As String
    
    
    Dim imon As Integer
    Dim iyear As Integer
    Dim cBis As String
    Dim cVon As String
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim iTop            As Integer
    Dim myarr(0 To 5)   As Long
    Dim iMax            As Integer
    Dim iBuffer         As Integer
    
'    gitop = Shape1(0).Top
    gitop = Label6(0).Top
    
    imon = Month(Now)
    iyear = Year(Now)
    
    For k = 2 To 0 Step -1
        If imon = 1 Then
            imon = 12
            iyear = Year(Now) - 1
        Else
            imon = imon - 1
        End If
        
        If imon = 1 Or imon = 3 Or imon = 5 Or imon = 7 Or imon = 8 Or imon = 10 Or imon = 12 Then
            lBis = 31
        ElseIf imon = 2 Then
            lBis = 28
        Else
            lBis = 30
        End If
        
        cVon = "01." & imon & "." & iyear
        cBis = lBis & "." & imon & "." & iyear
        
        lVon = DateValue(cVon)
        lBis = DateValue(cBis)
        
        
        
        sSQL = "Select count(*) as anzahl from GUTSCH where Status <> 'L' and Dat_Ausg >= " & Trim$(Str$(lVon)) & " And  Dat_Ausg <= " & Trim$(Str$(lBis))
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            myarr(k) = rsrs!ANZAHL
        End If
        rsrs.Close: Set rsrs = Nothing
        
        sSQL = "Select count(*) as anzahl from GUTSCH where Status <> 'L' and Dat_einl >= " & Trim$(Str$(lVon)) & " And  Dat_einl <= " & Trim$(Str$(lBis))
        Set rsrs = gdBase.OpenRecordset(sSQL)
        
        If Not rsrs.EOF Then
            myarr(k + 3) = rsrs!ANZAHL
        End If
        rsrs.Close: Set rsrs = Nothing
        
        Label3(k).Caption = MonthName(imon, True) & " " & Right(iyear, 2)
    Next k
    
    iBuffer = 0
    iMax = 0
    
    For i = 0 To 5
        iBuffer = myarr(i)
        If iBuffer > iMax Then
            iMax = iBuffer
        End If
    Next i
    
        iMax = IIf(iMax = 0, 1, iMax)
    
    For i = 0 To 5
        Label6(i).Top = gitop
        Label6(i).Height = (1900 / iMax) * IIf(myarr(i) < 0, 0, myarr(i))
        Label6(i).Top = gitop - ((1900 / iMax) * myarr(i))
        
        Label10(i).Top = Label6(i).Top - 250
        Label10(i).Caption = myarr(i)
        Label10(i).Refresh
    Next i
    
    

    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "diagramm"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub StueckErmittlung()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    Dim sVon As String
    Dim lVon As Long
    sVon = Format("01.01." & Year(DateValue(Now)), "DD.MM.YY")
    lVon = DateValue(sVon)
    
    
    
    'Stück
    sSQL = "Select count(*) as anzahl from GUTSCH where Status <> 'L' and  (DAT_EINL = 0 or DAT_EINL is NULL) " 'offene
    If Check38.Value = vbChecked Then
        sSQL = sSQL & " and Dat_Ausg >= " & Trim$(Str$(lVon)) & " "
    End If
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        Label2(2).Caption = rsrs!ANZAHL
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Select count(*) as anzahl from GUTSCH where Status <> 'L' and DAT_EINL <> 0 "    'eingelöste
    If Check38.Value = vbChecked Then
        sSQL = sSQL & " and DAT_EINL >= " & Trim$(Str$(lVon)) & " "
    End If
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        Label2(3).Caption = rsrs!ANZAHL
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    sSQL = "Select count(*)  as anzahl from GUTSCH where Status <> 'L' and wert > 0  "  'alle
    
    If Check38.Value = vbChecked Then
        sSQL = sSQL & " and Dat_Ausg >= " & Trim$(Str$(lVon)) & " "
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        Label2(4).Caption = rsrs!ANZAHL
    End If
    rsrs.Close: Set rsrs = Nothing
    
    'Werte
    sSQL = "Select sum(wert) as anzahl from GUTSCH where Status <> 'L' and (DAT_EINL = 0 or DAT_EINL is NULL) " 'offene
    If Check38.Value = vbChecked Then
        sSQL = sSQL & " and Dat_Ausg >= " & Trim$(Str$(lVon)) & " "
    End If
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        Label2(8).Caption = Format$(rsrs!ANZAHL, "#####0.00")
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Select sum(wert) as anzahl from GUTSCH where Status <> 'L' and DAT_EINL <> 0 "    'eingelöste
    If Check38.Value = vbChecked Then
        sSQL = sSQL & " and DAT_EINL >= " & Trim$(Str$(lVon)) & " "
    End If
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        Label2(7).Caption = Format$(rsrs!ANZAHL, "#####0.00")
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Select sum(wert)  as anzahl from GUTSCH where Status <> 'L' "  'alle
    
    If Check38.Value = vbChecked Then
        sSQL = sSQL & " and Dat_Ausg >= " & Trim$(Str$(lVon)) & " "
    End If
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        Label2(6).Caption = Format$(rsrs!ANZAHL, "#####0.00")
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "StueckErmittlung"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheGutscheineWKLal()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim cAusgabedat As String
    Dim cEinldat    As String
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim cOrderBy    As String
    Dim lVon        As Long
    Dim lBis        As Long
    Dim lAnz        As Long
    Dim lEinBis     As Long
    Dim lEinVon     As Long
    Dim dGesWert    As Double
    Dim dWert       As Double
    Dim rsrs        As Recordset
    
    lblanzeige.ForeColor = vbBlue
    lblanzeige.Caption = "Gutscheine werden ermittelt..."
    
    If Text1(0).Text <> "" Then
        cFeld = Text1(0).Text
        lVon = DateValue(cFeld)
    
        If Text1(1).Text <> "" Then
            cFeld = Text1(1).Text
            lBis = DateValue(cFeld)
        Else
            lBis = DateValue(Now)
        End If
        cAusgabedat = " and Dat_Ausg >= " & Trim$(Str$(lVon)) & " "
        cAusgabedat = cAusgabedat & "and Dat_Ausg <= " & Trim$(Str$(lBis)) & " "
    Else
        cAusgabedat = ""
    End If
    
    If Text1(2).Text <> "" Then
        cFeld = Text1(2).Text
        lEinVon = DateValue(cFeld)
    
        If Text1(3).Text <> "" Then
            cFeld = Text1(3).Text
            lEinBis = DateValue(cFeld)
        Else
            lEinBis = DateValue(Now)
        End If
        cEinldat = " and DAT_EINL >= " & Trim$(Str$(lEinVon)) & " "
        cEinldat = cEinldat & "and DAT_EINL <= " & Trim$(Str$(lEinBis)) & " "
    Else
        cEinldat = ""
    End If
    
    If Option2(0).Value = True Then
        cOrderBy = " order by GUTSCHNR"
    ElseIf Option2(1).Value = True Then
        cOrderBy = " order by DAT_AUSG, GUTSCHNR"
    ElseIf Option2(2).Value = True Then
        cOrderBy = " order by DAT_EINL, GUTSCHNR"
    ElseIf Option2(3).Value = True Then
        cOrderBy = " order by WERT, GUTSCHNR"
    Else
        cOrderBy = " order by GUTSCHNR"
    End If
    
    loeschNEW "Guttemp", gdBase
    
'    If Option1(0).Value = True Then
'        cSQL = "Select *  into Guttemp from GUTSCH  where Status <> 'L' and (DAT_EINL = 0 or DAT_EINL is NULL) and gutschnr < 10000000 " & cAusgabedat & cEinldat
'    ElseIf Option1(1).Value = True Then
'        cSQL = "Select * into Guttemp from GUTSCH where Status <> 'L' and DAT_EINL <> 0  and gutschnr < 10000000 " & cAusgabedat & cEinldat
'    Else
'        cSQL = "Select * into Guttemp from GUTSCH where Status <> 'L' and Wert > 0 and gutschnr < 10000000 " & cAusgabedat & cEinldat
'    End If
    
    If Option1(0).Value = True Then
        cSQL = "Select *  into Guttemp from GUTSCH  where Status <> 'L' and (DAT_EINL = 0 or DAT_EINL is NULL)  " & cAusgabedat & cEinldat
    ElseIf Option1(1).Value = True Then
        cSQL = "Select * into Guttemp from GUTSCH where Status <> 'L' and DAT_EINL <> 0   " & cAusgabedat & cEinldat
    Else
        cSQL = "Select * into Guttemp from GUTSCH where Status <> 'L' and Wert > 0  " & cAusgabedat & cEinldat
    End If
    
    
    
    If Text1(4).Text <> "" Then
        cFeld = Trim(Text1(4).Text)
        If IsNumeric(cFeld) Then
        
            If Len(cFeld) = 13 And Left(cFeld, 1) = "2" And gbGutschnrKomplett = True Then
                cFeld = Left(cFeld, 8)
                Text1(4).Text = cFeld
            End If
        
            cSQL = cSQL & " and gutschnr = " & cFeld & " "
        End If
    End If
    
    cSQL = cSQL & cOrderBy
    
    gdBase.Execute cSQL, dbFailOnError
    
    List2.Clear
    
    lAnz = 0
    dGesWert = 0
    
    Set rsrs = gdBase.OpenRecordset("Guttemp", dbOpenTable)
    If Not rsrs.EOF Then
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                lAnz = lAnz + 1
                cLBSatz = ""
                If Not IsNull(rsrs!gutschnr) Then
                    cFeld = rsrs!gutschnr
                Else
                    cFeld = ""
                End If
                cFeld = Trim$(cFeld)
                cFeld = Space$(11 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & "  "
                
                If Not IsNull(rsrs!Wert) Then
                    dWert = rsrs!Wert
                Else
                    dWert = 0
                End If
                
                If dWert > 100000 Then 'löschen
                    cSQL = "Update GUTSCH SET STATUS = 'L' where GUTSCHNR = " & cFeld
                    gdBase.Execute cSQL, dbFailOnError

                    cSQL = "Update GUTSCH SET SYNSTATUS = 'D' where GUTSCHNR = " & cFeld
                    gdBase.Execute cSQL, dbFailOnError
                    rsrs.Close: Set rsrs = Nothing
                    Screen.MousePointer = 0
                    lblanzeige.ForeColor = vbRed
                    lblanzeige.Caption = "Bitte nochmal 'Suche Daten' drücken!"
                    Exit Sub
                End If
                
                
                dGesWert = dGesWert + dWert
                cFeld = Format$(dWert, "#####0.00")
                cFeld = Space$(9 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & "   "
                
                If Not IsNull(rsrs!DAT_AUSG) Then
                    dWert = rsrs!DAT_AUSG
                Else
                    dWert = 0
                End If
                cFeld = Format$(dWert, "DD.MM.YYYY")
                cFeld = Space$(10 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & "   "
                
                If Not IsNull(rsrs!DAT_EINL) Then
                    dWert = rsrs!DAT_EINL
                Else
                    dWert = 0
                End If
                If dWert <> 0 Then
                    cFeld = Format$(dWert, "DD.MM.YYYY")
                Else
                    cFeld = "offen"
                End If
                cFeld = Space$(10 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & "   "
                
                If Not IsNull(rsrs!BEDNU) Then
                    dWert = rsrs!BEDNU
                Else
                    dWert = 0
                End If
                cFeld = Format$(dWert, "##0")
                cFeld = Space$(3 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & "   "
                
                If Not IsNull(rsrs!Kundnr) Then
                    dWert = rsrs!Kundnr
                Else
                    dWert = 0
                End If
                cFeld = Format$(dWert, "#####0")
                
                If Not IsNull(rsrs!Kundnr) Then
                    cFeld = Space$(9 - Len(cFeld)) & cFeld & " " & WhatIsXfromKu(rsrs!Kundnr, "Name")
                End If
                
                cLBSatz = cLBSatz & cFeld & " "
                
                List2.AddItem cLBSatz
                
                rsrs.MoveNext
            Loop
            
            cFeld = Format$(lAnz, "######0")
            cFeld = Space$(7 - Len(cFeld)) & cFeld
            cLBSatz = cFeld & " Gutschein(e) "
            
            cFeld = Format$(dGesWert, "###,##0.00")
            cFeld = Space$(12 - Len(cFeld)) & cFeld
            
            cLBSatz = cLBSatz & "im Wert von " & cFeld & "  " & gcWaehrung
            
            lblanzeige.ForeColor = vbBlue
            lblanzeige.Caption = cLBSatz
            
            Frame3.Visible = True
            
        End If
        
    Else
        lblanzeige.ForeColor = vbRed
        lblanzeige.Caption = "Keine Daten gefunden"
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 13 Then
        lblanzeige.ForeColor = vbRed
        lblanzeige.Caption = "Bitte überprüfen Sie Ihre Eingaben!"
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SucheGutscheineWKLal"
        Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Sub Kopfdaten_speichern()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim sGutschnr       As String
    Dim sGutschArt      As String
    
    Dim sAusgabe_von    As String
    Dim sAusgabe_bis    As String
    
    Dim sEinloese_von   As String
    Dim sEinloese_bis   As String
    
    loeschNEW "KOPF_GUTSCH", gdBase
    CreateTableT2 "KOPF_GUTSCH", gdBase
    
    sGutschnr = ""
    If Text1(4).Text <> "" Then
        sGutschnr = Text1(4).Text
    End If
    
    sGutschArt = ""
    If Option1(0).Value = True Then
        sGutschArt = "nur offene Gutscheine"
    
    ElseIf Option1(1).Value = True Then
        sGutschArt = "nur eingelöste Gutscheine"
    
    ElseIf Option1(2).Value = True Then
        sGutschArt = "alle Gutscheine (offene und eingelöste)"
    End If
    
    sAusgabe_von = ""
    If Text1(0).Text <> "" Then
        sAusgabe_von = Text1(0).Text
    End If
    
    sAusgabe_bis = ""
    If Text1(1).Text <> "" Then
        sAusgabe_bis = Text1(1).Text
    End If
    
    sEinloese_von = ""
    If Text1(2).Text <> "" Then
        sEinloese_von = Text1(2).Text
    End If
    
    sEinloese_bis = ""
    If Text1(3).Text <> "" Then
        sEinloese_bis = Text1(3).Text
    End If
    
    sSQL = "Insert into KOPF_GUTSCH "
    sSQL = sSQL & " ( "
    sSQL = sSQL & " GUTSCHNR "
    sSQL = sSQL & ", GUTSCH_ART "
    sSQL = sSQL & ", AUSGABE_VON  "
    sSQL = sSQL & ", AUSGABE_BIS  "
    sSQL = sSQL & ", EINLOESE_VON  "
    sSQL = sSQL & ", EINLOESE_BIS  "
    sSQL = sSQL & " ) values "
    sSQL = sSQL & " ( "
    sSQL = sSQL & " '" & sGutschnr & "' "
    sSQL = sSQL & " ,'" & sGutschArt & "' "
    sSQL = sSQL & " ,'" & sAusgabe_von & "' "
    sSQL = sSQL & " ,'" & sAusgabe_bis & "' "
    sSQL = sSQL & " ,'" & sEinloese_von & "' "
    sSQL = sSQL & " ,'" & sEinloese_bis & "' "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Kopfdaten_speichern"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."

    Fehlermeldung1
        
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim bFound As Boolean
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0
            SucheGutscheineWKLal
        Case Is = 1
            Unload frmWKLal
        Case Is = 2
            If List2.ListCount > 0 Then
            
                Kopfdaten_speichern
                reportbildschirm "WKL030", "aWKLala"
            Else
                lblanzeige.ForeColor = vbRed
                lblanzeige.Caption = "Keine Daten zum Drucken vorhanden!"
                Command1(0).SetFocus
            End If
        Case Is = 3
            newGutschStrichcods

        Case Is = 4 'Restgutscheine
            Screen.MousePointer = 11
            zeigeHilfeDabapfad "LPROTOK", "GEN_GUT.txt"
            Screen.MousePointer = 0
        Case Is = 5 'alte Gutscheine
            Screen.MousePointer = 11
            zeigeHilfeDabapfad "LPROTOK", "ALT_GUT.txt"
            Screen.MousePointer = 0
        Case Is = 6
            For lcount = 0 To List2.ListCount - 1
                If List2.Selected(lcount) = True Then
                    bFound = True
                End If
            Next lcount
        
            If Not bFound Then
                anzeige "rot", "Bitte einen Eintrag in der Liste markieren!", lblanzeige
                Exit Sub
            End If
            
            Frame5.Visible = True
            
            ermittleEinlösedetails Trim$(Left$(List2.list(List2.ListIndex), 11))
        Case 7
            Frame5.Visible = False
        Case 8
            Label1_Click 14
        Case Is = 9
            For lcount = 0 To List2.ListCount - 1
                If List2.Selected(lcount) = True Then
                    bFound = True
                End If
            Next lcount
        
            If Not bFound Then
                anzeige "rot", "Bitte einen Eintrag in der Liste markieren!", lblanzeige
                Exit Sub
            End If
            
            Frame6.Visible = True
            
            Text1(5).Text = ermittleGutschNotizen(Trim$(Left$(List2.list(List2.ListIndex), 11)))
            Text1(5).SetFocus
        Case Is = 10
            Frame6.Visible = False
            speicherGutschNotizen Trim$(Left$(List2.list(List2.ListIndex), 11)), Text1(5).Text
        
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub ermittleEinlösedetails(cGutschnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    
    Screen.MousePointer = 11
        
    sSQL = "Select * from GUHIS where gutschnrO = " & cGutschnr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        
        
        If Not IsNull(rsrs!FILIALE) Then
            Label1(7).Caption = rsrs!FILIALE
        Else
            Label1(7).Caption = "keine Daten"
        End If
        
        If Not IsNull(rsrs!kasnum) Then
            Label1(9).Caption = rsrs!kasnum
        Else
            Label1(9).Caption = "keine Daten"
        End If
        
        If Not IsNull(rsrs!BEDNU) Then
            Label1(12).Caption = rsrs!BEDNU
        Else
            Label1(12).Caption = "keine Daten"
        End If
        
        If Not IsNull(rsrs!BEDNU) Then
            Label1(11).Caption = rsrs!BEDNU
        Else
            Label1(11).Caption = "keine Daten"
        End If
        
        If Not IsNull(rsrs!BELEGNR) Then
            Label1(14).Caption = rsrs!BELEGNR
        Else
            Label1(14).Caption = "keine Daten"
        End If
        
        If Not IsNull(rsrs!ADATE) Then
            Label1(16).Caption = rsrs!ADATE
        Else
            Label1(16).Caption = "keine Daten"
        End If
        
        If Not IsNull(rsrs!AZEIT) Then
            Label1(18).Caption = rsrs!AZEIT
        Else
            Label1(18).Caption = "keine Daten"
        End If
        
        Label1(12).Caption = ermfromBed("BEDNAME", Label1(12).Caption) & "(" & Label1(12).Caption & ")"
    Else
        Label1(7).Caption = "keine Daten"
        
        Label1(9).Caption = "keine Daten"
 
        Label1(12).Caption = "keine Daten"

        Label1(14).Caption = "keine Daten"
 
        Label1(16).Caption = "keine Daten"
 
        Label1(18).Caption = "keine Daten"
        Label1(11).Caption = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittleEinlösedetails"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz     As String
    Dim iRet        As Integer
    Dim cSQL        As String
    Dim lcount      As Long
    Dim lHeute      As Long
    Dim bFound      As Boolean
        
    cLBSatz = List2.list(List2.ListIndex)
    cLBSatz = Trim$(Left(cLBSatz, 11))
    
    lHeute = Fix(Now)
    
    Select Case Index
        Case Is = 0     'löschen
            iRet = MsgBox("Wollen Sie den/die Gutschein/e " & "wirklich löschen?", vbQuestion + vbYesNo, "LÖSCHEN")
            If iRet = vbYes Then
                For lcount = 0 To List2.ListCount - 1
                    If List2.Selected(lcount) = True Then
                        cLBSatz = List2.list(lcount)
                        cLBSatz = Trim$(Left(cLBSatz, 11))
                        
                        If Val(gcFilNr) = 0 Then
                            cSQL = "Delete from GUTSCH where GUTSCHNR = " & cLBSatz
                            gdBase.Execute cSQL, dbFailOnError
                        Else
                        
                            cSQL = "Update GUTSCH SET STATUS = 'L' where GUTSCHNR = " & cLBSatz
                            gdBase.Execute cSQL, dbFailOnError
                            
                            cSQL = "Update GUTSCH SET SYNSTATUS = 'D' where GUTSCHNR = " & cLBSatz
                            gdBase.Execute cSQL, dbFailOnError
                        End If
                    End If
                Next lcount
                StueckErmittlung
                Command1_Click 0
            End If
        Case Is = 1     'ausbuchen
            iRet = MsgBox("Wollen Sie den/die Gutschein/e " & "wirklich ausbuchen?", vbQuestion + vbYesNo, "AUSBUCHEN")
            If iRet = vbYes Then    '//Felder LastDate und LastTime
                For lcount = 0 To List2.ListCount - 1
                    If List2.Selected(lcount) = True Then
                        cLBSatz = List2.list(lcount)
                        cLBSatz = Trim$(Left(cLBSatz, 11))
                        cSQL = "Update GUTSCH SET STATUS = 'E' where GUTSCHNR = " & cLBSatz
                        gdBase.Execute cSQL, dbFailOnError
                        cSQL = "Update GUTSCH SET SYNSTATUS = 'E' where GUTSCHNR = " & cLBSatz
                        gdBase.Execute cSQL, dbFailOnError
                        cSQL = "Update GUTSCH set DAT_EINL = " & Trim$(Str$(lHeute)) & " where GUTSCHNR = " & cLBSatz
                        gdBase.Execute cSQL, dbFailOnError
                    End If
                Next lcount
                Command1_Click 0
            End If
        
        Case Is = 2     'reaktivieren
            iRet = MsgBox("Wollen Sie den/die Gutschein/e " & "wirklich reaktivieren?", vbQuestion + vbYesNo, "REAKTIVIEREN")
            If iRet = vbYes Then
                '//LastDate und LastTime
                For lcount = 0 To List2.ListCount - 1
                    If List2.Selected(lcount) = True Then
                        cLBSatz = List2.list(lcount)
                        cLBSatz = Trim$(Left(cLBSatz, 11))
                        
                        cSQL = "Update GUTSCH SET STATUS = 'R' where GUTSCHNR = " & cLBSatz
                        gdBase.Execute cSQL, dbFailOnError
                        cSQL = "Update GUTSCH SET SYNSTATUS = 'E' where GUTSCHNR = " & cLBSatz
                        gdBase.Execute cSQL, dbFailOnError
                        cSQL = "Update GUTSCH set DAT_EINL = 0 where GUTSCHNR = " & cLBSatz
                        gdBase.Execute cSQL, dbFailOnError
                    End If
                Next lcount
                Command1_Click 0
            End If
        
    End Select
    
    
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    
    lblanzeige.Caption = ""
    lblanzeige.Refresh
    
    Label1(20).ForeColor = glS1
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_MouseMove"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    LogtoEnd Me
    voreinstellungspeichern
    loeschNEW "Guttemp", gdBase
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    
    Label1(14).ForeColor = glS1
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame5_MouseMove"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim iFil        As Integer
    Dim iKasnum     As Integer
    Dim lBelegnr    As Long
    Dim lbed        As Long
    Dim lDate       As Long

    Select Case Index
        Case 14
            If IsNumeric(Label1(14).Caption) = False Then
                Exit Sub
            End If
            lBelegnr = CLng(Trim(Label1(14).Caption))
            lbed = CLng(Trim(Label1(11).Caption))
            iKasnum = CInt(Trim(Label1(9).Caption))
            iFil = CInt(Trim(Label1(7).Caption))
            lDate = DateValue(Label1(16).Caption)
            
            loeschNEW "KAT" & srechnertab, gdBase
            
            sSQL = "Select * into  KAT" & srechnertab
            sSQL = sSQL & " from Kassjour where "
            sSQL = sSQL & " Filiale = " & iFil
            sSQL = sSQL & " and Kasnum = " & iKasnum
            sSQL = sSQL & " and belegnr = " & lBelegnr
            sSQL = sSQL & " and bediener = " & lbed
            sSQL = sSQL & " and adate = " & lDate
            gdBase.Execute sSQL, dbFailOnError
            
            If Datendrin("KAT" & srechnertab, gdBase) Then
                frmWKL123.Show 1
            End If
        Case Is = 20
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/78-gutscheine-mit-strichcode.html"
    End Select

    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
'Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error GoTo LOKAL_ERROR
'
'    If Index = 14 Then
'        Label1(14).ForeColor = glLink
'    End If
'
'    Exit Sub
'LOKAL_ERROR:
'
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Label1_MouseMove"
'    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
Private Sub List2_Click()
On Error GoTo LOKAL_ERROR

    Dim cNotz As String
        cNotz = ermittleGutschNotizen(Trim$(Left$(List2.list(List2.ListIndex), 11)))
        
        If cNotz <> "" Then
        
            Text1(5).Text = cNotz
            Command1(9).BackColor = vbRed
        Else
        
            Text1(5).Text = ""
            Command1(9).BackColor = Command1(2).BackColor
        
        End If
                                  
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_Click"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Option1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case Is = 5    'vormonat
            If Month(DateValue(Now)) = 1 Then
                Text1(0).Text = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
                Text1(1).Text = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
            Else
                Text1(0).Text = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                Select Case Month(DateValue(Now)) - 1
                    Case 1, 3, 5, 7, 8, 10, 12
                        Text1(1).Text = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                    
                    Case 2
                        If Year(DateValue(Now)) = 2016 Then
                            Text1(1).Text = Format("29." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                        Else
                            Text1(1).Text = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                        End If
                    
                    Case Else
                        Text1(1).Text = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                End Select
            End If
                
        Case Is = 6    'ak monat
            Text1(0).Text = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YY")
            
        Case Is = 4    'vormonat
            If Month(DateValue(Now)) = 1 Then
                Text1(2).Text = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
                Text1(3).Text = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YY")
            Else
                Text1(2).Text = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                Select Case Month(DateValue(Now)) - 1
                    Case 1, 3, 5, 7, 8, 10, 12
                        Text1(3).Text = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                    
                    Case 2
                        If Year(DateValue(Now)) = 2016 Then
                            Text1(3).Text = Format("29." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                        Else
                            Text1(3).Text = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                        End If
                    
                    Case Else
                        Text1(3).Text = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YY")
                End Select
            End If
                
        Case Is = 3    'ak monat
            Text1(2).Text = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YY")
            Text1(3).Text = Format(DateValue(Now), "DD.MM.YY")
        
        
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
'    Text1(Index).SelStart = 0
'    Text1(Index).SelLength = Len(Text1(Index).Text)
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    If Index = 4 Then
        cZeichen = Chr$(KeyAscii)
        cZeichen = UCase$(cZeichen)
        KeyAscii = Asc(cZeichen)
        
        cValid = "0123456789" & Chr$(8)
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
            
        End If
        
    ElseIf Index = 5 Then
    
    
'        cZeichen = Chr$(KeyAscii)
'        cZeichen = UCase$(cZeichen)
'        KeyAscii = Asc(cZeichen)
        
        cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
        cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
        cValid = cValid & "+äÄÜüÖöß%"
        
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
            
        End If
        
    Else
    
        cZeichen = Chr$(KeyAscii)
        If cZeichen = "," Then
            cZeichen = "."
        End If
        cZeichen = UCase$(cZeichen)
        KeyAscii = Asc(cZeichen)
        
        cValid = "0123456789." & Chr$(8)
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        End If
        
    End If
    
     Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If Index = 4 Then
        If KeyCode = vbKeyReturn Then
            Command1_Click 0
        End If
    End If
   
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "In der Gutscheinverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

