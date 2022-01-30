VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frmWKL133 
   Caption         =   "Reparaturverwaltung"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL133.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   7680
      MaxLength       =   50
      TabIndex        =   91
      Top             =   480
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Annahme"
      Height          =   7815
      Left            =   600
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   6975
      Begin sevCommand3.Command Command1 
         Height          =   240
         Index           =   15
         Left            =   6000
         TabIndex        =   90
         Top             =   2740
         Width           =   255
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
         Caption         =   "x"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   240
         Index           =   14
         Left            =   2280
         TabIndex        =   89
         Top             =   2740
         Width           =   255
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
         Caption         =   "x"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   240
         Index           =   13
         Left            =   10200
         TabIndex        =   88
         Top             =   580
         Width           =   255
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
         Caption         =   "x"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   240
         Index           =   12
         Left            =   7200
         TabIndex        =   87
         Top             =   580
         Width           =   255
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
         Caption         =   "x"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   240
         Index           =   11
         Left            =   2280
         TabIndex        =   86
         Top             =   580
         Width           =   255
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
         Caption         =   "x"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   240
         Index           =   10
         Left            =   6360
         TabIndex        =   84
         Top             =   2740
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
         Caption         =   "zeigen..."
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame3 
         Caption         =   "ausgewählter Kunde"
         Height          =   1815
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Visible         =   0   'False
         Width           =   3615
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Titel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   19
            Left            =   1320
            TabIndex        =   47
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fil"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   15
            Left            =   2520
            TabIndex        =   46
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Strasse"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   45
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ort"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   12
            Left            =   720
            TabIndex        =   44
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "PLZ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   43
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Kundenname"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   41
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Kundenname"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "KTEXT1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   39
            Top             =   1440
            Visible         =   0   'False
            Width           =   3255
         End
      End
      Begin sevCommand3.Command Command1 
         Height          =   240
         Index           =   9
         Left            =   2640
         TabIndex        =   83
         Top             =   2740
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
         Caption         =   "zeigen..."
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Filialbeleg"
         Height          =   255
         Index           =   2
         Left            =   9600
         TabIndex        =   82
         Top             =   5760
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Werkstattbeleg"
         Height          =   255
         Index           =   1
         Left            =   9600
         TabIndex        =   81
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Kundenbeleg"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   9600
         TabIndex        =   80
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "zu reparierender Gegenstand"
         Height          =   1815
         Left            =   3840
         TabIndex        =   66
         Top             =   840
         Visible         =   0   'False
         Width           =   3615
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   120
            MaxLength       =   50
            TabIndex        =   73
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   120
            MaxLength       =   50
            TabIndex        =   69
            Top             =   1440
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   120
            MaxLength       =   50
            TabIndex        =   68
            Top             =   1080
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   120
            MaxLength       =   50
            TabIndex        =   67
            Top             =   720
            Width           =   3375
         End
      End
      Begin sevCommand3.Command Command1 
         Height          =   480
         Index           =   5
         Left            =   9480
         TabIndex        =   65
         Top             =   4560
         Visible         =   0   'False
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
         Caption         =   "hinzufügen"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   7560
         TabIndex        =   64
         Top             =   3000
         Visible         =   0   'False
         Width           =   4095
      End
      Begin sevCommand3.Command Command1 
         Height          =   240
         Index           =   8
         Left            =   3840
         TabIndex        =   63
         Top             =   580
         Width           =   2055
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
         Caption         =   "Garantie überprüfen"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Garantieleistung"
         Height          =   255
         Left            =   3840
         TabIndex        =   62
         Top             =   240
         Width           =   1815
      End
      Begin sevCommand3.Command Command1 
         Height          =   240
         Index           =   7
         Left            =   10560
         TabIndex        =   60
         Top             =   580
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
         Caption         =   "suchen..."
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame5 
         Caption         =   "ausgewählter Reparaturbetrieb"
         Height          =   1815
         Left            =   7560
         TabIndex        =   50
         Top             =   840
         Visible         =   0   'False
         Width           =   4095
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Kürzel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   27
            Left            =   1320
            TabIndex        =   59
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   58
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Strasse"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   57
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ort"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   24
            Left            =   720
            TabIndex        =   56
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "PLZ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   55
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Bezeichnung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   53
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fax"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   21
            Left            =   1800
            TabIndex        =   52
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   51
            Top             =   1440
            Visible         =   0   'False
            Width           =   3255
         End
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   2
         Left            =   9480
         TabIndex        =   49
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   3
         Left            =   9480
         TabIndex        =   48
         Top             =   6840
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
         Caption         =   "Übersicht"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   240
         Index           =   4
         Left            =   2640
         TabIndex        =   37
         Top             =   580
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
         Caption         =   "suchen..."
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   240
         Index           =   2
         Left            =   3840
         TabIndex        =   36
         Top             =   4920
         Visible         =   0   'False
         Width           =   1335
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
            Size            =   6
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
         Caption         =   "löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   4920
         Visible         =   0   'False
         Width           =   1335
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
            Size            =   6
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
         Caption         =   "löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   240
         Index           =   6
         Left            =   120
         TabIndex        =   29
         Top             =   4680
         Visible         =   0   'False
         Width           =   1335
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
            Size            =   6
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
         Caption         =   "Zubehör speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   28
         Top             =   4680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         TabIndex        =   27
         Top             =   3960
         Visible         =   0   'False
         Width           =   3615
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
         Height          =   690
         Left            =   3840
         TabIndex        =   26
         Top             =   3960
         Visible         =   0   'False
         Width           =   3615
      End
      Begin sevCommand3.Command Command1 
         Height          =   225
         Index           =   1
         Left            =   3840
         TabIndex        =   25
         Top             =   4680
         Visible         =   0   'False
         Width           =   1335
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
            Size            =   6
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
         Caption         =   "Mängel speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   24
         Top             =   4680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   3840
         TabIndex        =   22
         Top             =   3000
         Visible         =   0   'False
         Width           =   3615
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid2 
         Height          =   1095
         Left            =   120
         TabIndex        =   76
         Top             =   6000
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   1931
         _Version        =   393216
         FocusRect       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Left            =   3480
         TabIndex        =   94
         Top             =   5280
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
         Picture         =   "frmWKL133.frx":0442
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   120
         X2              =   11520
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Ausgefüllt
         Height          =   195
         Index           =   2
         Left            =   11280
         Top             =   5760
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Ausgefüllt
         Height          =   195
         Index           =   1
         Left            =   11280
         Top             =   5520
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   200
         Index           =   0
         Left            =   11280
         Top             =   5280
         Width           =   200
      End
      Begin VB.Label Label2 
         Caption         =   "Autragsnummer"
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
         Left            =   1920
         TabIndex        =   75
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Autragsnummer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Dienstleistungen des Reparaturbetriebes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   7560
         TabIndex        =   72
         Top             =   2760
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Zusammenstellung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   71
         Top             =   5280
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Reparaturbetrieb"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   7680
         TabIndex        =   61
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Mängel:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   35
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Kunde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Zubehör:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Mängel:"
         Height          =   255
         Index           =   10
         Left            =   3840
         TabIndex        =   32
         Top             =   3720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Zubehör:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   31
         Top             =   2760
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12735
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   11655
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   4
         Left            =   9600
         TabIndex        =   16
         Top             =   6840
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
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   1
         Left            =   9600
         TabIndex        =   15
         Top             =   120
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
         Caption         =   "Suche"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   2
         Left            =   9600
         TabIndex        =   14
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
         Caption         =   "Annahme"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   5
         Left            =   9600
         TabIndex        =   13
         Top             =   720
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
         Left            =   9600
         TabIndex        =   12
         Top             =   1320
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
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   10
         Tag             =   "2"
         Top             =   330
         Width           =   1095
      End
      Begin sevCommand3.Command Command1 
         Height          =   355
         Index           =   0
         Left            =   2400
         TabIndex        =   9
         ToolTipText     =   "Auswahlhilfe"
         Top             =   340
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   120
         MaxLength       =   6
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "alle"
         Height          =   195
         Index           =   0
         Left            =   7800
         TabIndex        =   7
         Top             =   0
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "nicht bezahlte"
         Height          =   195
         Index           =   1
         Left            =   7800
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "bezahlte"
         Height          =   195
         Index           =   2
         Left            =   7800
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
         Height          =   3975
         Left            =   120
         TabIndex        =   85
         Top             =   720
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   7011
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid3 
         Height          =   1095
         Left            =   120
         TabIndex        =   92
         Top             =   5280
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   1931
         _Version        =   393216
         FocusRect       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin sevCommand3.Command Command4 
         Height          =   360
         Left            =   9600
         TabIndex        =   93
         Top             =   1920
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
         Picture         =   "frmWKL133.frx":0AD4
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   420
         Index           =   1
         Left            =   4800
         TabIndex        =   95
         ToolTipText     =   "Kalender"
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   741
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
      Begin sevCommand3.Command Command0 
         Height          =   420
         Index           =   0
         Left            =   7080
         TabIndex        =   96
         ToolTipText     =   "Kalender"
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   741
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
      Begin sevCommand3.Command Command2 
         Height          =   165
         Index           =   9
         Left            =   4440
         TabIndex        =   97
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   165
         Index           =   10
         Left            =   4440
         TabIndex        =   98
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   165
         Index           =   7
         Left            =   6720
         TabIndex        =   99
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   165
         Index           =   8
         Left            =   6720
         TabIndex        =   100
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   6840
         Width           =   7215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Datum von:"
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
         Left            =   3240
         TabIndex        =   19
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Datum bis:"
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
         Left            =   5520
         TabIndex        =   18
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "kein Lieferant"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   3015
      End
   End
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   3
      ToolTipText     =   "Hilfe"
      Top             =   360
      Width           =   375
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
         Size            =   6
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
      Caption         =   "?"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   0
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
      BackColor       =   &H00C0C0C0&
      Caption         =   "Es bedient Sie:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   8280
      TabIndex        =   79
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Es bedient Sie:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   8400
      TabIndex        =   78
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Es bedient Sie:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   6240
      TabIndex        =   77
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Kundenname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   7080
      TabIndex        =   70
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Width           =   9255
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
      Caption         =   "Reparaturverwaltung"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmWKL133"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerAuftragsnr As Byte
Dim SpaltennummerKUNDNR As Byte
Dim SpaltennummerLINR As Byte
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 6
            speicher "Kat1", List3, List1, Text1(12).Text
            Text1(12).Text = ""
        Case 1
            speicher "Kat2", List4, List2, Text1(0).Text
            Text1(0).Text = ""
        Case 2
            delete "Kat2", List2, Text1(0).Text
            removeList List4, Text1(0).Text
            removeList List2, Text1(0).Text
            Text1(0).Text = ""
        Case 3
            delete "Kat1", List1, Text1(12).Text
            removeList List3, Text1(12).Text
            removeList List1, Text1(12).Text
            Text1(12).Text = ""
        Case 4
            frmWKL134.Show 1
            
            Frame3.Visible = False
            
            If gckundnr <> "" Then
                If IsNumeric(gckundnr) Then
                    HoleKundenDatenWKL133 gckundnr
                    Command1(4).BackColor = &H8000000F
                    Command1(7).BackColor = &H8000000F
                    Command1(9).BackColor = &H8000000F
                    Command1(10).BackColor = &H8000000F
                    Command1(5).BackColor = &H8000000F
                    
                    Command1(8).BackColor = vbGreen
                End If
            End If
            gckundnr = ""
        Case 5
            SpeicherDenSatz CLng(Label2(12).Caption)
            zeige_Grid
        Case 7 'Lieferant wählen
        
            Frame5.Visible = False
            Label2(5).Visible = False
            List5.Visible = False
            
            gF2Prompt.cFeld = ""
            gF2Prompt.cWert = ""
            gF2Prompt.cWert2 = ""
            gF2Prompt.cWahl = ""
            gF2Prompt.bMultiple = False
            
            gF2Prompt.cFeld = "LINR"
                
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    glLiNr = gF2Prompt.cWahl
                    HoleLieferantenDatenWKL133 glLiNr
                    HoleArtikelDatenWKL133 glLiNr
                    Command1(7).BackColor = &H8000000F
                    Command1(4).BackColor = &H8000000F
                    Command1(8).BackColor = &H8000000F
                    Command1(10).BackColor = &H8000000F
                    Command1(5).BackColor = &H8000000F
                    
                    Command1(9).BackColor = vbGreen
                    glLiNr = 0
                End If
            End If
        Case 8
            frmWKL135.Show 1
            If glGarantienummer > 0 Then
                HoleGarantieDatenWKL133 glGarantienummer
                Command1(8).BackColor = &H8000000F
                Command1(4).BackColor = &H8000000F
                Command1(9).BackColor = &H8000000F
                Command1(10).BackColor = &H8000000F
                Command1(5).BackColor = &H8000000F
                
                Command1(7).BackColor = vbGreen
                
            End If
        Case 9
            zeigezub
            Command1(9).BackColor = &H8000000F
            Command1(7).BackColor = &H8000000F
            Command1(4).BackColor = &H8000000F
            Command1(8).BackColor = &H8000000F
            Command1(5).BackColor = &H8000000F
                    
            Command1(10).BackColor = vbGreen
        Case 10
            zeigemang
            Command1(4).BackColor = &H8000000F
            Command1(7).BackColor = &H8000000F
            Command1(9).BackColor = &H8000000F
            Command1(8).BackColor = &H8000000F
            Command1(5).BackColor = &H8000000F

        Case 11 'rück Kunde
            zeigekundenicht
            
        Case 12
            zeigegarantienicht
            
        Case 13
            zeigerepnicht
            
        Case 14
            zeigezubnicht
        Case 15
            zeigemangnicht
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeige_Grid()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim j           As Integer
    Dim recAnz      As Recordset
    Dim rsrs        As Recordset
    Dim ctmp        As String
    Dim siFak       As Single
    Dim cArtNr      As String
    Dim iStufe      As Integer
    Dim iRet        As Integer
    
    Set recAnz = gdBase.OpenRecordset("Reparatursatz")
    
    If recAnz.EOF Then
        MSFlexGrid2.Visible = False
        MSFlexGrid2.Clear
        anzeige "rot2", "Keine Daten gefunden!", Label1(6)
        Exit Sub
    Else
        
    End If
    recAnz.Close: Set recAnz = Nothing
    
    Screen.MousePointer = 11

    Tabcheck "REPSATZ"
    
    FormatGridOverTablay "REPSATZ"

    With MSFlexGrid2
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 2
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
        Next j
    
        FuellenMSFlex133
'        ermittlespalten
        .Redraw = False
        
        Tabellenbreiteanpassen MSFlexGrid2, 1.1 * gdTabfak
        
        .Visible = True
        .Redraw = True
        .Row = 1
'        .SetFocus
    End With
    
    Me.Refresh
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_Grid"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeige_GridD()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim j           As Integer
    Dim recAnz      As Recordset
    Dim rsrs        As Recordset
    Dim ctmp        As String
    Dim siFak       As Single
    Dim cArtNr      As String
    Dim iStufe      As Integer
    Dim iRet        As Integer
    
    Set recAnz = gdBase.OpenRecordset("Reparatursatz")
    
    If recAnz.EOF Then
        MSFlexGrid3.Visible = False
        MSFlexGrid3.Clear
        anzeige "rot2", "Keine Daten gefunden!", Label1(6)
        Exit Sub
    Else
        
    End If
    recAnz.Close: Set recAnz = Nothing
    
    Screen.MousePointer = 11

    Tabcheck "REPSATZ"
    
    FormatGridOverTablay "REPSATZ"

    With MSFlexGrid3
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 2
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
        Next j
    
        FuellenMSFlex133D
'        ermittlespalten
        .Redraw = False
        
        Tabellenbreiteanpassen MSFlexGrid3, 1.1 * gdTabfak
        
        .Visible = True
        .Redraw = True
        .Row = 1
'        .SetFocus
    End With
    
    Me.Refresh
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_GridD"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeige_GridKopf()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim j           As Integer
    Dim recAnz      As Recordset
    Dim rsrs        As Recordset
    Dim ctmp        As String
    Dim siFak       As Single
    Dim cArtNr      As String
    Dim iStufe      As Integer
    Dim iRet        As Integer
    
    Set recAnz = gdBase.OpenRecordset("ReparaturKopf")
    
    If recAnz.EOF Then
        MSFlexGrid1.Visible = False
        MSFlexGrid1.Clear
        anzeige "rot2", "Keine Daten gefunden!", Label1(6)
        Exit Sub
    Else
        
    End If
    recAnz.Close: Set recAnz = Nothing
    
    Screen.MousePointer = 11

    Tabcheck "REPKOPF"
    
    FormatGridOverTablay "REPKOPF"

    With MSFlexGrid1
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 2
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
        Next j
    
        FuellenMSFlex133KOPF
        ermittlespalten
        .Redraw = False
        
        Tabellenbreiteanpassen MSFlexGrid1, 1.1 * gdTabfak
        
        .Visible = True
        .Redraw = True
        .Row = 1
'        .SetFocus
    End With
    
    Me.Refresh
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_GridKopf"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "AUFTRAGNR"
                SpaltennummerAuftragsnr = i
            Case Is = "KUNDNR"
                SpaltennummerKUNDNR = i
            Case Is = "LINR"
                SpaltennummerKUNDNR = i
        End Select
    Next i
    
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Kunden bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub zeigezub()
On Error GoTo LOKAL_ERROR

    List3.Visible = True
    List1.Visible = True
    Label2(0).Visible = True
    Command1(6).Visible = True
    Command1(3).Visible = True
    Text1(12).Visible = True
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigezub"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigekundenicht()
On Error GoTo LOKAL_ERROR

    Frame3.Visible = False
    gckundnr = ""
    
    Label1(3).Caption = ""
    Label1(7).Caption = ""
    Label1(11).Caption = ""
    Label1(12).Caption = ""
    Label1(19).Caption = ""
    Label1(13).Caption = ""
    Label1(15).Caption = ""
    Label1(17).Caption = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigekundenicht"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigegarantienicht()
On Error GoTo LOKAL_ERROR

    Check1.Value = vbUnchecked

    Text1(3).Text = ""
    Text1(4).Text = ""
    Text1(2).Text = ""
    Text1(1).Text = ""
    
    Frame4.Visible = False
    glGarantienummer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigegarantienicht"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigerepnicht()
On Error GoTo LOKAL_ERROR

    Frame5.Visible = False
    List5.Visible = False
    Label2(5).Visible = False
    
    Label1(22).Caption = ""
    Label1(26).Caption = ""
    Label1(23).Caption = ""
    Label1(24).Caption = ""
    Label1(27).Caption = ""
    Label1(25).Caption = ""
    Label1(21).Caption = ""
    Label1(20).Caption = ""
    Label2(8).Caption = ""

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigerepnicht"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigezubnicht()
On Error GoTo LOKAL_ERROR

    List3.Visible = False
    List1.Visible = False
    Label2(0).Visible = False
    Command1(6).Visible = False
    Command1(3).Visible = False
    Text1(12).Visible = False
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigezubnicht"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigemang()
On Error GoTo LOKAL_ERROR

    List4.Visible = True
    List2.Visible = True
    Label2(10).Visible = True
    Command1(2).Visible = True
    Command1(1).Visible = True
    Text1(0).Visible = True
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigemang"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigemangnicht()
On Error GoTo LOKAL_ERROR

    List4.Visible = False
    List2.Visible = False
    Label2(10).Visible = False
    Command1(2).Visible = False
    Command1(1).Visible = False
    Text1(0).Visible = False
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigemangnicht"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuellenMSFlex133()
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
    Dim cSQL        As String
    
    cSQL = "Select * from Reparatursatz order by autopos "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSFlexGrid2
        .Redraw = False
        lrow = 1
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                lrow = lrow + 1
                .Rows = lrow + 1
                .Col = 0
                
                For i = 0 To byAnzahlSpalten - 1
                    .Row = 0
                    .Col = i
                    
                    If sSpaltenname(i) = .Text Then
                        Select Case sSpaltenname(i)
                            Case Is = "Kassen - VK"
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
                                    sWert = ""
                                End If
                                .Row = lrow
                                .Text = sWert
                        End Select
                        
                        If Len(.TextMatrix(lrow, i)) * 80 > aBreite(i) Then
                            aBreite(i) = Len(.TextMatrix(lrow, i)) * 80
                        End If
                        
                    End If
                Next i
                rsrs.MoveNext
            Loop
        End If
        
        For i = 0 To byAnzahlSpalten - 1
            .Col = i
            .ColWidth(i) = aBreite(i) * 1.8
        Next i
            
        rsrs.Close: Set rsrs = Nothing
        
        If byAnzahlSpalten < 2 Then
        
        Else
            .FixedCols = 1
        End If
        
        .RowHeight(1) = 0
        lrow = lrow - 1
        .Redraw = True
        .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMSFlex133"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
        
    Fehlermeldung1
   
End Sub
Private Sub FuellenMSFlex133D()
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
    Dim cSQL        As String
    
    cSQL = "Select * from Reparatursatz order by autopos "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSFlexGrid3
        .Redraw = False
        lrow = 1
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                lrow = lrow + 1
                .Rows = lrow + 1
                .Col = 0
                
                For i = 0 To byAnzahlSpalten - 1
                    .Row = 0
                    .Col = i
                    
                    If sSpaltenname(i) = .Text Then
                        Select Case sSpaltenname(i)
                            Case Is = "Kassen - VK"
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
                                    sWert = ""
                                End If
                                .Row = lrow
                                .Text = sWert
                        End Select
                        
                        If Len(.TextMatrix(lrow, i)) * 80 > aBreite(i) Then
                            aBreite(i) = Len(.TextMatrix(lrow, i)) * 80
                        End If
                        
                    End If
                Next i
                rsrs.MoveNext
            Loop
        End If
        
        For i = 0 To byAnzahlSpalten - 1
            .Col = i
            .ColWidth(i) = aBreite(i) * 1.8
        Next i
            
        rsrs.Close: Set rsrs = Nothing
        
        If byAnzahlSpalten < 2 Then
        
        Else
            .FixedCols = 1
        End If
        
        .RowHeight(1) = 0
        lrow = lrow - 1
        .Redraw = True
        .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMSFlex133D"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
        
    Fehlermeldung1
   
End Sub
Private Sub FuellenMSFlex133KOPF()
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
    Dim cSQL        As String
    
    cSQL = "Select * from Reparaturkopf order by kvdate "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSFlexGrid1
        .Redraw = False
        lrow = 1
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                lrow = lrow + 1
                .Rows = lrow + 1
                .Col = 0
                
                For i = 0 To byAnzahlSpalten - 1
                    .Row = 0
                    .Col = i
                    
                    If sSpaltenname(i) = .Text Then
                        Select Case sSpaltenname(i)
                            Case Is = "Kassen - VK"
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
                                    sWert = ""
                                End If
                                .Row = lrow
                                .Text = sWert
                        End Select
                        
                        If Len(.TextMatrix(lrow, i)) * 80 > aBreite(i) Then
                            aBreite(i) = Len(.TextMatrix(lrow, i)) * 80
                        End If
                        
                    End If
                Next i
                rsrs.MoveNext
            Loop
        End If
        
        For i = 0 To byAnzahlSpalten - 1
            .Col = i
            .ColWidth(i) = aBreite(i) * 1.8
        Next i
            
        rsrs.Close: Set rsrs = Nothing
        
        If byAnzahlSpalten < 2 Then
        
        Else
            .FixedCols = 1
        End If
        
        .RowHeight(1) = 0
        lrow = lrow - 1
        .Redraw = True
        .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMSFlex133KOPF"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
        
    Fehlermeldung1
   
End Sub
Private Sub HoleKundenDatenWKL133(cKdnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select * from KUNDEN where KUNDNR = " & cKdnr & " "
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst

        If Not IsNull(rsrs!name) Then
            Label1(3).Caption = rsrs!name
        Else
            Label1(3).Caption = ""
        End If
        If Not IsNull(rsrs!vorname) Then
            Label1(7).Caption = rsrs!vorname
        Else
            Label1(7).Caption = ""
        End If
        
        If Not IsNull(rsrs!Plz) Then
            Label1(11).Caption = rsrs!Plz
        Else
            Label1(11).Caption = ""
        End If
        
        If Not IsNull(rsrs!STADT) Then
            Label1(12).Caption = rsrs!STADT
        Else
            Label1(12).Caption = ""
        End If
        
        If Not IsNull(rsrs!titel) Then
            Label1(19).Caption = rsrs!titel
        Else
            Label1(19).Caption = ""
        End If
        
        
        If Not IsNull(rsrs!strasse) Then
            Label1(13).Caption = rsrs!strasse
        Else
            Label1(13).Caption = ""
        End If
        
        If Not IsNull(rsrs!FILIALNR) Then
            Label1(15).Caption = rsrs!FILIALNR
        Else
            Label1(15).Caption = ""
        End If
        
        If Not IsNull(rsrs!KurzTEXT2) Then
            Label1(17).Caption = rsrs!KurzTEXT2
        Else
            Label1(17).Caption = ""
        End If
        
        Label2(2).Caption = cKdnr
        Frame3.Visible = True
    Else
        Frame3.Visible = False
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleKundenDatenWKL133"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub HoleLieferantenDatenWKL133(lLinr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select * from LISRT where LINR = " & lLinr & " "
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst

        If Not IsNull(rsrs!LIEFBEZ) Then
            Label1(22).Caption = rsrs!LIEFBEZ
        Else
            Label1(22).Caption = ""
        End If
        
        If Not IsNull(rsrs!Tel) Then
            Label1(26).Caption = rsrs!Tel
        Else
            Label1(26).Caption = ""
        End If
        
        If Not IsNull(rsrs!Plz) Then
            Label1(23).Caption = rsrs!Plz
        Else
            Label1(23).Caption = ""
        End If
        
        If Not IsNull(rsrs!STADT) Then
            Label1(24).Caption = rsrs!STADT
        Else
            Label1(24).Caption = ""
        End If
        
        If Not IsNull(rsrs!Kuerzel) Then
            Label1(27).Caption = rsrs!Kuerzel
        Else
            Label1(27).Caption = ""
        End If
        
        
        If Not IsNull(rsrs!strasse) Then
            Label1(25).Caption = rsrs!strasse
        Else
            Label1(25).Caption = ""
        End If
        
        If Not IsNull(rsrs!Fax) Then
            Label1(21).Caption = rsrs!Fax
        Else
            Label1(21).Caption = ""
        End If
        
        If Not IsNull(rsrs!Email) Then
            Label1(20).Caption = rsrs!Email
        Else
            Label1(20).Caption = ""
        End If
        
        Label2(8).Caption = lLinr
        Frame5.Visible = True
        Label2(5).Visible = True
        List5.Visible = True
    Else
        Frame5.Visible = False
        Label2(5).Visible = False
        List5.Visible = False
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleLieferantenDatenWKL133"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub HoleGarantieDatenWKL133(lGarant As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select * from Garantie where lfnR = " & lGarant & " "
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst

        If Not IsNull(rsrs!Bemerk) Then
            Text1(3).Text = rsrs!Bemerk
        Else
            Text1(3).Text = ""
        End If
        
        If Not IsNull(rsrs!artnr) Then
            Text1(4).Text = rsrs!artnr
        Else
            Text1(4).Text = ""
        End If
        
        If Not IsNull(rsrs!Seriennr) Then
            Text1(2).Text = rsrs!Seriennr
        Else
            Text1(2).Text = ""
        End If
        
        If Not IsNull(rsrs!BEZEICH) Then
            Text1(1).Text = rsrs!BEZEICH
        Else
            Text1(1).Text = ""
        End If
        
        
        Frame4.Visible = True
    Else
        Frame4.Visible = False
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleGarantieDatenWKL133"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub HoleArtikelDatenWKL133(lLinr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cSatz As String
    
    List5.Clear
    
    cSQL = "Select Artikel.Artnr, Artikel.bezeich, Artikel.kvkpr1 from Artikel inner join Artlief on "
    cSQL = cSQL & " artikel.artnr = artlief.artnr where Artlief.LINR = " & lLinr & " "
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cSatz = ""
            If Not IsNull(rsrs!artnr) Then
                cFeld = Space(6 - Len(rsrs!artnr)) & rsrs!artnr & Space(2)
            Else
                cFeld = Space(8)
            End If
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH & Space(35 - Len(rsrs!BEZEICH)) & Space(2)
            Else
                cFeld = Space(37)
            End If
            cSatz = cSatz & cFeld
            
            If Not IsNull(rsrs!KVKPR1) Then
                cFeld = Format(rsrs!KVKPR1, "#,##0.00") & Space(2)
            Else
                cFeld = Space(10)
            End If
            cSatz = cSatz & cFeld
            
            List5.AddItem cSatz
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleArtikelDatenWKL133"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 2
            Frame1.Visible = False
            Frame2.Visible = True
            
            Command1(4).BackColor = vbGreen
            Label2(12).Caption = ermMaxRepnr
            Label2(12).Refresh
            
            Check2(0).Value = vbChecked
            Check2(1).Value = vbChecked
            Check2(2).Value = vbChecked
            
        Case 4
            Unload frmWKL133
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 11
            gsHelpstring = "Reparaturverwaltung"
            frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command4_Click()
On Error GoTo LOKAL_ERROR
    
    gsZSpalte = "Kundnr"
    gstab = "REPKOPF"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub

Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
Dim sSQL    As String

    Select Case Index
        Case 0
            Unload frmWKL133
            
            
        Case 2
            'speichern
            SpeicherDenAuftrag CLng(Label2(12).Caption)

            'drucken
            If Check2(0).Value = vbChecked Then
                DruckDenAuftragKunde CLng(Label2(12).Caption), "Kundenbeleg"
            End If
            
            If Check2(2).Value = vbChecked Then
                DruckDenAuftragKunde CLng(Label2(12).Caption), "Filialbeleg"
            End If
            
            If Check2(1).Value = vbChecked Then
                DruckDenAuftragKunde CLng(Label2(12).Caption), "Werkstattbeleg"
            End If

            zeigezubnicht
            zeigemangnicht
            zeigerepnicht
            zeigegarantienicht
            zeigekundenicht
            
            Command1(5).Visible = False
            
            sSQL = "Delete from Reparatursatz"
            gdBase.Execute sSQL, dbFailOnError
            
            MSFlexGrid2.Visible = False
            MSFlexGrid2.Clear
            
            Command1(4).BackColor = vbGreen
            Label2(12).Caption = ermMaxRepnr
            Label2(12).Refresh
            
            Me.Refresh
        Case 3
            zeige_GridKopf
            Frame1.Visible = True
            Frame2.Visible = False
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckDenAuftragKunde(lAuftragnr As Long, cBeleg As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    loeschNEW "REPPRINTKU", gdBase
    CreateTable "REPPRINTKU", gdBase
    
    sSQL = "Insert into REPPRINTKU select "
    sSQL = sSQL & " AUFTRAGNR"
    sSQL = sSQL & ", ARTNRDEF"
    sSQL = sSQL & ", ARTNRDIE"
    sSQL = sSQL & ", BEZEICHDEF"
    sSQL = sSQL & ", BEZEICHDIE"
    sSQL = sSQL & ", Seriennr"
    sSQL = sSQL & ", Bemerk"
    sSQL = sSQL & ", MENGEDEF"
    sSQL = sSQL & ", MENGEDIE"
    sSQL = sSQL & ", STATUSKV"
    sSQL = sSQL & ", STATUSAUTR"
    sSQL = sSQL & ", KVKPR1"
    sSQL = sSQL & ", kundnr"
    sSQL = sSQL & ", linr"
    sSQL = sSQL & ", Bednu"
    sSQL = sSQL & ", KVDATE"
    sSQL = sSQL & ", AUTRDATE"
    sSQL = sSQL & ", FILIALE"
    sSQL = sSQL & ", SENDOK"
    sSQL = sSQL & ", Mangel"
    sSQL = sSQL & ", ZUBEH"
    sSQL = sSQL & ", Garantie "
    sSQL = sSQL & ", '" & cBeleg & "' as belegart "
    sSQL = sSQL & " from Reparatur "
    sSQL = sSQL & " where AUFTRAGNR = " & lAuftragnr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update REPPRINTKU inner join Kunden on REPPRINTKU.kundnr = Kunden.kundnr"
    sSQL = sSQL & " set REPPRINTKU.KUNACHNAME = Kunden.name "
    sSQL = sSQL & " , REPPRINTKU.KUVORNAME = Kunden.vorname "
    sSQL = sSQL & " , REPPRINTKU.KUTITEL = Kunden.titel "
    sSQL = sSQL & " , REPPRINTKU.KUPLZ = Kunden.plz "
    sSQL = sSQL & " , REPPRINTKU.KUSTADT = Kunden.stadt "
    sSQL = sSQL & " , REPPRINTKU.KUANREDE = Kunden.anrede "
    sSQL = sSQL & " , REPPRINTKU.KUGESCHLECHT = Kunden.geschlecht "
    sSQL = sSQL & " , REPPRINTKU.KUFIRMA = Kunden.firma "
    sSQL = sSQL & " , REPPRINTKU.KUSTRASSE = Kunden.strasse "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update REPPRINTKU inner join LISRT on REPPRINTKU.LINR = LISRT.LINR"
    sSQL = sSQL & " set REPPRINTKU.LIBEZEICH = LISRT.LIEFBEZ "
    sSQL = sSQL & " , REPPRINTKU.LIPLZ = LISRT.PLZ "
    sSQL = sSQL & " , REPPRINTKU.LISTADT = LISRT.STADT "
    sSQL = sSQL & " , REPPRINTKU.LISTRASSE = LISRT.STRASSE "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update REPPRINTKU inner join Bedname on REPPRINTKU.Bednu = Bedname.Bednu "
    sSQL = sSQL & " set REPPRINTKU.bedname = Bedname.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    Dim cLeer As String
    cLeer = " "
    
    If UCase(cBeleg) = "WERKSTATTBELEG" Then
    
        sSQL = "Update REPPRINTKU "
        sSQL = sSQL & " set anredtit  = '' "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update REPPRINTKU "
        sSQL = sSQL & " set namefirma  = LIBEZEICH"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update REPPRINTKU "
        sSQL = sSQL & " set KUSTADT  = LISTADT "
        sSQL = sSQL & " , KUPLZ  = LIPLZ "
        sSQL = sSQL & " , KUSTRASSE  = LISTRASSE "
        gdBase.Execute sSQL, dbFailOnError
    Else
    
        sSQL = "Update REPPRINTKU "
        sSQL = sSQL & " set anredtit  = kuanrede + '" & cLeer & "' "
        sSQL = sSQL & " where kutitel <> ''"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update REPPRINTKU "
        sSQL = sSQL & " set anredtit  = anredtit  + kutitel "
        sSQL = sSQL & " where kutitel <> ''"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update REPPRINTKU "
        sSQL = sSQL & " set namefirma  = kuvorname + '" & cLeer & "' "
        sSQL = sSQL & " where kuvorname <> ''"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update REPPRINTKU "
        sSQL = sSQL & " set namefirma  = namefirma  + kunachname "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update REPPRINTKU "
        sSQL = sSQL & " set namefirma  = namefirma  + kufirma "
        sSQL = sSQL & " where kufirma <> ''"
        gdBase.Execute sSQL, dbFailOnError
        
    End If
    
    reportbildschirm "", "aWKL133a"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckDenAuftragKunde"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherDenAuftrag(lAuftragnr As Long)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim lbednu As Long
    Dim lLinr As Long
    Dim sLiefBez As String
    
    lLinr = Label2(8).Caption
    lbednu = Text1(5).Text
    sLiefBez = Label1(22).Caption
    
    
    
    sSQL = "Insert into Reparatur select "
    sSQL = sSQL & " AUFTRAGNR"
    sSQL = sSQL & ", ARTNRDEF"
    sSQL = sSQL & ", ARTNRDIE"
    sSQL = sSQL & ", BEZEICHDEF"
    sSQL = sSQL & ", BEZEICHDIE"
    sSQL = sSQL & ", Seriennr"
    sSQL = sSQL & ", Bemerk"
    sSQL = sSQL & ", MENGEDEF"
    sSQL = sSQL & ", MENGEDIE"
    sSQL = sSQL & ", 'aufgenommen' as STATUSKV"
    sSQL = sSQL & ", 'kein Auftrag' as STATUSAUTR"
    sSQL = sSQL & ", KVKPR1"
    sSQL = sSQL & ", kundnr"
    sSQL = sSQL & ", nachname"
    sSQL = sSQL & ", " & lbednu & " as bednu"
    sSQL = sSQL & ", " & lLinr & " as linr"
    sSQL = sSQL & ", '" & sLiefBez & "' as liefbez "
    sSQL = sSQL & ", KVDATE"
    sSQL = sSQL & ", AUTRDATE"
    sSQL = sSQL & ", FILIALE"
    sSQL = sSQL & ", SENDOK"
    sSQL = sSQL & ", Mangel"
    sSQL = sSQL & ", ZUBEH"
    sSQL = sSQL & ", Garantie "
    sSQL = sSQL & " from Reparatursatz "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Reparaturkopf select "
    sSQL = sSQL & " AUFTRAGNR "
    sSQL = sSQL & ", STATUSKV"
    sSQL = sSQL & ", STATUSAUTR"
    sSQL = sSQL & ", kundnr"
    sSQL = sSQL & ", nachname"
    sSQL = sSQL & ", bednu"
    sSQL = sSQL & ", linr"
    sSQL = sSQL & ", liefbez "
    sSQL = sSQL & ", KVDATE"
    sSQL = sSQL & ", AUTRDATE"
    sSQL = sSQL & ", FILIALE"
    sSQL = sSQL & ", SENDOK"
    sSQL = sSQL & " from Reparatur "
    sSQL = sSQL & " where AUFTRAGNR = " & lAuftragnr & " "
    sSQL = sSQL & " group by AUFTRAGNR"
    sSQL = sSQL & ", STATUSKV"
    sSQL = sSQL & ", STATUSAUTR"
    sSQL = sSQL & ", kundnr"
    sSQL = sSQL & ", nachname"
    sSQL = sSQL & ", bednu"
    sSQL = sSQL & ", linr"
    sSQL = sSQL & ", liefbez "
    sSQL = sSQL & ", KVDATE"
    sSQL = sSQL & ", AUTRDATE"
    sSQL = sSQL & ", FILIALE"
    sSQL = sSQL & ", SENDOK"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherDenAuftrag"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherDenSatz(lAuftragnr As Long)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    Dim lArtnrdef As Long
    Dim lArtnrDIE As Long
    Dim sBEZEICHDEF As String
    Dim sBEZEICHDIE As String
    Dim sSERIENNR As String
    Dim sBemerk As String
    Dim lMENGEDEF As Long
    Dim lMENGEDIE As Long
    Dim sSTATUSKV As String
    Dim sSTATUSAUTR As String
    Dim dEkpr As Double
    Dim dKVkPr1 As Double
    Dim sLiefBez As String
    Dim lLinr As Long
    Dim sNACHNAME As String
    Dim lKUNDNR As Long
    Dim lKVDATE As Long
    Dim lAUTRDATE As Long
    Dim byFILIALE As Byte
    Dim sMangel As String
    Dim sZUBEH As String
    Dim bGarantie As Boolean
    Dim ctemp As String
    
    Dim lcount As Long
    Dim bFound As Boolean

    bFound = False

    For lcount = 0 To List5.ListCount - 1
        If List5.Selected(lcount) = True Then
            bFound = True
        End If
    Next lcount
    
    If Not bFound Then
        anzeige "rot", "Bitte markieren Sie einen Artikel", Label1(6)
        Exit Sub
    End If
    
    For lcount = 0 To List5.ListCount - 1
        If List5.Selected(lcount) = True Then
        
            lArtnrDIE = Mid(List5.list(List5.ListIndex), 1, 6)
            sBEZEICHDIE = Mid(List5.list(List5.ListIndex), 8, 35)
            ctemp = CDbl(Mid(List5.list(List5.ListIndex), 45, 10))
            dKVkPr1 = CDbl(Mid(List5.list(List5.ListIndex), 45, 10))
            
        End If
    Next lcount
    
    sMangel = ""
    For lcount = 0 To List4.ListCount - 1
        sMangel = sMangel & List4.list(lcount) & " "
    Next lcount
    
    sZUBEH = ""
    For lcount = 0 To List3.ListCount - 1
        sZUBEH = sZUBEH & List3.list(lcount) & " "
    Next lcount
    
    If Frame4.Visible = False Then
        lArtnrdef = 0
        sBEZEICHDEF = ""
        sSERIENNR = ""
        sBemerk = ""
    Else
        lArtnrdef = CLng(Text1(4).Text)
        sBEZEICHDEF = Text1(1).Text
        sSERIENNR = Text1(2).Text
        sBemerk = Text1(3).Text
    End If
    
    lMENGEDEF = 1
    lMENGEDIE = 1
    
    If Check1.Value = vbChecked Then
        bGarantie = True
    Else
        bGarantie = False
    End If
    
    lKUNDNR = Label2(2).Caption
    sNACHNAME = Label1(3).Caption
    
    Set rsrs = gdBase.OpenRecordset("REPARATURSATZ")
    rsrs.AddNew
    
    rsrs!AUFTRAGNR = lAuftragnr
    rsrs!ARTNRDEF = lArtnrdef
    rsrs!ARTNRDIE = lArtnrDIE
    rsrs!BEZEICHDEF = sBEZEICHDEF
    rsrs!BEZEICHDIE = sBEZEICHDIE
    rsrs!Seriennr = sSERIENNR
    rsrs!Bemerk = sBemerk
    rsrs!MENGEDEF = lMENGEDEF
    rsrs!MENGEDIE = lMENGEDIE
    rsrs!STATUSKV = sSTATUSKV
    rsrs!STATUSAUTR = sSTATUSAUTR
'    rsrs!ekpr = dEKPR
    rsrs!KVKPR1 = dKVkPr1
    
    rsrs!linr = lLinr
    rsrs!LIEFBEZ = sLiefBez
    rsrs!Kundnr = lKUNDNR
    rsrs!nachname = sNACHNAME
    
    rsrs!KVDATE = CLng(DateValue(Now))
    rsrs!AUTRDATE = CLng(DateValue(Now))
    rsrs!FILIALE = gcFilNr
    rsrs!SENDOK = False
    rsrs!Mangel = sMangel
    rsrs!ZUBEH = sZUBEH
    rsrs!Garantie = bGarantie
    
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherDenSatz"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command98_Click()
On Error GoTo LOKAL_ERROR
    
    gsZSpalte = "Artnr"
    gstab = "REPSATZ"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command98_Click"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    WKL133Positionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    anzeige "normal", "", Label1(4)
    
    If Not NewTableSuchenDBKombi("KAT1", gdBase) Then
        CreateTable "KAT1", gdBase
    End If
    
    If Not NewTableSuchenDBKombi("KAT2", gdBase) Then
        CreateTable "KAT2", gdBase
    End If
    
    fuelleKatliste "KAT1", List1
    fuelleKatliste "KAT2", List2
    
    vorbereit
    
    zeige_GridKopf
    
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
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
Private Sub vorbereit()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Delete from Reparatursatz "
    gdBase.Execute sSQL, dbFailOnError
    
    Text1(5).Text = gcBedienerNr
'    Label1(8).Caption = gcBedienerNr
    Label1(9).Caption = gcUserName
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereit"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuelleKatliste(sKat As String, Listx As ListBox)
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim dWert       As Double
    Dim cFeld       As String
    Dim cLBSatz     As String
    
    Listx.Clear
    
    cSQL = "Select * from " & sKat & " order by BEZEICH "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cFeld = ""
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            End If
            Listx.AddItem cFeld
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuelleKatliste"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicher(sKat As String, listz As ListBox, Listx As ListBox, sBez As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    
'    listx.Clear
    sBez = Trim(sBez)
    If sBez <> "" Then
        sSQL = "Delete from " & sKat & " where Bezeich = '" & sBez & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into " & sKat & " (BEZEICH) values ('" & sBez & "')"
        gdBase.Execute sSQL, dbFailOnError
        
        If ZeigAn(listz, sBez) = False Then
            listz.AddItem sBez
        End If
    
'        fuelleKatliste sKat, listx
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub delete(sKat As String, Listx As ListBox, sBez As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    
'    listx.Clear
    sBez = Trim(sBez)
    If sBez <> "" Then
        sSQL = "Delete from " & sKat & " where Bezeich = '" & sBez & "'"
        gdBase.Execute sSQL, dbFailOnError
    
'        fuelleKatliste sKat, listx
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WKL133Positionieren()
On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 960
    Frame1.Left = 0
    Frame1.Width = 11775
    Frame1.Height = 7455
    
    Frame2.Top = 960
    Frame2.Left = 120
    Frame2.Width = 11775
    Frame2.Height = 7455
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL133Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
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

Private Sub List1_Click()
On Error GoTo LOKAL_ERROR

    BezInTextbox List1, Text1(12)
    listverschiebung List1, List3
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List1_Click"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub BezInTextbox(Listx As ListBox, textx As TextBox)
On Error GoTo LOKAL_ERROR

    Dim bFound As Boolean
    Dim lcount As Long
    
    bFound = False
    
    textx.Text = ""
    
    If Listx.ListCount = 0 Then
        Exit Sub
    End If
    
    For lcount = 0 To Listx.ListCount - 1
        If Listx.Selected(lcount) = True Then
            bFound = True
        End If
    Next lcount
    
    If bFound Then
        For lcount = 0 To Listx.ListCount - 1
            If Listx.Selected(lcount) Then
                textx.Text = Listx.list(lcount)
                Exit For
            End If
        Next lcount
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BezInTextbox"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List2_Click()
On Error GoTo LOKAL_ERROR

    BezInTextbox List2, Text1(0)
    listverschiebung List2, List4
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_Click"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List3_Click()
On Error GoTo LOKAL_ERROR

    BezInTextbox List3, Text1(12)
    listverschiebung List3, List1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List3_Click"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List4_Click()
On Error GoTo LOKAL_ERROR

    BezInTextbox List4, Text1(0)
    listverschiebung List4, List2
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List4_Click"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub removeList(listx1 As ListBox, cBez As String)
On Error GoTo LOKAL_ERROR

    Dim bFound As Boolean
    Dim lcount As Long
    
    bFound = False
    
    If listx1.ListCount = 0 Then
        Exit Sub
    End If
    
    For lcount = 0 To listx1.ListCount - 1
        If cBez = listx1.list(lcount) Then
            listx1.RemoveItem lcount
        End If
    Next lcount
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "removeList"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ZeigAn(listx1 As ListBox, cBez As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    
    ZeigAn = False
    
    For lcount = 0 To listx1.ListCount - 1
        If cBez = listx1.list(lcount) Then
            ZeigAn = True
            Exit For
        End If
    Next lcount
    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigAn"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub listverschiebung(quellList As ListBox, ZielList As ListBox)
On Error GoTo LOKAL_ERROR

    Dim bFound As Boolean
    Dim lcount As Long
    
    bFound = False
    
    If quellList.ListCount = 0 Then
        Exit Sub
    End If
    
    For lcount = 0 To quellList.ListCount - 1
        If quellList.Selected(lcount) = True Then
            bFound = True
        End If
    Next lcount
    
    If bFound Then
        
        For lcount = 0 To quellList.ListCount - 1
            If quellList.Selected(lcount) Then
                ZielList.AddItem quellList.list(lcount)
                quellList.RemoveItem lcount
                Exit For
            End If
        Next lcount
    
    Else
        ZielList.AddItem quellList.list(quellList.TopIndex)
        quellList.RemoveItem quellList.TopIndex
    End If
    
    quellList.Refresh
    ZielList.Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_Click"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List5_Click()
On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim bFound As Boolean

    bFound = False

    For lcount = 0 To List5.ListCount - 1
        If List5.Selected(lcount) = True Then
            bFound = True
        End If
    Next lcount
    
    If bFound Then
        Command1(5).Visible = True
        Command1(5).BackColor = vbGreen
    Else
        Command1(5).BackColor = &H8000000F
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List5_Click"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_DblClick()
On Error GoTo LOKAL_ERROR
    
    If MSFlexGrid1.Row > 1 Then
        detailanzeigen
    Else
        sortierenHGrid MSFlexGrid1
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub detailanzeigen()
    On Error GoTo LOKAL_ERROR
    
    Dim cSuch As String
    Dim cSQL As String
    
    If MSFlexGrid1.Row < 1 Then
        Screen.MousePointer = 0
        MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerAuftragsnr)
    cSuch = Trim$(cSuch)
    
    If IsNumeric(cSuch) Then
    
    Else
        Screen.MousePointer = 0
        MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    cSQL = "Delete from Reparatursatz"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into Reparatursatz Select * from Reparatur where Auftragnr = " & cSuch
    gdBase.Execute cSQL, dbFailOnError
    
    
    zeige_GridD


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "detailanzeigen"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    Text1(Index).BackColor = glSelBack1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
    
        If Index = 12 Then
            Command1_Click 6
        ElseIf Index = 0 Then
            Command1_Click 1
        ElseIf Index = 5 Then
            fnLogonBedienerWKL133
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnLogonBedienerWKL133()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp        As String
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    fnLogonBedienerWKL133 = 0
    
    ctmp = Text1(5).Text
    ctmp = Trim$(ctmp)
    
    Label1(9).Caption = ""
    anzeige "normal", "", Label1(6)
    
    If ctmp = "" Then
        anzeige "rot2", "Bitte Ihre Bediener-Nummer eingeben!", Label1(6)
        Text1(5).SetFocus
        fnLogonBedienerWKL133 = 1
        Exit Function
    End If
    
    cSQL = "Select * from BEDNAME where BEDNU = " & ctmp & " "
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    If rsrs.EOF Then
        anzeige "rot", "Die eingegebene Bediener-Nummer ist ungültig!", Label1(6)
        Text1(5).Text = ""
        Text1(5).SetFocus
        fnLogonBedienerWKL133 = 1
    Else
        rsrs.MoveFirst
        If Not IsNull(rsrs!bedname) Then
            gcBediener = rsrs!bedname
        Else
            gcBediener = ""
        End If
        Label1(9).Caption = gcBediener
        fnLogonBedienerWKL133 = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnLogonBedienerWKL133"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Reparaturverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
