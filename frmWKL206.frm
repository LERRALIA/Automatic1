VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL206 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Lagerplatz zuweisen"
   ClientHeight    =   8625
   ClientLeft      =   2145
   ClientTop       =   2655
   ClientWidth     =   11910
   Icon            =   "frmWKL206.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      Height          =   735
      Left            =   9960
      TabIndex        =   37
      Top             =   6720
      Visible         =   0   'False
      Width           =   1920
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   16
         Left            =   9480
         TabIndex        =   26
         Top             =   1320
         Width           =   2175
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
         Caption         =   "<<<"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   20
         Left            =   8640
         TabIndex        =   25
         Top             =   1320
         Width           =   840
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
         Caption         =   "F4"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   19
         Left            =   7800
         TabIndex        =   24
         Top             =   1320
         Width           =   840
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
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   18
         Left            =   6960
         TabIndex        =   23
         Top             =   1320
         Width           =   840
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
         Caption         =   ","
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   14
         Left            =   4440
         TabIndex        =   22
         Top             =   1320
         Width           =   2520
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
         Caption         =   "Rückgängig"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   13
         Left            =   1920
         TabIndex        =   21
         Top             =   1320
         Width           =   2520
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
         Caption         =   "Leeren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   12
         Left            =   1080
         TabIndex        =   20
         Top             =   1320
         Width           =   840
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
         Caption         =   "-"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   11
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   840
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
         Caption         =   "+"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   17
         Left            =   9480
         TabIndex        =   18
         Top             =   480
         Width           =   2175
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
         Caption         =   ">>>"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   10
         Left            =   8640
         TabIndex        =   17
         Top             =   480
         Width           =   840
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
         Caption         =   "00"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   9
         Left            =   7800
         TabIndex        =   16
         Top             =   480
         Width           =   840
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
         Caption         =   "0"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   8
         Left            =   6960
         TabIndex        =   15
         Top             =   480
         Width           =   840
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
         Caption         =   "9"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   7
         Left            =   6120
         TabIndex        =   14
         Top             =   480
         Width           =   840
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
         Caption         =   "8"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   6
         Left            =   5280
         TabIndex        =   13
         Top             =   480
         Width           =   840
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
         Caption         =   "7"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   5
         Left            =   4440
         TabIndex        =   12
         Top             =   480
         Width           =   840
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
         Caption         =   "6"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   4
         Left            =   3600
         TabIndex        =   11
         Top             =   480
         Width           =   840
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
         Caption         =   "5"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   3
         Left            =   2760
         TabIndex        =   10
         Top             =   480
         Width           =   840
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
         Caption         =   "4"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   2
         Left            =   1920
         TabIndex        =   9
         Top             =   480
         Width           =   840
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
         Caption         =   "3"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   1
         Left            =   1080
         TabIndex        =   8
         Top             =   480
         Width           =   840
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
         Caption         =   "2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   840
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
         Caption         =   "1"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Label3"
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
         Left            =   240
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      ItemData        =   "frmWKL206.frx":0442
      Left            =   240
      List            =   "frmWKL206.frx":0444
      TabIndex        =   42
      Top             =   7320
      Width           =   7695
   End
   Begin VB.ListBox List6 
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
      ItemData        =   "frmWKL206.frx":0446
      Left            =   240
      List            =   "frmWKL206.frx":0448
      TabIndex        =   45
      Top             =   7080
      Width           =   7695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'Kein
      Height          =   5655
      Left            =   120
      TabIndex        =   27
      Top             =   960
      Width           =   11775
      Begin VB.CheckBox Check1 
         Caption         =   "schneller Scanmodus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7440
         TabIndex        =   51
         Top             =   3600
         Width           =   4335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Fach + 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7440
         TabIndex        =   49
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   3
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   4
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   2
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   3
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   1
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   2
         Top             =   3240
         Width           =   1815
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Left            =   8760
         TabIndex        =   6
         Top             =   4560
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1085
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
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   15
         Left            =   5760
         TabIndex        =   5
         Top             =   4560
         Width           =   2880
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   2520
         MaxLength       =   13
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   5055
      End
      Begin sevCommand3.Command Command1 
         Height          =   615
         Left            =   8760
         TabIndex        =   1
         Top             =   240
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1085
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
         Caption         =   "Suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   22
         Left            =   2760
         TabIndex        =   52
         Top             =   4560
         Width           =   2880
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
         Caption         =   "Leeren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Fach:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   5160
         TabIndex        =   48
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Boden:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   3000
         TabIndex        =   47
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
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
         Left            =   240
         TabIndex        =   46
         Top             =   5400
         Width           =   9615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   240
         X2              =   11640
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 Euro"
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
         Height          =   375
         Index           =   5
         Left            =   9600
         TabIndex        =   40
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Kassen-Vk:"
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
         Index           =   9
         Left            =   8040
         TabIndex        =   39
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Regal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   840
         TabIndex        =   36
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Listen-Vk:"
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
         Left            =   8040
         TabIndex        =   35
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 Euro"
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
         Height          =   375
         Index           =   3
         Left            =   9600
         TabIndex        =   34
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Bestand:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   0
         Left            =   4920
         TabIndex        =   32
         Top             =   960
         Width           =   6735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   1
         Left            =   2520
         TabIndex        =   31
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Artikel:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   2
         Left            =   2520
         TabIndex        =   29
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "EAN / ArtNr:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   2175
      End
   End
   Begin sevCommand3.Command Command2 
      Height          =   360
      Index           =   21
      Left            =   11280
      TabIndex        =   50
      Top             =   240
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   635
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
      ToolTip         =   "Hier können Sie sich die Bildschirmtastatur ein- bzw. ausblenden."
      ToolTipTitle    =   "Tastatur"
      ButtonStyle     =   2
      Caption         =   ""
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin VB.Label Label14 
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
      Left            =   240
      TabIndex        =   44
      Top             =   6480
      Width           =   11415
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "zuletzt bearbeitete Artikel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   43
      Top             =   6840
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   11640
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Lagerplatz zuweisen"
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
      TabIndex        =   41
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmWKL206"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim new19Artikel As ArtikelTyp


'Private Sub Command7_Click(Index As Integer)
'On Error GoTo LOKAL_ERROR
'
'    Select Case Index
'        Case Is = 0
'            Unload frmWKL206
'        Case Is = 1
'
'        Case Is = 2
'
'
'            List5.Visible = False
'            List6.Visible = False
'            Label11(3).Visible = False
'            Frame1.Visible = False
'            Command2(21).Visible = False
'
'    End Select
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Command7_Click"
'    Fehler.gsFehlertext = "Im Programmteil Lagerplatz zuweisen ist ein Fehler aufgetreten. "
'
'    Fehlermeldung1
'End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    gsARTNR = ""
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
Private Sub PositionierenWKL206()
    On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 6240
    Frame1.Left = 0
    Frame1.Height = 2775
    Frame1.Width = 12000
    
    Frame3.Top = 840
    Frame3.Left = 0
    Frame3.Height = 6000
    Frame3.Width = 12000
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL206"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub LeereDialogWKL15(bAll As Boolean, Optional bPlusOne As Boolean = False)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    
    Text1(0).Text = ""
    If bAll Then
        Text1(1).Text = ""
        Text1(2).Text = ""
        Text1(3).Text = ""
    End If
    Label2(0).Caption = "unbekannt"
    Label2(1).Caption = "0"
    Label2(2).Caption = "0"
    Label2(3).Caption = "0,00 " & gcWaehrung
    Label2(5).Caption = "0,00 " & gcWaehrung
    Label3.Caption = "0"
    
    If bPlusOne = True Then
    
        Dim cNullen As String
        Dim iAnzNullen As Integer
    
        'führende Nullen?
        iAnzNullen = Len(Text1(3).Text) - Len(CStr(Val(Text1(3).Text)))
        
        If iAnzNullen > 0 Then
            For i = 1 To iAnzNullen
                cNullen = cNullen & "0"
            Next
            
            Text1(3).Text = cNullen & CStr(Val(Text1(3).Text) + 1)
        
        Else
    
            Text1(3).Text = Val(Text1(3).Text) + 1
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKL15"
    Fehler.gsFehlertext = "Im Programmteil Lagerplatz zuweisen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub SchreibeDatenWKL206()
    On Error GoTo LOKAL_ERROR
        
    Dim cFach          As String
    Dim cRegal         As String
    Dim cBoden         As String
    Dim cLAGERPLATZ    As String
    Dim sSQL           As String
    Dim cArtNr         As String
    Dim cBezeich       As String
    
    cArtNr = Label2(2).Caption
    
    cRegal = Trim$(Text1(1).Text)
    cBoden = Trim$(Text1(2).Text)
    cFach = Trim$(Text1(3).Text)
    cLAGERPLATZ = Trim(cRegal & cBoden & cFach)
    
    cArtNr = Label2(2).Caption
    cArtNr = Trim$(cArtNr)
    If cArtNr = "" Then
        MsgBox "Artikel-Nr fehlt! Daten speichern nicht möglich!", vbCritical, "Winkiss Hinweis:"
        Text1(0).SetFocus
        Exit Sub
    End If
    
    sSQL = "Delete * from Lagerplatz where artnr = " & cArtNr
    gdBase.Execute sSQL, dbFailOnError
    
    If cLAGERPLATZ <> "" Then
        sSQL = "Insert into LAGERPLATZ (ARTNR,LAGERP) Values (" & cArtNr & ", " & cLAGERPLATZ & ") "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    cBezeich = fnArtBezSuchen(cArtNr)
    
    new19Artikel.artnr = cArtNr
    new19Artikel.BEZEICH = cBezeich
    new19Artikel.LAGERPLATZ = Val(cLAGERPLATZ)

    If Check2.Value = vbChecked Then
        LeereDialogWKL15 False, True
    Else
        LeereDialogWKL15 False
    End If
    
    Text1(0).SetFocus

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWKL206"
    Fehler.gsFehlertext = "Im Programmteil Lagerplatz zuweisen ist ein Fehler aufgetreten. "

    Fehlermeldung1
'    Resume Next
End Sub
Private Sub SucheArtikelWKL206()
    On Error GoTo LOKAL_ERROR
    
    Dim bDebug As Boolean
    Dim iRet As Integer
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsRs2 As Recordset
    Dim cSuch As String
    Dim cArtNr As String
    Dim cArtBez As String
    Dim dBestand As Double
    Dim dVkPr As Double
    Dim dKVKPR As Double
    Dim dEkpr As Double
    Dim dLEKPR As Double
    
    Dim cLiBesNr As String
    Dim bgefunden As Boolean
    Dim cFeld As String
    Dim cLBSatz As String
    Dim lMinBest As Long
    Dim bEAN As Boolean
    Dim cEAN As String
    
    bDebug = False
    bgefunden = True
    bEAN = True
    
    cSuch = Text1(0).Text
    cSuch = Trim$(cSuch)
    
    If cSuch = "" Then
        anzeige "rot2", "Bitte Wert eingeben!", Label1(1)
'        MsgBox "Bitte Wert eingeben!", vbCritical, "STOP!"
        Text1(0).SetFocus
        Exit Sub
    End If
    
    
    If Len(cSuch) > 6 Then
        iRet = fnPruefeEANWert(cSuch)
        Select Case iRet
            Case Is = 0
                'alles okay
            Case Is = 1     'falsche Länge
                bEAN = False

            Case Is = 8     'falscher EAN-8
                bEAN = False

            Case Is = 12    'falscher UPC-A
                bEAN = False

            Case Is = 13    'falscher EAN-13
                bEAN = False

        End Select
        cSQL = "Select B.ARTNR, A.BEZEICH, B.LINR, A.BESTAND, A.VKPR, A.KVKPR1, A.MINBEST, B.LIBESNR, A.EAN "
        cSQL = cSQL & "from ARTIKEL A, ARTLIEF B where A.ARTNR = B.ARTNR "
    End If
    
    If Len(cSuch) <= 6 Then
        cSQL = "Select B.ARTNR, A.BEZEICH, B.LINR, A.BESTAND, A.VKPR, A.KVKPR1, A.MINBEST, B.LIBESNR, A.EAN "
        cSQL = cSQL & "from ARTIKEL A, ARTLIEF B where B.ARTNR = " & cSuch & " and A.ARTNR = B.ARTNR "
    Else
        If Len(cSuch) <= 8 And (Left(cSuch, 1) = "2") Then  'Or Left(cSuch, 1) = "0"
            cSuch = Mid(cSuch, 2, 6)
            cSQL = cSQL & "and B.ARTNR = " & cSuch & " "
            
        ElseIf Len(cSuch) <= 8 And (Left(cSuch, 1) = "0") Then
            
            cSQL = cSQL & "and (A.EAN = '" & cSuch & "' "
            cSQL = cSQL & "or A.EAN2 = '" & cSuch & "' "
            cSQL = cSQL & "or A.EAN3 = '" & cSuch & "' )"
            
        Else
            If bEAN Then
                cSQL = cSQL & "and (A.EAN = '" & cSuch & "' "
                cSQL = cSQL & "or A.EAN2 = '" & cSuch & "' "
                cSQL = cSQL & "or A.EAN3 = '" & cSuch & "' )"
            Else
                cSQL = cSQL & "and A.LIBESNR = '" & cSuch & " ' "
            End If
        End If
    End If
    
    
     cSQL = cSQL & " and ( A.SYNSTATUS is null or A.SYNSTATUS = 'E' or A.SYNSTATUS = 'A' ) "
    bgefunden = False
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        cSQL = "Select * from ARTIKEL where ARTNR = " & cSuch & " "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If rsrs.EOF Then
            If Len(cSuch) = 8 And Left(cSuch, 1) = "2" Then
                cSuch = Mid(cSuch, 2, 6)
            
                rsrs.Close: Set rsrs = Nothing
                cSQL = "Select * from ARTIKEL where ARTNR = " & cSuch & " "
                Set rsrs = gdBase.OpenRecordset(cSQL)
                If rsrs.EOF Then
                
                Else
                    bgefunden = True
                End If
            ElseIf Len(cSuch) = 8 And Left(cSuch, 1) = "0" Then
                cSuch = Mid(cSuch, 2, 6)
            
                rsrs.Close: Set rsrs = Nothing
                cSQL = "Select * from ARTIKEL where ARTNR = " & cSuch & " "
                Set rsrs = gdBase.OpenRecordset(cSQL)
                If rsrs.EOF Then
                
                Else
                    bgefunden = True
                End If

            End If
        Else
            bgefunden = True
        End If
    Else
        bgefunden = True
    End If
    
    If bgefunden = True Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!artnr) Then
            cArtNr = rsrs!artnr
        Else
            cArtNr = ""
        End If
        cArtNr = Trim$(cArtNr)
        Text1(0).Text = cArtNr
        
        If Not IsNull(rsrs!BEZEICH) Then
            cArtBez = rsrs!BEZEICH
        Else
            cArtBez = ""
        End If
        cArtBez = Trim$(cArtBez)
        
        If Not IsNull(rsrs!EAN) Then
            cEAN = rsrs!EAN
        Else
            cEAN = ""
        End If
        
        
        
    
        If Not IsNull(rsrs!BESTAND) Then
            dBestand = rsrs!BESTAND
        Else
            dBestand = 0
        End If
    
        If Not IsNull(rsrs!vkpr) Then
            dVkPr = rsrs!vkpr
        Else
            dVkPr = 0
        End If
    
        If Not IsNull(rsrs!KVKPR1) Then
            dKVKPR = rsrs!KVKPR1
        Else
            dKVKPR = 0
        End If
        
        If Not IsNull(rsrs!MINBEST) Then
            lMinBest = rsrs!MINBEST
        Else
            lMinBest = 0
        End If
        
        If Not IsNull(rsrs!LIBESNR) Then
            cLiBesNr = rsrs!LIBESNR
        Else
            cLiBesNr = ""
        End If
        cLiBesNr = Trim$(cLiBesNr)
        
    Else
        MsgBox "Artikel nicht gefunden!", vbInformation, "INFO"
        
    End If
    
    rsrs.Close: Set rsrs = Nothing

    If bgefunden Then
        

        Label2(0).Caption = cArtBez
        Label2(1).Caption = dBestand
        Label2(2).Caption = cArtNr
        Label2(3).Caption = Format$(dVkPr, "##,##0.00") & " " & gcWaehrung
        Label2(5).Caption = Format$(dKVKPR, "##,##0.00") & " " & gcWaehrung


        Text1(0).Text = cEAN
'        Text1(1).SetFocus
        'der Speichern Button bekommt den Fokus
        
        Command2(15).SetFocus
        
        
        If Check1.Value = vbChecked Then 'schneller Scan
            Command2_Click 15
        End If
        
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheArtikelWKL206"
    Fehler.gsFehlertext = "Im Programmteil Lagerplatz zuweisen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
    
End Sub
Private Sub Command1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cValid As String
    Dim cFeld As String
    Dim cZeichen As String
    Dim lcount As Long
    Dim bTextSuche As Boolean
    
    Screen.MousePointer = 11
    
    cValid = "1234567890"
    cFeld = Text1(0).Text
    
    bTextSuche = False
    
    For lcount = 1 To Len(cFeld)
        cZeichen = Mid(cFeld, lcount, 1)
        If InStr(cValid, cZeichen) = 0 Then
            bTextSuche = True
            Exit For
        End If
    Next lcount
    
    If bTextSuche Then
        gcSuch = Text1(0).Text
        gsARTNR = ""
        frmWKL70.Show 1
        Me.Refresh
        If gsARTNR <> "" Then
            Text1(0).Text = gsARTNR
            gsARTNR = ""
            Command1_Click
        End If

    Else
        SucheArtikelWKL206
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Lagerplatz zuweisen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub SchreibeListe()
    On Error GoTo LOKAL_ERROR
    
    Dim sTempstring As String

    sTempstring = new19Artikel.artnr & Space(8 - Len(CStr(new19Artikel.artnr)))
    sTempstring = sTempstring & new19Artikel.BEZEICH & Space(37 - Len(new19Artikel.BEZEICH))
    sTempstring = sTempstring & Space(15 - Len(Trim(CStr(new19Artikel.LAGERPLATZ)))) & new19Artikel.LAGERPLATZ

    List5.AddItem sTempstring, 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeListe"
    Fehler.gsFehlertext = "Im Programmteil Lagerplatz zuweisen ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    Dim lcount As Long
    Dim ctmp As String
    lcount = Val(Label3.Caption)
    
    Select Case Index
        Case 0 To 10
            If lcount >= 0 Then
                Text1(lcount).Text = Text1(lcount).Text & Command2(Index).Caption
                Text1(lcount).SetFocus
                Text1(lcount).SelLength = Len(Text1(lcount).Text)
            End If
            
        Case Is = 11        '** Plus-Zeichen **
            If lcount = 1 Then
                If InStr(1, Text1(lcount).Text, "+") > 0 Then
                    Exit Sub
                ElseIf InStr(1, Text1(lcount).Text, "-") > 0 Then
                    ctmp = Text1(lcount).Text
                    Mid(ctmp, 1, 1) = "+"
                    Text1(lcount).Text = ctmp
                Else
                    Text1(lcount).Text = "+"
                End If
            End If
            Text1(lcount).SetFocus
            
        Case Is = 12        '** Minus-Zeichen **
            If lcount = 1 Then
                If InStr(1, Text1(lcount).Text, "-") > 0 Then
                    Exit Sub
                ElseIf InStr(1, Text1(lcount).Text, "+") > 0 Then
                    ctmp = Text1(lcount).Text
                    Mid(ctmp, 1, 1) = "-"
                    Text1(lcount).Text = ctmp
                Else
                    Text1(lcount).Text = "-"
                End If
            End If
            Text1(lcount).SetFocus
        Case Is = 13        '** Löschen **
            Text1(lcount).Text = ""
            Text1(lcount).SetFocus
            
        Case Is = 14        '** Rückgängig **
            If Len(Text1(lcount).Text) > 0 Then
                ctmp = Text1(lcount).Text
                ctmp = Left(ctmp, Len(ctmp) - 1)
                Text1(lcount).Text = ctmp
            End If
            Text1(lcount).SetFocus
            
        Case Is = 15        'Speichern
            If Trim$(Text1(0).Text) = "" Then
                If Label2(2).Caption = "0" Then
                    MsgBox "Bitte einen Artikel festlegen!", vbCritical, "STOP!"
                    Text1(0).SetFocus
                    Exit Sub
                Else
                    Text1(0).Text = Label2(2).Caption
                End If
            End If
            
            SchreibeDatenWKL206
            SchreibeListe
        Case Is = 22 'Leeren
            LeereDialogWKL15 True
            Text1(0).SetFocus
            
        Case Is = 16        'Vorheriges Feld
            If lcount > 0 Then
                Text1(lcount - 1).SetFocus
            Else
                Text1(3).SetFocus
            End If
            
        Case Is = 17        'Nächstes Feld
            If lcount < 3 Then
                Text1(lcount + 1).SetFocus
            Else
                Text1(0).SetFocus
            End If
            
        Case Is = 18        'Komma
            If lcount = 2 Or lcount = 3 Then
                If InStr(Text1(lcount).Text, ",") = 0 Then
                    Text1(lcount).Text = Text1(lcount).Text & Command2(Index).Caption
                End If
                Text1(lcount).SetFocus
                Text1(lcount).SelLength = Len(Text1(lcount).Text)
            End If
        Case Is = 19        'F2
            Text1_KeyUp Val(Label3.Caption), vbKeyF2, 0
            
        Case Is = 20        'F4
            If Text1(0).Text = "" Then
                MsgBox "Bitte den Artikel eindeutig definieren (Artikelnummer oder EAN-Code)!", vbCritical, "STOP!"
                Text1(0).SetFocus
                Exit Sub
            Else
                Text1_KeyUp Val(Label3.Caption), vbKeyF4, 0
            End If
        Case Is = 21
            If Frame1.Visible Then
                Frame1.Visible = False
            Else
                Frame1.Visible = True
            End If
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Lagerplatz zuweisen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click()
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKL206
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Lagerplatz zuweisen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    PositionierenWKL206
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    LeereDialogWKL15 False
    gF2Prompt.lLastPos = -1
    List6.AddItem " Artnr  Artikelbezeichnung                               Lagerplatz"
    Screen.MousePointer = 0
    
    Frame1.Visible = False
    
    List5.Visible = False
    List6.Visible = False
    Label11(3).Visible = False
    Command2(21).Visible = False
    
    'Manuell
    Frame3.Visible = True

    If gbBILDTAST = False Then
        Frame1.Visible = False
    Else
        Frame1.Visible = True
    End If

    List5.Visible = True
    List6.Visible = True
    Label11(3).Visible = True
    Command2(21).Visible = True
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Lagerplatz zuweisen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    Label3.Caption = Format$(Index, "##0")
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Lagerplatz zuweisen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
    Select Case Index
        Case Is = 0
            'wegen Volltextsuche nicht mehr gültig
            cValid = "1234567890" & Chr$(8)
        Case Is = 1
            cValid = "1234567890+-" & Chr$(8)
        Case Is = 2
            cValid = "1234567890," & Chr$(8)
        Case Is = 3
            cValid = "1234567890," & Chr$(8)
        Case Is = 4
            cValid = "1234567890" & Chr$(8)
        Case Is = 5
            cValid = "1234567890" & Chr$(8)
    End Select
    
    If Index <> 0 And Index <> 6 Then
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        End If
    End If
    If Index = 2 And cZeichen = "," Then
        If InStr(Text1(Index).Text, ",") > 0 Then
            KeyAscii = 0
        End If
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Lagerplatz zuweisen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            Command1_Click
        End If
        If Index >= 1 Then
            Command2_Click 15
        End If
    End If
    
    If KeyCode = vbKeyF4 Then
        If Index = 0 Then
            ctmp = Trim$(Text1(4).Text)
            If ctmp = "" Then
                MsgBox "Bitte einen Lieferanten angeben!", vbCritical, "STOP!"
                Text1(4).SetFocus
                Exit Sub
            End If
            
            gF2Prompt.cFeld = "ARTNRPOS"
            gF2Prompt.cWert = ctmp
            gF2Prompt.cWert2 = ""
            gF2Prompt.cWahl = ""
            gF2Prompt.bMultiple = False
            
            Command1_Click
            ctmp = Trim$(Text1(0).Text)
            If ctmp = "" Then
                MsgBox "Bitte den Artikel eindeutig bestimmen (Artikelnummer oder EAN-Code)!", vbCritical, "STOP!"
                Text1(0).SetFocus
                Exit Sub
            End If
            gF2Prompt.cWert2 = ctmp
        
        If gF2Prompt.cFeld <> "" Then
            
            frmWK00a.Show 1
        
            If gF2Prompt.cWahl <> "" Then
                Text1(Index).Text = gF2Prompt.cWahl
                If Index = 0 Then
                    Command1_Click
                End If
            End If
            
        End If
        
        End If
    End If
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False

        Select Case Index
            Case Is = 0     'Artikel
                ctmp = Trim$(Text1(4).Text)
                If ctmp = "" Then
                    MsgBox "Bitte einen Lieferanten angeben!", vbCritical, "STOP!"
                    Text1(4).SetFocus
                    Exit Sub
                Else
                    gF2Prompt.cFeld = "ARTNRPOS"
                    gF2Prompt.cWert = ctmp
                End If
            
            Case Is = 4     'Lieferant
                gF2Prompt.cFeld = "LINR"
        End Select
        
        If gF2Prompt.cFeld <> "" Then
            
            frmWK00a.Show 1
        
            If gF2Prompt.cWahl <> "" Then
                Text1(Index).Text = gF2Prompt.cWahl
                If Index = 0 Then
                    Command1_Click
                End If
            End If
            
        End If
        
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Lagerplatz zuweisen ist ein Fehler aufgetreten. "
    
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
    Fehler.gsFehlertext = "Im Programmteil Lagerplatz zuweisen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub


