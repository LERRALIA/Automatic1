VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWKL19 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Bestandskorrektur"
   ClientHeight    =   8625
   ClientLeft      =   2145
   ClientTop       =   2655
   ClientWidth     =   11910
   Icon            =   "frmWKL19.frx":0000
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
      Left            =   4680
      TabIndex        =   37
      Top             =   5880
      Visible         =   0   'False
      Width           =   1200
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
      ItemData        =   "frmWKL19.frx":0442
      Left            =   240
      List            =   "frmWKL19.frx":0444
      TabIndex        =   47
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
      ItemData        =   "frmWKL19.frx":0446
      Left            =   240
      List            =   "frmWKL19.frx":0448
      TabIndex        =   83
      Top             =   7080
      Width           =   7695
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'Kein
      Height          =   735
      Left            =   8640
      TabIndex        =   64
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "addieren"
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
         Index           =   2
         Left            =   6600
         TabIndex        =   93
         Top             =   6000
         Width           =   2655
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   82
         Top             =   4560
         Width           =   1095
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
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "ersetzen"
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
         Left            =   6600
         TabIndex        =   81
         Top             =   5520
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "subtrahieren"
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
         Index           =   0
         Left            =   6600
         TabIndex        =   80
         Top             =   5760
         Width           =   2655
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   6000
         TabIndex        =   65
         Top             =   1080
         Width           =   5535
      End
      Begin VB.ListBox List11 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   120
         TabIndex        =   66
         Top             =   1080
         Width           =   5535
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   8
         Left            =   10440
         TabIndex        =   72
         Top             =   4560
         Width           =   1095
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
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List4 
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
         Left            =   6000
         TabIndex        =   71
         Top             =   840
         Width           =   5535
      End
      Begin VB.ListBox List12 
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
         Left            =   120
         TabIndex        =   70
         Top             =   840
         Width           =   5535
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   11
         Left            =   9360
         TabIndex        =   69
         Top             =   6240
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
         Height          =   495
         Index           =   12
         Left            =   9360
         TabIndex        =   68
         Top             =   5520
         Visible         =   0   'False
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
         Caption         =   "Einlesen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Rechts ausgerichtet
         BackColor       =   &H00C0C000&
         Caption         =   "mit Etikettenerstellung"
         Height          =   210
         Left            =   9360
         TabIndex        =   67
         Top             =   5280
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComctlLib.ProgressBar pbr1 
         Height          =   375
         Left            =   120
         TabIndex        =   73
         Top             =   6360
         Visible         =   0   'False
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   "Daten aus dem MDE Gerät "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "insgesamt:"
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
         Index           =   4
         Left            =   6000
         TabIndex        =   77
         Top             =   4560
         Width           =   4335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "insgesamt:"
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
         TabIndex        =   76
         Top             =   4560
         Width           =   4335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel, die nicht zugeordnet werden können"
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
         Left            =   6000
         TabIndex        =   75
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel aus dem MDE - Gerät"
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
         TabIndex        =   74
         Top             =   600
         Width           =   5295
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'Kein
      Height          =   1695
      Left            =   9960
      TabIndex        =   57
      Top             =   7320
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CheckBox Check7 
         Caption         =   "Original-EAN verwendet"
         Height          =   255
         Left            =   9720
         TabIndex        =   87
         Top             =   6000
         Width           =   2295
      End
      Begin sevCommand3.Command Command7 
         Height          =   495
         Index           =   3
         Left            =   9720
         TabIndex        =   61
         Top             =   6360
         Width           =   1815
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
      Begin sevCommand3.Command Command7 
         Height          =   495
         Index           =   2
         Left            =   9720
         TabIndex        =   60
         Top             =   6960
         Width           =   1815
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
      Begin VB.PictureBox picprogress 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   8835
         TabIndex        =   59
         Top             =   6360
         Width           =   8895
      End
      Begin VB.TextBox txtStatus 
         Height          =   315
         Left            =   120
         TabIndex        =   58
         Top             =   6000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   79
         Top             =   6840
         Width           =   8895
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   "Daten aus dem MDE Gerät "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   6135
      End
      Begin VB.Image Image2 
         Height          =   915
         Left            =   10800
         MouseIcon       =   "frmWKL19.frx":044A
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL19.frx":0754
         ToolTipText     =   "Klicken Sie hier, wenn Sie Daten aus dem MDE - Gerät einlesen möchten"
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   $"frmWKL19.frx":0D37
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   5
         Left            =   2040
         TabIndex        =   62
         Top             =   2160
         Width           =   7935
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      Height          =   735
      Left            =   7320
      TabIndex        =   52
      Top             =   120
      Width           =   975
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "manuell mit Scanner"
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
         Left            =   240
         TabIndex        =   55
         Top             =   1560
         Width           =   6615
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "mit dem MDE - Gerät"
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
         Left            =   240
         TabIndex        =   54
         Top             =   2280
         Width           =   6615
      End
      Begin sevCommand3.Command Command7 
         Height          =   495
         Index           =   1
         Left            =   9720
         TabIndex        =   0
         Top             =   6360
         Width           =   1815
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
         Caption         =   "Weiter"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   495
         Index           =   0
         Left            =   9720
         TabIndex        =   53
         Top             =   6960
         Width           =   1815
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
      Begin VB.Label Label6 
         Caption         =   "Wie möchten Sie bei der Bestandskorrektur vorgehen?"
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
         Left            =   240
         TabIndex        =   56
         Top             =   360
         Width           =   10935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'Kein
      Height          =   5415
      Left            =   960
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   11175
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   9600
         MaxLength       =   3
         TabIndex        =   92
         Top             =   3120
         Width           =   855
      End
      Begin VB.CheckBox Check5 
         Caption         =   "als Lieferantenbestellnummer suchen"
         Height          =   255
         Left            =   4560
         TabIndex        =   90
         Top             =   0
         Width           =   3375
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Original-EAN"
         Height          =   255
         Left            =   2520
         TabIndex        =   89
         Top             =   20
         Width           =   1815
      End
      Begin VB.CheckBox Check3 
         Caption         =   "direkt hierher"
         Height          =   255
         Left            =   240
         TabIndex        =   86
         Top             =   4200
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Druckdaten löschen"
         Height          =   255
         Left            =   3120
         TabIndex        =   84
         Top             =   4200
         Width           =   2895
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   0
         Left            =   6120
         TabIndex        =   51
         Top             =   3600
         Visible         =   0   'False
         Width           =   1800
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
         Caption         =   "Bestandshistorie"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command8 
         Height          =   375
         Left            =   5880
         TabIndex        =   50
         Top             =   2760
         Visible         =   0   'False
         Width           =   1320
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
         Caption         =   "in Filialen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   615
         Left            =   3120
         TabIndex        =   46
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "LiefNr halten"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   9480
         TabIndex        =   45
         Top             =   4440
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "LiefNr leeren"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   9480
         TabIndex        =   44
         Top             =   4800
         Value           =   -1  'True
         Width           =   2055
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
         Index           =   4
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   3
         Top             =   960
         Width           =   2295
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
         Height          =   495
         Index           =   1
         Left            =   3960
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3360
         Width           =   2055
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Left            =   240
         TabIndex        =   6
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   15
         Left            =   6000
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
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   5055
      End
      Begin sevCommand3.Command Command1 
         Height          =   615
         Left            =   8040
         TabIndex        =   2
         Top             =   360
         Width           =   2775
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
         Caption         =   "Suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   4200
         TabIndex        =   96
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   95
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   94
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "MB:"
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
         Index           =   6
         Left            =   7560
         TabIndex        =   91
         Top             =   3120
         Width           =   1935
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
         TabIndex        =   85
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
            Size            =   14.25
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
         TabIndex        =   42
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Kassen-Vk:"
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
         Index           =   9
         Left            =   7560
         TabIndex        =   41
         Top             =   2640
         Width           =   1935
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
         Index           =   4
         Left            =   5040
         TabIndex        =   40
         Top             =   1080
         Width           =   6735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferant:"
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
         Index           =   8
         Left            =   120
         TabIndex        =   39
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "neuer Bestand:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   36
         Top             =   3360
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Listen-Vk:"
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
         Index           =   5
         Left            =   7560
         TabIndex        =   35
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 Euro"
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
         Caption         =   "alter Bestand:"
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
         Left            =   360
         TabIndex        =   33
         Top             =   2640
         Width           =   3495
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
         Left            =   5040
         TabIndex        =   32
         Top             =   1560
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
         Left            =   3960
         TabIndex        =   31
         Top             =   2640
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
         Top             =   1560
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
         Top             =   1560
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
         Top             =   480
         Width           =   2175
      End
   End
   Begin sevCommand3.Command Command2 
      Height          =   360
      Index           =   21
      Left            =   11280
      TabIndex        =   88
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
      TabIndex        =   49
      Top             =   6480
      Width           =   11415
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "zuletzt bearbeitete Bestände"
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
      TabIndex        =   48
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
      Caption         =   "Bestandskorrektur"
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
      TabIndex        =   43
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmWKL19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim new19Artikel As ArtikelTyp

Private Sub Check1_Click()
On Error GoTo LOKAL_ERROR

If Check1.Value = vbChecked Then
    loeschapp "beTemp"
    Command5.BackColor = Command3.BackColor
    Check1.Value = vbUnchecked
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub UpdateBenex()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    If Check3.Value = vbChecked Then
        sSQL = "Update BEKENX set Ind2 = 0 "
        gdApp.Execute sSQL, dbFailOnError
    ElseIf Check3.Value = vbUnchecked Then
        sSQL = "Update BEKENX set Ind2 = -1 "
        gdApp.Execute sSQL, dbFailOnError
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UpdateBenex"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Command5_Click()
On Error GoTo LOKAL_ERROR

Dim lLoeschen As Long
Dim cSQL As String

If tableSuchenDBKombi("beTemp", 2) Then



    If Not SpalteInTabellegefundenNEW("Betemp", "AENGRUND", gdApp) Then
        SpalteAnfuegenNEW "Betemp", "AENGRUND", "Text(20)", gdApp
        
        cSQL = "Update Betemp set AENGRUND =''"
        gdApp.Execute cSQL, dbFailOnError
        
    End If





    lLoeschen = MsgBox("Druckdaten nach dem Drucken löschen?", vbQuestion + vbYesNo, "Winkiss Frage:")
    
    loeschNEW "FILA", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "FILA"
    
    reportbildschirmApp "dWKL19", "aWKL19"
Else
    MsgBox "Es sind keine Druckdaten vorhanden.", vbInformation, "Winkiss Hinweis:"
End If
    
If lLoeschen = vbYes Then
    loeschapp "beTemp"
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command6_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            gsARTNR = Label2(2).Caption
            frmWKL78.Show 1
        Case Is = 1
            einlesenausmdeVorschlag
        Case Is = 8
            reportbildschirm "umv1a", "aWKL19c"
        Case Is = 12
            einlesenausMDE
        Case Is = 11 'zurück Dateien Zentrale 1
            Frame8.Visible = False
            Frame5.Visible = True
            
            
            
            
            
            anzeigeNew "normal", "", Label5
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command7_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0
            Unload frmWKL19
        Case Is = 1
            Zeigeauswahlframe
        Case Is = 2
            Frame5.Visible = False
            Frame4.Visible = True
            
            List5.Visible = False
            List6.Visible = False
            Label11(3).Visible = False
            Frame1.Visible = False
            Command2(21).Visible = False
        Case Is = 3
            If MDEeinlesenOhneLinr(Label5, txtStatus, picprogress, frmWKL19) = False Then
                anzeigeNew "rot", "Es konnten keine Daten aus dem MDE - Gerät ausgelesen werden.", Label5
            Else
                Frame5.Visible = False
                Frame8.Visible = True
                anzeigeNew "normal", "", Label5

                MdeVerarbeitung1
            End If
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub einlesenausMDE()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rs          As Recordset
    Dim lCounter    As Long
    Dim dBestand    As Double
    Screen.MousePointer = 11
    
    loeschNEW "KORRY", gdBase
    
    sSQL = "Select sum(bestvor) as lmaxanz"
    sSQL = sSQL & ", ARTNR "
    sSQL = sSQL & ", BEZEICH "
    sSQL = sSQL & ", LINR "
    sSQL = sSQL & ", LPZ "
    sSQL = sSQL & ", LIBESNR "
    sSQL = sSQL & ", KVKPR1 "
    sSQL = sSQL & ", BESTAND "
    sSQL = sSQL & ", FILIALE "
    sSQL = sSQL & ", 0 as BESTANDN "
    sSQL = sSQL & ", 0 as FARBNR "
    sSQL = sSQL & ", 0 as FARBwert "
    sSQL = sSQL & ", 0 as FARBwertS "
    sSQL = sSQL & ", '' as FARBTEXT "
    sSQL = sSQL & " into KORRY from KORREKB where Status = 'vorhanden' "
    sSQL = sSQL & " group by "
    sSQL = sSQL & " ARTNR "
    sSQL = sSQL & ", BEZEICH "
    sSQL = sSQL & ", LINR "
    sSQL = sSQL & ", LPZ "
    sSQL = sSQL & ", LIBESNR "
    sSQL = sSQL & ", KVKPR1 "
    sSQL = sSQL & ", BESTAND "
    sSQL = sSQL & ", FILIALE "
    sSQL = sSQL & " order by linr,lpz "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Update KORRY inner join Artikel on KORRY.ARTNR = Artikel.Artnr "
    sSQL = sSQL & " set  KORRY.FARBNR = val(ARTIKEL.awm) "
    gdBase.Execute sSQL, dbFailOnError

    BringFarbeInsSpiel "KORRY", gdBase
    
    Set rs = gdBase.OpenRecordset("KORRY")
    
    If rs.EOF Then
        anzeigeNew "rot", "Keine Artikeldaten zur Verarbeitung vorhanden", Label5
        Screen.MousePointer = 0
        rs.Close: Set rs = Nothing
        Exit Sub
    End If
    
    pbr1.Max = 50
    pbr1.Visible = True
    
    lCounter = 0
    rs.MoveFirst
    If Not rs.EOF Then
        anzeigeNew "normal", "Die Bestandskorrektur wird jetzt eingelesen...", Label5
        Do While Not rs.EOF
            If lCounter = 50 Then
                lCounter = 0
            End If
            lCounter = lCounter + 1
            pbr1.Value = lCounter
            
            If Not IsNull(rs!artnr) Then
                If Not IsNull(rs!lmaxanz) Then
                    If Option3(1).Value = True Then
                        Bestandsveraenderung rs!artnr, CLng(rs!lmaxanz), "Bestandskorrektur MD"
                    ElseIf Option3(0).Value = True Then
                        If Not IsNull(rs!BESTAND) Then
                            dBestand = rs!BESTAND
                        Else
                            dBestand = 0
                        End If
                        
                        dBestand = dBestand - CLng(rs!lmaxanz)
                        rs.Edit
                        rs!BESTANDN = dBestand
                        rs.Update
                        
                        Bestandsveraenderung rs!artnr, CLng(dBestand), "Bestandskorrektur MD"
                    ElseIf Option3(2).Value = True Then
                        If Not IsNull(rs!BESTAND) Then
                            dBestand = rs!BESTAND
                        Else
                            dBestand = 0
                        End If
                        
                        dBestand = dBestand + CLng(rs!lmaxanz)
                        rs.Edit
                        rs!BESTANDN = dBestand
                        rs.Update
                        
                        Bestandsveraenderung rs!artnr, CLng(dBestand), "Bestandskorrektur MD"
                    End If
                End If
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    
    pbr1.Visible = False
    Screen.MousePointer = 0
    anzeigeNew "normal", "Die Aktualisierung wurde erfolgreich durchgeführt.", Label5
    
    Check2.Visible = False
    Command6(12).Visible = False
    
    If Option3(1).Value = True Then
        reportbildschirm "umv1", "aWKL19d"
    ElseIf Option3(0).Value = True Then
        reportbildschirm "umv1", "aWKL19e"
    ElseIf Option3(2).Value = True Then
        reportbildschirm "umv1", "aWKL19e"
    End If
    
    Exit Sub
LOKAL_ERROR:
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "einlesenausmde"
        Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten."
        
        Fehlermeldung1
End Sub
Private Sub einlesenausmdeVorschlag()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rs          As Recordset
    Dim lCounter    As Long
    Dim dBestand    As Double
    Screen.MousePointer = 11
    
    
    loeschNEW "KORRYVOR", gdBase
    
    sSQL = "Select sum(bestvor) as lmaxanz"
    sSQL = sSQL & ", ARTNR "
    sSQL = sSQL & ", BEZEICH "
    sSQL = sSQL & ", LINR "
    sSQL = sSQL & ", LPZ "
    sSQL = sSQL & ", LIBESNR "
    sSQL = sSQL & ", KVKPR1 "
    sSQL = sSQL & ", BESTAND "
    sSQL = sSQL & ", FILIALE "
    sSQL = sSQL & ", 0 as BESTANDN "
    
    sSQL = sSQL & " into KORRYVOR from KORREKB where Status = 'vorhanden' "
    sSQL = sSQL & " group by "
    sSQL = sSQL & " ARTNR "
    sSQL = sSQL & ", BEZEICH "
    sSQL = sSQL & ", LINR "
    sSQL = sSQL & ", LPZ "
    sSQL = sSQL & ", LIBESNR "
    sSQL = sSQL & ", KVKPR1 "
    sSQL = sSQL & ", BESTAND "
    sSQL = sSQL & ", FILIALE "
    sSQL = sSQL & " order by linr,lpz "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Set rs = gdBase.OpenRecordset("KORRYVOR")
    
    If rs.EOF Then
        anzeigeNew "rot", "Keine Artikeldaten zur Druckansicht vorhanden", Label5
        Screen.MousePointer = 0
        rs.Close: Set rs = Nothing
        Exit Sub
    End If
    
    
    pbr1.Max = 50
    pbr1.Visible = True
    
    lCounter = 0
    rs.MoveFirst
    If Not rs.EOF Then
        anzeigeNew "normal", "Die Druckansicht wird erstellt...", Label5
        Do While Not rs.EOF
            If lCounter = 50 Then
                lCounter = 0
            End If
            lCounter = lCounter + 1
            pbr1.Value = lCounter
            
            If Not IsNull(rs!artnr) Then
                If Not IsNull(rs!lmaxanz) Then
                    If Option3(1).Value = True Then

                    ElseIf Option3(0).Value = True Then
                        If Not IsNull(rs!BESTAND) Then
                            dBestand = rs!BESTAND
                        Else
                            dBestand = 0
                        End If
                        
                        dBestand = dBestand - CLng(rs!lmaxanz)
                        rs.Edit
                        rs!BESTANDN = dBestand
                        rs.Update
                        
                    ElseIf Option3(2).Value = True Then
                        If Not IsNull(rs!BESTAND) Then
                            dBestand = rs!BESTAND
                        Else
                            dBestand = 0
                        End If
                        
                        dBestand = dBestand + CLng(rs!lmaxanz)
                        rs.Edit
                        rs!BESTANDN = dBestand
                        rs.Update
                        
                    End If
                End If
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    
    pbr1.Visible = False
    Screen.MousePointer = 0
    anzeigeNew "normal", "", Label5
    
    
    
    If Option3(1).Value = True Then
        reportbildschirm "umv1", "aWKL19f" '"aWKL19d"
    Else
        reportbildschirm "umv1", "aWKL19g" '"aWKL19e"
    End If
    
    Exit Sub
LOKAL_ERROR:
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "einlesenausmdeVorschlag"
        Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten."
        
        Fehlermeldung1
End Sub

Private Sub Zeigeauswahlframe()
    On Error GoTo LOKAL_ERROR
    
    Frame4.Visible = False
    
    If Option2(0).Value = True Then         'Manuell
        Frame3.Visible = True
        Text1(0).SetFocus
        If gbBILDTAST = False Then
            Frame1.Visible = False
        Else
            Frame1.Visible = True
        End If

        List5.Visible = True
        List6.Visible = True
        Label11(3).Visible = True
        Command2(21).Visible = True

    ElseIf Option2(2).Value = True Then     'Mde
        Frame5.Visible = True
        
        List5.Visible = False
        List6.Visible = False
        Label11(3).Visible = False
        Frame1.Visible = False
        Command2(21).Visible = False
        
        Select Case gsMDEGERAET
        
                Case "SCANPAL"
                
                    lbl6(5).Caption = "Gerät richtig einstellen! Am Scanpal 2 auf 'Daten senden' navigieren und mit der Enter - Taste auf dem Scanpal 2 bestätigen. Wenn auf dem Display des Scanpal 2 'Verbindung....' steht, dann können Sie den hier unten aufgeführten Button 'Einlesen' klicken."
            
                Case "CIPHERLAB"
            
                    lbl6(5).Caption = "Gerät richtig einstellen! Am Cipherlab auf 'Daten senden' navigieren und mit der Enter - Taste auf dem Cipherlab bestätigen. Wenn auf dem Display des Cipherlab 'Verbindung....' steht, dann können Sie den hier unten aufgeführten Button 'Einlesen' klicken."
        
                Case "FORCOM"
                    
                    lbl6(5).Caption = "Das Formula in die Station stecken - dann im Menü 'Übertragen' anwählen"
                    lbl6(5).Caption = lbl6(5).Caption & " dann Enter auf dem Formula drücken und danach hier im Programm auf 'Einlesen' klicken."
        
                Case Else
                    lbl6(5).Caption = ""
            End Select
            
            lbl6(5).Refresh

    End If
    
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
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
Private Sub PositionierenWKL15()
    On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 6240
    Frame1.Left = 0
    Frame1.Height = 2775
    Frame1.Width = 12000
    
   
    Frame3.Top = 840
    Frame3.Left = 0
    Frame3.Height = 6000
    Frame3.Width = 12000
    
    Frame4.Top = 840
    Frame4.Left = 120
    Frame4.Height = 7695
    Frame4.Width = 11655
    
    Frame5.Top = 840
    Frame5.Left = 120
    Frame5.Height = 7695
    Frame5.Width = 11655
    
    Frame8.Top = 840
    Frame8.Left = 120
    Frame8.Height = 7695
    Frame8.Width = 11655
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL15"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub LeereDialogWKL15()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""

    If Option1(1).Value Then
        Text1(4).Text = ""
        Label2(4).Caption = ""
    End If
    
    Label2(6).Caption = ""
    Label2(7).Caption = ""
    Label2(8).Caption = ""
    
    
    
    Label2(0).Caption = "unbekannt"
    Label2(1).Caption = "0"
    Label2(2).Caption = "0"
    Label2(3).Caption = "0,00 " & gcWaehrung
    Label2(5).Caption = "0,00 " & gcWaehrung
    
    Label3.Caption = "0"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKL15"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub SchreibeDatenWKL15()
    On Error GoTo LOKAL_ERROR
        
    Dim ctmp                    As String
    Dim cSQL                    As String
    Dim cArtNr                  As String
    Dim cBezeich                As String
    Dim dAnzahl                 As Double
    Dim dBestand                As Double
    Dim dMindestBestand         As Double
    Dim dBWert                  As Double
    Dim dAltBestand             As Double
    Dim iFehlerstufe            As Integer
    Dim iRet                    As Integer
    Dim bTrans                  As Boolean
    Dim rsrs                    As Recordset
    Dim rsEti                   As Recordset
    Dim rsArt                   As Recordset
    Dim izubuchmenge            As Long
    Dim rsbetemp                As Recordset
    Dim cJetzt                  As String

    cArtNr = Label2(2).Caption
    
    ctmp = Trim$(Text1(1).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    dBWert = Val(ctmp)
    
    iFehlerstufe = 0
    bTrans = False

    ctmp = Trim$(Text1(1).Text)
    
    If ctmp = "" Then
        ctmp = "0"
    End If
    dAnzahl = Val(ctmp)
    
    If Abs(dAnzahl) > 99 Then
        iRet = MsgBox("Neue Bestandsangabe von " & ctmp & " fraglich! Trotzdem speichern?", vbQuestion + vbYesNo + vbDefaultButton2, "Neuer Bestand")
        If iRet = vbNo Then
            Text1(1).SetFocus
            Exit Sub
        End If
    End If
    
    cArtNr = Label2(2).Caption
    cArtNr = Trim$(cArtNr)
    If cArtNr = "" Then
        MsgBox "Artikel-Nr fehlt! Daten speichern nicht möglich!", vbCritical, "FEHLER2"
        Text1(0).SetFocus
        Exit Sub
    End If
    
    Bestandsveraenderung cArtNr, CLng(dAnzahl), "Bestandskorrektur"

    If Not tableSuchenDBKombi("Betemp", 2) Then
        cSQL = "Create Table Betemp "
        cSQL = cSQL & "( "
        cSQL = cSQL & "ARTNR LONG"
        cSQL = cSQL & ", BEZEICH Text (35) "
        cSQL = cSQL & ", LEKPR DOUBLE "
        cSQL = cSQL & ", ADATE DATETIME "
        cSQL = cSQL & ", UHRZEIT TEXT (5) "
        cSQL = cSQL & ", BEDNU long "
        cSQL = cSQL & ", BEDNAME TEXT (32) "
        cSQL = cSQL & ", FILIALNR BYTE "
        cSQL = cSQL & ", BESTANDALT INTEGER "
        cSQL = cSQL & ", BEWEGUNG INTEGER "
        cSQL = cSQL & ", BESTANDNEU INTEGER "
        cSQL = cSQL & ", AENGRUND Text(20)"
        cSQL = cSQL & ") "
        gdApp.Execute cSQL, dbFailOnError
    Else
        If Not SpalteInTabellegefundenNEW("Betemp", "AENGRUND", gdApp) Then
            SpalteAnfuegenNEW "Betemp", "AENGRUND", "Text(20)", gdApp
            
            cSQL = "Update Betemp set AENGRUND =''"
            gdApp.Execute cSQL, dbFailOnError
            
        End If
    End If
               
    Set rsbetemp = gdApp.OpenRecordset("betemp", dbOpenTable)
    
    cSQL = "Select * from Artikel where ARTNR = " & cArtNr
    Set rsArt = gdBase.OpenRecordset(cSQL)
    
    If IsNumeric(Text1(2).Text) Then
        rsArt.Edit
        rsArt!MINBEST = Text1(2).Text
        rsArt.Update
    End If
    
    cSQL = "Select * from ETIDRU where ARTNR = " & cArtNr
    cSQL = cSQL & " and FILNR = " & gcFilNr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
       rsrs.Edit
       rsrs!artnr = cArtNr
        rsrs!BEZEICH = cBezeich
        rsrs!vkpr = rsArt!KVKPR1
        
        If Not IsNull(rsrs!BESTAND) Then
            rsrs!BESTAND = dAnzahl
        Else
            rsrs!BESTAND = dAnzahl
        End If
        
        If Not IsNull(rsrs!ANZAHL) Then
            rsrs!ANZAHL = dAnzahl
        Else
            rsrs!ANZAHL = dAnzahl
        End If
        rsrs!BEZEICH = rsArt!BEZEICH
        rsrs!LIBESNR = rsArt!LIBESNR
        
        rsrs!EAN = rsArt!EAN
        rsrs!linr = rsArt!linr
        rsrs!LPZ = rsArt!LPZ
        rsrs!filnr = gcFilNr
        rsrs!Pcname = srechnertab
        rsrs.Update
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    izubuchmenge = CLng(rsArt!BESTAND) - CLng(Label2(1).Caption)
    
    new19Artikel.artnr = rsArt!artnr
    new19Artikel.BEZEICH = rsArt!BEZEICH
    new19Artikel.BESTAND = rsArt!BESTAND
    new19Artikel.OLDBESTAND = CLng(Label2(1).Caption)
    new19Artikel.ZubuchMe = izubuchmenge
    
    cJetzt = Format$(Now, "HH:MM")
    
    rsbetemp.AddNew
    rsbetemp!artnr = rsArt!artnr
    rsbetemp!BEZEICH = rsArt!BEZEICH
    rsbetemp!lekpr = rsArt!lekpr
    rsbetemp!Adate = DateValue(Now)
    rsbetemp!Uhrzeit = cJetzt
    rsbetemp!BEDNU = Val(gcBedienerNr)
    rsbetemp!bedname = gcUserName
    rsbetemp!FILIALNR = Val(gcFilNr)
    rsbetemp!bestandalt = CLng(Label2(1).Caption)
    rsbetemp!BEWEGUNG = izubuchmenge
    rsbetemp!BESTANDneu = rsArt!BESTAND
    rsbetemp!AENGRUND = sGlobAenderGRUND
    rsbetemp.Update
    rsbetemp.Close
    
    rsArt.Close: Set rsArt = Nothing

    LeereDialogWKL15
    Text1(0).SetFocus

Exit Sub
LOKAL_ERROR:
    If bTrans Then
        Rollback
    End If
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWKL15"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
    Resume Next
    
End Sub
Private Sub SucheArtikelWKL15()
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
    Dim dMindestBestand As Double
    Dim dVkPr As Double
    Dim dKVKPR As Double
    Dim dEkpr As Double
    Dim dLEKPR As Double
    Dim cLinr As String
    Dim cLiefBez As String
    Dim cLiBesNr As String
    Dim bgefunden As Boolean
    Dim cFeld As String
    Dim cLBSatz As String
    Dim lMinBest As Long
    Dim bEAN As Boolean
    Dim cEAN As String
    
    
    
    Dim cEAN2 As String
    Dim cEAN3 As String
    
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
    
    cLinr = Text1(4).Text
    cLinr = Trim$(cLinr)
    
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
        cSQL = "Select B.ARTNR, A.BEZEICH, B.LINR, A.BESTAND, A.VKPR, A.KVKPR1, A.MINBEST, B.LIBESNR, A.EAN,A.EAN2,A.EAN3 "
        cSQL = cSQL & "from ARTIKEL A, ARTLIEF B where A.ARTNR = B.ARTNR "
    End If
    
    If Len(cSuch) <= 6 Then
        cSQL = "Select B.ARTNR, A.BEZEICH, B.LINR, A.BESTAND, A.VKPR, A.KVKPR1, A.MINBEST, B.LIBESNR, A.EAN,A.EAN2,A.EAN3 "
        cSQL = cSQL & "from ARTIKEL A, ARTLIEF B where B.ARTNR = " & cSuch & " and A.ARTNR = B.ARTNR "
    Else
        If Len(cSuch) <= 8 And (Left(cSuch, 1) = "2") Then  'Or Left(cSuch, 1) = "0"
        
        
            If Check4.Value = vbChecked Then
                cSQL = cSQL & " and (A.EAN = '" & cSuch & "' "
                cSQL = cSQL & "or A.EAN2 = '" & cSuch & "' "
                cSQL = cSQL & "or A.EAN3 = '" & cSuch & "' )"
            
            Else
                cSuch = Mid(cSuch, 2, 6)
                cSQL = cSQL & " and B.ARTNR = " & cSuch & " "
            End If
            
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
    
    If Len(cLinr) > 0 Then
        cSQL = cSQL & "and B.LINR = " & cLinr & " "
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
        
        If Not IsNull(rsrs!EAN2) Then
            cEAN2 = rsrs!EAN2
        Else
            cEAN2 = ""
        End If
        
        If Not IsNull(rsrs!EAN3) Then
            cEAN3 = rsrs!EAN3
        Else
            cEAN3 = ""
        End If
        
        If Not IsNull(rsrs!linr) Then
            cLinr = rsrs!linr
        Else
            cLinr = "-1"
        End If
        cLinr = Trim$(cLinr)
        
    
        If Not IsNull(rsrs!BESTAND) Then
            dBestand = rsrs!BESTAND
        Else
            dBestand = 0
        End If
        
        If Not IsNull(rsrs!MINBEST) Then
            dMindestBestand = rsrs!MINBEST
        Else
            dMindestBestand = 0
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
        cLiefBez = ermLiefBez(CLng(cLinr))

        Label2(0).Caption = cArtBez
        Label2(1).Caption = dBestand
        
        
        Label2(6).Caption = cEAN
        Label2(7).Caption = cEAN2
        Label2(8).Caption = cEAN3
        
        
        
        Text1(2).Text = dMindestBestand
        Label2(2).Caption = cArtNr
        Label2(3).Caption = Format$(dVkPr, "##,##0.00") & " " & gcWaehrung
        Label2(5).Caption = Format$(dKVKPR, "##,##0.00") & " " & gcWaehrung
        If Trim$(Text1(4).Text) <> "" Then
            If Option1(1).Value Then
                Text1(4).Text = cLinr
                Label2(4).Caption = cLiefBez
            End If
        Else
            Text1(4).Text = cLinr
            Label2(4).Caption = cLiefBez
        End If

        Text1(0).Text = cEAN
        Text1(1).SetFocus
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheArtikelWKL15"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
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
    Command8.Visible = False
    
    
    bTextSuche = False
    
    If Check5.Value = vbChecked Then
        gsARTNR = ermartnrausLIBESNR(Trim$(Text1(0).Text), Val(Text1(4).Text))
        If gsARTNR <> "" Then
            Text1(0).Text = gsARTNR
            gsARTNR = ""
            bTextSuche = False
        Else
            bTextSuche = True
            gbLibesnrSeek = True
        End If
    Else
        gbLibesnrSeek = False
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
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
'    For lcount = 1 To Len(cFeld)
'        cZeichen = Mid(cFeld, lcount, 1)
'        If InStr(cValid, cZeichen) = 0 Then
'            bTextSuche = True
'            Exit For
'        End If
'    Next lcount
    
    If bTextSuche Then
        gcSuch = Text1(0).Text
        gsARTNR = ""
        frmWKL70.Show 1
        Me.Refresh
        If gsARTNR <> "" Then
            Text1(0).Text = gsARTNR
            gsARTNR = ""
            Command1_Click
            gbLibesnrSeek = False
        End If

    Else
        SucheArtikelWKL15
        gbLibesnrSeek = False
    End If
    
    
    If CInt(gcFilNr) > 0 And Label2(2).Caption <> "" Then
        Command8.Visible = True
    End If
    
    If Label2(2).Caption <> "" Then
        Command6(0).Visible = True
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub SchreibeListe()
    On Error GoTo LOKAL_ERROR
    
    Dim sTempstring As String

    sTempstring = new19Artikel.artnr & Space(8 - Len(CStr(new19Artikel.artnr)))
    
    sTempstring = sTempstring & new19Artikel.BEZEICH & Space(37 - Len(new19Artikel.BEZEICH))
    
    
    
    
    sTempstring = sTempstring & new19Artikel.ZubuchMe & Space(10 - Len(Trim(CStr(new19Artikel.ZubuchMe))))
    
    sTempstring = sTempstring & new19Artikel.BESTAND & Space(15 - Len(Trim(CStr(new19Artikel.BESTAND))))
    sTempstring = sTempstring & new19Artikel.OLDBESTAND

    List5.AddItem sTempstring, 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeListe"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
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
            
        Case Is = 15        '** Leeren/Speichern **
            If Trim$(Text1(1).Text) = "" Or Command2(15).Caption = "Leeren" Then
                LeereDialogWKL15
                Text1(0).SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            If Trim$(Text1(0).Text) = "" Then
                If Label2(2).Caption = "0" Then
                    MsgBox "Bitte einen Artikel festlegen!", vbCritical, "STOP!"
                    Text1(0).SetFocus
                    Exit Sub
                Else
                    Text1(0).Text = Label2(2).Caption
                End If
            End If
            glBestandNeu = Val(Text1(1).Text)
            If glBestandNeu < 0 Then
                If glLevel < 7 Then
                    MsgBox "Mengen-Reduzierung nicht möglich!" & vbCrLf & vbCrLf & "Bestandsminderungen sind nur mit Zugriffs-Level 7 oder höher erlaubt!", vbInformation, "INFO"
                    Exit Sub
                End If
            End If
            SchreibeDatenWKL15
            SchreibeListe
            
            If NewTableSuchenDBKombi("beTemp", gdApp) Then
                If Datendrin("beTemp", gdApp) Then
                    Command5.BackColor = vbRed
                End If
            End If
            
        
        Case Is = 16        'Vorheriges Feld
            If lcount > 0 Then
                Text1(0).SetFocus
            Else
                Text1(lcount).SetFocus
            End If
            
        Case Is = 17        'Nächstes Feld
            If lcount < 1 Then
                Text1(1).SetFocus
            Else
                Text1(lcount).SetFocus
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
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click()
    On Error GoTo LOKAL_ERROR
    
    
    UpdateBenex
    
    Unload frmWKL19
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
    Fehlermeldung1

End Sub
Private Sub Command8_Click()
    On Error GoTo LOKAL_ERROR
    
    gcArtNrFiliale = Trim(Label2(2).Caption)
    frmWKLae.Show 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    PositionierenWKL15
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    LeereDialogWKL15
    gF2Prompt.lLastPos = -1
    List6.AddItem " Artnr  Artikelbezeichnung                   Menge neuer Bestand alter Bestand"
    Screen.MousePointer = 0
    
    Frame1.Visible = False
    
    gbLibesnrSeek = False
    

    List5.Visible = False
    List6.Visible = False
    Label11(3).Visible = False
    Command2(21).Visible = False
    
    If NewTableSuchenDBKombi("beTemp", gdApp) Then
        If Datendrin("beTemp", gdApp) Then
            Command5.BackColor = vbRed
        End If
    End If
    
    Option2(Leselast19Einstellung).Value = True
    Option2(2).Caption = Option2(2).Caption & " (" & gsMDEGERAET & ")"
    
    If Check3.Value = vbChecked Then
        Frame4.Visible = False
    
        'Manuell
        Frame3.Visible = True
'        Text1(0).SetFocus
        If gbBILDTAST = False Then
            Frame1.Visible = False
        Else
            Frame1.Visible = True
        End If

        List5.Visible = True
        List6.Visible = True
        Label11(3).Visible = True
        Command2(21).Visible = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub anzeigeMDE()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim sSQL        As String
    Dim cLBSatz     As String
    Dim cArtNr      As String
    Dim cBez        As String
    Dim ckPr        As String
    Dim cMenge      As String
    Dim cLinr       As String
    Dim iZaehler    As Integer
    
    List12.Clear
    List11.Clear
    List12.AddItem "Artnr  Bezeichnung      VK-Preis Menge  Lieferant"
    
    sSQL = "Select * from KORREKB where Status = 'vorhanden' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    iZaehler = 0
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iZaehler = iZaehler + 1
            cArtNr = IIf(IsNull(rsrs!artnr), "", rsrs!artnr)
            cBez = IIf(IsNull(rsrs!BEZEICH), "", rsrs!BEZEICH)
            ckPr = IIf(IsNull(rsrs!KVKPR1), "0,00", Format$(rsrs!KVKPR1, "#####0.00"))
            cMenge = IIf(IsNull(rsrs!BESTVOR), "0", rsrs!BESTVOR)
            cLinr = IIf(IsNull(rsrs!linr), "0", rsrs!linr)
            
            cLBSatz = cArtNr & Space$(7 - Len(cArtNr))
            If Len(cBez) > 15 Then
                cBez = Left(cBez, 15) & "..."
            End If
            cLBSatz = cLBSatz & cBez & Space$(19 - Len(cBez))
            
            cLBSatz = cLBSatz & ckPr & Space$(7 - Len(ckPr))
            cLBSatz = cLBSatz & cMenge & Space$(7 - Len(cMenge)) & cLinr
            List11.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
        Label7(5).Caption = "insgesamt: " & iZaehler & " verschiedene Artikel"
        Label7(5).Refresh
    End If
    rsrs.Close: Set rsrs = Nothing
    
    List4.Clear
    List3.Clear
    List4.AddItem "EANCODE       Menge Scanreihenfolge"
    
    sSQL = "Select * from KORREKB where Status = 'nicht vorhanden' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    iZaehler = 0
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iZaehler = iZaehler + 1
            
            cBez = IIf(IsNull(rsrs!BEZEICH), "", rsrs!BEZEICH)
            cMenge = IIf(IsNull(rsrs!BESTVOR), "0", rsrs!BESTVOR)
            cLinr = IIf(IsNull(rsrs!lfnr), "0", rsrs!lfnr)
        
            cLBSatz = cBez & Space$(14 - Len(cBez))
            cLBSatz = cLBSatz & cMenge & Space$(6 - Len(cMenge)) & cLinr
            List3.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
        Command6(8).Visible = True
        Label7(4).Caption = "insgesamt: " & iZaehler & " verschiedene Artikel"
        Label7(4).Refresh
    Else
        Command6(8).Visible = False
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Frame8.Visible = True
    Frame5.Visible = False
    
    
    Exit Sub
LOKAL_ERROR:
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "anzeigeMDE"
        Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten."
        
        Fehlermeldung1
End Sub
Private Sub MdeVerarbeitung1()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsMDE       As Recordset
    Dim rsFilB      As Recordset
    Dim rsArt       As Recordset
    Dim seekEAN     As String
    
    Screen.MousePointer = 11
    
    Check2.Visible = False
    Command6(12).Visible = False
    
    loeschNEW "KORREKB", gdBase
    CreateTable "KORREKB", gdBase
    
    Set rsFilB = gdBase.OpenRecordset("KORREKB")
    
    Set rsMDE = gdBase.OpenRecordset("mdeinh")
    If Not rsMDE.EOF Then
        rsMDE.MoveFirst
        Do While Not rsMDE.EOF
            If Not IsNull(rsMDE!eancode) Then
            
                seekEAN = Trim(rsMDE!eancode)
                seekEAN = checkean(seekEAN)
                
                If Ist_in_ARTEAN_K(seekEAN) Then
                
                End If
                
                If Len(seekEAN) = 11 Then
                    seekEAN = "0" & seekEAN
            
                    sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                ElseIf Len(seekEAN) = 8 Then
                
                
                
                    If Left(seekEAN, 1) = "2" Then
                    
                        If Check7.Value = vbChecked Then
                            sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                            sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                            sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                        Else
                            seekEAN = Mid$(seekEAN, 2, 6)
                            sSQL = "select * from artikel where artnr = " & seekEAN
                        End If
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
                
                If Not rsArt.EOF Then 'hier die bekannten
                    rsFilB.AddNew
                    
                    rsFilB!artnr = rsArt!artnr
                    rsFilB!BEZEICH = rsArt!BEZEICH
                    rsFilB!linr = rsArt!linr
                    rsFilB!LIBESNR = rsArt!LIBESNR
                    rsFilB!LPZ = rsArt!LPZ
                    rsFilB!KVKPR1 = rsArt!KVKPR1
                    rsFilB!BESTVOR = rsMDE!Menge
                    
                    If IsNull(rsArt!BESTAND) Then
                        rsArt.Edit
                        rsArt!BESTAND = 0
                        rsArt.Update
                    End If
                        
                    rsFilB!BESTAND = rsArt!BESTAND
                    rsFilB!BESTANDN = CLng(rsArt!BESTAND) + CLng(rsMDE!Menge)
                    rsFilB!FILIALE = CByte(gcFilNr)
                    rsFilB!Status = "vorhanden"

                    rsFilB.Update
                Else 'hier die unbekannten
                
                    rsFilB.AddNew
                    rsFilB!BEZEICH = seekEAN
                    rsFilB!BESTVOR = rsMDE!Menge
                    rsFilB!Status = "nicht vorhanden"
                    rsFilB!FILIALE = CByte(gcFilNr)
                    rsFilB.Update
                    
                End If
                rsArt.Close: Set rsArt = Nothing
            End If
            rsMDE.MoveNext
        Loop
    
    End If
    
    rsMDE.Close: Set rsMDE = Nothing
    rsFilB.Close: Set rsFilB = Nothing
    
    anzeigeMDE
    
    anzeigeNew "normal", "Wollen Sie die Bestandskorrektur jetzt einlesen?", Label5

    Command6(12).Visible = True 'Einlese Button aktiv

    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3010 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "MdeVerarbeitung1"
        Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub

Private Sub Option2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    speicherlast19Einstellung Index
     
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option2_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherlast19Einstellung(i As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
'    loeschapp "BEKENX"
'    CreateTable "BEKENX", gdApp

    sSQL = "Delete from BEKENX "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into BEKENX (Ind) values (" & i & ")"
    gdApp.Execute sSQL, dbFailOnError
    
     
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherlast19Einstellung"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Leselast19Einstellung() As Byte
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    
    Leselast19Einstellung = 0
    
    If Not NewTableSuchenDBKombi("BEKENX", gdApp) Then
        CreateTable "BEKENX", gdApp
        
        sSQL = "Insert into BEKENX (Ind) values (0)"
        gdApp.Execute sSQL, dbFailOnError
    Else
        If Not SpalteInTabellegefundenNEW("BEKENX", "ind2", gdApp) Then
            SpalteAnfuegenNEW "BEKENX", "ind2", "byte", gdApp
        End If
    End If
    
    Set rsrs = gdApp.OpenRecordset("BEKENX")
    If Not rsrs.EOF Then
        Leselast19Einstellung = rsrs!ind
        
        If Not IsNull(rsrs!ind2) Then
            If rsrs!ind2 = 0 Then
                Check3.Value = vbChecked
            Else
                Check3.Value = vbUnchecked
            End If
        Else
            Check3.Value = vbUnchecked
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Leselast19Einstellung"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Text1_Change(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If Text1(1).Text <> "" And Label2(0).Caption <> "unbekannt" Then
        Command2(15).Caption = "Speichern"
    Else
        Command2(15).Caption = "Leeren"
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
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
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
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
            cValid = "1234567890" & Chr$(8)
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
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
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
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim ctmp As String
    
    If Index = 4 Then
        ctmp = Text1(4).Text
        ctmp = Trim$(Str$(Val(ctmp)))
        
        cSQL = "Select * from LISRT where LINR = " & ctmp & " "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            If Not IsNull(rsrs!LIEFBEZ) Then
                Label2(4).Caption = rsrs!LIEFBEZ
            Else
                Label2(4).Caption = ""
            End If
        Else
            Label2(4).Caption = ""
        End If
        rsrs.Close: Set rsrs = Nothing
        
    End If
    
    Text1(Index).BackColor = vbWhite

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Bestandskorrektur ist ein Fehler aufgetreten. "
    
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
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub
