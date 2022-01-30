VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL40 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Artikelliste nach Lieferanten"
   ClientHeight    =   8910
   ClientLeft      =   645
   ClientTop       =   2610
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWKL40.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8910
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame0 
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
      Height          =   855
      Left            =   120
      TabIndex        =   36
      Top             =   7800
      Width           =   12015
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   15
         Left            =   7320
         TabIndex        =   30
         Top             =   0
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
         Caption         =   "."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   14
         Left            =   10920
         TabIndex        =   35
         Top             =   0
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
         Caption         =   ">>"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   13
         Left            =   10200
         TabIndex        =   34
         Top             =   0
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
         Caption         =   "<<"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   12
         Left            =   9480
         TabIndex        =   33
         Top             =   0
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
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   11
         Left            =   8760
         TabIndex        =   32
         Top             =   0
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   10
         Left            =   8040
         TabIndex        =   31
         Top             =   0
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   9
         Left            =   6600
         TabIndex        =   29
         Top             =   0
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   8
         Left            =   5880
         TabIndex        =   28
         Top             =   0
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   7
         Left            =   5160
         TabIndex        =   27
         Top             =   0
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   6
         Left            =   4440
         TabIndex        =   26
         Top             =   0
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   5
         Left            =   3720
         TabIndex        =   25
         Top             =   0
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   4
         Left            =   3000
         TabIndex        =   24
         Top             =   0
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   3
         Left            =   2280
         TabIndex        =   23
         Top             =   0
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   2
         Left            =   1560
         TabIndex        =   22
         Top             =   0
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   1
         Left            =   840
         TabIndex        =   21
         Top             =   0
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
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   0
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
         Index           =   1
         Left            =   2040
         TabIndex        =   38
         Top             =   1800
         Visible         =   0   'False
         Width           =   3735
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
         Index           =   0
         Left            =   2040
         TabIndex        =   37
         Top             =   1440
         Visible         =   0   'False
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00000000&
      Height          =   6615
      Left            =   360
      TabIndex        =   40
      Top             =   840
      Width           =   11775
      Begin VB.CheckBox Check6 
         BackColor       =   &H00C0C000&
         Caption         =   "nur mit Verkäufen"
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
         TabIndex        =   59
         Top             =   4200
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikelnummer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   6600
         TabIndex        =   58
         Top             =   3960
         Width           =   4335
      End
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
         Height          =   375
         Index           =   5
         Left            =   8160
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   840
         Width           =   1095
      End
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
         Height          =   375
         Index           =   4
         Left            =   8160
         MaxLength       =   6
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0C000&
         Caption         =   "nur Bestand > 0"
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
         TabIndex        =   54
         Top             =   3480
         Width           =   3255
      End
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
         Height          =   375
         Index           =   3
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0C000&
         Caption         =   "EK - Preise drucken"
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
         TabIndex        =   10
         Top             =   3120
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferantennummer, Artikelnummer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   6600
         TabIndex        =   16
         Top             =   3600
         Width           =   4335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferantennummer, Linie, Lieferantenbestellnummer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   6600
         TabIndex        =   15
         Top             =   3240
         Width           =   4935
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferantennummer, Linie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   6600
         TabIndex        =   14
         Top             =   2880
         Width           =   3135
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferantennummer, Lieferantenbestellnummer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   13
         Top             =   2520
         Width           =   4815
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikelgruppennummer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   12
         Top             =   2160
         Width           =   4335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Bezeichnung "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   6600
         TabIndex        =   11
         Top             =   1800
         Width           =   4335
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   1
         Left            =   9480
         TabIndex        =   19
         Top             =   720
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   0
         Left            =   9480
         TabIndex        =   17
         Top             =   120
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
         Caption         =   "Suche Daten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
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
         Height          =   375
         Index           =   0
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
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
         Height          =   375
         Index           =   1
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   840
         Width           =   1095
      End
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
         Height          =   375
         Index           =   2
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "nur geräumte Artikel (RKZ = 'J')"
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
         TabIndex        =   8
         Top             =   2400
         Width           =   3135
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C000&
         Caption         =   "nur geführte Artikel (Geführt = 'J')"
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
         TabIndex        =   9
         Top             =   2760
         Width           =   3255
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   2
         Left            =   4080
         TabIndex        =   18
         Top             =   3240
         Width           =   2415
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
         Caption         =   "Leeren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   4080
         Sorted          =   -1  'True
         TabIndex        =   44
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text2 
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
         Left            =   5040
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text2 
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
         Left            =   5040
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   360
         Width           =   855
      End
      Begin sevCommand3.Command Command0 
         Height          =   420
         Index           =   16
         Left            =   2880
         TabIndex        =   43
         Top             =   360
         Width           =   480
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
      Begin sevCommand3.Command Command0 
         Height          =   420
         Index           =   17
         Left            =   2880
         TabIndex        =   42
         Top             =   840
         Width           =   480
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
      Begin sevCommand3.Command Command0 
         Height          =   420
         Index           =   18
         Left            =   6000
         TabIndex        =   41
         Top             =   1440
         Width           =   480
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
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0C000&
         Caption         =   "Ex Artikel ausblenden"
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
         TabIndex        =   55
         Top             =   3840
         Width           =   3255
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   20
         Left            =   2880
         TabIndex        =   61
         ToolTipText     =   "Kalender"
         Top             =   1440
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
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   21
         Left            =   2880
         TabIndex        =   62
         ToolTipText     =   "Kalender"
         Top             =   1920
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
         Caption         =   "verschiedene Linien :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   4080
         TabIndex        =   60
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "bis Artnr:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   6720
         TabIndex        =   57
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "von Artnr:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   6720
         TabIndex        =   56
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "bis VK-Datum :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   53
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblUeberschrift 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   6600
         TabIndex        =   51
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "von Lieferantennr.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   -120
         TabIndex        =   50
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "bis Lieferantennr.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   49
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "ab VK-Datum :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   48
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblUeberschrift 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   47
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "bis Linie :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   46
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "von Linie :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   45
         Top             =   480
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   3960
         X2              =   3960
         Y1              =   3840
         Y2              =   240
      End
   End
   Begin VB.Label lblAnzeige 
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
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   52
      Top             =   7320
      Width           =   10815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Artikelliste nach Lieferanten"
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
      Left            =   120
      TabIndex        =   39
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmWKL40"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub PositionierenWKL40()
    On Error GoTo LOKAL_ERROR
    
    With Frame2
        .Top = 720
        .Left = 120
        .Height = 6615
        .Width = 11775
    End With
    
    With Frame0
        .Top = 7800
        .Left = 120
        .Height = 855
        .Width = 12015
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL40"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "adrute", gdBase
    loeschNEW "TELINR", gdBase
    loeschNEW "aLite", gdBase
    loeschNEW "temp", gdBase
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
Private Function bInhaltfound() As Boolean
    On Error GoTo LOKAL_ERROR
    
    bInhaltfound = False
    
    If Text1(0).Text <> "" Then
        bInhaltfound = True
        Exit Function
    ElseIf Text1(1).Text <> "" Then
        bInhaltfound = True
        Exit Function
    ElseIf Text1(4).Text <> "" Then
        bInhaltfound = True
        Exit Function
    ElseIf Text1(5).Text <> "" Then
        bInhaltfound = True
        Exit Function
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "bInhaltfound"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub LeereDialogWKL40()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    Text1(5).Text = ""
    Text2(0).Text = ""
    Text2(1).Text = ""
    Option2(0).Value = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKL40"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseDatenWKL40()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim cSQLMulti       As String
    Dim cSQLMulti2      As String
    Dim cSQLMulti1      As String
    Dim sSQLArtnr       As String
    Dim cArtNr          As String
    Dim cArtNrb         As String
    Dim rsrs            As Recordset
    Dim cLiNr1          As String
    Dim cLiNr2          As String
    Dim cLinie1         As String
    Dim cLinie2         As String
    ReDim cOrderBy(0 To 6) As String
    Dim lcount          As Long
    Dim lOrderBy        As Long
    Dim cLBSatz         As String
    Dim lJahr1          As Long
    Dim lMonat1         As Long
    Dim lJahr2          As Long
    Dim lMonat2         As Long
    Dim lDatum1         As Long
    Dim cDatum1         As String
    Dim lDatum2         As Long
    Dim cDatum2         As String
    Dim lVKZeitraum     As Long
    Dim dUmsZeitraum    As Double
    Dim lListSatz       As Long
    Dim cSelect         As String
    Dim iStufe          As Integer
    iStufe = 1
    
    anzeige "normal", "Artikeldaten werden ermittelt, bitte warten...", lblAnzeige
    
    cSelect = "Select A.ARTNR  "
    cSelect = cSelect & ", A.BEZEICH  "
    cSelect = cSelect & ", A.AGN  "
    cSelect = cSelect & ", B.LEKPR  "
    cSelect = cSelect & ", A.VKPR  "
    cSelect = cSelect & ", A.MWST  "
    cSelect = cSelect & ", B.LINR  "
    cSelect = cSelect & ", B.LIBESNR  "
    cSelect = cSelect & ", A.EAN  "
    cSelect = cSelect & ", A.EAN2  "
    
    cSelect = cSelect & ", A.EAN3  "
    cSelect = cSelect & ", A.ETIMERK  "
    cSelect = cSelect & ", A.MOPREIS  "
    cSelect = cSelect & ", B.RKZ  "
    cSelect = cSelect & ", A.LPZ  "
    cSelect = cSelect & ", A.NOTIZEN  "
    cSelect = cSelect & ", A.BESTAND  "
    
    cSelect = cSelect & ", B.MINMEN  "
    cSelect = cSelect & ", A.INHALT  "
    cSelect = cSelect & ", A.INHALTBEZ  "
    cSelect = cSelect & ", A.GRUNDPREIS  "
    cSelect = cSelect & ", A.MINBEST  "
    cSelect = cSelect & ", A.RABATT_OK  "
    cSelect = cSelect & ", A.GEFUEHRT  "
    cSelect = cSelect & ", A.KVKPR1  "
    cSelect = cSelect & ", A.EKPR  "

    cSelect = cSelect & ", A.PREISSCHU  "
    cSelect = cSelect & ", A.BONUS_OK  "
    cSelect = cSelect & ", A.UMS_OK  "
    cSelect = cSelect & ", A.AWM  "
    cSelect = cSelect & ", A.LASTDATE  "
    cSelect = cSelect & ", A.LASTTIME  "
    cSelect = cSelect & ", A.AUFDAT  "
    cSelect = cSelect & ", B.EXDAT   "

    cSelect = cSelect & ", A.GROESSE  "
    cSelect = cSelect & ", A.SPANNE  "
    cSelect = cSelect & ", A.AUFSCHLAG  "
    cSelect = cSelect & ", A.SYNSTATUS  "
    
    cSelect = cSelect & ", A.VKPR as VKMLF  "
    cSelect = cSelect & ", A.VKPR as VKMVM  "
    cSelect = cSelect & ", A.VKPR as VKWLF  "
    cSelect = cSelect & ", A.VKPR as VKWVM  "
    cSelect = cSelect & ", A.VKPR as VKMZR  "
    cSelect = cSelect & ", A.VKPR as VKWZR  "
    
    cSelect = cSelect & ", A.PGN "
    
    
    'artnr von
    If Trim$(Text1(4).Text) <> "" Then
        cArtNr = Val(Trim$(Text1(4).Text))
    End If
    
    'artnr bis
    If Trim$(Text1(5).Text) <> "" Then
        cArtNrb = Val(Trim$(Text1(5).Text))
    End If
    
    sSQLArtnr = ""
    
    If cArtNrb <> "" Then
        If cArtNr <> "" Then
            sSQLArtnr = sSQLArtnr & " and a.ARTNR between " & cArtNr & " and " & cArtNrb & " "
        Else
            sSQLArtnr = sSQLArtnr & " and a.ARTNR = " & cArtNrb & " "
        End If
    Else
        If cArtNr <> "" Then
            sSQLArtnr = sSQLArtnr & " and a.ARTNR = " & cArtNr & " "
        End If
    End If
   

    If List1.ListCount = 0 Then
        iStufe = 2
        cSQLMulti1 = ""
    Else
        iStufe = 3
        For lListSatz = 0 To List1.ListCount - 1
            cLBSatz = List1.list(lListSatz)
            cLiNr1 = Left(cLBSatz, 6)
            cLiNr1 = Trim$(cLiNr1)
            cLinie1 = Right(cLBSatz, 3)
            cLinie1 = Trim$(cLinie1)
            
            iStufe = 4
            cSQLMulti = "(B.LINR = " & cLiNr1 & " "
            cSQLMulti = cSQLMulti & "and A.LPZ = " & cLinie1 & " )"
            
            iStufe = 5
            If lListSatz = 0 Then
                iStufe = 6
                cSQLMulti2 = cSQLMulti
            Else
                iStufe = 7
                cSQLMulti = " or " & cSQLMulti
                cSQLMulti2 = cSQLMulti2 & cSQLMulti
            End If
            
            iStufe = 8
        Next lListSatz
        cSQLMulti1 = " Where " & cSQLMulti2
    End If
    loeschNEW "adrute", gdBase
    
    iStufe = 9
    
    cDatum1 = Text1(2).Text
    If cDatum1 <> "" Then
    
        iStufe = 10
        If IsDate(cDatum1) Then
            lDatum1 = DateValue(cDatum1)
        Else
            Text1(2).SetFocus
            
            anzeige "rot", "Bitte richtiges Datum eingeben!", lblAnzeige
            Exit Sub
        End If
       
    End If
    iStufe = 11
    
    cDatum2 = Text1(3).Text
    
    If cDatum2 <> "" Then
        iStufe = 12
        If IsDate(cDatum2) Then
            lDatum2 = DateValue(cDatum2)
        Else
            Text1(3).SetFocus
            
            anzeige "rot", "Bitte richtiges Datum eingeben!", lblAnzeige
            Exit Sub
        End If
    End If
    iStufe = 13
    If cDatum1 <> "" And cDatum2 = "" Then
        cDatum2 = DateValue(Now)
        lDatum2 = DateValue(Now)
    End If
    iStufe = 14
    lJahr1 = Year(Now)
    lMonat1 = Month(Now)
    
    lMonat2 = lMonat1 - 1
    lJahr2 = lJahr1
    If lMonat2 < 1 Then
        lMonat2 = 12
        lJahr2 = lJahr1 - 1
    End If
    iStufe = 15
    cOrderBy(0) = "order by A.BEZEICH "
    cOrderBy(1) = "order by A.AGN, A.BEZEICH "
    cOrderBy(2) = "order by A.LINR, A.LIBESNR "
    cOrderBy(3) = "order by A.LINR, A.LPZ, A.BEZEICH "
    cOrderBy(4) = "order by A.LINR, A.LPZ, A.LIBESNR "
    cOrderBy(5) = "order by A.LINR, A.ARTNR "
    cOrderBy(6) = "order by A.ARTNR "
    iStufe = 16
    cLiNr1 = Text1(0).Text
    cLiNr2 = Text1(1).Text
    cLiNr1 = Trim$(cLiNr1)
    cLiNr2 = Trim$(cLiNr2)
    iStufe = 17
    loeschNEW "TELINR", gdBase
    cSQL = "Create Table TELINR ( LINR1 Text(6), LINR2 Text(6)) "
    SQL_Befehl_ausführen cSQL
    
    iStufe = 18
    cSQL = "Insert Into TELINR ( LINR1 , LINR2) Values ('" & cLiNr1 & "','" & cLiNr2 & "')"
    SQL_Befehl_ausführen cSQL
    
    iStufe = 19
    
    For lcount = 0 To 6
        If Option2(lcount).Value = True Then
            lOrderBy = lcount
            Exit For
        End If
    Next lcount
    
    iStufe = 20
    
    If cSQLMulti1 <> "" Then
        iStufe = 21
        
        cSQL = cSelect & " , B.LEKPR as LEKPR2 into adrute "
        cSQL = cSQL & " from ARTIKEL A inner join ARTLIEF B on A.ARTNR = B.ARTNR "
        cSQL = cSQL & cSQLMulti1
        cSQL = cSQL & " and (A.SYNSTATUS = 'E' or A.SYNSTATUS = 'A' or A.SYNSTATUS is null) "
        
        If cLiNr1 <> "" And cLiNr2 <> "" Then
            cSQL = cSQL & " and B.LINR >= " & cLiNr1 & " and B.LINR <= " & cLiNr2 & " "
        End If
        
    Else
        iStufe = 22
        cSQL = cSelect & " , B.LEKPR as LEKPR2 into adrute "
        cSQL = cSQL & " from ARTIKEL A inner join ARTLIEF B on A.ARTNR = B.ARTNR "
        cSQL = cSQL & " where (A.SYNSTATUS = 'E' or A.SYNSTATUS = 'A' or A.SYNSTATUS is null) "
        
        If cLiNr1 <> "" And cLiNr2 <> "" Then
            cSQL = cSQL & " and B.LINR >= " & cLiNr1 & " and B.LINR <= " & cLiNr2 & " "
        End If
    End If
    
    If List1.ListCount = 0 Then
        iStufe = 23
        cLinie1 = Text2(0).Text
        cLinie1 = Trim$(cLinie1)
        
        cLinie2 = Text2(1).Text
        cLinie2 = Trim$(cLinie2)
        
        If cLinie1 = "" Then
            cLinie1 = "0"
        End If
        
        If cLinie2 = "" Then
            cLinie2 = "999"
        End If
        
        cSQL = cSQL & "and A.LPZ >= " & cLinie1 & " and A.LPZ <= " & cLinie2 & " "
    End If
    
    'EX
    
    If Check1.Value = vbChecked Then
        iStufe = 24
        cSQL = cSQL & " and B.RKZ = 'J' "
    End If
    
    If Check5.Value = vbChecked Then
        iStufe = 24
        cSQL = cSQL & " and not B.RKZ = 'J' "
    End If
    
    'EX Ende
    
    If Check2.Value = vbChecked Then
        iStufe = 25
        cSQL = cSQL & " and A.GEFUEHRT = 'J' "
    End If
    
    cSQL = cSQL & " " & sSQLArtnr & " "
    cSQL = cSQL & cOrderBy(lOrderBy)
'    MsgBox cSQL
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Update adrute set bestand = 0 where Bestand is null "
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Update adrute set VKMLF = 0 ,VKMVM =0 , VKWLF = 0 "
    cSQL = cSQL & " ,VKWVM = 0 "
    cSQL = cSQL & " ,VKMZR = 0 "
    cSQL = cSQL & " , VKWZR = 0 "
    SQL_Befehl_ausführen cSQL
    
    
    
    
    
    
    If Check4.Value = vbChecked Then
        cSQL = "Delete from adrute where Bestand <= 0 "
        SQL_Befehl_ausführen cSQL
    End If
    
    iStufe = 27
    If cDatum1 = "" Then
        'Vergleich
        cSQL = "Update ADRUTE inner join ums_art on ADRUTE.artnr = ums_art.artnr "
        cSQL = cSQL & " Set ADRUTE.VKMLF = ums_art.anzahl, ADRUTE.VKWLF = ums_art.umsatz "
        cSQL = cSQL & "where monat = " & lMonat1
        cSQL = cSQL & "and Jahr = " & lJahr1
        SQL_Befehl_ausführen cSQL

        cSQL = "Update ADRUTE inner join ums_art on ADRUTE.artnr = ums_art.artnr "
        cSQL = cSQL & " Set ADRUTE.VKMVM = ums_art.anzahl, ADRUTE.VKWVM = ums_art.umsatz "
        cSQL = cSQL & "where monat = " & lMonat2
        cSQL = cSQL & "and Jahr = " & lJahr2
        SQL_Befehl_ausführen cSQL
        
    Else
        'Zeitraum
        loeschNEW "temp", gdBase
        
        iStufe = 30
        
        cSQL = "Select sum(kassjour.menge) as VKmenge,sum(kassjour.preis) as VKPreis, kassjour.artnr into Temp from Kassjour inner join ADRUTE on ADRUTE.artnr = Kassjour.artnr "
        cSQL = cSQL & "  where Kassjour.ADATE BETWEEN " & lDatum1 & " "
        cSQL = cSQL & "  and " & lDatum2 & " "
        cSQL = cSQL & " and Filiale = " & gcFilNr
        cSQL = cSQL & " group by kassjour.artnr "
        SQL_Befehl_ausführen cSQL
        
        
        iStufe = 31

        cSQL = "Update ADRUTE inner join Temp on ADRUTE.artnr = Temp.artnr "
        cSQL = cSQL & " Set ADRUTE.VKWZR = Temp.VKPreis"
        SQL_Befehl_ausführen cSQL
        
        iStufe = 32
        cSQL = "Update ADRUTE inner join Temp on ADRUTE.artnr = Temp.artnr "
        cSQL = cSQL & " Set ADRUTE.VKMZR = Temp.VKmenge"
        SQL_Befehl_ausführen cSQL
        iStufe = 33
        cSQL = "Update ADRUTE "
        cSQL = cSQL & " Set ADRUTE.EAN2 = '" & cDatum1 & "' "
        SQL_Befehl_ausführen cSQL
        iStufe = 34
        cSQL = "Update ADRUTE "
        cSQL = cSQL & " Set ADRUTE.EAN3 = '" & cDatum2 & "' "
        SQL_Befehl_ausführen cSQL
    End If
    
    If Check6.Value = vbChecked Then
        cSQL = "Delete from adrute where VKMZR <= 0 "
        SQL_Befehl_ausführen cSQL
    End If
    iStufe = 35
    
    loeschNEW "aLite", gdBase
    CreateTable "ALITE", gdBase
    iStufe = 37
    
    cSQL = "Insert into alite Select * from ADRUTE A "
    cSQL = cSQL & cOrderBy(lOrderBy)
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Select count(*) from alite "
    Set rsrs = gdBase.OpenRecordset(cSQL, dbOpenDynaset)
    If Not rsrs.EOF Then
    
        rsrs.MoveLast
        If rsrs.RecordCount = 0 Then
            iStufe = 39
            lblAnzeige.ForeColor = vbRed
            lblAnzeige.Caption = "Es wurden keine Daten ermittelt."
            Exit Sub
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If gbSQLSERVER Then
    
        Hintergrundtabelle_kopieren "ALITE"
        Hintergrundtabelle_kopieren "FILA"
        dbPrintOpen

    End If
    
    anzeige "normal", "Druckvorschau wird erstellt, bitte warten...", lblAnzeige
    If cDatum1 = "" Then
        'Vergleich
        iStufe = 41
        If Check3.Value = vbChecked Then
            reportbildschirm "dWKL001", "aWKL40"
        Else
            reportbildschirm "dWKL001b", "aWKL40b"
        End If
         
    Else
        'Zeitraum
        iStufe = 42
        If Check3.Value = vbChecked Then
            reportbildschirm "dWKL001c", "aWKL40c"
        Else
            reportbildschirm "dWKL001d", "aWKL40d"
        End If
    End If
    
    anzeige "normal", "Fertig", lblAnzeige
    
    If gbSQLSERVER Then
    
        dbPrintClose

    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseDatenWKL40"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten. " & iStufe
    
    Fehlermeldung1
'    Resume Next
   
End Sub
Private Sub KopfDaten()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim cLiNr           As String
    Dim cLiefbez        As String
    
    loeschNEW "KOPFDATEN40", gdBase

    cSQL = "Create Table KOPFDATEN40 "
    cSQL = cSQL & "("
    cSQL = cSQL & " LINR TEXT(6)"
    cSQL = cSQL & ", LIEFBEZ TEXT(50)"
    cSQL = cSQL & ")"
    gdBase.Execute cSQL, dbFailOnError
    
    cLiNr = "0"
    cLiefbez = ""
    
    'linr von
    If Trim$(Text1(0).Text) = Trim$(Text1(1).Text) Then
        If Trim$(Text1(0).Text) <> "" Then
            cLiNr = Val(Trim$(Text1(0).Text))
            cLiefbez = ermLiefBez(CLng(cLiNr))
        End If
    End If
    
    
    cSQL = "Insert into KOPFDATEN40 "
    cSQL = cSQL & "("
    cSQL = cSQL & " LINR "
    cSQL = cSQL & ", LIEFBEZ "
    cSQL = cSQL & ") values "
    
    cSQL = cSQL & "("
    cSQL = cSQL & " " & cLiNr & " "
    cSQL = cSQL & ", '" & cLiefbez & "'"
    cSQL = cSQL & ")"
    gdBase.Execute cSQL, dbFailOnError

    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KopfDaten"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten. "
    
    Fehlermeldung1

   
End Sub
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Set WshShell = CreateObject("WScript.Shell")
    'WshShell.SendKeys "+{Tab}", True
    
    Select Case Index
        Case 0 To 9         'Ziffern
            If Label0(0).Caption <> "" Then
                If Label0(0).Caption = "Text1" Then
                    Text1(Val(Label0(1).Caption)).Text = Text1(Val(Label0(1).Caption)).Text & Command0(Index).Caption
                    Text1(Val(Label0(1).Caption)).SetFocus
                ElseIf Label0(0).Caption = "Text2" Then
                    Text2(Val(Label0(1).Caption)).Text = Text2(Val(Label0(1).Caption)).Text & Command0(Index).Caption
                    Text2(Val(Label0(1).Caption)).SetFocus
                End If
            End If
        Case Is = 10        'Backspace
            If Label0(0).Caption <> "" Then
                If Label0(0).Caption = "Text1" Then
                    If Len(Text1(Val(Label0(1).Caption)).Text) > 0 Then
                        Text1(Val(Label0(1).Caption)).Text = Left(Text1(Val(Label0(1).Caption)).Text, Len(Text1(Val(Label0(1).Caption)).Text) - 1)
                    End If
                    Text1(Val(Label0(1).Caption)).SetFocus
                ElseIf Label0(0).Caption = "Text2" Then
                    If Len(Text2(Val(Label0(1).Caption)).Text) > 0 Then
                        Text2(Val(Label0(1).Caption)).Text = Left(Text2(Val(Label0(1).Caption)).Text, Len(Text2(Val(Label0(1).Caption)).Text) - 1)
                    End If
                    Text2(Val(Label0(1).Caption)).SetFocus
                End If
            End If
        Case Is = 11        'Clear
            If Label0(0).Caption <> "" Then
                If Label0(0).Caption = "Text1" Then
                    Text1(Val(Label0(1).Caption)).Text = ""
                    Text1(Val(Label0(1).Caption)).SetFocus
                ElseIf Label0(0).Caption = "Text2" Then
                    Text2(Val(Label0(1).Caption)).Text = ""
                    Text2(Val(Label0(1).Caption)).SetFocus
                End If
            End If
        Case Is = 12, 16, 17   'F2
            Select Case Index
                Case Is = 16
                    Text1_GotFocus 0
                    Text1(0).SetFocus
                Case Is = 17
                    Text1_GotFocus 1
                    Text1(1).SetFocus
            End Select
            
            If Label0(0).Caption <> "" Then
                If Label0(0).Caption = "Text1" Then
                    Text1_KeyUp Val(Label0(1).Caption), vbKeyF2, 0
'                ElseIf Label0(0).Caption = "Text2" Then
'                    Text2_KeyUp Val(Label0(1).Caption), vbKeyF2, 0
                End If
            End If
        Case Is = 13        'vorheriges Element
            WshShell.SendKeys "+{Tab}", True
        Case Is = 14        'nachfolgendes Element
            WshShell.SendKeys "{Tab}", True
        Case Is = 15        'Punkt für Datum
            If Label0(0).Caption <> "" Then
                If Label0(0).Caption = "Text1" Then
                    If Val(Label0(1).Caption) = 2 Then
                        Text1(Val(Label0(1).Caption)).Text = Text1(Val(Label0(1).Caption)).Text & Command0(Index).Caption
                        Text1(Val(Label0(1).Caption)).SetFocus
                    Else
                        Text1(Val(Label0(1).Caption)).SetFocus
                    End If
                ElseIf Label0(0).Caption = "Text2" Then
                    Text2(Val(Label0(1).Caption)).SetFocus
                End If
            End If
        Case Is = 18 'Linienauswahl
            Linienauswahl
        Case Is = 20        ' Kalender
            Text1(2).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
        Case Is = 21        ' Kalender
            Text1(3).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
    End Select
    
    If Label0(0).Caption = "Text1" Then
        Text1(Val(Label0(1).Caption)).SetFocus
    ElseIf Label0(0).Caption = "Text2" Then
        Text2(Val(Label0(1).Caption)).SetFocus
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
        
    Screen.MousePointer = 11
        
    Select Case Index
        Case Is = 0
            If bInhaltfound Then
            
                KopfDaten
                LeseDatenWKL40
            Else
                anzeige "rot", "Bitte geben Sie einen Lieferant oder eine Artikelnummer ein!", lblAnzeige
            End If
            

        Case Is = 1
            Unload frmWKL40
        Case Is = 2
            List1.Clear
            
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnz As Long
    Dim lcount As Long
    
    PositionierenWKL40
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift(0)
    
    
    Screen.MousePointer = 11

    LeereDialogWKL40
    Label0(0).Caption = "Text1"
    Label0(1).Caption = "0"
    Check3.Value = vbChecked
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String

    Select Case Index
        Case 2 To 3  'dat
            cValid = "1234567890." & Chr$(8)
        
        Case 0 To 1  'linr
            cValid = "1234567890" & Chr$(8)
            
        Case 4 To 5  'artnr
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
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyF2 Or KeyCode = vbKeyF4 Then
        
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
   
        Select Case Index
            Case Is = 0
                gF2Prompt.cFeld = "LINR"
                
            Case Is = 1
                gF2Prompt.cFeld = "LINR"
                    
        End Select
        
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
        End If
        
        If gF2Prompt.cWahl <> "" Then
            Text1(Index).Text = gF2Prompt.cWahl
        End If
        Text1(Index).SetFocus
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Linienauswahl()
    On Error GoTo LOKAL_ERROR
    
    Dim cLiNr1 As String
    Dim cLiNr2 As String
    Dim lcount As Long
    Dim ctmp As String
    
    
    cLiNr1 = Text1(0).Text
    cLiNr2 = Text1(1).Text
    
    If cLiNr1 = "" And cLiNr2 = "" Then
        MsgBox "Eingabehilfe nur möglich, wenn Lieferantennummern vorliegen!", vbInformation, "Winkiss Hinweis:"
        Text1(0).SetFocus
        Exit Sub
    End If
    If cLiNr1 = "" And cLiNr2 <> "" Then
        Text1(0).Text = cLiNr2
        cLiNr1 = cLiNr2
    End If
    If cLiNr2 = "" And cLiNr1 <> "" Then
        Text1(1).Text = cLiNr1
        cLiNr2 = cLiNr1
    End If
    
    gF2Prompt.cFeld = ""
    gF2Prompt.cWert = ""
    gF2Prompt.cWert2 = ""
    gF2Prompt.cWahl = ""
    gF2Prompt.bMultiple = True
    
    gF2Prompt.cFeld = "LPZ_VB"
    gF2Prompt.cWert = cLiNr1
    gF2Prompt.cWert2 = cLiNr2
            
    frmWK00a.Show 1
    
    List1.Clear
    For lcount = 0 To 100
        If gF2Prompt.cArray(lcount) <> "" Then
            Text2(0).Text = ""
            Text2(1).Text = ""
            ctmp = gF2Prompt.cArray(lcount)
            ctmp = Space$(10 - Len(ctmp)) & ctmp
            ctmp = Right(ctmp, 6) & " " & Left(ctmp, 3)
            List1.AddItem ctmp
        End If
    Next lcount
    
'    Text2(Index).SetFocus
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Linienauswahl"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Option2_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Label0(0).Caption = ""
    Label0(1).Caption = ""
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    Label0(0).Caption = Text1(Index).name
    Label0(1).Caption = Trim$(Str$(Index))
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text2(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text2(Index).BackColor = glSelBack1
    Label0(0).Caption = Text2(Index).name
    Label0(1).Caption = Trim$(Str$(Index))
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikellisten nach Lieferanten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


