VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWKL102 
   Caption         =   "Kundenkonto Korrektur"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      Caption         =   "neue Kunden Nr"
      Height          =   1335
      Left            =   9600
      TabIndex        =   33
      Top             =   2760
      Width           =   2055
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   120
         MaxLength       =   7
         TabIndex        =   35
         Top             =   840
         Width           =   1815
      End
      Begin sevCommand3.Command Command1 
         Height          =   360
         Index           =   6
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1815
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
         Caption         =   "Kunde suchen..."
         PictureAlign    =   2
         Version3        =   -1  'True
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   31
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Frame Frame0 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   6960
      Width           =   9615
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   0
         Left            =   120
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
         Caption         =   "1"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   1
         Left            =   840
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
         Caption         =   "2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   2
         Left            =   1560
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
         Caption         =   "3"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   3
         Left            =   2280
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
         Caption         =   "4"
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
         Index           =   5
         Left            =   3720
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
         Caption         =   "6"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   6
         Left            =   4440
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
         Caption         =   "7"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   7
         Left            =   5160
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
         Caption         =   "8"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   8
         Left            =   5880
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
         Caption         =   "9"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   720
         Index           =   9
         Left            =   6600
         TabIndex        =   19
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
         Index           =   11
         Left            =   7320
         TabIndex        =   18
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
         Index           =   13
         Left            =   8040
         TabIndex        =   17
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
         Index           =   14
         Left            =   8760
         TabIndex        =   16
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
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   30
         Top             =   1440
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   3735
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   5
      Left            =   9600
      TabIndex        =   14
      Top             =   1560
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
      Caption         =   "Protokoll"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   4
      Left            =   9600
      TabIndex        =   13
      Top             =   4680
      Width           =   2055
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
      Caption         =   "Verkäufe"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   3
      Left            =   9600
      TabIndex        =   12
      Top             =   4200
      Width           =   2055
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
      Caption         =   "Kundendaten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   9375
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   9
      Top             =   960
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
      Caption         =   "Suchen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command3 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   8
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
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      MaxLength       =   6
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   2
      Left            =   9600
      TabIndex        =   2
      Top             =   6360
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
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   3
      Top             =   7920
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
   Begin MSComCtl2.DTPicker Text2 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Format          =   74448897
      UpDown          =   -1  'True
      CurrentDate     =   38453
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   9375
   End
   Begin sevCommand3.Command Command1 
      Height          =   360
      Index           =   21
      Left            =   3240
      TabIndex        =   36
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
      Image           =   20
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "Kassennummer"
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
      Index           =   0
      Left            =   3960
      TabIndex        =   32
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "Verkaufsdatum"
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
      Index           =   5
      Left            =   1920
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "Artikel Nr"
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
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kundenkonto Korrektur"
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
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   10575
   End
   Begin VB.Label lblAnzeige 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   7920
      Width           =   9255
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
End
Attribute VB_Name = "frmWKL102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0 To 9         'Ziffern
            Text1(Val(Label0(1).Caption)).Text = Text1(Val(Label0(1).Caption)).Text & Command0(Index).Caption
            Text1(Val(Label0(1).Caption)).SetFocus
        Case Is = 11        'Clear
            Text1(Val(Label0(1).Caption)).Text = ""
            Text1(Val(Label0(1).Caption)).SetFocus
        Case Is = 13        'vorheriges Element
            
            If Val(Label0(1).Caption) = 4 Then
                Text1(1).SetFocus
            Else
                Text1(1).SetFocus
            End If
            
        Case Is = 14        'nachfolgendes Element
            If Val(Label0(1).Caption) = 1 Then
                Text1(4).SetFocus
            Else
                Text1(4).SetFocus
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenkonto Korrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
        
    Select Case Index
        Case Is = 0
            SucheBon
        Case Is = 1
            Unload frmWKL102
        Case Is = 2
            gcKundenNr = Trim(Text1(4).Text)
            If gcKundenNr <> "" Then
                If DatendrinSQL("select * from Kunden where kundnr = " & gcKundenNr, gdBase) Then
                    
                Else
                    Screen.MousePointer = 0
                    anzeige "rot", "Dieser Kunde ist unbekannt", lblanzeige
                    Exit Sub
                End If
            End If
            gcKundenNr = ""
            
            If SpeicherBonNeu Then
                SucheBon
            End If
        Case 3
            gcKundenNr = Trim(Text1(4).Text)
            If gcKundenNr <> "" Then
                If gcKundenNr <> "0" Then
                    If DatendrinSQL("select * from Kunden where kundnr = " & gcKundenNr, gdBase) Then
                        anzeige "normal", "Die Kundendaten(" & gcKundenNr & ") werden angezeigt...", lblanzeige
                        iKasse = 2
                        frmWKL13.Show 1
                        setzedrucker gcBonDrucker
                        anzeige "", "", lblanzeige
                    Else
                        anzeige "rot", "Dieser Kunde ist unbekannt", lblanzeige
                    End If
                End If
            End If
            gcKundenNr = ""
        Case 4
            gckundnr = Trim$(Text1(4).Text)
            gsARTNR = ""
            If gckundnr <> "" Then
                If gckundnr <> "0" Then
                    If DatendrinSQL("select * from Kunden where kundnr = " & gckundnr, gdBase) Then
                        frmWKL74.Show 1
                    Else
                        anzeige "rot", "Dieser Kunde ist unbekannt", lblanzeige
                    End If
                End If
            End If
            gckundnr = ""
        Case 5
            Screen.MousePointer = 11
            zeigeHilfeDabapfad "LPROTOK", "Kundenkontokorrektur.txt"
            Screen.MousePointer = 0
        Case 6
            Screen.MousePointer = 0
            frmWKL134.Show 1
        
            If gckundnr <> "" Then
                If IsNumeric(gckundnr) Then
                    Text1(4).Text = gckundnr
                End If
            End If
            gckundnr = ""
        Case 21
            Screen.MousePointer = 0
            Text2(0).Value = Format(Datumschreiben11a(3100, 4000), "DD.MM.YYYY")
            'fertig
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenkonto Korrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function SpeicherBonNeu() As Boolean
    On Error GoTo LOKAL_ERROR

    Dim bFound      As Boolean
    Dim lFound      As Long
    Dim lcount      As Long
    Dim lBelegnr    As Long
    Dim sSQL        As String
    Dim iRet        As Integer
    Dim lKund       As Long
    Dim lkundvorher As Long
    Dim lDatum      As Long
    Dim cVon        As String
    Dim dBonusf     As Double
    Dim dUmsatzf    As Double
    Dim cKasnum     As String
    Dim cUhrzeit    As String
    
    lcount = 0
    
    SpeicherBonNeu = False
    
    ctmp = Text1(0).Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        anzeige "rot", "Bitte eine Kassennummer angeben!", lblanzeige
        Text1(0).SetFocus
        Exit Function
    Else
        cKasnum = Val(ctmp)
    End If
    
    ctmp = Text1(4).Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        anzeige "rot", "Bitte eine Kundennummer angeben!", lblanzeige
        Text1(4).SetFocus
        Exit Function
    Else
        lKund = Val(ctmp)
    End If
    
    If Text2(0).Value <> "" Then
        cVon = Text2(0).Value
    Else
        cVon = DateValue(Now)
        Text2(0).Value = DateValue(Now) - 1
    End If
    lDatum = DateValue(cVon)
    
    bFound = False
    lFound = 0
    If List3.ListCount = 0 Then
        anzeige "rot", "Bitte einen Satz markieren!", lblanzeige
        Exit Function
    End If
    
    For lcount = 0 To List3.ListCount - 1
        If List3.Selected(lcount) = True Then
            bFound = True
            lFound = lFound + 1
        End If
    Next lcount
    
    If bFound Then
    
        If lFound > 1 Then
            anzeige "rot", "Bitte nur einen Satz markieren!", lblanzeige
            Exit Function
        Else
            For lcount = 0 To List3.ListCount - 1
                If List3.Selected(lcount) Then
                    
                    
                    cUhrzeit = Trim(Mid(List3.list(lcount), 9, 12))
                    lkundvorher = Mid(List3.list(lcount), 66, 10)
                    lBelegnr = Left(List3.list(lcount), 8)
                    cKasnum = Trim(Right(List3.list(lcount), 6))
                    
                    
                    iRet = MsgBox("Möchten Sie wirklich die Verkäufe des Bons: " & lBelegnr & " vom: " & Text2(0).Value & " dem Kunden: " & lKund & " übertragen?", vbYesNo + vbQuestion, "Winkiss Frage:")
                    If iRet = vbYes Then
                        schreibeProtokollKundenkorrektur "Bon: " & lBelegnr & " Datum: " & Text2(0).Value & " für Kunde: " & lKund & " Kunde vorher: " & lkundvorher
                        
                        dBonusf = ermbonusf(lBelegnr, gcKasNum, lDatum, lkundvorher)
                        dUmsatzf = ermUmsatzf(lBelegnr, gcKasNum, lDatum, lkundvorher)
                        
                        sSQL = "Update AFCBUCH set aKunum = " & lKund
                        sSQL = sSQL & " where Belegnr = " & lBelegnr
                        sSQL = sSQL & " and  adate = " & lDatum
                        sSQL = sSQL & " and  aKunum = " & lkundvorher
                        sSQL = sSQL & " and KASNUM = " & cKasnum
                        sSQL = sSQL & " and AZEIT = '" & cUhrzeit & "'"
                        gdBase.Execute sSQL, dbFailOnError
                        
                        sSQL = "Update Kassjour set Kundnr = " & lKund
                        sSQL = sSQL & " where Belegnr = " & lBelegnr
                        sSQL = sSQL & " and  adate = " & lDatum
                        sSQL = sSQL & " and  Kundnr = " & lkundvorher
                        sSQL = sSQL & " and KASNUM = " & cKasnum
                        sSQL = sSQL & " and AZEIT = '" & cUhrzeit & "'"
                        gdBase.Execute sSQL, dbFailOnError
                        
                        BonusVeränderung "negativ", lkundvorher, dBonusf, dUmsatzf
                        
                        BonusVeränderung "positiv", lKund, dBonusf, dUmsatzf
                        
                        anzeige "normal", "erfolgreich", lblanzeige
                    Else
                        anzeige "normal", "Abbruch", lblanzeige
                    End If
                    
                    Exit For
                End If
                
            Next lcount
        End If
    Else
        anzeige "rot", "Bitte einen Satz markieren!", lblanzeige
        Exit Function
    End If
    
    SpeicherBonNeu = True

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherBonNeu"
    Fehler.gsFehlertext = "Im Programmteil Kundenkonto Korrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function
Private Function ermbonusf(lBelegnr As Long, cKas As String, lDatum As Long, lkundvorher As Long) As Double
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermbonusf = 0
    
    sSQL = "Select sum(apreis) as maxi from afcbuch"
    sSQL = sSQL & " where Belegnr = " & lBelegnr
    sSQL = sSQL & " and  adate = " & lDatum
    sSQL = sSQL & " and  aKunum = " & lkundvorher
    sSQL = sSQL & " and KASNUM = " & cKas
    sSQL = sSQL & " and bonus_ok = 'J'"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermbonusf = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermbonusf"
    Fehler.gsFehlertext = "Im Programmteil Kundenkonto Korrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermUmsatzf(lBelegnr As Long, cKas As String, lDatum As Long, lkundvorher As Long) As Double
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermUmsatzf = 0
    
    sSQL = "Select sum(apreis) as maxi  from afcbuch"
    sSQL = sSQL & " where Belegnr = " & lBelegnr
    sSQL = sSQL & " and  adate = " & lDatum
    sSQL = sSQL & " and  aKunum = " & lkundvorher
    sSQL = sSQL & " and KASNUM = " & cKas
    sSQL = sSQL & " and Ums_ok = 'J'"
    sSQL = sSQL & " and aartnr <> 666666"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermUmsatzf = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermUmsatzf"
    Fehler.gsFehlertext = "Im Programmteil Kundenkonto Korrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub SucheBon()
    On Error GoTo LOKAL_ERROR
    
    Dim lartnr      As Long
    Dim lDatum      As Long
    Dim cVon        As String
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cKunde      As String
    Dim czeit       As String
    Dim cBeleg      As String
    Dim cArtNr      As String
    Dim cSatz       As String
    Dim cKasnum     As String
    Dim cBezeich    As String
    
    If Text2(0).Value <> "" Then
        cVon = Text2(0).Value
    Else
        cVon = DateValue(Now)
        Text2(0).Value = DateValue(Now) - 1
    End If
    lDatum = DateValue(cVon)
    
    ctmp = Text1(1).Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        lartnr = 0
    Else
        lartnr = Val(ctmp)
    End If
    
    ctmp = Text1(0).Text
    ctmp = Trim$(ctmp)
    If ctmp = "" Then
        cKasnum = "1"
    Else
        cKasnum = Val(ctmp)
    End If
    List3.Clear

    If gcFilNr > 0 Then
        If lartnr = 0 Then
            cSQL = "Select * from AFCBUCH  where not aArtNR  is null "
        Else
            cSQL = "Select * from AFCBUCH  where aArtNR = " & Trim$(Str$(lartnr)) & " "
        End If
        cSQL = cSQL & " and adate =  " & lDatum
        cSQL = cSQL & " and KASNUM = " & cKasnum
        cSQL = cSQL & " order by kasnum, belegnr "
        
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                cSatz = ""
                If Not IsNull(rsrs!BELEGNR) Then
                    cBeleg = rsrs!BELEGNR
                    If Not IsNull(rsrs!aartnr) Then
                        cArtNr = rsrs!aartnr
                        
                        If Not IsNull(rsrs!ABEZEICH) Then
                            cBezeich = rsrs!ABEZEICH
                            If Not IsNull(rsrs!AZEIT) Then
                                czeit = rsrs!AZEIT
                                If Not IsNull(rsrs!AKUNUM) Then
                                    cKunde = rsrs!AKUNUM
                                    If Not IsNull(rsrs!kasnum) Then
                                        cKasnum = rsrs!kasnum
                                        
                                        cSatz = cSatz & Space(8 - Len(cBeleg)) & cBeleg & Space(12 - Len(czeit)) & czeit
                                        cSatz = cSatz & Space(8 - Len(cArtNr)) & cArtNr & Space(37 - Len(cBezeich)) & cBezeich & Space(10 - Len(cKunde)) & cKunde
                                        cSatz = cSatz & Space(6 - Len(cKasnum)) & cKasnum
                                        List3.AddItem cSatz
                                    End If
                                End If
                            
                            End If
                        End If
                    End If
                End If
            
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
    Else
        If lartnr = 0 Then
            cSQL = "Select * from Kassjour  where not ArtNR  is null "
        Else
            cSQL = "Select * from Kassjour  where ArtNR = " & Trim$(Str$(lartnr)) & " "
        End If
        cSQL = cSQL & " and adate =  " & lDatum
        cSQL = cSQL & " and KASNUM = " & cKasnum
        cSQL = cSQL & " order by kasnum, belegnr "
        
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                cSatz = ""
                If Not IsNull(rsrs!BELEGNR) Then
                    cBeleg = rsrs!BELEGNR
                    If Not IsNull(rsrs!artnr) Then
                        cArtNr = rsrs!artnr
                        
                        If Not IsNull(rsrs!BEZEICH) Then
                            cBezeich = rsrs!BEZEICH
                            If Not IsNull(rsrs!AZEIT) Then
                                czeit = rsrs!AZEIT
                                If Not IsNull(rsrs!Kundnr) Then
                                    cKunde = rsrs!Kundnr
                                    
                                    If Not IsNull(rsrs!kasnum) Then
                                        cKasnum = rsrs!kasnum
                                        
                                        cSatz = cSatz & Space(8 - Len(cBeleg)) & cBeleg & Space(12 - Len(czeit)) & czeit
                                        cSatz = cSatz & Space(8 - Len(cArtNr)) & cArtNr & Space(37 - Len(cBezeich)) & cBezeich & Space(10 - Len(cKunde)) & cKunde
                                        cSatz = cSatz & Space(6 - Len(cKasnum)) & cKasnum
                                        List3.AddItem cSatz
                                    End If
                                    
                                End If
                            
                            End If
                        End If
                    End If
                End If
            
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    anzeige "normal", "", lblanzeige
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheBon"
    Fehler.gsFehlertext = "Im Programmteil Kundenkonto Korrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case 11
            gsHelpstring = "Kundenkonto Korrektur"
            frmWKL110.Show 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenkonto Korrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Farbform Me, lblUeberschrift
    
    List1.AddItem "Beleg Nr    Uhrzeit   Artikel                                         Kunde Kasse"
    
    Text2(0).Value = DateValue(Now)
    Label0(1).Caption = "1"
    Text1(0).Text = gcKasNum
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kundenkonto Korrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub List3_Click()
On Error GoTo LOKAL_ERROR

    Text1(4).Text = Trim(Mid(List3.list(List3.ListIndex), 66, 10))
    Text1(4).Refresh
    
    Text1(0).Text = Trim(Right(List3.list(List3.ListIndex), 6))
    Text1(0).Refresh
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List3_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    Label0(1).Caption = Trim$(Str$(Index))
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            
            Case Is = 4
                
                gF2Prompt.cEsFeld = gcFilNr
            
                gF2Prompt.cFeld = "KUN"
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                End If
                
                If gF2Prompt.cWahl <> "" Then
                    ctmp = gF2Prompt.cWahl
                    Text1(Index).Text = ctmp
                End If
                Text1(Index).SetFocus
                
        End Select
        
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundenkonto Korrektur ist ein Fehler aufgetreten."

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
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
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
        Case 1, 4, 0 'artnr und fil kassnr
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
    End Select
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundenkonto Korrektur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

