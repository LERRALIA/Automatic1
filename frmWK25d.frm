VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWK25d 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Verkaufsprotokoll"
   ClientHeight    =   8625
   ClientLeft      =   1140
   ClientTop       =   1800
   ClientWidth     =   11910
   Icon            =   "frmWK25d.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Timer Timer4 
      Interval        =   200
      Left            =   9360
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   200
      Left            =   8880
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   7920
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   8400
      Top             =   0
   End
   Begin VB.Frame Frame0 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Height          =   855
      Left            =   0
      TabIndex        =   37
      Top             =   7560
      Width           =   11895
      Begin sevCommand3.Command Command0 
         Height          =   855
         Index           =   13
         Left            =   11040
         TabIndex        =   69
         Top             =   0
         Width           =   855
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
         Height          =   855
         Index           =   12
         Left            =   10200
         TabIndex        =   50
         Top             =   0
         Width           =   840
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
         Height          =   855
         Index           =   11
         Left            =   9360
         TabIndex        =   49
         Top             =   0
         Width           =   840
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
         Caption         =   ">"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   855
         Index           =   10
         Left            =   8520
         TabIndex        =   48
         Top             =   0
         Width           =   840
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
         Height          =   855
         Index           =   9
         Left            =   7680
         TabIndex        =   47
         Top             =   0
         Width           =   840
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
         Height          =   855
         Index           =   8
         Left            =   6840
         TabIndex        =   46
         Top             =   0
         Width           =   840
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
         Height          =   855
         Index           =   7
         Left            =   6000
         TabIndex        =   45
         Top             =   0
         Width           =   840
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
         Height          =   855
         Index           =   6
         Left            =   5160
         TabIndex        =   44
         Top             =   0
         Width           =   840
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
         Height          =   855
         Index           =   5
         Left            =   4320
         TabIndex        =   43
         Top             =   0
         Width           =   840
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
         Height          =   855
         Index           =   4
         Left            =   3480
         TabIndex        =   42
         Top             =   0
         Width           =   840
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
         Height          =   855
         Index           =   3
         Left            =   2640
         TabIndex        =   41
         Top             =   0
         Width           =   840
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
         Height          =   855
         Index           =   2
         Left            =   1800
         TabIndex        =   40
         Top             =   0
         Width           =   840
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
         Height          =   855
         Index           =   1
         Left            =   960
         TabIndex        =   39
         Top             =   0
         Width           =   840
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
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   0
         Width           =   840
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
         BackStyle       =   0  'Transparent
         Caption         =   "-1"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   6735
      Left            =   120
      TabIndex        =   29
      Top             =   720
      Width           =   11895
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
         Index           =   13
         Left            =   1320
         MaxLength       =   35
         TabIndex        =   2
         Tag             =   "1"
         Top             =   960
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
         Height          =   405
         Index           =   12
         Left            =   1320
         MaxLength       =   13
         TabIndex        =   103
         Tag             =   "9"
         Top             =   2880
         Width           =   1575
      End
      Begin sevCommand3.Command Command0 
         Height          =   405
         Index           =   16
         Left            =   5520
         TabIndex        =   86
         Top             =   4680
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
         Caption         =   "L"
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
         Height          =   405
         Index           =   11
         Left            =   4560
         MaxLength       =   6
         TabIndex        =   12
         Top             =   2640
         Width           =   1095
      End
      Begin sevCommand3.Command Command0 
         Height          =   405
         Index           =   15
         Left            =   4920
         TabIndex        =   84
         Top             =   3120
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
         Caption         =   "V"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   3000
         MultiSelect     =   2  'Erweitert
         TabIndex        =   83
         Top             =   3840
         Width           =   2415
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
         Height          =   405
         Index           =   9
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   3
         Tag             =   "1"
         Top             =   1440
         Width           =   1095
      End
      Begin sevCommand3.Command Command1 
         Height          =   405
         Index           =   5
         Left            =   2520
         TabIndex        =   80
         Top             =   1440
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
      Begin VB.ListBox List3 
         Height          =   1230
         Left            =   120
         MultiSelect     =   2  'Erweitert
         TabIndex        =   78
         Top             =   3840
         Visible         =   0   'False
         Width           =   2775
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
         Height          =   405
         Index           =   10
         Left            =   1320
         TabIndex        =   6
         Text            =   "alle"
         Top             =   3360
         Width           =   1095
      End
      Begin sevCommand3.Command Command0 
         Height          =   405
         Index           =   14
         Left            =   2520
         TabIndex        =   77
         Top             =   3360
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
      Begin sevCommand3.Command Command1 
         Height          =   405
         Index           =   4
         Left            =   5760
         TabIndex        =   73
         Top             =   1200
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
      Begin sevCommand3.Command Command1 
         Height          =   405
         Index           =   3
         Left            =   5760
         TabIndex        =   72
         Top             =   720
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
      Begin sevCommand3.Command Command1 
         Height          =   405
         Index           =   1
         Left            =   5760
         TabIndex        =   71
         Top             =   240
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   4560
         MaxLength       =   3
         TabIndex        =   9
         Tag             =   "1"
         Top             =   1200
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
         Height          =   405
         Index           =   5
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   8
         Tag             =   "1"
         Top             =   720
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
         Height          =   405
         Index           =   4
         Left            =   4560
         MaxLength       =   6
         TabIndex        =   7
         Tag             =   "1"
         Top             =   240
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
         Height          =   405
         Index           =   3
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "1"
         Top             =   1920
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
         Height          =   405
         Index           =   2
         Left            =   4560
         MaxLength       =   6
         TabIndex        =   11
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0FF&
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
         Height          =   1455
         Left            =   120
         TabIndex        =   62
         Top             =   5160
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Vorjahr"
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
            Index           =   14
            Left            =   2400
            TabIndex        =   101
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "aktuelles Jahr"
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
            Left            =   2400
            TabIndex        =   87
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Heute"
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
            Index           =   7
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Gestern"
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
            Index           =   6
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "aktueller Monat"
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
            Index           =   5
            Left            =   120
            TabIndex        =   64
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Vormonat"
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
            Index           =   2
            Left            =   2400
            TabIndex        =   63
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Datum Voreinstellung"
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
            Index           =   12
            Left            =   120
            TabIndex        =   67
            Top             =   120
            Width           =   2175
         End
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
         Height          =   405
         Index           =   8
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   5
         Tag             =   "9"
         Top             =   2400
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
         Height          =   405
         Index           =   7
         Left            =   4560
         MaxLength       =   5
         TabIndex        =   10
         Tag             =   "8"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Frame Frame5 
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
         Height          =   2055
         Left            =   7680
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "YabandPay"
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
            Index           =   11
            Left            =   2040
            TabIndex        =   110
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "PayPal"
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
            Index           =   10
            Left            =   2040
            TabIndex        =   109
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "Google Pay"
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
            Index           =   9
            Left            =   2040
            TabIndex        =   108
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "Apple Pay"
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
            Index           =   8
            Left            =   2040
            TabIndex        =   107
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "AliPay"
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
            Index           =   7
            Left            =   2040
            TabIndex        =   106
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "Sonstige"
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
            Index           =   6
            Left            =   2040
            TabIndex        =   94
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "EC-Karte"
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
            Index           =   5
            Left            =   2040
            TabIndex        =   26
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "Barclay Card"
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
            Index           =   4
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "Diners Club"
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
            Index           =   3
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "American Express"
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
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "Euro-Card"
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
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "VISA"
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
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Kreditkarten"
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
            Index           =   8
            Left            =   120
            TabIndex        =   57
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   1935
         Left            =   6360
         TabIndex        =   52
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "g. Zahlung"
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
            Index           =   5
            Left            =   120
            TabIndex        =   76
            Top             =   1560
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Kredit"
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
            Index           =   4
            Left            =   120
            TabIndex        =   20
            Top             =   1320
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Kreditkarte"
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
            Index           =   3
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Scheck"
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
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "Bar"
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
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C000&
            Caption         =   "alle"
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
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Zahlungsart"
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
            Index           =   7
            Left            =   120
            TabIndex        =   56
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0080C0FF&
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
         Height          =   3375
         Left            =   6000
         TabIndex        =   36
         Top             =   2760
         Width           =   5655
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Artikel kumuliert - Bestand = 0"
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
            Index           =   15
            Left            =   2520
            TabIndex        =   102
            Top             =   1320
            Width           =   3015
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "kumuliert"
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
            Index           =   18
            Left            =   1800
            TabIndex        =   93
            Top             =   1800
            Visible         =   0   'False
            Width           =   1335
         End
         Begin sevCommand3.Command Command1 
            Height          =   360
            Index           =   11
            Left            =   2760
            TabIndex        =   91
            Top             =   2880
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
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "mit Kollegen VK"
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
            Index           =   17
            Left            =   1800
            TabIndex        =   90
            Top             =   2520
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "aufgeklappt"
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
            Index           =   16
            Left            =   1800
            TabIndex        =   89
            Top             =   2160
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "abgewogene Ware"
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
            Index           =   13
            Left            =   2520
            TabIndex        =   88
            Top             =   960
            Width           =   2775
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Artikel kum. preisgetrennt"
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
            Index           =   11
            Left            =   2520
            TabIndex        =   82
            Top             =   600
            Width           =   2775
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Artikel kumuliert"
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
            Index           =   10
            Left            =   120
            TabIndex        =   75
            Top             =   2520
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Stornoliste"
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
            Index           =   9
            Left            =   2520
            TabIndex        =   74
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Bon - Übersicht"
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
            Index           =   8
            Left            =   120
            TabIndex        =   70
            Top             =   1320
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Kundenprotokoll"
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
            Left            =   120
            TabIndex        =   59
            Top             =   2160
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Retouren"
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
            Index           =   4
            Left            =   120
            TabIndex        =   15
            Top             =   1800
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Kollegen-Verkauf"
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
            Index           =   3
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Tagesprotokoll "
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
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Rechts
            BackColor       =   &H00C0C000&
            Caption         =   "alle Farben"
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
            Index           =   16
            Left            =   1320
            TabIndex        =   92
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Anzeige:"
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
            Index           =   6
            Left            =   120
            TabIndex        =   55
            Top             =   120
            Width           =   1095
         End
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   2
         Left            =   9600
         TabIndex        =   28
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   0
         Left            =   9600
         TabIndex        =   27
         Top             =   2160
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
         Height          =   405
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "2"
         Top             =   480
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
         Height          =   405
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Top             =   0
         Width           =   1095
      End
      Begin sevCommand3.Command Command0 
         Height          =   420
         Index           =   20
         Left            =   3000
         TabIndex        =   95
         ToolTipText     =   "Kalender"
         Top             =   0
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
         Index           =   21
         Left            =   3000
         TabIndex        =   96
         ToolTipText     =   "Kalender"
         Top             =   480
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
      Begin sevCommand3.Command Command8 
         Height          =   165
         Left            =   2640
         TabIndex        =   97
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
         ToolTip         =   "Zurück"
         ToolTipTitle    =   "Zurück"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   165
         Left            =   2640
         TabIndex        =   98
         Top             =   0
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
      Begin sevCommand3.Command Command3 
         Height          =   165
         Left            =   2640
         TabIndex        =   99
         Top             =   720
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
         Left            =   2640
         TabIndex        =   100
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
         ToolTip         =   "Vor"
         ToolTipTitle    =   "Vor"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Bezeichnung:"
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
         Index           =   18
         Left            =   0
         TabIndex        =   105
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "EAN:"
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
         Index           =   17
         Left            =   120
         TabIndex        =   104
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Artnr bis:"
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
         Index           =   13
         Left            =   3240
         TabIndex        =   85
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "AGN:"
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
         Index           =   15
         Left            =   240
         TabIndex        =   81
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "PGN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   79
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "besonderes Merkmal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   120
         TabIndex        =   61
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Kassen Nr.:"
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
         Index           =   10
         Left            =   0
         TabIndex        =   60
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Bon Nr.:"
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
         Index           =   9
         Left            =   3600
         TabIndex        =   58
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Bed.Nr.:"
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
         Left            =   3600
         TabIndex        =   35
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Kund-Nr.:"
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
         Left            =   3480
         TabIndex        =   34
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Lief-Nr.:"
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
         Left            =   3600
         TabIndex        =   33
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Artnr von:"
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
         Left            =   3240
         TabIndex        =   32
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
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
         Index           =   1
         Left            =   0
         TabIndex        =   31
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
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
         Index           =   0
         Left            =   0
         TabIndex        =   30
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte geben Sie ein Suchkriterium ein!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5880
      TabIndex        =   68
      Top             =   240
      Width           =   5895
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
      Caption         =   "Verkaufsprotokoll"
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
      Left            =   120
      TabIndex        =   54
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmWK25d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iPrueD          As Integer
Dim iZaehler        As Integer





Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "KILOVK", gdBase
    
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
Private Sub voreinstellungspeichern()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim bo1 As Integer
    Dim bo2 As Integer
    Dim bo3 As Integer
    Dim bo4 As Integer
    Dim bo5 As Integer
    Dim bo6 As Integer
    
    loeschNEW "VKPE", gdApp
    CreateTableT2 "VKPE", gdApp
    
    bo1 = Option1(7).Value
    bo2 = Option1(6).Value
    bo3 = Option1(5).Value
    bo4 = Option1(2).Value
    bo5 = Option1(12).Value
    bo6 = Option1(14).Value
    
    sSQL = "Insert into VKPE (BO1,BO2,BO3,BO4,BO5,BO6) "
    sSQL = sSQL & " values (" & bo1 & "," & bo2 & "," & bo3 & "," & bo4 & "," & bo5 & "," & bo6
    sSQL = sSQL & ")"
    gdApp.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub WK25dPositionieren()
    On Error GoTo LOKAL_ERROR
    
    With Frame0
        .Top = 7560
        .Left = 0
        .Height = 855
        .Width = 11895
    End With
    
    With Frame1
        .Top = 720
        .Left = 0
        .Height = 6735
        .Width = 11895
    End With
    
    With Frame2
        .Top = 5160
        .Left = 120
        .Height = 1455
        .Width = 4215
    End With
    
    With Frame3
        .Top = 2760
        .Left = 6000
        .Height = 3375
        .Width = 5655
    End With
    
    With Frame4
        .Top = 120
        .Left = 6360
        .Height = 1935
        .Width = 1335
    End With
    
    With Frame5
        .Top = 0
        .Left = 7680
        .Height = 2055
        .Width = 3855
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Function fnPruefeEingabeWK25d()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    fnPruefeEingabeWK25d = 1
    
     For lcount = 0 To 8
        If Trim$(Text1(lcount).Text) <> "" Then
            fnPruefeEingabeWK25d = 0
            Exit Function
        End If
    Next lcount
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEinhabeWK25d"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub LeseDatenWK25d()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim sSQLAGN         As String
    Dim sSQLArtnr       As String
    Dim sSQLRETOUREArtnr As String
    Dim rsrs            As Recordset
    Dim cFeld           As String
    Dim dWert           As Double
    Dim lVon            As Long
    Dim lBis            As Long
    Dim lDatum          As Long
    Dim cArtBez         As String
    Dim cArtNr          As String
    Dim cArtNrb         As String
    Dim cBonNr          As String
    Dim cAWM            As String
    Dim cLinr           As String
    Dim cKundnr         As String
    Dim cDatum          As String
    Dim cAgn            As String
    Dim cEAN            As String
    Dim cBedNr          As String
    Dim ckassnr         As String
    Dim cMopreis        As String
    Dim iFehler         As Integer
    Dim cLBSatz         As String
    Dim lcount          As Long
    Dim bgefunden       As Boolean
    Dim cPfad           As String
    Dim i               As Integer
    Dim cKassen         As String
    Dim cPGNWahl        As String
    Dim dSummeUmsatz    As Double
    Dim lSummeAnzahl    As Long
    Dim lAnzeigeArt     As Long
    
    
    If Option1(1).Value = True Or Option1(8).Value = True Then
        lAnzeigeArt = 1
    ElseIf Option1(1).Value = True Or Option2(5).Value = True Then 'Tagesprotokoll der gemischten
        lAnzeigeArt = 8
    ElseIf Option1(3).Value = True Then
        lAnzeigeArt = 3
    ElseIf Option1(4).Value = True Then
        lAnzeigeArt = 4
    ElseIf Option1(0).Value = True Then
        lAnzeigeArt = 5
    ElseIf Option1(9).Value = True Then
        lAnzeigeArt = 6
    ElseIf Option1(10).Value = True Then
        lAnzeigeArt = 7
    ElseIf Option1(11).Value = True Then
        lAnzeigeArt = 9
    ElseIf Option1(13).Value = True Then
        lAnzeigeArt = 10
    ElseIf Option1(15).Value = True Then
        lAnzeigeArt = 11
    End If
    
    iFehler = 1
    
    cFeld = Text1(0).Text
    If cFeld <> "" Then
        If IsDate(cFeld) Then
            lVon = DateValue(cFeld)
        Else
            anzeigeNew "rot", "Bitte ein gültiges Datum eingeben!", Label6
            Text1(0).SetFocus
            Exit Sub
        End If
    Else
        lVon = 0
    End If
    
    cFeld = Text1(1).Text
    If cFeld <> "" Then
        If IsDate(cFeld) Then
            lBis = DateValue(cFeld)
        Else
            anzeigeNew "rot", "Bitte ein gültiges Datum eingeben!", Label6
            Text1(1).SetFocus
            Exit Sub
        End If
    Else
        lBis = 0
    End If
    
    Screen.MousePointer = 11
    anzeigeNew "normal", "Daten werden ermittelt, bitte warten...", Label6
    
    'artnr von
    cFeld = Trim$(Text1(2).Text)
    If cFeld <> "" Then
        cArtNr = Val(cFeld)
    End If
    
    'artnr bis
    cFeld = Trim$(Text1(11).Text)
    If cFeld <> "" Then
        cArtNrb = Val(cFeld)
    End If
    
    'kasnr
    cFeld = Trim$(Text1(3).Text)
    If cFeld <> "" Then
        ckassnr = Val(cFeld)
    End If
    
    'linr
    cFeld = Trim$(Text1(4).Text)
    If cFeld <> "" Then
        cLinr = Val(cFeld)
    End If
    
    'cKundNr
    cFeld = Trim$(Text1(5).Text)
    If cFeld <> "" Then
        cKundnr = Val(cFeld)
    End If
    
    'cEAN
    cFeld = Trim$(Text1(12).Text)
    If cFeld <> "" Then
        cEAN = Val(cFeld)
    End If
    
    'cArtbez
    cFeld = Trim$(Text1(13).Text)
    If cFeld <> "" Then
        cArtBez = cFeld
    End If
    
    'cAGN
    cFeld = Trim$(Text1(9).Text)
    If cFeld <> "" Then
        cAgn = Val(cFeld)
    End If

    'bed
    cFeld = Trim$(Text1(6).Text)
    If cFeld <> "" Then
        cBedNr = Val(cFeld)
    End If

    'Bonnr
    cFeld = Trim$(Text1(7).Text)
    If cFeld <> "" Then
        cBonNr = Val(cFeld)
    End If
    
    'mopreis
    cFeld = Trim$(Text1(8).Text)
    If cFeld <> "" Then
        cMopreis = Val(cFeld)
    End If
    
    
    
    'AWM
    cFeld = Trim$(Label1(16).Tag)
    If cFeld <> "" Then
        cAWM = Val(cFeld)
    End If
    
    
    sSQLArtnr = ""
    If List1.ListCount > 0 Then
        sSQLArtnr = sSQLArtnr & " and (ARTNR= " & Left(List1.list(0), 6) & " "
        For i = 1 To List1.ListCount - 1
            sSQLArtnr = sSQLArtnr & " or ARTNR= " & Left(List1.list(i), 6) & " "
        Next i
        sSQLArtnr = sSQLArtnr & ") "
    Else
        If cArtNrb <> "" Then
            If cArtNr <> "" Then
                sSQLArtnr = sSQLArtnr & " and ARTNR between " & cArtNr & " and " & cArtNrb & " "
            Else
                sSQLArtnr = sSQLArtnr & " and ARTNR = " & cArtNrb & " "
            End If
        Else
            If cArtNr <> "" Then
                sSQLArtnr = sSQLArtnr & " and ARTNR = " & cArtNr & " "
            End If
        End If
    End If
    
    sSQLRETOUREArtnr = ""
    If List1.ListCount > 0 Then
        sSQLRETOUREArtnr = sSQLRETOUREArtnr & " and (RETOURE.ARTNR= " & Left(List1.list(0), 6) & " "
        For i = 1 To List1.ListCount - 1
            sSQLRETOUREArtnr = sSQLRETOUREArtnr & " or RETOURE.ARTNR= " & Left(List1.list(i), 6) & " "
        Next i
        sSQLRETOUREArtnr = sSQLRETOUREArtnr & ") "
    Else
        If cArtNrb <> "" Then
            If cArtNr <> "" Then
                sSQLRETOUREArtnr = sSQLRETOUREArtnr & " and RETOURE.ARTNR between " & cArtNr & " and " & cArtNrb & " "
            Else
                sSQLRETOUREArtnr = sSQLRETOUREArtnr & " and RETOURE.ARTNR = " & cArtNrb & " "
            End If
        Else
            If cArtNr <> "" Then
                sSQLRETOUREArtnr = sSQLRETOUREArtnr & " and RETOURE.ARTNR = " & cArtNr & " "
            End If
        End If
    End If
    
    'Datenbank wechsel
    
    If lAnzeigeArt = 10 Then
        If NewTableSuchenDBKombi("KILOJOUR", gdBase) = False Then
            CreateTableT2 "KILOJOUR", gdBase
        End If
        
        insert_KiloJour ermBizerbaPfad()
        
        Zeige_VKPROT_GEWICHT lVon, lBis
        
        Exit Sub
    Else
        Datenbankwechsel
        Me.Refresh
    End If
    
    loeschNEW "vkpro4", dabalokal
    loeschNEW "vk271", dabalokal
    loeschNEW "vk272", dabalokal
    loeschNEW "vk277", dabalokal
    loeschNEW "vk278", dabalokal
    loeschNEW "vk279A", dabalokal
    
    If lAnzeigeArt = 1 Then
    
        If cEAN <> "" Then
            loeschNEW "vkEAN", dabalokal
            
            cSQL = "Create Table VKEAN (Artnr long) "
            dabalokal.Execute cSQL, dbFailOnError
            
            cSQL = "Insert into vkean Select ARTNR from Artikel where "
            cSQL = cSQL & " (ean like '" & cEAN & "*'"
            cSQL = cSQL & " or EAN2 like '" & cEAN & "*' "
            cSQL = cSQL & " or EAN3 like '" & cEAN & "*' )"
            dabalokal.Execute cSQL, dbFailOnError
            
            cSQL = "Insert into vkean Select ARTNR from ARTEAN_K where "
            cSQL = cSQL & " ean like '" & cEAN & "*'"
            dabalokal.Execute cSQL, dbFailOnError
            
        End If
        
    
        CreateTable "VK271", dabalokal

        cSQL = "Insert into vk271 Select "
        cSQL = cSQL & " artnr "
        cSQL = cSQL & ", BEZEICH "
        cSQL = cSQL & ", MENGE "
        cSQL = cSQL & ", PREIS "
        cSQL = cSQL & ", adate "
        cSQL = cSQL & ", azeit "
        cSQL = cSQL & ", KUNDNR "
        cSQL = cSQL & ", FILIALE "
        cSQL = cSQL & ", KASNUM "
        cSQL = cSQL & ", LINR "
        cSQL = cSQL & ", LPZ "
        cSQL = cSQL & ", AGN "
        cSQL = cSQL & ", EAN "
        cSQL = cSQL & ", MWST "
        cSQL = cSQL & ", EKPR "
        cSQL = cSQL & ", VKPR "
        cSQL = cSQL & ", MOPREIS "
        cSQL = cSQL & ", BELEGNR "
        cSQL = cSQL & ", best1 "
        cSQL = cSQL & ", RABKENN "
        cSQL = cSQL & ", KK_ART "
        cSQL = cSQL & ", BEDIENER "
        cSQL = cSQL & ", UMS_OK "
        cSQL = cSQL & " from Kassjour  "
    
        cSQL = cSQL & " where Filiale = " & gcFilNr & " "
    
        If lVon > 0 Then
            cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lVon)) & " "
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            Else
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
            End If
        Else
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lBis)) & " "
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            End If
            
        End If
        
        If cLinr <> "" Then
            cSQL = cSQL & " and artnr in (Select artnr from artlief where linr = " & cLinr & ")"
        End If
        
        If cKundnr <> "" Then
            cSQL = cSQL & "and KUNDNR = " & cKundnr & " "
        End If
        
        If cEAN <> "" Then
            cSQL = cSQL & " and artnr in (Select artnr from VKEAN)"

        End If
        
        
        If cArtBez <> "" Then
            cArtBez = SwapStr(cArtBez, " ", "*")
            cSQL = cSQL & " and Bezeich like '*" & cArtBez & "*' "
        End If
        
        If cAgn <> "" Then
            cSQL = cSQL & "and agn = " & cAgn & " "
        End If
        
        If cBedNr <> "" Then
            cSQL = cSQL & "and BEDIENER = " & cBedNr & " "
        End If
        
        
        If cBonNr <> "" Then
            cSQL = cSQL & "and belegnr = " & cBonNr & " "
        End If
        

        If ckassnr <> "" Then
            cSQL = cSQL & "and kasnum = " & ckassnr & " "
        End If
        
        If cMopreis <> "" Then
            cSQL = cSQL & "and mopreis = " & cMopreis & " "
        End If
        
        If Option2(0).Value = True Then         'Alle Zahlungsarten
        
        ElseIf Option2(1).Value = True Then     'BAR-Zahlung
            cSQL = cSQL & "and KK_ART = 'BA' "
        ElseIf Option2(2).Value = True Then     'SCHECK-Zahlung
            cSQL = cSQL & "and KK_ART = 'SC' "
        ElseIf Option2(3).Value = True Then     'KREDITKARTE-Zahlung
            bgefunden = False
            For lcount = 0 To 11
                If Check1(lcount).Value = vbChecked Then
                    bgefunden = True
                    Exit For
                End If
            Next lcount
            If Not bgefunden Then
                Screen.MousePointer = 0
                anzeigeNew "rot", "Bitte mindestens eine Kreditkarte auswählen!", Label6
                
                Check1(0).SetFocus
                Exit Sub
            End If
            bgefunden = False
            If Check1(0).Value = vbChecked Then
                cSQL = cSQL & "and (KK_ART = 'VI'"
                bgefunden = True
            End If
            If Check1(1).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'EU'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'EU'"
                End If
                bgefunden = True
            End If
            
            If Check1(2).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'AE'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'AE'"
                End If
                bgefunden = True
            End If
    
            If Check1(3).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'DC'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'DC'"
                End If
                bgefunden = True
            End If
    
            If Check1(4).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'BC'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'BC'"
                End If
                bgefunden = True
            End If
    
            If Check1(5).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'EC'"
                Else
                    cSQL = cSQL & "and (kK_ART = 'EC'"
                End If
                bgefunden = True
            End If
            
            If Check1(6).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'SO'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'SO'"
                End If
                bgefunden = True
            End If
            
            
            If Check1(7).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'AL'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'AL'"
                End If
                bgefunden = True
            End If
            
            If Check1(8).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'AP'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'AP'"
                End If
                bgefunden = True
            End If
            
            If Check1(9).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'GP'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'GP'"
                End If
                bgefunden = True
            End If
            
            If Check1(10).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'PP'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'PP'"
                End If
                bgefunden = True
            End If
            
            If Check1(11).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'YP'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'YP'"
                End If
                bgefunden = True
            End If
            
            
            
            cSQL = cSQL & ") "
                
        ElseIf Option2(4).Value = True Then 'KREDIT-Zahlung
            cSQL = cSQL & "and KK_ART = 'KR' "
        End If
        
        cSQL = cSQL & BildePGnSQL
        cSQL = cSQL & sSQLArtnr
        cSQL = cSQL & " order by ADATE, AZEIT, ARTNR "
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = "Update vk271 inner join MARKIERUNG on vk271.artnr = MARKIERUNG.artnr "
        cSQL = cSQL & " and vk271.MENGE = MARKIERUNG.MENGE "
        cSQL = cSQL & " and vk271.adate = MARKIERUNG.adate "
        cSQL = cSQL & " and vk271.BELEGNR = MARKIERUNG.BELEGNR "
        cSQL = cSQL & " Set vk271.RABKENN = 'x' "
        dabalokal.Execute cSQL, dbFailOnError
        
        If cLinr <> "" Then
        
        
            cSQL = "Update vk271 inner join Artlief on vk271.artnr = Artlief.artnr "
            cSQL = cSQL & " Set vk271.linr = Artlief.linr "
            cSQL = cSQL & " where artlief.linr = " & cLinr & " "
            dabalokal.Execute cSQL, dbFailOnError
        
    
        End If
        
        cSQL = "Update vk271 inner join Artlief on vk271.linr = Artlief.linr "
        cSQL = cSQL & " and vk271.artnr = Artlief.artnr "
        cSQL = cSQL & " Set vk271.Libesnr = Artlief.Libesnr "
        dabalokal.Execute cSQL, dbFailOnError
        
        loeschNEW "temp_vk271", dabalokal
        
        cSQL = "Select * into temp_vk271 from vk271 "
        dabalokal.Execute cSQL, dbFailOnError
        
        loeschNEW "vk271", dabalokal
        CreateTable "VK271", dabalokal
        
        cSQL = "Insert into vk271 Select * from temp_vk271 "
        cSQL = cSQL & " order by ADATE, AZEIT, ARTNR "
        dabalokal.Execute cSQL, dbFailOnError
        
        
        loeschNEW "KOPF271", dabalokal
        CreateTable "KOPF271", dabalokal
        
        
        If cMopreis = "" Then
            cMopreis = "0"
        End If
        cSQL = "Insert into KOPF271 (MOPREIS) values ( " & cMopreis & ")"
        dabalokal.Execute cSQL, dbFailOnError
        
        If Not Datendrin("vk271", dabalokal) Then
            anzeigeNew "rot", "Es wurden keine Daten ermittelt.", Label6
            Exit Sub
        Else
            
            If Option1(1).Value = True Then
                reportbildschirmApp "dWKL27", "aWKL271"
            Else
                reportbildschirmApp "dWKL27", "aWKL274"
            End If
'            anzeigeNew "normal", "Bitte geben Sie ein Suchkriterium ein!", Label6
        End If
        
            
    End If
    
    
    iFehler = 3
    
    If lAnzeigeArt = 5 Then
    
        CreateTable "VK272", dabalokal
        
        cSQL = "Insert into vk272 Select "
        
        cSQL = cSQL & " artnr "
        cSQL = cSQL & ", BEZEICH "
        cSQL = cSQL & ", MENGE "
        cSQL = cSQL & ", PREIS "
        cSQL = cSQL & ", adate "
        cSQL = cSQL & ", azeit "
        cSQL = cSQL & ", KUNDNR "
        cSQL = cSQL & ", FILIALE "
        cSQL = cSQL & ", KASNUM "
        cSQL = cSQL & ", LINR "
        cSQL = cSQL & ", LPZ "
        cSQL = cSQL & ", AGN "
        cSQL = cSQL & ", EAN "
        cSQL = cSQL & ", MWST "
        cSQL = cSQL & ", EKPR "
        cSQL = cSQL & ", VKPR "
        cSQL = cSQL & ", MOPREIS "
        cSQL = cSQL & ", BELEGNR "
        cSQL = cSQL & ", best1 "
        cSQL = cSQL & ", RABKENN "
        cSQL = cSQL & ", KK_ART "
        cSQL = cSQL & ", BEDIENER "
        cSQL = cSQL & ", UMS_OK "
        
        cSQL = cSQL & ", '' as KuName "
        cSQL = cSQL & ", '' as KuVName  "
        
        
        cSQL = cSQL & " from Kassjour  "
    
        cSQL = cSQL & " where Filiale = " & gcFilNr & " "
        cSQL = cSQL & " and KUNDNR <> 0 "
    
        If lVon > 0 Then
            cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lVon)) & " "
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            Else
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
            End If
        Else
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lBis)) & " "
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            End If
            
        End If
    
'        If cArtNr <> "" Then
'            cSQL = cSQL & "and ARTNR = " & cArtNr & " "
'        End If
        If cLinr <> "" Then
            cSQL = cSQL & " and artnr in (Select artnr from artlief where linr = " & cLinr & ")"
'            cSQL = cSQL & "and LINR = " & cLinr & " "
        End If
        If cKundnr <> "" Then
            cSQL = cSQL & "and KUNDNR = " & cKundnr & " "
        End If
        If cEAN <> "" Then
            cSQL = cSQL & "and EAN = '" & cEAN & "' "
        End If
        
        
        If cArtBez <> "" Then
            cArtBez = SwapStr(cArtBez, " ", "*")
            cSQL = cSQL & " and Bezeich like '*" & cArtBez & "*' "
        End If
        
        
        
        If cAgn <> "" Then
            cSQL = cSQL & "and agn = " & cAgn & " "
        End If
        If cBedNr <> "" Then
            cSQL = cSQL & "and BEDIENER = " & cBedNr & " "
        End If
        If cBonNr <> "" Then
            cSQL = cSQL & "and belegnr = " & cBonNr & " "
        End If
        If ckassnr <> "" Then
            cSQL = cSQL & "and kasnum = " & ckassnr & " "
        End If
        
        If cMopreis <> "" Then
            cSQL = cSQL & "and mopreis = " & cMopreis & " "
        End If
        
        
        If Option2(0).Value = True Then         'Alle Zahlungsarten
        
        ElseIf Option2(1).Value = True Then     'BAR-Zahlung
            cSQL = cSQL & "and KK_ART = 'BA' "
        ElseIf Option2(2).Value = True Then     'SCHECK-Zahlung
            cSQL = cSQL & "and KK_ART = 'SC' "
        ElseIf Option2(3).Value = True Then     'KREDITKARTE-Zahlung
            bgefunden = False
            For lcount = 0 To 11
                If Check1(lcount).Value = vbChecked Then
                    bgefunden = True
                    Exit For
                End If
            Next lcount
            If Not bgefunden Then
                Screen.MousePointer = 0
                anzeigeNew "rot", "Bitte mindestens eine Kreditkarte auswählen!", Label6
                Check1(0).SetFocus
                Exit Sub
            End If
            bgefunden = False
            If Check1(0).Value = vbChecked Then
                cSQL = cSQL & "and (KK_ART = 'VI'"
                bgefunden = True
            End If
            If Check1(1).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'EU'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'EU'"
                End If
                bgefunden = True
            End If
            
            If Check1(2).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'AE'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'AE'"
                End If
                bgefunden = True
            End If
    
            If Check1(3).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'DC'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'DC'"
                End If
                bgefunden = True
            End If
    
            If Check1(4).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'BC'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'BC'"
                End If
                bgefunden = True
            End If
    
            If Check1(5).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'EC'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'EC'"
                End If
                bgefunden = True
            End If
            
            If Check1(6).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'SO'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'SO'"
                End If
                bgefunden = True
            End If
            
            
            
            If Check1(7).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'AL'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'AL'"
                End If
                bgefunden = True
            End If
            
            If Check1(8).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'AP'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'AP'"
                End If
                bgefunden = True
            End If
            
            If Check1(9).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'GP'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'GP'"
                End If
                bgefunden = True
            End If
            
            If Check1(10).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'PP'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'PP'"
                End If
                bgefunden = True
            End If
            
            If Check1(11).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'YP'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'YP'"
                End If
                bgefunden = True
            End If
            
            
            cSQL = cSQL & ") "
                
        ElseIf Option2(4).Value = True Then 'KREDIT-Zahlung
            cSQL = cSQL & "and KK_ART = 'KR' "
        End If
        
        cSQL = cSQL & BildePGnSQL
        cSQL = cSQL & sSQLArtnr
        
        cSQL = cSQL & " order by ADATE, AZEIT, ARTNR "
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = "Update vk272 inner join kunden on vk272.kundnr = kunden.Kundnr "
        cSQL = cSQL & " Set vk272.kuname = kunden.Name , vk272.KuvName=kunden.Vorname"
        dabalokal.Execute cSQL, dbFailOnError
        
        If Not Datendrin("vk272", dabalokal) Then
            anzeigeNew "rot", "Es wurden keine Daten ermittelt.", Label6
            Exit Sub
        Else
            If Check1(16).Value = vbChecked Then
                reportbildschirmApp "dWKL27c", "aWKL272b"
            Else
                reportbildschirmApp "dWKL27c", "aWKL272"
            End If

        End If
        
            
    End If
    
    
    iFehler = 3
    
    
    If lAnzeigeArt = 8 Then
    
        CreateTable "VK278", dabalokal

        cSQL = "Insert into vk278 Select "
        
        cSQL = cSQL & " artnr "
        cSQL = cSQL & ", BEZEICH "
        cSQL = cSQL & ", MENGE "
        cSQL = cSQL & ", PREIS "
        cSQL = cSQL & ", adate "
        cSQL = cSQL & ", azeit "
        cSQL = cSQL & ", KUNDNR "
        cSQL = cSQL & ", FILIALE "
        cSQL = cSQL & ", KASNUM "
        cSQL = cSQL & ", LINR "
        cSQL = cSQL & ", LPZ "
        cSQL = cSQL & ", AGN "
        cSQL = cSQL & ", EAN "
        cSQL = cSQL & ", MWST "
        cSQL = cSQL & ", EKPR "
        cSQL = cSQL & ", VKPR "
        cSQL = cSQL & ", MOPREIS "
        cSQL = cSQL & ", BELEGNR "
        cSQL = cSQL & ", best1 "
        cSQL = cSQL & ", RABKENN "
        cSQL = cSQL & ", KK_ART "
        cSQL = cSQL & ", BEDIENER "
        cSQL = cSQL & ", UMS_OK "
        
        
        cSQL = cSQL & " from Kassjour "
    
        cSQL = cSQL & " where Filiale = " & gcFilNr & " "
    
        If lVon > 0 Then
            cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lVon)) & " "
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            Else
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
            End If
        Else
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lBis)) & " "
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            End If
            
        End If
    
'        If cArtNr <> "" Then
'            cSQL = cSQL & "and ARTNR = " & cArtNr & " "
'        End If
        If cLinr <> "" Then
            cSQL = cSQL & " and artnr in (Select artnr from artlief where linr = " & cLinr & ")"
'            cSQL = cSQL & "and LINR = " & cLinr & " "
        End If
        If cKundnr <> "" Then
            cSQL = cSQL & "and KUNDNR = " & cKundnr & " "
        End If
        If cEAN <> "" Then
            cSQL = cSQL & "and EAN = '" & cEAN & "' "
        End If
        
        If cArtBez <> "" Then
            cArtBez = SwapStr(cArtBez, " ", "*")
            cSQL = cSQL & " and Bezeich like '*" & cArtBez & "*' "
        End If
        
        
        
        If cAgn <> "" Then
            cSQL = cSQL & "and agn = " & cAgn & " "
        End If
        If cBedNr <> "" Then
            cSQL = cSQL & "and BEDIENER = " & cBedNr & " "
        End If
        
        
        If cBonNr <> "" Then
            cSQL = cSQL & "and belegnr = " & cBonNr & " "
        End If
        

        If ckassnr <> "" Then
            cSQL = cSQL & "and kasnum = " & ckassnr & " "
        End If
        
        If cMopreis <> "" Then
            cSQL = cSQL & "and mopreis = " & cMopreis & " "
        End If
        cSQL = cSQL & "and KK_ART = 'GZ' "
        
        cSQL = cSQL & BildePGnSQL
        cSQL = cSQL & sSQLArtnr
        
        cSQL = cSQL & " order by ADATE, AZEIT, ARTNR "
        dabalokal.Execute cSQL, dbFailOnError
 
        
        loeschNEW "KOPF271", dabalokal
        CreateTable "KOPF271", dabalokal
        
        
        If cMopreis = "" Then
            cMopreis = "0"
        End If
        cSQL = "Insert into KOPF271 (MOPREIS) values ( " & cMopreis & ")"
        dabalokal.Execute cSQL, dbFailOnError
        
        If Not Datendrin("vk278", dabalokal) Then
            anzeigeNew "rot", "Es wurden keine Daten ermittelt.", Label6
            Exit Sub
        Else
        
            Zahldetails "VK278"
            
            If Option1(1).Value = True Then
                reportbildschirmApp "dWKL27", "aWKL278"
            End If
'            anzeigeNew "normal", "Bitte geben Sie ein Suchkriterium ein!", Label6
        End If
        
            
    End If
    
    
    iFehler = 3
'    End If
    
    If lAnzeigeArt = 3 Then
    
        CreateTable "VK272", dabalokal
        
        cSQL = "Insert into vk272 Select "
        
        cSQL = cSQL & " artnr "
        cSQL = cSQL & ", BEZEICH "
        cSQL = cSQL & ", MENGE "
        cSQL = cSQL & ", PREIS "
        cSQL = cSQL & ", adate "
        cSQL = cSQL & ", azeit "
        cSQL = cSQL & ", KUNDNR "
        cSQL = cSQL & ", FILIALE "
        cSQL = cSQL & ", KASNUM "
        cSQL = cSQL & ", LINR "
        cSQL = cSQL & ", LPZ "
        cSQL = cSQL & ", AGN "
        cSQL = cSQL & ", EAN "
        cSQL = cSQL & ", MWST "
        cSQL = cSQL & ", EKPR "
        cSQL = cSQL & ", VKPR "
        cSQL = cSQL & ", BELEGNR "
        cSQL = cSQL & ", best1 "
        cSQL = cSQL & ", RABKENN "
        cSQL = cSQL & ", BEDIENER "
        
        cSQL = cSQL & ", '' as KuName "
        cSQL = cSQL & ", '' as KuVName  "
        
        cSQL = cSQL & " from kollverk  "
        cSQL = cSQL & " where Filiale = " & gcFilNr & " "
        cSQL = cSQL & " and KUNDNR <> 0 "
    
        If lVon > 0 Then
            cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lVon)) & " "
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            Else
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
            End If
        Else
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lBis)) & " "
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            End If
            
        End If
    
'        If cArtNr <> "" Then
'            cSQL = cSQL & "and ARTNR = " & cArtNr & " "
'        End If
        If cLinr <> "" Then
            cSQL = cSQL & " and artnr in (Select artnr from artlief where linr = " & cLinr & ")"
'            cSQL = cSQL & "and LINR = " & cLinr & " "
        End If
        If cKundnr <> "" Then
            cSQL = cSQL & "and KUNDNR = " & cKundnr & " "
        End If
        If cEAN <> "" Then
            cSQL = cSQL & "and EAN = '" & cEAN & "' "
        End If
        
        
        If cArtBez <> "" Then
            cArtBez = SwapStr(cArtBez, " ", "*")
            cSQL = cSQL & " and Bezeich like '*" & cArtBez & "*' "
        End If
        
        
        If cAgn <> "" Then
            cSQL = cSQL & "and agn = " & cAgn & " "
        End If
        If cBedNr <> "" Then
            cSQL = cSQL & "and BEDIENER = " & cBedNr & " "
        End If
        If cBonNr <> "" Then
            cSQL = cSQL & "and belegnr = " & cBonNr & " "
        End If
        If ckassnr <> "" Then
            cSQL = cSQL & "and kasnum = " & ckassnr & " "
        End If
        
        If cMopreis <> "" Then
            cSQL = cSQL & "and mopreis = " & cMopreis & " "
        End If
        
        cSQL = cSQL & BildePGnSQL
        cSQL = cSQL & sSQLArtnr
        
        cSQL = cSQL & " order by ADATE, AZEIT, ARTNR "
        dabalokal.Execute cSQL, dbFailOnError
    

    
        cSQL = "Update vk272 inner join kunden on vk272.kundnr = kunden.Kundnr "
        cSQL = cSQL & " Set vk272.kuname = kunden.Name , vk272.KuvName=kunden.Vorname"
        dabalokal.Execute cSQL, dbFailOnError
        
        If Not Datendrin("vk272", dabalokal) Then
            anzeigeNew "rot", "Es wurden keine Daten ermittelt.", Label6
            Exit Sub
        Else
            reportbildschirmApp "dWKL27c", "aWKL273"
'            anzeigeNew "normal", "Bitte geben Sie ein Suchkriterium ein!", Label6
        End If
        
    End If

    iFehler = 4
    If lAnzeigeArt = 6 Then
    
    
        CreateTable "VK271", dabalokal

        cSQL = "Insert into vk271 Select "
        
        cSQL = cSQL & " artnr "
        cSQL = cSQL & ", BEZEICH "
        cSQL = cSQL & ", MENGE "
        cSQL = cSQL & ", PREIS "
        cSQL = cSQL & ", adate "
        cSQL = cSQL & ", azeit "
        cSQL = cSQL & ", KUNDNR "
        cSQL = cSQL & ", FILIALE "
        cSQL = cSQL & ", KASNUM "
        cSQL = cSQL & ", LINR "
        cSQL = cSQL & ", LPZ "
        cSQL = cSQL & ", AGN "
        cSQL = cSQL & ", EAN "
        cSQL = cSQL & ", MWST "
        cSQL = cSQL & ", EKPR "
        cSQL = cSQL & ", VKPR "
        cSQL = cSQL & ", MOPREIS "
        cSQL = cSQL & ", BELEGNR "
        cSQL = cSQL & ", best1 "
        cSQL = cSQL & ", RABKENN "
        cSQL = cSQL & ", KK_ART "
        cSQL = cSQL & ", BEDIENER "
        cSQL = cSQL & ", UMS_OK "
        
        cSQL = cSQL & " from Kassjour  "
    
        cSQL = cSQL & " where Filiale = " & gcFilNr & " "
'        cSQL = cSQL & " and menge < 0 "
        cSQL = cSQL & " and preis < 0 "
    
        If lVon > 0 Then
            cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lVon)) & " "
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            Else
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
            End If
        Else
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lBis)) & " "
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            End If
            
        End If
    
'        If cArtNr <> "" Then
'            cSQL = cSQL & "and ARTNR = " & cArtNr & " "
'        End If
        If cLinr <> "" Then
            cSQL = cSQL & " and artnr in (Select artnr from artlief where linr = " & cLinr & ")"
'            cSQL = cSQL & "and LINR = " & cLinr & " "
        End If
        If cKundnr <> "" Then
            cSQL = cSQL & "and KUNDNR = " & cKundnr & " "
        End If
        If cEAN <> "" Then
            cSQL = cSQL & "and EAN = '" & cEAN & "' "
        End If
        
        If cArtBez <> "" Then
            cArtBez = SwapStr(cArtBez, " ", "*")
            cSQL = cSQL & " and Bezeich like '*" & cArtBez & "*' "
        End If
        
        If cAgn <> "" Then
            cSQL = cSQL & "and agn = " & cAgn & " "
        End If
        If cBedNr <> "" Then
            cSQL = cSQL & "and BEDIENER = " & cBedNr & " "
        End If
        If cBonNr <> "" Then
            cSQL = cSQL & "and belegnr = " & cBonNr & " "
        End If
        If ckassnr <> "" Then
            cSQL = cSQL & "and kasnum = " & ckassnr & " "
        End If
        
        If cMopreis <> "" Then
            cSQL = cSQL & "and mopreis = " & cMopreis & " "
        End If
        
        If Option2(0).Value = True Then         'Alle Zahlungsarten
        
        ElseIf Option2(1).Value = True Then     'BAR-Zahlung
            cSQL = cSQL & "and KK_ART = 'BA' "
        ElseIf Option2(2).Value = True Then     'SCHECK-Zahlung
            cSQL = cSQL & "and KK_ART = 'SC' "
        ElseIf Option2(3).Value = True Then     'KREDITKARTE-Zahlung
            bgefunden = False
            For lcount = 0 To 11
                If Check1(lcount).Value = vbChecked Then
                    bgefunden = True
                    Exit For
                End If
            Next lcount
            If Not bgefunden Then
                Screen.MousePointer = 0
                anzeigeNew "rot", "Bitte mindestens eine Kreditkarte auswählen!", Label6
                
                Check1(0).SetFocus
                Exit Sub
            End If
            bgefunden = False
            If Check1(0).Value = vbChecked Then
                cSQL = cSQL & "and (KK_ART = 'VI'"
                bgefunden = True
            End If
            If Check1(1).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'EU'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'EU'"
                End If
                bgefunden = True
            End If
            
            If Check1(2).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'AE'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'AE'"
                End If
                bgefunden = True
            End If
    
            If Check1(3).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'DC'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'DC'"
                End If
                bgefunden = True
            End If
    
            If Check1(4).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'BC'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'BC'"
                End If
                bgefunden = True
            End If
    
            If Check1(5).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'EC'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'EC'"
                End If
                bgefunden = True
            End If
            
            If Check1(6).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'SO'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'SO'"
                End If
                bgefunden = True
            End If
            
            If Check1(7).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'AL'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'AL'"
                End If
                bgefunden = True
            End If
            
            If Check1(8).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'AP'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'AP'"
                End If
                bgefunden = True
            End If
            
            If Check1(9).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'GP'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'GP'"
                End If
                bgefunden = True
            End If
            
            If Check1(10).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'PP'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'PP'"
                End If
                bgefunden = True
            End If
            
            If Check1(11).Value = vbChecked Then
                If bgefunden Then
                    cSQL = cSQL & " or KK_ART = 'YP'"
                Else
                    cSQL = cSQL & "and (KK_ART = 'YP'"
                End If
                bgefunden = True
            End If
            
            
            cSQL = cSQL & ") "
                
        ElseIf Option2(4).Value = True Then 'KREDIT-Zahlung
            cSQL = cSQL & "and KK_ART = 'KR' "
        End If
        
        cSQL = cSQL & BildePGnSQL
        cSQL = cSQL & sSQLArtnr
        
        cSQL = cSQL & " order by ADATE, AZEIT, ARTNR "
        dabalokal.Execute cSQL, dbFailOnError
        
        loeschNEW "KOPF271", dabalokal
        CreateTable "KOPF271", dabalokal
        
        
        If cMopreis = "" Then
            cMopreis = "0"
        End If
        cSQL = "Insert into KOPF271 (MOPREIS) values ( " & cMopreis & ")"
        dabalokal.Execute cSQL, dbFailOnError
        
        If Not Datendrin("vk271", dabalokal) Then
            anzeigeNew "rot", "Es wurden keine Daten ermittelt.", Label6
            Exit Sub
        Else
            reportbildschirmApp "dWKL27", "aWKL275"
        End If
    End If
       
    If lAnzeigeArt = 4 Then
    
        CreateTableT2 "VKPRO4", dabalokal

        cSQL = "Insert into VKPRO4 Select "
        cSQL = cSQL & " ARTNR "
        cSQL = cSQL & ", BEZEICH "
        cSQL = cSQL & ", MENGE "
        cSQL = cSQL & ", PREIS "
        cSQL = cSQL & ", ADATE "
        cSQL = cSQL & ", AZEIT "
        cSQL = cSQL & ", BEDIENER "
        cSQL = cSQL & ", KUNDNR "
        cSQL = cSQL & ", FILIALE "
        cSQL = cSQL & ", KASNUM "
        cSQL = cSQL & ", LINR "
        cSQL = cSQL & ", LPZ "
        cSQL = cSQL & ", AGN "
        cSQL = cSQL & ", EAN "
        cSQL = cSQL & ", MWST "
        cSQL = cSQL & ", 0.0 as LEKPR "
        cSQL = cSQL & ", EKPR "
        cSQL = cSQL & ", VKPR "
        cSQL = cSQL & ", MOPPREIS "
        cSQL = cSQL & ", BELEGNR "
        cSQL = cSQL & ", IN_BESTELL "
        cSQL = cSQL & ", BEST1 "
        cSQL = cSQL & ", RABKENN "
        cSQL = cSQL & " from RETOURE "
        cSQL = cSQL & " where Retoure.Filiale = " & gcFilNr & " "
        
        If lVon > 0 Then
            cSQL = cSQL & "and Retoure.ADATE >= " & Trim$(Str$(lVon)) & " "
            If lBis > 0 Then
                cSQL = cSQL & "and Retoure.ADATE <= " & Trim$(Str$(lBis)) & " "
            Else
                cSQL = cSQL & "and Retoure.ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
            End If
        Else
            If lBis > 0 Then
                cSQL = cSQL & "and Retoure.ADATE >= " & Trim$(Str$(lBis)) & " "
                cSQL = cSQL & "and Retoure.ADATE <= " & Trim$(Str$(lBis)) & " "
            End If
        End If
        
        If cLinr <> "" Then
            cSQL = cSQL & "and Retoure.LINR = " & cLinr & " "
        End If
        If cKundnr <> "" Then
            cSQL = cSQL & "and Retoure.KUNDNR = " & cKundnr & " "
        End If
        If cEAN <> "" Then
            cSQL = cSQL & "and Retoure.EAN = '" & cEAN & "' "
        End If
        
        If cArtBez <> "" Then
            cArtBez = SwapStr(cArtBez, " ", "*")
            cSQL = cSQL & " and Retoure.Bezeich like '*" & cArtBez & "*' "
        End If
        
        
        If cAgn <> "" Then
            cSQL = cSQL & "and artikel.agn = " & cAgn & " "
        End If
        If cBedNr <> "" Then
            cSQL = cSQL & "and Retoure.BEDIENER = " & cBedNr & " "
        End If
        If cBonNr <> "" Then
            cSQL = cSQL & "and Retoure.belegnr = " & cBonNr & " "
        End If
        If ckassnr <> "" Then
            cSQL = cSQL & "and kasnum = " & ckassnr & " "
        End If
        
        If cMopreis <> "" Then
            cSQL = cSQL & "and Mopreis = " & cMopreis & " "
        End If
        
        cSQL = cSQL & sSQLRETOUREArtnr
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = " Update vkpro4 inner join artikel on vkpro4.artnr = artikel.artnr "
        cSQL = cSQL & " Set vkpro4.Preis  = Round(artikel.KVKPR1,2)"
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = "Update vkpro4 inner join Artlief on vkpro4.linr = Artlief.linr "
        cSQL = cSQL & " and vkpro4.artnr = Artlief.artnr "
        cSQL = cSQL & " Set vkpro4.Libesnr = Artlief.Libesnr "
        cSQL = cSQL & " , vkpro4.LEKPR = Artlief.LEKPR "
        dabalokal.Execute cSQL, dbFailOnError
        
        loeschNEW "KOPFRET", dabalokal
        CreateTableT2 "KOPFRET", dabalokal
        
        If cMopreis = "" Then
            cMopreis = "0"
        End If

        cSQL = "Insert into KOPFRET (MOPREIS,Von,Bis,FILIALE) values (" & cMopreis & ",'" & Text1(0).Text & "','" & Text1(1).Text & "','')"
        dabalokal.Execute cSQL, dbFailOnError
        
        If Not Datendrin("vkpro4", dabalokal) Then
            anzeigeNew "rot", "Es wurden keine Daten ermittelt.", Label6
            Exit Sub
        Else
            loeschNEW "vk27b", dabalokal
            
            cSQL = "Select * into vk27b from vkpro4 "
            dabalokal.Execute cSQL, dbFailOnError
            
            If Check1(18).Value = vbChecked Then
                reportbildschirmApp "dWKL27b", "aWKL27c"
            Else
                reportbildschirmApp "dWKL27b", "aWKL27b"
            End If
        End If
    End If
    
    
    If lAnzeigeArt = 7 Then
    
        CreateTable "VK277", dabalokal

        cSQL = "Insert into vk277 Select "
        cSQL = cSQL & " artnr "
        cSQL = cSQL & ", sum(MENGE) as SumMenge "
        cSQL = cSQL & ", sum(Preis) as SumPreis "
        cSQL = cSQL & " from Kassjour  "
        cSQL = cSQL & " where Filiale = " & gcFilNr & " "
    
        If lVon > 0 Then
            cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lVon)) & " "
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            Else
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
            End If
        Else
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lBis)) & " "
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            End If
        End If
        
        If cLinr <> "" Then
            cSQL = cSQL & " and artnr in (Select artnr from artlief where linr = " & cLinr & ")"
'            cSQL = cSQL & "and LINR = " & cLinr & " "
        End If
        If cKundnr <> "" Then
            cSQL = cSQL & "and KUNDNR = " & cKundnr & " "
        End If
        If cEAN <> "" Then
            cSQL = cSQL & "and EAN = '" & cEAN & "' "
        End If
        
        If cArtBez <> "" Then
            cArtBez = SwapStr(cArtBez, " ", "*")
            cSQL = cSQL & " and Bezeich like '*" & cArtBez & "*' "
        End If
        
        
        
        If cAgn <> "" Then
            cSQL = cSQL & "and agn = " & cAgn & " "
        End If
        If cBedNr <> "" Then
            cSQL = cSQL & "and BEDIENER = " & cBedNr & " "
        End If
        
        
        If cBonNr <> "" Then
            cSQL = cSQL & "and belegnr = " & cBonNr & " "
        End If
        
        If ckassnr <> "" Then
            cSQL = cSQL & "and kasnum = " & ckassnr & " "
        End If
        
        If cMopreis <> "" Then
            cSQL = cSQL & "and mopreis = " & cMopreis & " "
        End If
        
        cSQL = cSQL & BildePGnSQL
        
        cSQL = cSQL & sSQLArtnr
        
        cSQL = cSQL & " group by  "
        cSQL = cSQL & " artnr "
'        MsgBox cSQL
        dabalokal.Execute cSQL, dbFailOnError
        
        'mit Kollegen
        If Check1(17).Value = vbChecked Then
            cSQL = "Insert into vk277 Select "
            cSQL = cSQL & " artnr "
            cSQL = cSQL & ", sum(MENGE) as SumMenge "
            cSQL = cSQL & ", sum(Preis) as SumPreis "
            cSQL = cSQL & " from Kollverk  "
            cSQL = cSQL & " where Filiale = " & gcFilNr & " "
        
            If lVon > 0 Then
                cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lVon)) & " "
                If lBis > 0 Then
                    cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
                Else
                    cSQL = cSQL & "and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
                End If
            Else
                If lBis > 0 Then
                    cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lBis)) & " "
                    cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
                End If
            End If
            
            If cLinr <> "" Then
                cSQL = cSQL & " and artnr in (Select artnr from artlief where linr = " & cLinr & ")"
'                cSQL = cSQL & "and LINR = " & cLinr & " "
            End If
            If cKundnr <> "" Then
                cSQL = cSQL & "and KUNDNR = " & cKundnr & " "
            End If
            If cEAN <> "" Then
                cSQL = cSQL & "and EAN = '" & cEAN & "' "
            End If
        
        
            If cArtBez <> "" Then
                cArtBez = SwapStr(cArtBez, " ", "*")
                cSQL = cSQL & " and Bezeich like '*" & cArtBez & "*' "
            End If
        
        
            If cAgn <> "" Then
                cSQL = cSQL & "and agn = " & cAgn & " "
            End If
            If cBedNr <> "" Then
                cSQL = cSQL & "and BEDIENER = " & cBedNr & " "
            End If
            
            
            If cBonNr <> "" Then
                cSQL = cSQL & "and belegnr = " & cBonNr & " "
            End If
            
            If ckassnr <> "" Then
                cSQL = cSQL & "and kasnum = " & ckassnr & " "
            End If
            
            If cMopreis <> "" Then
                cSQL = cSQL & "and mopreis = " & cMopreis & " "
            End If
            
            cSQL = cSQL & BildePGnSQL
            
            cSQL = cSQL & sSQLArtnr
            
            cSQL = cSQL & " group by  "
            cSQL = cSQL & " artnr "
    '        MsgBox cSQL
            dabalokal.Execute cSQL, dbFailOnError
        End If
            
        cSQL = " Update VK277 "
        cSQL = cSQL & " set best1 = 0 "
        cSQL = cSQL & " , LUG = 0 "
        dabalokal.Execute cSQL, dbFailOnError
        
        
        loeschNEW "wKass", dabalokal
        
        cSQL = " Select distinct(kasnum) as wKasse "
        cSQL = cSQL & " into wKass from Kassjour  "
        cSQL = cSQL & " where Filiale = " & gcFilNr & " "
        If lVon > 0 Then
            cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lVon)) & " "
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            Else
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
            End If
        Else
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lBis)) & " "
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            End If
        End If

        If cLinr <> "" Then
            cSQL = cSQL & " and artnr in (Select artnr from artlief where linr = " & cLinr & ")"
'            cSQL = cSQL & "and LINR = " & cLinr & " "
        End If
        If cKundnr <> "" Then
            cSQL = cSQL & "and KUNDNR = " & cKundnr & " "
        End If
        If cEAN <> "" Then
            cSQL = cSQL & "and EAN = '" & cEAN & "' "
        End If
        If cAgn <> "" Then
            cSQL = cSQL & "and agn = " & cAgn & " "
        End If
        If cBedNr <> "" Then
            cSQL = cSQL & "and BEDIENER = " & cBedNr & " "
        End If
        
        
        If cBonNr <> "" Then
            cSQL = cSQL & "and belegnr = " & cBonNr & " "
        End If
        
        If ckassnr <> "" Then
            cSQL = cSQL & "and kasnum = " & ckassnr & " "
        End If
        
        If cMopreis <> "" Then
            cSQL = cSQL & "and mopreis = " & cMopreis & " "
        End If
        
        cSQL = cSQL & BildePGnSQL
        
        cSQL = cSQL & sSQLArtnr
        
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = " Update VK277 "
        cSQL = cSQL & " set best1 = 0 "
        cSQL = cSQL & " , LUG = 0 "
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = " Update VK277 inner join artikel on vk277.artnr = artikel.artnr"
        cSQL = cSQL & " set vk277.bezeich = artikel.bezeich "
        cSQL = cSQL & " , vk277.KVKPR1 = artikel.KVKPR1 "
        cSQL = cSQL & " , vk277.VKPR = artikel.VKPR "
        cSQL = cSQL & " , vk277.PGN = artikel.PGN "
        cSQL = cSQL & " , vk277.EKPR = artikel.EKPR "
        cSQL = cSQL & " , vk277.LINR = artikel.LINR "
        cSQL = cSQL & " , vk277.LPZ = artikel.LPZ "
        cSQL = cSQL & " , vk277.MWST = artikel.MWST "
        cSQL = cSQL & " , vk277.best1 = artikel.bestand "
        cSQL = cSQL & " , vk277.farbnr = val(artikel.awm) "
        dabalokal.Execute cSQL, dbFailOnError
        
        If cAWM <> "" Then
            cSQL = "Delete from VK277 where farbnr <> " & Val(cAWM)
            dabalokal.Execute cSQL, dbFailOnError
        End If
        
        cSQL = "Update VK277 inner join ARTLIEF on VK277.Linr = ARTLIEF.Linr and VK277.ARTNR = ARTLIEF.ARTNR "
        cSQL = cSQL & " set VK277.LIBESNR = ARTLIEF.LIBESNR "
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = "Update VK277 inner join LISRT on VK277.Linr = LISRT.Linr "
        cSQL = cSQL & " set VK277.LIEFBEZ = LISRT.LIEFBEZ "
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = "Update VK277 inner join ALLARTLU on VK277.artnr = ALLARTLU.artnummer "
        cSQL = cSQL & " set VK277.LUG = ALLARTLU.LUG "
        dabalokal.Execute cSQL, dbFailOnError
        
        loeschNEW "KOPF277", dabalokal
        CreateTable "KOPF277", dabalokal
        
        
        
        cSQL = "Insert into KOPF277 (Von,Bis,Bednu,Bedname) values ('" & Text1(0).Text & "','" & Text1(1).Text & "'," & Val(Text1(6).Text) & ",'" & ermBEDbez(Val(Text1(6).Text)) & "')"
        dabalokal.Execute cSQL, dbFailOnError
        
        
        cKassen = ""
        Set rsrs = dabalokal.OpenRecordset("wKass")
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!wkasse) Then
                    If cKassen = "" Then
                        cKassen = rsrs!wkasse
                    Else
                        cKassen = cKassen & ", " & rsrs!wkasse
                    End If
                End If
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        
        'pgnwahl
        cPGNWahl = ""
        If Text1(10).Text <> "alle" Or IsNumeric(Text1(10).Text) Then
            If Text1(10).Text = "" Then
                If List3.ListCount = 0 Then
                    'leer
                    cPGNWahl = "alle Produktgruppen"
                Else
                    For i = 0 To List3.ListCount - 1
                        cPGNWahl = cPGNWahl & List3.list(i) & vbCrLf
                    Next i
                End If
            Else
                If Trim$(Text1(10).Text) <> "" Then
                    cPGNWahl = Trim$(Text1(10).Text)
                End If
            End If
        Else
            'leer
             cPGNWahl = "alle Produktgruppen"
        End If
        'pgnwahl ende
        
        BringFarbeInsSpiel "vk277", dabalokal
        
        cSQL = "UPDATE KOPF277 set wKassen = '" & cKassen & "'"
        cSQL = cSQL & " , PGNwahl =  '" & cPGNWahl & "'"
        dabalokal.Execute cSQL, dbFailOnError
        
        loeschNEW "wKass", dabalokal
        
        If Not Datendrin("vk277", dabalokal) Then
            anzeigeNew "rot", "Es wurden keine Daten ermittelt.", Label6
            Exit Sub
        Else
            reportbildschirmApp "dWKL27", "aWKL277"
        End If
    End If
    
    If lAnzeigeArt = 9 Then
    
        CreateTableT2 "VK279A", dabalokal

        cSQL = "Insert into vk279a Select "
        cSQL = cSQL & " artnr "
        cSQL = cSQL & ", Bezeich "
        cSQL = cSQL & ", Menge as meng "
        cSQL = cSQL & ", Preis "
        cSQL = cSQL & ", Preis/Menge as Epreis "
        cSQL = cSQL & " from Kassjour  "
    
        cSQL = cSQL & " where Filiale = " & gcFilNr & " "
    
        If lVon > 0 Then
            cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lVon)) & " "
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            Else
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
            End If
        Else
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lBis)) & " "
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            End If
        End If
        
'        If cArtNr <> "" Then
'            cSQL = cSQL & "and ARTNR = " & cArtNr & " "
'        End If
        If cLinr <> "" Then
            cSQL = cSQL & " and artnr in (Select artnr from artlief where linr = " & cLinr & ")"
'            cSQL = cSQL & "and LINR = " & cLinr & " "
        End If
        If cKundnr <> "" Then
            cSQL = cSQL & "and KUNDNR = " & cKundnr & " "
        End If
        If cEAN <> "" Then
            cSQL = cSQL & "and EAN = '" & cEAN & "' "
        End If
        
        If cArtBez <> "" Then
            cArtBez = SwapStr(cArtBez, " ", "*")
            cSQL = cSQL & " and Bezeich like '*" & cArtBez & "*' "
        End If
        
        If cAgn <> "" Then
            cSQL = cSQL & "and agn = " & cAgn & " "
        End If
        If cBedNr <> "" Then
            cSQL = cSQL & "and BEDIENER = " & cBedNr & " "
        End If
        
        If cBonNr <> "" Then
            cSQL = cSQL & "and belegnr = " & cBonNr & " "
        End If
        
        If ckassnr <> "" Then
            cSQL = cSQL & "and kasnum = " & ckassnr & " "
        End If
        
        If cMopreis <> "" Then
            cSQL = cSQL & "and mopreis = " & cMopreis & " "
        End If
        
        cSQL = cSQL & BildePGnSQL
        
        cSQL = cSQL & sSQLArtnr
        
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = " Update VK279a "
        cSQL = cSQL & " set best1 = 0 "
        dabalokal.Execute cSQL, dbFailOnError
        
        loeschNEW "VK279", dabalokal
        CreateTableT2 "VK279", dabalokal
        
        cSQL = " Insert into VK279 select artnr,EPreis,preis,Bezeich "
        cSQL = cSQL & " , sum(Meng) as Menge from VK279a group by artnr, EPreis,preis,Bezeich "
        dabalokal.Execute cSQL, dbFailOnError
        
        
        loeschNEW "wKass", dabalokal
        
        cSQL = " Select distinct(kasnum) as wKasse "
        cSQL = cSQL & " into wKass from Kassjour  "
        cSQL = cSQL & " where Filiale = " & gcFilNr & " "
        If lVon > 0 Then
            cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lVon)) & " "
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            Else
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
            End If
        Else
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lBis)) & " "
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            End If
            
        End If
'        If cArtNr <> "" Then
'            cSQL = cSQL & "and ARTNR = " & cArtNr & " "
'        End If
        If cLinr <> "" Then
            cSQL = cSQL & " and artnr in (Select artnr from artlief where linr = " & cLinr & ")"
'            cSQL = cSQL & "and LINR = " & cLinr & " "
        End If
        If cKundnr <> "" Then
            cSQL = cSQL & "and KUNDNR = " & cKundnr & " "
        End If
        If cEAN <> "" Then
            cSQL = cSQL & "and EAN = '" & cEAN & "' "
        End If
        
        
        If cArtBez <> "" Then
            cArtBez = SwapStr(cArtBez, " ", "*")
            cSQL = cSQL & " and Bezeich like '*" & cArtBez & "*' "
        End If
        
        
        If cAgn <> "" Then
            cSQL = cSQL & "and agn = " & cAgn & " "
        End If
        If cBedNr <> "" Then
            cSQL = cSQL & "and BEDIENER = " & cBedNr & " "
        End If
        
        If cBonNr <> "" Then
            cSQL = cSQL & "and belegnr = " & cBonNr & " "
        End If
        
        If ckassnr <> "" Then
            cSQL = cSQL & "and kasnum = " & ckassnr & " "
        End If
        
        If cMopreis <> "" Then
            cSQL = cSQL & "and mopreis = " & cMopreis & " "
        End If
        
        cSQL = cSQL & BildePGnSQL
        cSQL = cSQL & sSQLArtnr
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = " Update VK279 "
        cSQL = cSQL & " set best1 = 0 "
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = " Update VK279 inner join artikel on VK279.artnr = artikel.artnr "
        cSQL = cSQL & " set VK279.best1 = artikel.bestand "
        dabalokal.Execute cSQL, dbFailOnError
        
        loeschNEW "KOPF277", dabalokal
        CreateTable "KOPF277", dabalokal
        
        cSQL = "Insert into KOPF277 (Von,Bis,Bednu,Bedname) values ('" & Text1(0).Text & "','" & Text1(1).Text & "'," & Val(Text1(6).Text) & ",'" & ermBEDbez(Val(Text1(6).Text)) & "')"
        dabalokal.Execute cSQL, dbFailOnError
        
        
        
        
        cKassen = ""
        Set rsrs = dabalokal.OpenRecordset("wKass")
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!wkasse) Then
                    If cKassen = "" Then
                        cKassen = rsrs!wkasse
                    Else
                        cKassen = cKassen & ", " & rsrs!wkasse
                    End If
                End If
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        
        'pgnwahl
        cPGNWahl = ""
        If Text1(10).Text <> "alle" Or IsNumeric(Text1(10).Text) Then
            If Text1(10).Text = "" Then
                If List3.ListCount = 0 Then
                    'leer
                    cPGNWahl = "alle Produktgruppen"
                Else
                    For i = 0 To List3.ListCount - 1
                        cPGNWahl = cPGNWahl & List3.list(i) & vbCrLf
                    Next i
                End If
            Else
                If Trim$(Text1(10).Text) <> "" Then
                    cPGNWahl = Trim$(Text1(10).Text)
                End If
            End If
        Else
            'leer
             cPGNWahl = "alle Produktgruppen"
        End If
        'pgnwahl ende
        
        cSQL = "UPDATE KOPF277 set wKassen = '" & cKassen & "'"
        cSQL = cSQL & " , PGNwahl =  '" & cPGNWahl & "'"
        dabalokal.Execute cSQL, dbFailOnError
        
        loeschNEW "wKass", dabalokal
        
        If Not Datendrin("VK279", dabalokal) Then
            anzeigeNew "rot", "Es wurden keine Daten ermittelt.", Label6
            Exit Sub
        Else
            reportbildschirmApp "", "aWKL279"
        End If
    End If
    
    
    If lAnzeigeArt = 11 Then
    
        CreateTable "VK277", dabalokal

        cSQL = "Insert into vk277 Select "
        cSQL = cSQL & " artnr "
        cSQL = cSQL & ", sum(MENGE) as SumMenge "
        cSQL = cSQL & ", sum(Preis) as SumPreis "
        cSQL = cSQL & " from Kassjour  "
        cSQL = cSQL & " where Filiale = " & gcFilNr & " "
    
        If lVon > 0 Then
            cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lVon)) & " "
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            Else
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
            End If
        Else
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lBis)) & " "
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            End If
        End If
        
        If cLinr <> "" Then
            cSQL = cSQL & " and artnr in (Select artnr from artlief where linr = " & cLinr & ")"
'            cSQL = cSQL & "and LINR = " & cLinr & " "
        End If
        If cKundnr <> "" Then
            cSQL = cSQL & "and KUNDNR = " & cKundnr & " "
        End If
        If cEAN <> "" Then
            cSQL = cSQL & "and EAN = '" & cEAN & "' "
        End If
        
        
        If cArtBez <> "" Then
            cArtBez = SwapStr(cArtBez, " ", "*")
            cSQL = cSQL & " and Bezeich like '*" & cArtBez & "*' "
        End If
        
        
        If cAgn <> "" Then
            cSQL = cSQL & "and agn = " & cAgn & " "
        End If
        If cBedNr <> "" Then
            cSQL = cSQL & "and BEDIENER = " & cBedNr & " "
        End If
        
        
        If cBonNr <> "" Then
            cSQL = cSQL & "and belegnr = " & cBonNr & " "
        End If
        
        If ckassnr <> "" Then
            cSQL = cSQL & "and kasnum = " & ckassnr & " "
        End If
        
        If cMopreis <> "" Then
            cSQL = cSQL & "and mopreis = " & cMopreis & " "
        End If
        
        cSQL = cSQL & BildePGnSQL
        
        cSQL = cSQL & sSQLArtnr
        
        cSQL = cSQL & " group by  "
        cSQL = cSQL & " artnr "
        dabalokal.Execute cSQL, dbFailOnError
        
        
            
        cSQL = " Update VK277 "
        cSQL = cSQL & " set best1 = 0 "
        cSQL = cSQL & " , LUG = 0 "
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = "Delete from VK277 "
        cSQL = cSQL & " where artnr in (Select artnr from Warengru) "
        dabalokal.Execute cSQL, dbFailOnError
        
        
        loeschNEW "wKass", dabalokal
        
        cSQL = " Select distinct(kasnum) as wKasse "
        cSQL = cSQL & " into wKass from Kassjour  "
        cSQL = cSQL & " where Filiale = " & gcFilNr & " "
        If lVon > 0 Then
            cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lVon)) & " "
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            Else
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
            End If
        Else
            If lBis > 0 Then
                cSQL = cSQL & "and ADATE >= " & Trim$(Str$(lBis)) & " "
                cSQL = cSQL & "and ADATE <= " & Trim$(Str$(lBis)) & " "
            End If
        End If

        If cLinr <> "" Then
            cSQL = cSQL & " and artnr in (Select artnr from artlief where linr = " & cLinr & ")"
        End If
        If cKundnr <> "" Then
            cSQL = cSQL & "and KUNDNR = " & cKundnr & " "
        End If
        If cEAN <> "" Then
            cSQL = cSQL & "and EAN = '" & cEAN & "' "
        End If
        
        If cArtBez <> "" Then
            cArtBez = SwapStr(cArtBez, " ", "*")
            cSQL = cSQL & " and Bezeich like '*" & cArtBez & "*' "
        End If
        
        If cAgn <> "" Then
            cSQL = cSQL & "and agn = " & cAgn & " "
        End If
        If cBedNr <> "" Then
            cSQL = cSQL & "and BEDIENER = " & cBedNr & " "
        End If
        
        
        If cBonNr <> "" Then
            cSQL = cSQL & "and belegnr = " & cBonNr & " "
        End If
        
        If ckassnr <> "" Then
            cSQL = cSQL & "and kasnum = " & ckassnr & " "
        End If
        
        If cMopreis <> "" Then
            cSQL = cSQL & "and mopreis = " & cMopreis & " "
        End If
        
        cSQL = cSQL & BildePGnSQL
        
        cSQL = cSQL & sSQLArtnr
        
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = " Update VK277 "
        cSQL = cSQL & " set best1 = 0 "
        cSQL = cSQL & " , LUG = 0 "
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = " Update VK277 inner join artikel on vk277.artnr = artikel.artnr"
        cSQL = cSQL & " set vk277.bezeich = artikel.bezeich "
        cSQL = cSQL & " , vk277.KVKPR1 = artikel.KVKPR1 "
        cSQL = cSQL & " , vk277.VKPR = artikel.VKPR "
        cSQL = cSQL & " , vk277.PGN = artikel.PGN "
        cSQL = cSQL & " , vk277.EKPR = artikel.EKPR "
        cSQL = cSQL & " , vk277.LINR = artikel.LINR "
        cSQL = cSQL & " , vk277.LPZ = artikel.LPZ "
        cSQL = cSQL & " , vk277.MWST = artikel.MWST "
        cSQL = cSQL & " , vk277.best1 = artikel.bestand "
        cSQL = cSQL & " , vk277.farbnr = val(artikel.awm) "
        cSQL = cSQL & " , vk277.ean = artikel.ean "
        dabalokal.Execute cSQL, dbFailOnError
        
        
        cSQL = "Delete from VK277 where best1 > 0 "
        dabalokal.Execute cSQL, dbFailOnError
        
        
        cSQL = "Update VK277 inner join ARTLIEF on VK277.Linr = ARTLIEF.Linr and VK277.ARTNR = ARTLIEF.ARTNR "
        cSQL = cSQL & " set VK277.LIBESNR = ARTLIEF.LIBESNR "
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = "Update VK277 inner join LISRT on VK277.Linr = LISRT.Linr "
        cSQL = cSQL & " set VK277.LIEFBEZ = LISRT.LIEFBEZ "
        dabalokal.Execute cSQL, dbFailOnError
        
        cSQL = "Update VK277 inner join ALLARTLU on VK277.artnr = ALLARTLU.artnummer "
        cSQL = cSQL & " set VK277.LUG = ALLARTLU.LUG "
        dabalokal.Execute cSQL, dbFailOnError
        
        loeschNEW "KOPF277", dabalokal
        CreateTable "KOPF277", dabalokal
        
        
        
        cSQL = "Insert into KOPF277 (Von,Bis,Bednu,Bedname) values ('" & Text1(0).Text & "','" & Text1(1).Text & "'," & Val(Text1(6).Text) & ",'" & ermBEDbez(Val(Text1(6).Text)) & "')"
        dabalokal.Execute cSQL, dbFailOnError
        
        
        cKassen = ""
        Set rsrs = dabalokal.OpenRecordset("wKass")
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!wkasse) Then
                    If cKassen = "" Then
                        cKassen = rsrs!wkasse
                    Else
                        cKassen = cKassen & ", " & rsrs!wkasse
                    End If
                End If
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        
        'pgnwahl
        cPGNWahl = ""
        If Text1(10).Text <> "alle" Or IsNumeric(Text1(10).Text) Then
            If Text1(10).Text = "" Then
                If List3.ListCount = 0 Then
                    'leer
                    cPGNWahl = "alle Produktgruppen"
                Else
                    For i = 0 To List3.ListCount - 1
                        cPGNWahl = cPGNWahl & List3.list(i) & vbCrLf
                    Next i
                End If
            Else
                If Trim$(Text1(10).Text) <> "" Then
                    cPGNWahl = Trim$(Text1(10).Text)
                End If
            End If
        Else
            'leer
             cPGNWahl = "alle Produktgruppen"
        End If
        'pgnwahl ende
        
        BringFarbeInsSpiel "vk277", dabalokal
        
        cSQL = "UPDATE KOPF277 set wKassen = '" & cKassen & "'"
        cSQL = cSQL & " , PGNwahl =  '" & cPGNWahl & "'"
        dabalokal.Execute cSQL, dbFailOnError
        
        loeschNEW "wKass", dabalokal
        
        If Not Datendrin("vk277", dabalokal) Then
            anzeigeNew "rot", "Es wurden keine Daten ermittelt.", Label6
            Exit Sub
        Else
'            reportbildschirmApp "dWKL27", "aWKL277"
            
            reportbildschirmApp "dWKL27", "aWKL277a"
        End If
    End If
    
    anzeigeNew "normal", "", Label6
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseDatenWK25d"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten. Fehlerstufe = " & Trim$(Str$(iFehler))

    Fehlermeldung1
End Sub
Private Function insert_KiloJour(sQuellpfad As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    insert_KiloJour = False
    
    'dBase Import
    
    loeschNEW "P613011", gdBase
    
    If FileExists(sQuellpfad & "\P613011.DBF") = False Then
        Exit Function
    End If
    
    sSQL = "Select * into P613011 from P613011 IN '" & sQuellpfad & "' 'dBase IV;'"
    gdBase.Execute sSQL, dbFailOnError
    
    'löschen der Datei
    Kill sQuellpfad & "\P613011.DBF"
    
    'Insert
    
    sSQL = "Insert into KILOJOUR Select Feld1 as Artnr, Feld2 as BEZEICH"
    sSQL = sSQL & " ,Feld3 as ADATE, Feld4 as GEWICHTKG from P613011 "
    gdBase.Execute sSQL, dbFailOnError
    
    'Bestände/Gewicht in Kg von der Tabelle "Kiloart" abziehen runter
    
    loeschNEW "KILOTMP", gdBase
    CreateTableT2 "KILOTMP", gdBase
    
    sSQL = "Insert into KILOTMP Select Feld1 as Artnr, Feld2 as BEZEICH "
    sSQL = sSQL & " ,Feld3 as ADATE, Feld4 as GEWICHTKG from P613011 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KILOART inner join KILOTMP on KILOART.Artnr = KILOTMP.Artnr  "
    sSQL = sSQL & " set KILOART.GEWICHTKG = KILOART.GEWICHTKG - KILOTMP.GEWICHTKG "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "KILOTMP", gdBase
    loeschNEW "P613011", gdBase
    
    
    
    insert_KiloJour = True

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "insert_KiloJour"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub Zeige_VKPROT_GEWICHT(lVon As Long, lBis As Long)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim bAnd As String
    
    bAnd = False
    
    loeschNEW "KILOVK", gdBase
    CreateTableT2 "KILOVK", gdBase

    cSQL = "Insert into KILOVK Select "
    cSQL = cSQL & " K.artnr "
    cSQL = cSQL & ", K.BEZEICH "
    cSQL = cSQL & ", K.GEWICHTKG "
    cSQL = cSQL & ", K.adate "
    
    
'    cSQL = cSQL & ", A.LINR "
'    cSQL = cSQL & ", A.LPZ "
'    cSQL = cSQL & ", A.AGN "
'    cSQL = cSQL & ", A.MWST "
    
    cSQL = cSQL & ", 0 as GEWICHTKGIST "
    cSQL = cSQL & " from KILOJOUR K "  ',ARTIKEL A "
    'cSQL = cSQL & " where K.artnr = A.artnr "
    
    If lVon > 0 Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & "  ADATE >= " & Trim$(Str$(lVon)) & " "
        If lBis > 0 Then
            cSQL = cSQL & " and ADATE <= " & Trim$(Str$(lBis)) & " "
        Else
            cSQL = cSQL & " and ADATE <= " & Trim$(Str$(CLng(DateValue(Now)))) & " "
        End If
        bAnd = True
    Else
        If lBis > 0 Then
            If bAnd Then
                cSQL = cSQL & " and "
            Else
                cSQL = cSQL & " where "
            End If
            cSQL = cSQL & "  ADATE >= " & Trim$(Str$(lBis)) & " "
            cSQL = cSQL & " and ADATE <= " & Trim$(Str$(lBis)) & " "
            bAnd = True
        End If
    End If
    
'    If cLinr <> "" Then
'        If bAnd Then
'            cSQL = cSQL & " and "
'        Else
'            cSQL = cSQL & " where "
'        End If
'        cSQL = cSQL & " LINR = " & cLinr & " "
'        bAnd = True
'    End If
'
'    If cAgn <> "" Then
'        If bAnd Then
'            cSQL = cSQL & " and "
'        Else
'            cSQL = cSQL & " where "
'        End If
'        cSQL = cSQL & " agn = " & cAgn & " "
'        bAnd = True
'    End If
    
'    cSQL = cSQL & BildePGnSQL
'
'    cSQL = cSQL & sSQLArtnr

    cSQL = cSQL & " and artnr > 0 "
    
    cSQL = cSQL & " order by ADATE, K.ARTNR "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update KILOVK inner join KILOART on KILOVK.Artnr = KILOART.Artnr"
    cSQL = cSQL & " set KILOVK.GewichtKGIST  = KILOART.GewichtKG"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update KILOVK set von = '" & Text1(0).Text & "'"
    cSQL = cSQL & " ,bis = '" & Text1(1).Text & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    
    Screen.MousePointer = 0
        
    If Not Datendrin("KILOVK", gdBase) Then
        anzeigeNew "rot", "Es wurden keine Daten ermittelt.", Label6
        Exit Sub
    Else
        reportbildschirm "dWKL27", "aWKL274f"
    End If

    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zeige_VKPROT_GEWICHT"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function ermBizerbaPfad() As String
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Dim sSQL As String

    ermBizerbaPfad = ""
        
    sSQL = "Select BIZPFAD from WKEINSTE"
    Set rsrs = gdApp.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!BIZPFAD) Then
            ermBizerbaPfad = rsrs!BIZPFAD
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermBizerbaPfad"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Function BildePGnSQL() As String
On Error GoTo LOKAL_ERROR

    Dim i As Integer

    BildePGnSQL = ""
    
    If Text1(10).Text <> "alle" Or IsNumeric(Text1(10).Text) Then
        If Text1(10).Text = "" Then
            If List3.ListCount = 0 Then
                'leer
                BildePGnSQL = ""
            Else
                If LoesePGNInArtnr(Mid(List3.list(0), 1, InStr(1, List3.list(0), " ")), False, dabalokal) = True Then
                    BildePGnSQL = " and artnr in (select artnr from my" & srechnertab & ") "
                Else
                    BildePGnSQL = " and artnr = 11111111 "
                End If
            
                For i = 1 To List3.ListCount - 1
                    If LoesePGNInArtnr(Mid(List3.list(i), 1, InStr(1, List3.list(i), " ")), True, dabalokal) = True Then
                        BildePGnSQL = " and artnr in (select artnr from my" & srechnertab & ") "
                    Else
                        BildePGnSQL = " and artnr = 11111111 "
                    End If
                Next i
            End If
        Else
            If Trim$(Text1(10).Text) <> "" Then
                If LoesePGNInArtnr(Trim$(Text1(10).Text), False, dabalokal) = True Then
                    BildePGnSQL = " and artnr in (select artnr from my" & srechnertab & ") "
                Else
                    BildePGnSQL = " and artnr = 11111111 "
                End If
            End If
        End If
    Else
        'leer
         BildePGnSQL = ""
    End If
     
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BildePGnSQL"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten. "

    Fehlermeldung1


End Function
Private Sub Zahldetails(sTab As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String

    loeschNEW "GEMZD", dabalokal
    CreateTable "GEMZD", dabalokal
    
    
    loeschNEW sTab & "A", dabalokal
    
    
    
    cSQL = " Select distinct(Belegnr) as bon "
    cSQL = cSQL & ", adate "
    cSQL = cSQL & ", azeit "
    cSQL = cSQL & " into  " & sTab & "A from " & sTab
    dabalokal.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into GEMZD Select  "
    cSQL = cSQL & " GEMZ.geldwert "
    cSQL = cSQL & ", GEMZ.belegnr  "
    cSQL = cSQL & ", GEMZ.kasnum "
    cSQL = cSQL & ", GEMZ.kk_art  "
    cSQL = cSQL & ", GEMZ.adate "
    cSQL = cSQL & ", GEMZ.azeit "
    cSQL = cSQL & " from GEMZ inner join " & sTab & "A on " & sTab & "A.bon = GEMZ.BELEGNR and " & sTab & "A.adate = GEMZ.adate and " & sTab & "A.azeit = GEMZ.azeit  "
    dabalokal.Execute cSQL, dbFailOnError
    
    If Datendrin("GEMZD", dabalokal) Then
        reportbildschirmApp "dWKL27", "aWKLGZ"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zahldetails"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR

    
    Dim iFeld As Integer
    Dim iCount As Integer
    Dim ctmp As String
    Dim cZiel As String
    Dim cZeichen As String
    
    Set WshShell = CreateObject("WScript.Shell")
    'WshShell.SendKeys "+{Tab}", True
    
    iFeld = Val(Label0.Caption)
    
    Select Case Index
        Case 0 To 9     'Ziffern
            Text1(iFeld).Text = Text1(iFeld).Text & Command0(Index).Caption
        Case Is = 10
            WshShell.SendKeys "+{Tab}", True
'            SendKeys "+{TAB}"

        Case Is = 11
            WshShell.SendKeys "{Tab}", True
'            SendKeys "{TAB}"

        Case Is = 12    'C
            Text1(iFeld).Text = ""
        Case 13     'punkt
            If iFeld = 1 Or iFeld = 0 Then
                Text1(iFeld).Text = Text1(iFeld).Text & Command0(Index).Caption
            End If
        
        Case 14
            Text1_KeyUp 10, vbKeyF2, 0
        
        Case 15
        
            If Trim(Text1(2).Text) <> "" Then
                List1.AddItem Text1(2).Text
            End If
            Text1(2).Text = ""
            Text1(11).Text = ""
            Text1(2).SetFocus
        Case 16
            List1.Clear
        Case Is = 20        ' Kalender
            Text1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            Text1(1).SetFocus
            
        Case Is = 21        ' Kalender
            Text1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
        End Select
        
        Text1(iFeld).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    

    
    Screen.MousePointer = 11
    
    Dim iRet As Integer

    Select Case Index
        Case Is = 0     'Zeige
            iRet = fnPruefeEingabeWK25d()
            Select Case iRet
                Case Is = 0     'alles okay!
                    Command1(0).Enabled = False
                    LeseDatenWK25d
                    Command1(0).Enabled = True
                Case Is = 1
                    anzeigeNew "rot", "Bitte ein Suchkritrium eingeben!", Label6
            End Select
        Case 1
            Text1_KeyUp 4, vbKeyF2, 0
        Case Is = 2     'schließen
            voreinstellungspeichern
            Unload frmWK25d
        Case 3
            Text1_KeyUp 5, vbKeyF2, 0
        Case 4
            Text1_KeyUp 6, vbKeyF2, 0
        Case 5
            Text1_KeyUp 9, vbKeyF2, 0
        Case 11
            Screen.MousePointer = 0
            
            gsBackcolor = Label1(16).BackColor
            gsForecolor = Label1(16).ForeColor
            gsArtikelFarbe = Label1(16).Tag
            
            frmWKL49.Show 1
            
            Label1(16).BackColor = gsBackcolor
            Label1(16).ForeColor = gsForecolor
            Label1(16).Tag = gsArtikelFarbe
            
            If gsArtikelFarbe <> "" Then
                Label1(16).Caption = "Farbauswahl"
            Else
                Label1(16).Caption = "alle Farben"
            End If
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    WK25dPositionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    iZaehler = 0
    Timer1.Enabled = False
    Timer2.Enabled = False
    
    Timer3.Enabled = False
    Timer4.Enabled = False
    
    Frame4.Visible = True

    Label0.Caption = 1
    
    
   
    If NewTableSuchenDBKombi("VKPE", gdApp) Then
    
        If SpalteInTabellegefundenNEW("VKPE", "bo5", gdApp) = False Then
            SpalteAnfuegenNEW "VKPE", "bo5", "BIT", gdApp
        End If
        
        If SpalteInTabellegefundenNEW("VKPE", "bo6", gdApp) = False Then
            SpalteAnfuegenNEW "VKPE", "bo6", "BIT", gdApp
        End If
        voreinstellungladen
    End If
    
    
'    If gsWAAGE <> "keine Waage" Then
'        Option1(10).Visible = True
'        Option1(10).Value = True
'    End If
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten. " 'Fehlerstufe = " & Trim$(Str$(iFehler))

    Fehlermeldung1
   
End Sub
Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

If Index = 16 Then
    Label1(Index).Caption = "alle Farben"
    Label1(Index).Tag = ""
    Label1(Index).BackColor = Label1(11).BackColor
    Label1(Index).ForeColor = Label1(11).ForeColor
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."""
    
    Fehlermeldung1
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String

    Select Case Index
        Case 0 To 1   'dat
            cValid = "1234567890." & Chr$(8)
        
        Case 2 To 10  'alles
            cValid = "1234567890" & Chr$(8)
            
        Case 11  'alles
            cValid = "1234567890" & Chr$(8)
            List1.Clear
            
        Case 12  'alles
            cValid = "1234567890" & Chr$(8)
        Case 13 'bezeich
        
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß"
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
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim lcount As Long
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            Case Is = 2
                ctmp = Text1(3).Text
                ctmp = Trim$(Str$(Val(ctmp)))
                If Text1(4).Text = "" Then
                    MsgBox "Bitte einen Lieferanten angeben!", vbInformation, "Winkiss Hinweis:"
                    Text1(3).SetFocus
                    Exit Sub
                End If
                gF2Prompt.cFeld = "ARTNR"
                gF2Prompt.cWert = ctmp
                
                If gF2Prompt.cFeld <> "" Then
                    Screen.MousePointer = 0
                    frmWK00a.Show 1
                    
                    If gF2Prompt.cWahl <> "" Then
                        ctmp = gF2Prompt.cWahl
                        ctmp = ctmp & String$(6 - Len(ctmp), "_")
                        Text1(Index).Text = ctmp
                    End If
                    Text1(Index).SetFocus
                End If
                
            Case Is = 4
                gF2Prompt.cFeld = "LINR"
                
                If gF2Prompt.cFeld <> "" Then
                    Screen.MousePointer = 0
                    frmWK00a.Show 1
                    
                    If gF2Prompt.cWahl <> "" Then
                        ctmp = gF2Prompt.cWahl
                        ctmp = ctmp & String$(6 - Len(ctmp), "_")
                        Text1(Index).Text = ctmp
                    End If
                    Text1(Index).SetFocus
                End If
            Case Is = 5
                gF2Prompt.cFeld = "KUN"
                
                If gF2Prompt.cFeld <> "" Then
                    Screen.MousePointer = 0
                    frmWK00a.Show 1
                    
                    If gF2Prompt.cWahl <> "" Then
                        ctmp = gF2Prompt.cWahl
                        ctmp = ctmp & String$(7 - Len(ctmp), "_")
                        Text1(Index).Text = ctmp
                    End If
                    Text1(Index).SetFocus
                End If
            Case Is = 6
                gF2Prompt.cFeld = "BED"
            
            
            
                If gF2Prompt.cFeld <> "" Then
                    Screen.MousePointer = 0
                    frmWK00a.Show 1
                    
                    If gF2Prompt.cWahl <> "" Then
                        ctmp = gF2Prompt.cWahl
                        ctmp = ctmp & String$(6 - Len(ctmp), "_")
                        Text1(Index).Text = ctmp
                    End If
                    Text1(Index).SetFocus
                End If
            Case Is = 9
                gF2Prompt.cFeld = "AGN"
                
                If gF2Prompt.cFeld <> "" Then
                    Screen.MousePointer = 0
                    frmWK00a.Show 1
                    
                    Text1(Index).Text = gF2Prompt.cWahl
                    Text1(Index).SetFocus
                End If
            Case Is = 10
                gF2Prompt.bMultiple = True
                gF2Prompt.cFeld = "PGN"
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    If gF2Prompt.cWahl <> "" Then
                        Text1(Index).Text = gF2Prompt.cWahl
                    End If
                End If
                
                List3.Visible = False
                List3.Clear
                For lcount = 0 To 100
                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                        List3.Visible = True
                        Text1(Index).Text = ""
                        
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List3.AddItem gF2Prompt.cArray(lcount)
                        End If
                    
                    Else
                        If gF2Prompt.cArray(lcount) <> "" Then
                           
                            List3.AddItem gF2Prompt.cArray(lcount)
                            Text1(Index).Text = Left(gF2Prompt.cArray(lcount), 2)
                        End If
                        
                    End If
                Next lcount
                
        End Select
        
        
    ElseIf KeyCode = vbKeyReturn Then
        Command1_Click 0
    
    End If
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Option1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    'Retouren kumuliert checkbox standard unsichtbar
    Check1(18).Visible = False
    Check1(18).Value = vbUnchecked
    
    Select Case Index
    
        Case Is = 2    'vormonat
        
            If Month(DateValue(Now)) = 1 Then
                Text1(0).Text = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
                Text1(1).Text = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Else
                Text1(0).Text = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                Select Case Month(DateValue(Now)) - 1
                    Case 1, 3, 5, 7, 8, 10, 12
                        Text1(1).Text = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                    
                    Case 2
                        If Year(DateValue(Now)) = 2016 Then
                            Text1(1).Text = Format("29." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        Else
                            Text1(1).Text = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        End If
                    
                    Case Else
                        Text1(1).Text = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                End Select
            End If
                
        Case Is = 5     'ak monat
            Text1(0).Text = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YYYY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
        
        Case Is = 6     'gestern
            Text1(0).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
            Text1(1).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
        
        Case Is = 7     'heute
            Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
        Case 12 'aktuelles Jahr
            Text1(0).Text = Format("01.01." & Year(DateValue(Now)), "DD.MM.YYYY")
            Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
        Case 14 'Vorjahr
        
            Text1(0).Text = Format("01.01." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Text1(1).Text = Format("31.12." & Year(Now) - 1, "DD.MM.YYYY")
        
        
            
        Case Is = 4 'Retoure
            If Option1(4).Value = True Then
                Check1(18).Visible = True
            Else
                Check1(18).Visible = False
            End If
        
        Case Is = 1, 0, 8
'            Option2(0).Value = True
'            Frame4.Visible = True
        Case Else
'            Frame4.Visible = False
'            Frame5.Visible = False
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten. " 'Fehlerstufe = " & Trim$(Str$(iFehler))

    Fehlermeldung1
   
End Sub
Private Sub Option2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    

    If Index = 3 Then
        Check1(0).Value = vbUnchecked
        Check1(1).Value = vbUnchecked
        Check1(2).Value = vbUnchecked
        Check1(3).Value = vbUnchecked
        Check1(4).Value = vbUnchecked
        Check1(5).Value = vbUnchecked
        Check1(6).Value = vbUnchecked
        
        Check1(7).Value = vbUnchecked
        Check1(8).Value = vbUnchecked
        Check1(9).Value = vbUnchecked
        Check1(10).Value = vbUnchecked
        Check1(11).Value = vbUnchecked
        Frame5.Visible = True
    Else
        Frame5.Visible = False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option2_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten." ' Fehlerstufe = " & Trim$(Str$(iFehler))

    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    
    Label0.Caption = Index
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten. " 'Fehlerstufe = " & Trim$(Str$(iFehler))

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
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten. " 'Fehlerstufe = " & Trim$(Str$(iFehler))

    Fehlermeldung1
End Sub
Private Sub voreinstellungladen()
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Set rsrs = gdApp.OpenRecordset("VKPE")
    
    If Not rsrs.EOF Then
        
        Option1(7).Value = rsrs!bo1
        If Option1(7).Value Then Option1_Click 7
            
        Option1(6).Value = rsrs!bo2
        If Option1(6).Value Then Option1_Click 6
            
        Option1(5).Value = rsrs!bo3
        If Option1(5).Value Then Option1_Click 5
            
        Option1(2).Value = rsrs!bo4
        If Option1(2).Value Then Option1_Click 2
        
        Option1(12).Value = rsrs!bo5
        If Option1(12).Value Then Option1_Click 12
            
        Option1(14).Value = rsrs!bo6
        If Option1(14).Value Then Option1_Click 14
            
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command7_Click()
On Error GoTo LOKAL_ERROR
    
    Dim lDat As Long
    If IsDate(Text1(0).Text) = False Then
        Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Else
    
        If IsDate(Text1(0).Text) = True Then
            lDat = CLng(DateValue(Text1(0).Text))
        Else
            
        End If
        
        
        lDat = lDat + 1
        
        
        Text1(0).Text = Format(lDat, "DD.MM.YYYY")
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command8_Click()
On Error GoTo LOKAL_ERROR

    Dim lDat As Long

    If IsDate(Text1(0).Text) = False Then
        Text1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
    
    
    Else
    
        If IsDate(Text1(0).Text) = True Then
            lDat = CLng(DateValue(Text1(0).Text))
        Else
            
        End If
        
        
        lDat = lDat - 1
        
        
        Text1(0).Text = Format(lDat, "DD.MM.YYYY")
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command2_Click()
On Error GoTo LOKAL_ERROR
    
    Dim lDat As Long
    If IsDate(Text1(1).Text) = False Then
        Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Else
    
        If IsDate(Text1(1).Text) = True Then
            lDat = CLng(DateValue(Text1(1).Text))
        Else
           
        End If
        
        
        lDat = lDat + 1
        
        
        Text1(1).Text = Format(lDat, "DD.MM.YYYY")
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command3_Click()
On Error GoTo LOKAL_ERROR

    Dim lDat As Long
    If IsDate(Text1(1).Text) = False Then
        Text1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Else
        If IsDate(Text1(1).Text) = True Then
            lDat = CLng(DateValue(Text1(1).Text))
        Else
            
        End If
        
        lDat = lDat - 1
        
        Text1(1).Text = Format(lDat, "DD.MM.YYYY")
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer1.Enabled = True
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_MouseDown"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer1.Enabled = False
    iZaehler = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer2.Enabled = True
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_MouseDown"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer2.Enabled = False
    iZaehler = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer3.Enabled = True
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_MouseDown"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer3.Enabled = False
    iZaehler = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer4.Enabled = True
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_MouseDown"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer4.Enabled = False
    iZaehler = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
'Private Sub Text2_GotFocus(Index As Integer)
'On Error GoTo LOKAL_ERROR
'
'    Text2(Index).BackColor = glSelBack1
'    Text2(Index).SelStart = Len(Text2(Index).Text)
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Text2_GotFocus"
'    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
'Private Sub Text2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'On Error GoTo LOKAL_ERROR
'
'Dim lCount As Long
'Dim cTmp As String
'
'
'    If KeyCode = vbKeyF2 Then
'        gF2Prompt.cFeld = ""
'        gF2Prompt.cWert = ""
'        gF2Prompt.cWert2 = ""
'        gF2Prompt.cWahl = ""
'        gF2Prompt.bMultiple = True
'
'        Select Case Index
'
'
'            Case 0
'                gF2Prompt.cFeld = "PGN"
'                If gF2Prompt.cFeld <> "" Then
'                    frmWK00a.Show 1
'                    If gF2Prompt.cWahl <> "" Then
'                        Text2(Index).Text = gF2Prompt.cWahl
'                    End If
'                End If
'
'                List3.Visible = False
'                List3.Clear
'                For lCount = 0 To 100
'                    If lCount > 0 And gF2Prompt.cArray(lCount) <> "" Then
'                        List3.Visible = True
'                        Text2(Index).Text = ""
'
'                        If gF2Prompt.cArray(lCount) <> "" Then
'                            List3.AddItem gF2Prompt.cArray(lCount)
'                        End If
'
'                    Else
'                        If gF2Prompt.cArray(lCount) <> "" Then
'
'                            List3.AddItem gF2Prompt.cArray(lCount)
'                            Text2(Index).Text = Left(gF2Prompt.cArray(lCount), 2)
'                        End If
'
'                    End If
'                Next lCount
'
'        End Select
'        Text2(Index).SetFocus
'    End If
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Text2_KeyUp"
'    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
'Private Sub Text2_LostFocus(Index As Integer)
'On Error GoTo LOKAL_ERROR
'
'    Text2(Index).BackColor = vbWhite
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "Text2_LostFocus"
'    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'
'End Sub


Private Sub Timer1_Timer()
    On Error GoTo LOKAL_ERROR
    
    If iZaehler > 10 Then
        Timer1.Interval = 50
    ElseIf iZaehler > 100 Then
        Timer1.Interval = 10
    Else
        Timer1.Interval = 200
    End If
    
    iZaehler = iZaehler + 1
    
    Command7_Click
    
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Timer2_Timer()
    On Error GoTo LOKAL_ERROR
    
    If iZaehler > 10 Then
        Timer2.Interval = 50
    ElseIf iZaehler > 100 Then
        Timer2.Interval = 10
    Else
        Timer2.Interval = 200
    End If
    
    iZaehler = iZaehler + 1
    Command8_Click
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer2_Timer"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Timer3_Timer()
    On Error GoTo LOKAL_ERROR
    
    If iZaehler > 10 Then
        Timer3.Interval = 50
    ElseIf iZaehler > 100 Then
        Timer3.Interval = 10
    Else
        Timer3.Interval = 200
    End If
    
    iZaehler = iZaehler + 1
    
    Command2_Click
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer3_Timer"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Timer4_Timer()
    On Error GoTo LOKAL_ERROR
    
    If iZaehler > 10 Then
        Timer4.Interval = 50
    ElseIf iZaehler > 100 Then
        Timer4.Interval = 10
    Else
        Timer4.Interval = 200
    End If
    
    iZaehler = iZaehler + 1
    Command3_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer4_Timer"
    Fehler.gsFehlertext = "Im Programmteil Verkaufsprotokoll ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
