VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmWKL10 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Artikeldaten bearbeiten"
   ClientHeight    =   8565
   ClientLeft      =   1935
   ClientTop       =   2190
   ClientWidth     =   11910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWKL10.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8565
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
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
      Height          =   1095
      Left            =   4200
      TabIndex        =   4
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
      Begin sevCommand3.Command Command4 
         Height          =   600
         Index           =   1
         Left            =   9960
         TabIndex        =   71
         Top             =   2400
         Width           =   1455
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   58
         Left            =   9240
         TabIndex        =   70
         Top             =   2400
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   57
         Left            =   8400
         TabIndex        =   69
         Top             =   2400
         Width           =   840
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   56
         Left            =   7680
         TabIndex        =   68
         Top             =   2400
         Width           =   720
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
         Caption         =   "_"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   55
         Left            =   6960
         TabIndex        =   67
         Top             =   2400
         Width           =   720
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
         Caption         =   ":"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   54
         Left            =   6240
         TabIndex        =   66
         Top             =   2400
         Width           =   720
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
         Caption         =   ";"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   53
         Left            =   5520
         TabIndex        =   65
         Top             =   2400
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   52
         Left            =   4800
         TabIndex        =   64
         Top             =   2400
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   51
         Left            =   4080
         TabIndex        =   63
         Top             =   2400
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   50
         Left            =   3360
         TabIndex        =   62
         Top             =   2400
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   49
         Left            =   2640
         TabIndex        =   61
         Top             =   2400
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   48
         Left            =   1920
         TabIndex        =   60
         Top             =   2400
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   47
         Left            =   1200
         TabIndex        =   59
         Top             =   2400
         Width           =   720
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
      Begin sevCommand3.Command Command4 
         Height          =   600
         Index           =   0
         Left            =   9480
         TabIndex        =   58
         Top             =   1800
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
         Caption         =   "LEEREN"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   46
         Left            =   8760
         TabIndex        =   57
         Top             =   1800
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   45
         Left            =   8040
         TabIndex        =   56
         Top             =   1800
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   44
         Left            =   7320
         TabIndex        =   55
         Top             =   1800
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   43
         Left            =   6600
         TabIndex        =   54
         Top             =   1800
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   42
         Left            =   5880
         TabIndex        =   53
         Top             =   1800
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   41
         Left            =   5160
         TabIndex        =   52
         Top             =   1800
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   40
         Left            =   4440
         TabIndex        =   51
         Top             =   1800
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   39
         Left            =   3720
         TabIndex        =   50
         Top             =   1800
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   38
         Left            =   3000
         TabIndex        =   49
         Top             =   1800
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   37
         Left            =   2280
         TabIndex        =   48
         Top             =   1800
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   35
         Left            =   840
         TabIndex        =   46
         Top             =   1800
         Width           =   720
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
      Begin sevCommand3.Command Command4 
         Height          =   600
         Index           =   4
         Left            =   9720
         TabIndex        =   45
         Top             =   1200
         Width           =   1455
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
         Caption         =   "A -> a"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   34
         Left            =   9000
         TabIndex        =   44
         Top             =   1200
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   33
         Left            =   8280
         TabIndex        =   43
         Top             =   1200
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   32
         Left            =   7560
         TabIndex        =   42
         Top             =   1200
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   31
         Left            =   6840
         TabIndex        =   41
         Top             =   1200
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   30
         Left            =   6120
         TabIndex        =   40
         Top             =   1200
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   29
         Left            =   5400
         TabIndex        =   39
         Top             =   1200
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   28
         Left            =   4680
         TabIndex        =   38
         Top             =   1200
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   27
         Left            =   3960
         TabIndex        =   37
         Top             =   1200
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   26
         Left            =   3240
         TabIndex        =   36
         Top             =   1200
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   25
         Left            =   2520
         TabIndex        =   35
         Top             =   1200
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   24
         Left            =   1800
         TabIndex        =   34
         Top             =   1200
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   23
         Left            =   1080
         TabIndex        =   33
         Top             =   1200
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   22
         Left            =   360
         TabIndex        =   32
         Top             =   1200
         Width           =   720
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
      Begin sevCommand3.Command Command4 
         Height          =   600
         Index           =   3
         Left            =   8040
         TabIndex        =   31
         Top             =   600
         Width           =   1455
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   21
         Left            =   7320
         TabIndex        =   30
         Top             =   600
         Width           =   720
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
         Caption         =   "ß"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   20
         Left            =   6600
         TabIndex        =   29
         Top             =   600
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   19
         Left            =   5880
         TabIndex        =   28
         Top             =   600
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   18
         Left            =   5160
         TabIndex        =   27
         Top             =   600
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   17
         Left            =   4440
         TabIndex        =   26
         Top             =   600
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   16
         Left            =   3720
         TabIndex        =   25
         Top             =   600
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   15
         Left            =   3000
         TabIndex        =   24
         Top             =   600
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   14
         Left            =   2280
         TabIndex        =   23
         Top             =   600
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   13
         Left            =   1560
         TabIndex        =   22
         Top             =   600
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   12
         Left            =   840
         TabIndex        =   21
         Top             =   600
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   11
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   720
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
      Begin sevCommand3.Command Command4 
         Height          =   600
         Index           =   2
         Left            =   8040
         TabIndex        =   19
         Top             =   0
         Width           =   1455
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   10
         Left            =   7320
         TabIndex        =   18
         Top             =   0
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   9
         Left            =   6600
         TabIndex        =   17
         Top             =   0
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   8
         Left            =   5880
         TabIndex        =   16
         Top             =   0
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   7
         Left            =   5160
         TabIndex        =   15
         Top             =   0
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   6
         Left            =   4440
         TabIndex        =   14
         Top             =   0
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   5
         Left            =   3720
         TabIndex        =   13
         Top             =   0
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   4
         Left            =   3000
         TabIndex        =   12
         Top             =   0
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   3
         Left            =   2280
         TabIndex        =   11
         Top             =   0
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   2
         Left            =   1560
         TabIndex        =   10
         Top             =   0
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   1
         Left            =   840
         TabIndex        =   9
         Top             =   0
         Width           =   720
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
         Caption         =   "´"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   720
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   36
         Left            =   1560
         TabIndex        =   47
         Top             =   1800
         Width           =   720
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
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "-1"
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
         Left            =   9600
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
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
         Left            =   9600
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Zielfeld:"
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
         Left            =   9600
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame0 
      BackColor       =   &H00FF80FF&
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
      Height          =   4095
      Left            =   840
      TabIndex        =   272
      Top             =   2280
      Width           =   10815
      Begin VB.CheckBox Check14 
         BackColor       =   &H00C0C000&
         Caption         =   "nur Shop-Artikel"
         Height          =   255
         Left            =   9840
         TabIndex        =   367
         Top             =   3240
         Width           =   1935
      End
      Begin VB.CheckBox Check13 
         BackColor       =   &H00C0C000&
         Caption         =   "nur EX"
         Height          =   255
         Left            =   9840
         TabIndex        =   365
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00C0C000&
         Caption         =   "nur neg. Bestände"
         Height          =   255
         Left            =   9840
         TabIndex        =   364
         Top             =   2940
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   47
         Left            =   3480
         TabIndex        =   361
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   46
         Left            =   1800
         MaxLength       =   13
         TabIndex        =   323
         Text            =   "1234567890123"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   45
         Left            =   120
         MaxLength       =   13
         TabIndex        =   322
         Text            =   "1234567890123"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   8040
         MultiSelect     =   2  'Erweitert
         TabIndex        =   310
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   44
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   321
         Text            =   "1234567890123"
         Top             =   1440
         Width           =   975
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   1
         Left            =   8880
         TabIndex        =   295
         ToolTipText     =   "Hier legen Sie einen neuen Artikel an"
         Top             =   3200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
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
         Caption         =   "Neu"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   0
         Left            =   9840
         TabIndex        =   294
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         Caption         =   "Suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   2
         Left            =   9840
         TabIndex        =   293
         Top             =   3600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   3
         Left            =   5400
         TabIndex        =   292
         Top             =   3600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         Caption         =   "Umverpackung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   1
         Left            =   120
         MaxLength       =   13
         TabIndex        =   311
         Text            =   "1234567890123"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   3
         Left            =   8040
         MaxLength       =   5
         TabIndex        =   317
         Text            =   "1234567890123"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   4
         Left            =   6960
         MaxLength       =   13
         TabIndex        =   313
         Text            =   "1234567890123"
         Top             =   240
         Width           =   1695
      End
      Begin sevCommand3.Command Command1 
         Height          =   355
         Index           =   6
         Left            =   1440
         TabIndex        =   291
         Top             =   840
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
         Height          =   355
         Index           =   7
         Left            =   8760
         TabIndex        =   290
         Top             =   840
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
         Height          =   355
         Index           =   11
         Left            =   8760
         TabIndex        =   289
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
      Begin sevCommand3.Command Command98 
         Height          =   360
         Left            =   9360
         TabIndex        =   288
         Top             =   240
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
         Picture         =   "frmWKL10.frx":0442
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   6240
         MultiSelect     =   2  'Erweitert
         TabIndex        =   287
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin sevCommand3.Command Command1 
         Height          =   355
         Index           =   55
         Left            =   7080
         TabIndex        =   286
         Top             =   840
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   41
         Left            =   120
         MaxLength       =   13
         TabIndex        =   318
         Text            =   "1234567890123"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   42
         Left            =   1800
         MaxLength       =   13
         TabIndex        =   319
         Text            =   "1234567890123"
         Top             =   1440
         Width           =   1575
      End
      Begin sevCommand3.Command Command1 
         Height          =   355
         Index           =   10
         Left            =   4560
         TabIndex        =   285
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
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   0
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   320
         Text            =   "1234567890123"
         Top             =   1440
         Width           =   975
      End
      Begin sevCommand3.Command Command1 
         Height          =   355
         Index           =   16
         Left            =   5760
         TabIndex        =   284
         Top             =   840
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   2
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   315
         Text            =   "1234567890123"
         ToolTipText     =   "Haben Sie schon Ihre Markenkürzel gepflegt?. Bestimmen Sie selbst welche Marken in der F2 - Suchmaske erscheinen sollen."
         Top             =   840
         Width           =   2175
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   17
         Left            =   6960
         TabIndex        =   283
         ToolTipText     =   "Hier sehen Sie ein Protokoll der gelöschten Artikel."
         Top             =   3600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
         Caption         =   "gelöschte Artikel"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   19
         Left            =   6960
         TabIndex        =   282
         ToolTipText     =   "Hier sehen Sie ein Protokoll der Preisänderungen, die Sie vorgenommen haben."
         Top             =   3200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
         Caption         =   "Preisänderungen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0C000&
         Caption         =   "nur Penner"
         Height          =   255
         Left            =   9840
         TabIndex        =   281
         ToolTipText     =   "seit einem Jahr nicht mehr verkauft"
         Top             =   2340
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   7
         Left            =   120
         MaxLength       =   6
         TabIndex        =   314
         Text            =   "1234567890123"
         ToolTipText     =   "Sind Ihre Lieferantenkürzel ordentlich gepflegt, so können Sie mit dem Kürzel arbeiten. Geben Sie z.B. Joop ein!"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   35
         Left            =   6240
         MaxLength       =   3
         TabIndex        =   316
         Text            =   "1234567890123"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Index           =   36
         Left            =   2280
         MaxLength       =   35
         TabIndex        =   312
         Text            =   "1234567890123"
         ToolTipText     =   "Kein Sternchen(*) mehr. Trennen Sie die Textblöcke mit einem Leerzeichen z.B.: Cool edt 100"
         Top             =   240
         Width           =   3375
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00C0C000&
         Caption         =   "nur geführte"
         Height          =   255
         Left            =   9840
         TabIndex        =   280
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00C0C000&
         Caption         =   "nur mit Bestand"
         Height          =   255
         Left            =   9840
         TabIndex        =   279
         Top             =   1750
         Width           =   1935
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00C0C000&
         Caption         =   "EX ausblenden"
         Height          =   255
         Left            =   9840
         TabIndex        =   278
         Top             =   1140
         Value           =   1  'Aktiviert
         Width           =   1935
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   24
         Left            =   8880
         TabIndex        =   277
         Top             =   3600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
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
         Caption         =   "Export"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00C0C000&
         Caption         =   "nur zZ bestellte A."
         Height          =   255
         Left            =   9840
         TabIndex        =   276
         Top             =   2040
         Width           =   1935
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   25
         Left            =   3720
         TabIndex        =   275
         Top             =   3200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         Caption         =   "Strichcodes"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00C0C000&
         Caption         =   "mit Detailzahlen"
         Height          =   255
         Left            =   9840
         TabIndex        =   274
         ToolTipText     =   "seit einem Jahr nicht mehr verkauft"
         Top             =   2640
         Width           =   1935
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   6
         Left            =   3720
         TabIndex        =   273
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         Caption         =   "auto Abgleich"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   35
         Left            =   5400
         TabIndex        =   355
         Top             =   3200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         Caption         =   "Lagerplatz"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command30 
         Height          =   360
         Left            =   5280
         TabIndex        =   360
         ToolTipText     =   "Kalender"
         Top             =   2160
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
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label Label7 
         Caption         =   "angelegt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   3480
         TabIndex        =   362
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "bis KVK"
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
         Index           =   12
         Left            =   1800
         TabIndex        =   357
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "von KVK"
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
         Left            =   120
         TabIndex        =   356
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Filteroptionen"
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
         Left            =   9840
         TabIndex        =   345
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Größe"
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
         Index           =   47
         Left            =   5040
         TabIndex        =   308
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Lief.Best.Nr"
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
         Left            =   6960
         TabIndex        =   307
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "AGN"
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
         Left            =   8040
         TabIndex        =   306
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "EAN-Code / ArtNr"
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
         Left            =   120
         TabIndex        =   305
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikelbezeichnung"
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
         Left            =   2280
         TabIndex        =   304
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Linie"
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
         Left            =   6240
         TabIndex        =   303
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "alle Farben"
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
         Left            =   8400
         TabIndex        =   302
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "von Lagerplatz"
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
         Index           =   41
         Left            =   120
         TabIndex        =   301
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "bis Lagerplatz"
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
         Index           =   42
         Left            =   1800
         TabIndex        =   300
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "PGN"
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
         Index           =   8
         Left            =   3480
         TabIndex        =   299
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Marke"
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
         Left            =   3480
         TabIndex        =   298
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "kein Lieferant"
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
         Index           =   10
         Left            =   120
         TabIndex        =   297
         Top             =   600
         Width           =   3375
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
         Index           =   59
         Left            =   120
         TabIndex        =   296
         Top             =   2520
         Width           =   9255
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   1215
      Left            =   6960
      TabIndex        =   88
      Top             =   6120
      Visible         =   0   'False
      Width           =   1935
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   3
         Left            =   0
         TabIndex        =   92
         Top             =   720
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   14
         Left            =   960
         TabIndex        =   126
         Top             =   720
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   19
         Left            =   1440
         TabIndex        =   131
         Top             =   720
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   8
         Left            =   480
         TabIndex        =   97
         Top             =   720
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   9
         Left            =   480
         TabIndex        =   98
         Top             =   960
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   15
         Left            =   960
         TabIndex        =   127
         Top             =   960
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   4
         Left            =   0
         TabIndex        =   93
         Top             =   960
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   0
         Left            =   0
         TabIndex        =   89
         Top             =   0
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   16
         Left            =   1440
         TabIndex        =   128
         Top             =   0
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   5
         Left            =   480
         TabIndex        =   94
         Top             =   0
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   11
         Left            =   960
         TabIndex        =   123
         Top             =   0
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   1
         Left            =   0
         TabIndex        =   90
         Top             =   240
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   12
         Left            =   960
         TabIndex        =   124
         Top             =   240
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   17
         Left            =   1440
         TabIndex        =   129
         Top             =   240
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   6
         Left            =   480
         TabIndex        =   95
         Top             =   240
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   2
         Left            =   0
         TabIndex        =   91
         Top             =   480
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   13
         Left            =   960
         TabIndex        =   125
         Top             =   480
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   18
         Left            =   1440
         TabIndex        =   130
         Top             =   480
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command10 
         Height          =   235
         Index           =   7
         Left            =   480
         TabIndex        =   96
         Top             =   480
         Width           =   475
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
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
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
      Height          =   8040
      Left            =   8880
      TabIndex        =   145
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CheckBox Check15 
         BackColor       =   &H00C0C000&
         Caption         =   "Shop-Artikel"
         Height          =   255
         Left            =   8880
         TabIndex        =   368
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   48
         Left            =   6960
         MaxLength       =   6
         TabIndex        =   366
         ToolTipText     =   "LEK Abschlag in % eingeben -> Enter"
         Top             =   1060
         Width           =   615
      End
      Begin sevCommand3.Command Command1 
         Height          =   420
         Index           =   33
         Left            =   3510
         TabIndex        =   344
         Top             =   3830
         Width           =   420
         _ExtentX        =   741
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
         ToolTipTitle    =   "Wechseln"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   420
         Index           =   32
         Left            =   3510
         TabIndex        =   343
         Top             =   3360
         Width           =   420
         _ExtentX        =   741
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
         ToolTipTitle    =   "Wechseln"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   405
         Index           =   30
         Left            =   9240
         TabIndex        =   338
         Top             =   5640
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
      Begin VB.TextBox Text2 
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
         Index           =   10
         Left            =   7680
         MaxLength       =   6
         TabIndex        =   331
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox Text2 
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
         Index           =   9
         Left            =   5160
         MaxLength       =   35
         TabIndex        =   330
         Top             =   5640
         Width           =   2415
      End
      Begin VB.TextBox Text2 
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
         Index           =   8
         Left            =   2640
         MaxLength       =   35
         TabIndex        =   329
         Top             =   5640
         Width           =   2415
      End
      Begin VB.TextBox Text2 
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
         Index           =   7
         Left            =   120
         MaxLength       =   35
         TabIndex        =   328
         Top             =   5640
         Width           =   2415
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   29
         Left            =   1680
         TabIndex        =   309
         Top             =   120
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
         Caption         =   "..."
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
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
         Index           =   43
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   333
         Top             =   5640
         Width           =   735
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   28
         Left            =   8640
         TabIndex        =   271
         Top             =   120
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
         Caption         =   "..."
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   27
         Left            =   7320
         TabIndex        =   270
         Top             =   1305
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
         Caption         =   "..."
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
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
         Index           =   40
         Left            =   7440
         MaxLength       =   1
         TabIndex        =   195
         ToolTipText     =   "stornierfähig = J , Ausnahme = N wenn dieser Artikel nicht in der Stornosumme des Z-Bon enthalten sein sollte"
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox Text1 
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
         Index           =   39
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   194
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text1 
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
         Index           =   38
         Left            =   5040
         MaxLength       =   1
         TabIndex        =   193
         ToolTipText     =   "umsatzrelevant"
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox Text1 
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
         Index           =   5
         Left            =   840
         MaxLength       =   6
         TabIndex        =   176
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox Text1 
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
         Index           =   6
         Left            =   2880
         MaxLength       =   35
         TabIndex        =   177
         Top             =   120
         Width           =   4815
      End
      Begin VB.TextBox Text1 
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
         Index           =   8
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   185
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text1 
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
         Index           =   10
         Left            =   9840
         MaxLength       =   1
         TabIndex        =   179
         Top             =   1300
         Width           =   255
      End
      Begin VB.TextBox Text1 
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
         Index           =   11
         Left            =   6120
         MaxLength       =   9
         TabIndex        =   182
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text1 
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
         Index           =   12
         Left            =   3000
         MaxLength       =   9
         TabIndex        =   186
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox Text1 
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
         Index           =   13
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   190
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
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
         Index           =   14
         Left            =   11040
         MaxLength       =   5
         TabIndex        =   197
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text1 
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
         Index           =   15
         Left            =   11040
         MaxLength       =   1
         TabIndex        =   189
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text1 
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
         Index           =   16
         Left            =   5280
         MaxLength       =   25
         TabIndex        =   203
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox Text1 
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
         Index           =   17
         Left            =   8760
         MaxLength       =   6
         TabIndex        =   201
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text1 
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
         Index           =   18
         Left            =   1080
         MaxLength       =   13
         TabIndex        =   202
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox Text1 
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
         Index           =   19
         Left            =   3840
         MaxLength       =   13
         TabIndex        =   183
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Text1 
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
         Index           =   20
         Left            =   1080
         MaxLength       =   13
         TabIndex        =   204
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox Text1 
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
         Index           =   21
         Left            =   1080
         MaxLength       =   13
         TabIndex        =   205
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox Text1 
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
         Index           =   22
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   198
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text1 
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
         Index           =   23
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   199
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Text1 
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
         Index           =   24
         Left            =   6000
         MaxLength       =   1
         TabIndex        =   200
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Text1 
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
         Index           =   25
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   191
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text1 
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
         Index           =   26
         Left            =   3960
         MaxLength       =   1
         TabIndex        =   192
         ToolTipText     =   "rabattierfähig"
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox Text1 
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
         Index           =   27
         Left            =   9120
         MaxLength       =   1
         TabIndex        =   196
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox Text1 
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
         Index           =   28
         Left            =   9000
         MaxLength       =   3
         TabIndex        =   184
         Top             =   1300
         Width           =   735
      End
      Begin VB.TextBox Text1 
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
         Index           =   29
         Left            =   9000
         MaxLength       =   9
         TabIndex        =   188
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text1 
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
         Index           =   30
         Left            =   6120
         MaxLength       =   8
         TabIndex        =   187
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text1 
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
         Index           =   31
         Left            =   6000
         MaxLength       =   1
         TabIndex        =   206
         Top             =   3720
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
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
         Index           =   32
         Left            =   7920
         MaxLength       =   5
         TabIndex        =   207
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Text1 
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
         Index           =   33
         Left            =   10320
         MaxLength       =   1
         TabIndex        =   208
         Top             =   3720
         Width           =   255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   34
         Left            =   11280
         MaxLength       =   2
         TabIndex        =   175
         Top             =   6120
         Visible         =   0   'False
         Width           =   495
      End
      Begin sevCommand3.Command Command8 
         Height          =   615
         Left            =   2040
         TabIndex        =   174
         Top             =   4680
         Visible         =   0   'False
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   1085
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
         Caption         =   "Bestände in Filialen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command9 
         Height          =   615
         Left            =   4800
         TabIndex        =   173
         Top             =   4680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1085
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
         Caption         =   "Info"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   615
         Left            =   6960
         TabIndex        =   172
         Top             =   4680
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1085
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
         Caption         =   "Calc"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   615
         Index           =   0
         Left            =   8400
         TabIndex        =   171
         Top             =   4680
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1085
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
         Height          =   615
         Index           =   1
         Left            =   10080
         TabIndex        =   170
         Top             =   4680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
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
      Begin VB.ComboBox cbo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   720
         TabIndex        =   181
         Top             =   1340
         Width           =   1215
      End
      Begin sevCommand3.Command Command5 
         Height          =   300
         Index           =   2
         Left            =   1935
         TabIndex        =   169
         Top             =   1080
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   529
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
         Caption         =   "neu"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   300
         Index           =   3
         Left            =   2760
         TabIndex        =   168
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
         Caption         =   "löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   167
         Top             =   1305
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Index           =   9
         Left            =   10080
         TabIndex        =   166
         Top             =   120
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   11160
         MaxLength       =   6
         TabIndex        =   180
         Top             =   120
         Width           =   615
      End
      Begin sevCommand3.Command Command1 
         Height          =   405
         Index           =   12
         Left            =   9840
         TabIndex        =   165
         Top             =   2760
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
         Height          =   285
         Index           =   13
         Left            =   11160
         TabIndex        =   164
         Top             =   3840
         Width           =   615
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
         Caption         =   "LUG"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   285
         Index           =   14
         Left            =   8040
         TabIndex        =   163
         Top             =   4365
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
         Height          =   285
         Index           =   15
         Left            =   11040
         TabIndex        =   162
         Top             =   4365
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
      Begin VB.TextBox Text2 
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
         Index           =   2
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   161
         Top             =   7080
         Width           =   855
      End
      Begin VB.TextBox Text2 
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
         Index           =   3
         Left            =   120
         MaxLength       =   13
         TabIndex        =   160
         Top             =   7080
         Width           =   2415
      End
      Begin VB.TextBox Text2 
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
         Index           =   4
         Left            =   120
         MaxLength       =   13
         TabIndex        =   159
         Top             =   7800
         Width           =   1095
      End
      Begin VB.TextBox Text2 
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
         Index           =   5
         Left            =   1440
         MaxLength       =   13
         TabIndex        =   158
         Top             =   7800
         Width           =   1095
      End
      Begin sevCommand3.Command Command6 
         Height          =   255
         Index           =   0
         Left            =   9000
         TabIndex        =   157
         Top             =   7200
         Width           =   1800
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
         Caption         =   "Bestandshistorie"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   255
         Index           =   1
         Left            =   9000
         TabIndex        =   156
         Top             =   7560
         Width           =   1800
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
         Caption         =   "EAN Historie"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   255
         Index           =   2
         Left            =   9000
         TabIndex        =   155
         Top             =   7920
         Width           =   1800
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
         Caption         =   "Kassenpreis Historie"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   285
         Index           =   18
         Left            =   3960
         TabIndex        =   154
         Top             =   3240
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
         Height          =   285
         Index           =   20
         Left            =   3960
         TabIndex        =   153
         Top             =   3600
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
         Height          =   285
         Index           =   21
         Left            =   3960
         TabIndex        =   152
         Top             =   3960
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
      Begin VB.TextBox Text2 
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
         Index           =   6
         Left            =   4080
         MaxLength       =   13
         TabIndex        =   151
         Top             =   7800
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         Left            =   5400
         TabIndex        =   150
         Text            =   "Combo1"
         Top             =   7560
         Width           =   735
      End
      Begin sevCommand3.Command Command6 
         Height          =   255
         Index           =   4
         Left            =   6240
         TabIndex        =   149
         Top             =   7920
         Width           =   1560
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
         Caption         =   "Konditionen +"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   255
         Index           =   5
         Left            =   6240
         TabIndex        =   148
         Top             =   7560
         Width           =   1560
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
         Caption         =   "Kond. löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdfarbe 
         Height          =   375
         Left            =   11040
         TabIndex        =   147
         Top             =   2760
         Width           =   735
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
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
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
         Index           =   37
         Left            =   7800
         MaxLength       =   4
         TabIndex        =   178
         ToolTipText     =   "Merkmal: 4 Zeichen zulässig"
         Top             =   120
         Width           =   735
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   26
         Left            =   7320
         TabIndex        =   146
         Top             =   1800
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
         Caption         =   "..."
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   300
         Index           =   4
         Left            =   2760
         TabIndex        =   349
         Top             =   1395
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
         Caption         =   "Etikett"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command11 
         Height          =   615
         Left            =   5670
         TabIndex        =   350
         Top             =   4680
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1085
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
         Caption         =   "Email"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   300
         Index           =   5
         Left            =   1935
         TabIndex        =   351
         Top             =   1395
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   529
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
         Caption         =   "in BV"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   34
         Left            =   7680
         TabIndex        =   354
         Top             =   1800
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
         ToolTip         =   "Staffelpreise"
         ButtonStyle     =   2
         Caption         =   "S"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   21
         Left            =   4080
         TabIndex        =   363
         Top             =   4680
         Width           =   695
         _ExtentX        =   1217
         _ExtentY        =   1085
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
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "weitere EAN..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   49
         Left            =   1080
         MouseIcon       =   "frmWKL10.frx":0AD4
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   353
         Top             =   4360
         Width           =   1815
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   7920
         TabIndex        =   348
         Top             =   1370
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferantenzuordnung:"
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
         Index           =   48
         Left            =   240
         TabIndex        =   347
         Top             =   780
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "geräumt am:"
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
         Index           =   33
         Left            =   10320
         TabIndex        =   346
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Gruppenbezeichnung"
         Height          =   255
         Index           =   25
         Left            =   7680
         TabIndex        =   340
         Top             =   6000
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Größe"
         Height          =   375
         Index           =   24
         Left            =   10440
         TabIndex        =   339
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Gruppe"
         Height          =   375
         Index           =   23
         Left            =   7680
         TabIndex        =   337
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Farbenbezeichnung"
         Height          =   255
         Index           =   22
         Left            =   5160
         TabIndex        =   335
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
         Height          =   255
         Index           =   21
         Left            =   2640
         TabIndex        =   334
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Modell"
         Height          =   375
         Index           =   20
         Left            =   120
         TabIndex        =   332
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         Index           =   4
         X1              =   120
         X2              =   11760
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Umverpackung"
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
         TabIndex        =   327
         Top             =   6600
         Width           =   3375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "stonf:"
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
         Index           =   47
         Left            =   6720
         TabIndex        =   268
         ToolTipText     =   "stornierfähig = J , Ausnahme = N wenn dieser Artikel nicht in der Stornosumme des Z-Bon enthalten sein sollte"
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "PGN:"
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
         Index           =   46
         Left            =   5400
         TabIndex        =   267
         ToolTipText     =   "Produktgruppe Details mit F2"
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "umsr:"
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
         Index           =   45
         Left            =   4320
         TabIndex        =   266
         ToolTipText     =   "umsatzrelevant"
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "auto Kalkulation = ja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   30
         Left            =   9360
         MouseIcon       =   "frmWKL10.frx":0DDE
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   263
         Top             =   2240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "ArtNr:"
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
         Left            =   120
         TabIndex        =   262
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Art-Bez:"
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
         Left            =   1920
         TabIndex        =   261
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Lief.-Nr:"
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
         Index           =   2
         Left            =   720
         TabIndex        =   260
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Linie:"
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
         Left            =   240
         TabIndex        =   259
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "AGN:"
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
         Left            =   10440
         TabIndex        =   258
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "EX:"
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
         Left            =   9840
         TabIndex        =   257
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "List-Ek:"
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
         Left            =   6120
         TabIndex        =   256
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Listen-Vk:"
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
         Left            =   1800
         TabIndex        =   255
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Bestand:"
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
         Index           =   8
         Left            =   0
         TabIndex        =   254
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Nettospanne % :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   9
         Left            =   9360
         MouseIcon       =   "frmWKL10.frx":10E8
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   253
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "MWSt:"
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
         Index           =   10
         Left            =   10200
         TabIndex        =   252
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Notiz:"
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
         Left            =   4440
         TabIndex        =   251
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "letzter EK:"
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
         Index           =   12
         Left            =   7560
         TabIndex        =   250
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "EAN:"
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
         Left            =   360
         TabIndex        =   249
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "LiefBestNr:"
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
         Index           =   14
         Left            =   3840
         TabIndex        =   248
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "2.EAN:"
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
         Left            =   120
         TabIndex        =   247
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "3.EAN:"
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
         Index           =   16
         Left            =   120
         TabIndex        =   246
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Kopie eines Artikels!     ArtNr wurde neu ermittelt!"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   245
         Top             =   4680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Left            =   6720
         TabIndex        =   244
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Inhalt:"
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
         Left            =   240
         TabIndex        =   243
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Einheit:"
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
         Index           =   18
         Left            =   2280
         TabIndex        =   242
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Grundpreis:"
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
         Left            =   4680
         TabIndex        =   241
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "MB:"
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
         Index           =   20
         Left            =   1800
         TabIndex        =   240
         ToolTipText     =   "Mindestbestand"
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "rabf:"
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
         Left            =   3240
         TabIndex        =   239
         ToolTipText     =   "rabattierfähig"
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "gef:"
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
         Index           =   22
         Left            =   8400
         TabIndex        =   238
         ToolTipText     =   "geführt"
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "VPE:"
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
         Index           =   23
         Left            =   9000
         TabIndex        =   237
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "S-Ek:"
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
         Index           =   24
         Left            =   8160
         TabIndex        =   236
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Kassen-Vk:"
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
         Left            =   4800
         TabIndex        =   235
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   11760
         Y1              =   1720
         Y2              =   1720
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Preisschutz:"
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
         Index           =   26
         Left            =   4680
         TabIndex        =   234
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "in Bestell:"
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
         Index           =   27
         Left            =   6720
         TabIndex        =   233
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Bonus-fähig:"
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
         Index           =   28
         Left            =   8880
         TabIndex        =   232
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Left            =   6720
         TabIndex        =   231
         Top             =   2880
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Farbe:"
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
         Index           =   29
         Left            =   10320
         TabIndex        =   230
         Top             =   2880
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   2
         X1              =   120
         X2              =   11760
         Y1              =   700
         Y2              =   700
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   3
         X1              =   120
         X2              =   120
         Y1              =   720
         Y2              =   1720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   11760
         X2              =   11760
         Y1              =   705
         Y2              =   1720
      End
      Begin VB.Label lblLiefbez 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2880
         TabIndex        =   229
         Top             =   780
         Width           =   8415
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   9840
         TabIndex        =   228
         Top             =   500
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "angelegt am:"
         Height          =   255
         Index           =   31
         Left            =   3000
         TabIndex        =   227
         Top             =   4365
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "angelegt am:"
         Height          =   255
         Index           =   34
         Left            =   10320
         TabIndex        =   226
         Top             =   1395
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "angelegt am:"
         Height          =   255
         Index           =   32
         Left            =   4440
         TabIndex        =   225
         Top             =   4365
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "angelegt am:"
         Height          =   255
         Index           =   35
         Left            =   6960
         TabIndex        =   224
         Top             =   4365
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "letzter Zugang:"
         Height          =   255
         Index           =   36
         Left            =   8400
         TabIndex        =   223
         Top             =   4365
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "angelegt am:"
         Height          =   255
         Index           =   37
         Left            =   9720
         TabIndex        =   222
         Top             =   4365
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "letzter Vk:"
         Height          =   255
         Index           =   38
         Left            =   5880
         TabIndex        =   221
         Top             =   4365
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Inhalt"
         Height          =   255
         Index           =   12
         Left            =   2640
         TabIndex        =   220
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "EAN - Code"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   219
         Top             =   6840
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Lagerplatz"
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
         Index           =   14
         Left            =   120
         TabIndex        =   218
         Top             =   7560
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Konditionen bei Bestellung  z.B.: (6 + 1)"
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
         Left            =   1440
         TabIndex        =   217
         Top             =   7560
         Width           =   4095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "+"
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
         Index           =   16
         Left            =   2760
         TabIndex        =   216
         Top             =   7800
         Width           =   255
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "6"
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
         Left            =   2640
         TabIndex        =   215
         Top             =   7800
         Width           =   255
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Index           =   18
         Left            =   3000
         TabIndex        =   214
         Top             =   7800
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H000000FF&
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
         Index           =   39
         Left            =   4080
         TabIndex        =   213
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "LUGD"
         Height          =   255
         Index           =   44
         Left            =   11160
         TabIndex        =   264
         ToolTipText     =   "Lagerumschlagsdauer in Tagen"
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "LR"
         Height          =   255
         Index           =   40
         Left            =   10560
         TabIndex        =   212
         ToolTipText     =   "Lagerreichweite in Tagen"
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "LUG"
         Height          =   255
         Index           =   41
         Left            =   10560
         TabIndex        =   211
         ToolTipText     =   "Umschlagshäufigkeit"
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "LUGD"
         Height          =   255
         Index           =   42
         Left            =   11160
         TabIndex        =   210
         ToolTipText     =   "Lagerumschlagsdauer in Tagen"
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "LR"
         Height          =   255
         Index           =   43
         Left            =   10560
         TabIndex        =   209
         ToolTipText     =   "Lagerreichweite in Tagen"
         Top             =   3240
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Undurchsichtig
         BorderWidth     =   2
         Height          =   345
         Left            =   7920
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.PictureBox picprogress 
      Height          =   375
      Left            =   7080
      ScaleHeight     =   315
      ScaleWidth      =   4155
      TabIndex        =   115
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   10440
      TabIndex        =   114
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H0080C0FF&
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
      Height          =   495
      Left            =   600
      TabIndex        =   99
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
      Begin VB.ComboBox cboGp 
         Height          =   330
         Left            =   4200
         TabIndex        =   110
         Top             =   1920
         Width           =   2895
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   5
         Left            =   9360
         TabIndex        =   112
         Top             =   3240
         Width           =   1935
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
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   1
         Left            =   8400
         MaxLength       =   4
         TabIndex        =   111
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   0
         Left            =   120
         MaxLength       =   13
         TabIndex        =   109
         Top             =   1920
         Width           =   2895
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   4
         Left            =   9360
         TabIndex        =   113
         Top             =   3720
         Width           =   1935
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
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikelbezeichnung"
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
         Index           =   11
         Left            =   3000
         TabIndex        =   108
         Top             =   2760
         Visible         =   0   'False
         Width           =   7695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikelnummer"
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
         Index           =   10
         Left            =   120
         TabIndex        =   107
         Top             =   2760
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikelbezeichnung"
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
         Index           =   9
         Left            =   3000
         TabIndex        =   106
         Top             =   2400
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikelnummer:"
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
         Index           =   8
         Left            =   120
         TabIndex        =   105
         Top             =   2400
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
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
         Index           =   7
         Left            =   240
         TabIndex        =   104
         Top             =   4800
         Width           =   11055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Anlegen von Umverpackungen"
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
         Left            =   120
         TabIndex        =   103
         Top             =   120
         Width           =   6255
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "EAN - Codes des Einzelprodukts"
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
         Index           =   5
         Left            =   120
         TabIndex        =   102
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Inhalt der  Umverpackung"
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
         Index           =   4
         Left            =   8400
         TabIndex        =   101
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "EAN - Code der Umverpackung"
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
         Index           =   3
         Left            =   4200
         TabIndex        =   100
         Top             =   1560
         Width           =   3375
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
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   12135
      Begin sevCommand3.Command Command1 
         Height          =   255
         Index           =   31
         Left            =   10680
         TabIndex        =   342
         Top             =   480
         Width           =   1095
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
         Caption         =   "einfügen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   10680
         TabIndex        =   341
         Top             =   120
         Width           =   1095
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   8
         Left            =   1350
         TabIndex        =   336
         Top             =   7800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         Caption         =   "Gruppieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00C0C000&
         Caption         =   "ohne Grund"
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
         TabIndex        =   326
         ToolTipText     =   "ohne Angabe von Gründen die Bestände minimieren"
         Top             =   720
         Width           =   1575
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'Kein
         FontTransparent =   0   'False
         Height          =   855
         Left            =   1920
         MouseIcon       =   "frmWKL10.frx":13F2
         MousePointer    =   99  'Benutzerdefiniert
         ScaleHeight     =   855
         ScaleWidth      =   975
         TabIndex        =   325
         Top             =   360
         Width           =   975
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   7
         Left            =   4780
         TabIndex        =   324
         Top             =   7800
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
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
         Caption         =   "Shop"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikelgruppe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   5400
         TabIndex        =   144
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Farbmerkmal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   5400
         TabIndex        =   143
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bestand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   5400
         TabIndex        =   142
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Produktgruppe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   5400
         TabIndex        =   141
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Linie, Bezeichnung"
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
         Left            =   3120
         TabIndex        =   140
         Top             =   480
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bezeichnung"
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
         Left            =   3120
         TabIndex        =   139
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "LiefBestNr"
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
         Left            =   3120
         TabIndex        =   138
         Top             =   960
         Width           =   1575
      End
      Begin sevCommand3.Command Command1 
         Height          =   255
         Index           =   23
         Left            =   10680
         TabIndex        =   136
         Top             =   960
         Width           =   1095
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
         Caption         =   "weitere..."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   255
         Index           =   22
         Left            =   9480
         TabIndex        =   135
         Top             =   960
         Width           =   1095
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
         Caption         =   "anzeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C000&
         Caption         =   "nicht geführte  ausblenden"
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
         Left            =   7440
         TabIndex        =   134
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "nur mit Bestand"
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
         Left            =   7440
         TabIndex        =   133
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0C000&
         Caption         =   "Ex ausblenden"
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
         Left            =   7440
         TabIndex        =   132
         Top             =   720
         Value           =   1  'Aktiviert
         Width           =   2055
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   6
         Left            =   6585
         TabIndex        =   122
         Top             =   7800
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   873
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
         Caption         =   "Extras"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   1
         Left            =   8190
         TabIndex        =   121
         Top             =   7800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
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
         Caption         =   "Listen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   5
         Left            =   9080
         TabIndex        =   120
         Top             =   7800
         Width           =   1455
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
         Caption         =   "t. VK-Preise"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   4
         Left            =   10560
         TabIndex        =   72
         Top             =   7800
         Width           =   1215
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   3
         Left            =   3655
         TabIndex        =   3
         Top             =   7800
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   873
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
         Caption         =   "Kopieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   2
         Left            =   2610
         TabIndex        =   2
         Top             =   7800
         Width           =   995
         _ExtentX        =   1746
         _ExtentY        =   873
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   7800
         Width           =   1190
         _ExtentX        =   2090
         _ExtentY        =   873
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
         Caption         =   "Auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0C000&
         Caption         =   "überschreiben"
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
         TabIndex        =   118
         ToolTipText     =   "Tabelleninhalt mit Eingabe überschreiben"
         Top             =   960
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6255
         Left            =   120
         TabIndex        =   269
         Top             =   1320
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   11033
         _Version        =   393216
         Cols            =   18
         FixedCols       =   2
         ForeColorSel    =   8454143
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   2
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
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   9
         Left            =   7350
         TabIndex        =   352
         Top             =   7800
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   873
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
         Caption         =   "in BV"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   10
         Left            =   5560
         TabIndex        =   359
         Top             =   7800
         Width           =   980
         _ExtentX        =   1720
         _ExtentY        =   873
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
         Caption         =   "Etiketten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         Caption         =   "rote Lieferanten = mehrere Lieferanten pro Artikel hinterlegt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   13
         Left            =   120
         TabIndex        =   358
         Top             =   7620
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   1440
         Top             =   480
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "sortiert nach:"
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
         Left            =   3120
         TabIndex        =   137
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0FF&
         Caption         =   "anzeige"
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
         TabIndex        =   119
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label0 
         BackColor       =   &H00C0C000&
         Caption         =   "Label0"
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
         Index           =   3
         Left            =   1200
         TabIndex        =   76
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label0 
         BackColor       =   &H00C0C000&
         Caption         =   "Label0"
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
         Index           =   2
         Left            =   840
         TabIndex        =   75
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label0 
         BackColor       =   &H00C0C000&
         Caption         =   "Label0"
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
         Index           =   1
         Left            =   120
         TabIndex        =   74
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label0 
         BackColor       =   &H00C0C000&
         Caption         =   "Label0"
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
         Index           =   0
         Left            =   480
         TabIndex        =   73
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H000000FF&
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
      Height          =   1695
      Left            =   960
      TabIndex        =   77
      Top             =   6600
      Visible         =   0   'False
      Width           =   1935
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   3
         Left            =   9840
         TabIndex        =   84
         Top             =   4200
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
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   120
         TabIndex        =   79
         Top             =   1200
         Width           =   5295
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   78
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bestände in den Filialen"
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
         Left            =   120
         TabIndex        =   87
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   10200
         TabIndex        =   85
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
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
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   83
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   82
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikelbezeichnung"
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
         Left            =   2160
         TabIndex        =   81
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikelnummer"
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
         Left            =   120
         TabIndex        =   80
         Top             =   0
         Width           =   1815
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   345
      Index           =   11
      Left            =   11400
      TabIndex        =   265
      Top             =   120
      Width           =   345
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
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C000&
      Caption         =   "EAN-Code / ArtNr"
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
      Left            =   6240
      TabIndex        =   117
      Top             =   8280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C000&
      Caption         =   "EAN-Code / ArtNr"
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
      Left            =   6240
      TabIndex        =   116
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label0 
      BackColor       =   &H00C0C000&
      Caption         =   "Artikeldaten bearbeiten"
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
      Index           =   4
      Left            =   120
      TabIndex        =   86
      Top             =   0
      Width           =   6825
   End
End
Attribute VB_Name = "frmWKL10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bAusblenden As Boolean
Dim bBest As Boolean

Dim binbest As Boolean
Dim bgef As Boolean
Dim bBeimSpeichern As Boolean
Dim iFocus As Integer

Dim gbNew As Boolean
Dim gbcomefromwoa As Boolean
Dim SpaltennummerAWM As Byte
Dim SpaltennummerArtnr As Byte
Dim SpaltennummerBEZEICH  As Byte
Dim SpaltennummerKVKPR1  As Byte
Dim SpaltennummerHS  As Byte

Dim SpaltennummerGEFUEHRT  As Byte
Dim SpaltennummerRABATT_OK  As Byte
Dim SpaltennummerBONUS_OK  As Byte
Dim SpaltennummerRKZ  As Byte
Dim SpaltennummerPREISSCHU As Byte
Dim SpaltennummerBESTAND As Byte

Dim SpaltennummerPGN As Byte
Dim SpaltennummerEAN As Byte
Dim SpaltennummerEAN2 As Byte
Dim SpaltennummerEAN3 As Byte
Dim SpaltennummerLPZ As Byte
Dim SpaltennummerNOTIZEN As Byte
Dim SpaltennummerAGN As Byte
Dim SpaltennummerLINR As Byte
Dim SpaltennummerLEKPR As Byte
Dim SpaltennummerLVKPR As Byte
Dim SpaltennummerLIBESNR As Byte
Dim SpaltennummerLagerP As Byte
Dim SpaltennummerGROESSE As Byte
Dim SpaltennummerMB As Byte
Dim SpaltennummerSHOP As Byte

Dim SpaltennummerModell As Byte
Dim SpaltennummerMaterial As Byte
Dim SpaltennummerFarbbez As Byte
Dim SpaltennummerGRUPPE As Byte
Dim SpaltennummerMWST As Byte

Dim gbAender As Boolean 'Bestand
Dim gbAenderKVK As Boolean
Dim gbAenderAWM As Boolean
Dim gbAenderPGN As Boolean
Dim gbAenderBEZEICH As Boolean
Dim gbAenderEAN As Boolean
Dim gbAenderEAN2 As Boolean
Dim gbAenderEAN3 As Boolean
Dim gbAenderLPZ As Boolean
Dim gbAendergefuehrt As Boolean
Dim gbAenderpreisSchu As Boolean
Dim gbAenderRABATT_OK As Boolean
Dim gbAenderBONUS_OK As Boolean
Dim gbAenderRKZ As Boolean
Dim gbAenderMWST As Boolean
Dim gbAenderNOTIZEN As Boolean
Dim gbAenderAGN As Boolean
Dim gbAenderLINR As Boolean
Dim gbAenderLEKPR As Boolean
Dim gbAenderLVKPR As Boolean
Dim gbAenderLIBESNR As Boolean
Dim gbAenderLAGERP As Boolean
Dim gbAenderGROESSE As Boolean
Dim gbAenderMB As Boolean
Dim gbAenderSHOP As Boolean
Dim gbAenderMODELL As Boolean
Dim gbAenderMATERIAL As Boolean
Dim gbAenderFARBBEZ As Boolean
Dim gbAenderGRUPPE As Boolean

Private Sub WKL10Positionieren()
    On Error GoTo LOKAL_ERROR
    
    Frame0.Top = 600
    Frame0.Left = 0
    Frame0.Height = 6615
    Frame0.Width = 11895
    
    Frame1.Top = 0
    Frame1.Left = 0
    Frame1.Height = 8655
    Frame1.Width = 11895
    
    Frame2.Top = 5400
    Frame2.Left = 0
    Frame2.Height = 3135
    Frame2.Width = 11895
    
    Frame3.Top = 0
    Frame3.Left = 0
    Frame3.Height = 8520
    Frame3.Width = 11895
    
    Frame6.Top = 0
    Frame6.Left = 0
    Frame6.Height = 8200
    Frame6.Width = 11895
    
    Frame4.Top = 0
    Frame4.Left = 0
    Frame4.Height = 5500
    Frame4.Width = 11895
    
    Frame5.Top = 3350
    Frame5.Left = 10800
    Frame5.Height = 1215
    Frame5.Width = 1935
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL10Positionieren"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnPruefeAGNWKL10() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim cAgn As String
    Dim cSQL As String
    Dim rsrs As Recordset
    
    fnPruefeAGNWKL10 = 0
    
    cAgn = Trim$(Str$(Val(Text1(9).Text)))
    
    cSQL = "Select * from AGNDBF where AGN = " & cAgn
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        fnPruefeAGNWKL10 = 1
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeAGNWKL10"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Sub fuellecombo1(cSuchi As String)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    
    Combo1.Clear
    
    cSQL = "Select * from KONDITIONEN where ARTNR = " & cSuchi & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!kondi) Then
            Combo1.AddItem rsrs!kondi
        End If
        rsrs.MoveNext
        Loop
        Combo1.Text = Text2(5).Text
    Else
        Combo1.Text = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo1"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Function fnPruefeLINRWKL10() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim cLinr As String
    Dim cSQL As String
    Dim rsrs As Recordset
    
    fnPruefeLINRWKL10 = 0
    
    cLinr = Trim$(Str$(Val(cbo1.Text)))
    
    If cLinr = "0" Then
        fnPruefeLINRWKL10 = 1
        Exit Function
    End If
    
    If cLinr = "" Then
        fnPruefeLINRWKL10 = 1
        Exit Function
    End If
    
    cSQL = "Select * from LISRT where LINR = " & cLinr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        fnPruefeLINRWKL10 = 1
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeLINRWKL10"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub FlexGrid_Delete(oGrid As MSFlexGrid)
On Error GoTo LOKAL_ERROR

    Dim nRow As Long
    Dim nCol As Long
    Dim nRowSel As Long
    Dim nColSel As Long
    Dim nDelRow As Long
    Dim lBig As Long
    
    Dim cArtNr As String
  

    With oGrid
        ' aktuelle Selektion merken
      
        nRow = .Row
        nCol = .Col
        nRowSel = .RowSel
        nColSel = .ColSel
      
      
        If nRow > nRowSel Then
            lBig = nRow
            nDelRow = nRowSel - 1
        Else
            lBig = nRowSel
            nDelRow = nRow - 1
        End If
      
      
        Do While nDelRow < lBig
        
            nDelRow = nDelRow + 1
            
            If nDelRow > 1 Then
            
                cArtNr = Trim(.TextMatrix(nDelRow, SpaltennummerArtnr))
                LoescheArtikelWKL10 cArtNr
                .RowHeight(nDelRow) = 0
                
            End If
        Loop

  
  End With
  

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FlexGrid_Delete"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub FlexGrid_Update(oGrid As MSFlexGrid)
On Error GoTo LOKAL_ERROR

    Dim nRow As Long
    Dim nCol As Long
    Dim nRowSel As Long
    Dim nColSel As Long
    Dim nDelRow As Long
    Dim lBig As Long
    
    Dim cArtNr As String
    Dim cWert As String
    cWert = Text3.Text
    
    With oGrid
        ' aktuelle Selektion merken
      
        nRow = .Row
        nCol = .Col
        nRowSel = .RowSel
        nColSel = .ColSel
      
      
        If nRow > nRowSel Then
            lBig = nRow
            nDelRow = nRowSel - 1
        Else
            lBig = nRowSel
            nDelRow = nRow - 1
        End If
      
        Do While nDelRow < lBig
        
            nDelRow = nDelRow + 1
            
            If nDelRow > 1 Then
            
                cArtNr = Trim(.TextMatrix(nDelRow, SpaltennummerArtnr))
                If UpdateArtikelWKL10(cArtNr, nCol, cWert) = True Then
                
                    .TextMatrix(nDelRow, nCol) = cWert
                End If

            End If
        Loop

  
  End With
  

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FlexGrid_Update"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FlexGrid_Etiketten(oGrid As MSFlexGrid)
On Error GoTo LOKAL_ERROR

    Dim nRow As Long
    Dim nCol As Long
    Dim nRowSel As Long
    Dim nColSel As Long
    Dim nDelRow As Long
    Dim lBig As Long
    Dim sSQL As String
    
    Dim cArtNr As String
    Dim cWert As String
    cWert = Text3.Text
    
    With oGrid
        ' aktuelle Selektion merken
      
        nRow = .Row
        nCol = .Col
        nRowSel = .RowSel
        nColSel = .ColSel
      
        If nRow > nRowSel Then
            lBig = nRow
            nDelRow = nRowSel - 1
        Else
            lBig = nRowSel
            nDelRow = nRow - 1
        End If
      
        Do While nDelRow < lBig
        
            nDelRow = nDelRow + 1
            
            If nDelRow > 1 Then
            
                cArtNr = Trim(.TextMatrix(nDelRow, SpaltennummerArtnr))
                
                sSQL = "Insert into LSTEETI select Artnr "
                sSQL = sSQL & ",   BEZEICH "
                sSQL = sSQL & ", 1 as BESTAND "
                sSQL = sSQL & ", 1 as ANZAHL "
                sSQL = sSQL & ", KVKPR1 as VKPR "
                
                sSQL = sSQL & ", '' as LIBESNR "
                sSQL = sSQL & ",  EAN "
                sSQL = sSQL & ",  LPZ "
                sSQL = sSQL & ", 0 as LINR "
                
                sSQL = sSQL & ", '" & gcFilNr & "' as FILNR "
                sSQL = sSQL & " from Artikel "
                sSQL = sSQL & " where Artnr = " & cArtNr
                gdBase.Execute sSQL, dbFailOnError

            End If
        Loop

  
  End With
  

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FlexGrid_Etiketten"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FlexGrid_Gruppieren(oGrid As MSFlexGrid, lGruppNr As Long)
On Error GoTo LOKAL_ERROR

    Dim nRow As Long
    Dim nCol As Long
    Dim nRowSel As Long
    Dim nColSel As Long
    Dim nDelRow As Long
    Dim lBig As Long
    
    Dim cArtNr As String
  

    With oGrid
        ' aktuelle Selektion merken
      
        nRow = .Row
        nCol = .Col
        nRowSel = .RowSel
        nColSel = .ColSel
      
      
        If nRow > nRowSel Then
            lBig = nRow
            nDelRow = nRowSel - 1
        Else
            lBig = nRowSel
            nDelRow = nRow - 1
        End If
      
      
        Do While nDelRow < lBig
        
            nDelRow = nDelRow + 1
            
            If nDelRow > 1 Then
            
                cArtNr = Trim(.TextMatrix(nDelRow, SpaltennummerArtnr))

                Gruppiere_ArtikelWKL10 cArtNr, lGruppNr
'                .RowHeight(nDelRow) = 0
                
            End If
        Loop

  
  End With
  

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FlexGrid_Gruppieren"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnPruefeDialogEingabenWKL10() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim iCount As Integer
    Dim ctmp As String
    Dim cZeichen As String
    Dim cValid As String
    Dim lRet As Long
    Dim lPos As Long
    Dim cMeld As String
    Dim iRet As Integer
    
    If Trim$(Text1(6).Text) = "" Or Trim$(Text1(5).Text) = "" Then
        fnPruefeDialogEingabenWKL10 = 99
        Exit Function
    End If
    
    
    For lcount = 5 To 40 '33
        If lcount = 7 Then lcount = 8

        ctmp = Text1(lcount).Text
        ctmp = Trim$(ctmp)
        Select Case lcount
            Case 5, 7, 8, 9, 13, 25, 28
                cValid = "1234567890-"
                If ctmp <> "" Then
                    For iCount = 1 To Len(ctmp)
                        cZeichen = Mid(ctmp, iCount, 1)
                        If InStr(cValid, cZeichen) = 0 Then
                            fnPruefeDialogEingabenWKL10 = lcount
                            cMeld = "Das Feld enthält ein ungültiges Zeichen!" & vbCrLf & vbCrLf
                            cMeld = cMeld & "Gültig sind: " & cValid
                            MsgBox cMeld, vbCritical, "STOP!"
                            fnPruefeDialogEingabenWKL10 = lcount
                            Exit Function
                        End If
                    Next iCount
                    If lcount = 13 Then
                        If Abs(Val(ctmp)) >= 1000 Then
                            iRet = MsgBox("Mengeneingabe von " & ctmp & " fraglich! Wirklich speichern?", vbQuestion + vbYesNo + vbDefaultButton1, "MENGE")
                            If iRet = vbNo Then
                                fnPruefeDialogEingabenWKL10 = lcount
                            End If
                        End If
                    End If
                    
                Else
                    If lcount = 8 Then
                        Text1(lcount).Text = "1"
                    End If
                End If
                
                
            Case 11, 12, 22, 29, 30
                cValid = "-1234567890,"
                If ctmp <> "" Then
                    For iCount = 1 To Len(ctmp)
                        cZeichen = Mid(ctmp, iCount, 1)
                        lPos = InStr(cValid, cZeichen)
                        If lPos = 0 Then
                            cMeld = "Das Feld enthält ein ungültiges Zeichen!" & vbCrLf & vbCrLf
                            cMeld = cMeld & "Gültig sind: " & cValid
                            MsgBox cMeld, vbCritical, "STOP!"
                            fnPruefeDialogEingabenWKL10 = lcount
                            Exit Function
                        End If
                        If cZeichen = "," Then
                            lPos = InStr(lPos + 1, ctmp, cZeichen)
                            If lPos <> 0 Then
                                cMeld = "Das Feld enthält ein ungültiges Zeichen!" & vbCrLf & vbCrLf
                                cMeld = cMeld & "Gültig sind: " & cValid
                                MsgBox cMeld, vbCritical, "STOP!"
                                fnPruefeDialogEingabenWKL10 = lcount
                                Exit Function
                            End If
                        End If
                    Next iCount
                End If
            Case 10, 24, 26, 27, 31, 33, 38, 40
                cValid = "JN"
                If ctmp <> "" Then
                    For iCount = 1 To Len(ctmp)
                        cZeichen = Mid(ctmp, iCount, 1)
                        If InStr(cValid, cZeichen) = 0 Then
                            cMeld = "Das Feld enthält ein ungültiges Zeichen!" & vbCrLf & vbCrLf
                            cMeld = cMeld & "Gültig sind: " & cValid
                            MsgBox cMeld, vbCritical, "STOP!"
                            fnPruefeDialogEingabenWKL10 = lcount
                            Exit Function
                        End If
                    Next iCount
                Else
                    If lcount = 10 Then
                        Text1(lcount).Text = "N"
                    End If
                    
                    If lcount = 24 Then
                        Text1(lcount).Text = "N"
                    End If
                    
                    If lcount = 31 Then
                        Text1(lcount).Text = "N"
                    End If
                    
                    If lcount = 33 Then
                        Text1(lcount).Text = "J"
                    End If
                    
                    If lcount = 26 Then
                        Text1(lcount).Text = "J"
                    End If
                    
                    If lcount = 27 Then
                        Text1(lcount).Text = "J"
                    End If
                    
                    If lcount = 38 Then
                        Text1(lcount).Text = "J"
                    End If
                    
                    If lcount = 40 Then
                        Text1(lcount).Text = "J"
                    End If
                End If
            
            Case Is = 14
                cValid = "1234567890,-"

                If ctmp <> "" Then
                    For iCount = 1 To Len(ctmp)
                        cZeichen = Mid(ctmp, iCount, 1)
                        If InStr(cValid, cZeichen) = 0 Then
                            cMeld = "Das Feld enthält ein ungültiges Zeichen!" & vbCrLf & vbCrLf
                            cMeld = cMeld & "Gültig sind: " & cValid
                            MsgBox cMeld, vbCritical, "STOP!"
                            fnPruefeDialogEingabenWKL10 = lcount
                            Exit Function
                        End If
                    Next iCount
                End If
            
            Case Is = 15
                cValid = "VEO"
                If ctmp <> "" Then
                    For iCount = 1 To Len(ctmp)
                        cZeichen = Mid(ctmp, iCount, 1)
                        If InStr(cValid, cZeichen) = 0 Then
                            cMeld = "Das Feld enthält ein ungültiges Zeichen!" & vbCrLf & vbCrLf
                            cMeld = cMeld & "Gültig sind: " & cValid
                            MsgBox cMeld, vbCritical, "STOP!"
                            fnPruefeDialogEingabenWKL10 = lcount
                            Exit Function
                        End If
                    Next iCount
                Else
                    Text1(15).Text = "V"
                End If
            
            Case 18, 20, 21
                If ctmp <> "" Then
                
                
                    lRet = fnPruefeEANWert(ctmp)
                    If lRet <> 0 Then
                        fnPruefeDialogEingabenWKL10 = lcount
                        Exit Function
                    End If
                End If
                
            Case Is = 23
                If ctmp <> "" Then
                    ctmp = UCase$(ctmp)
                    Text1(23).Text = ctmp
                End If
        End Select
        
    Next lcount
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeDialogEingabenWKL10"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub HoleDatenWKL10(cSuch As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSpanne         As String
    Dim se              As String
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim rsrs1           As Recordset
    Dim ctmp            As String
    Dim dWert           As Double
    Dim lJahr           As Long
    Dim sKalkulierbar   As String
    Dim iFehlerstufe    As Integer
    
    Label4(30).Visible = False
    Label4(30).ForeColor = glS1
    Label4(30).Refresh
    
    iFehlerstufe = 0
    lJahr = Year(Now)
    
    glBestandNeu = 0
    glBestandAlt = 0
    
    iFehlerstufe = 2
    
    If cSuch = "" Then
        Exit Sub
    End If
    
    If IsNumeric(cSuch) = False Then
        Exit Sub
    End If
    
    'Lagerumschlag
    Dim dlug As Double
    Dim dlugd As Double

    dlug = HoleLagerumschlag1(cSuch)
    Label4(41).Caption = Format$(dlug, "###0.00")
    Label4(41).Refresh

    dlugd = 0
    If dlug > 0 Then
        dlugd = 360 / dlug
    End If

    Label4(42).Caption = Val(dlugd)
    Label4(42).Refresh
    
    'Odayy START
        Dim returnTage As Double
        returnTage = Val(wievieleTage(cSuch))
    'Odayy ENDE
    
    Label4(40).Caption = CStr(returnTage)
    Label4(40).Refresh
    'Lagerumschlag ende
    
    cSQL = "Select * from ARTIKEL where ARTNR = " & cSuch & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        giDlgZustand = giUPD
        rsrs.MoveFirst
        
        iFehlerstufe = 3
        If Not IsNull(rsrs!artnr) Then
            ctmp = rsrs!artnr
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(5).Text = ctmp
        
        If LeseGeschwisterArt(cSuch) Then
            Command1(29).BackColor = vbRed
        Else
            Command1(29).BackColor = Command1(9).BackColor
        End If

        If LeseInterArt(cSuch) Then
            Command1(28).BackColor = vbRed
        Else
            Command1(28).BackColor = Command1(9).BackColor
        End If
        
        iFehlerstufe = 4
        If Not IsNull(rsrs!ETIMERK) Then
            sKalkulierbar = rsrs!ETIMERK
        Else
            sKalkulierbar = ""
        End If
        
        iFehlerstufe = 5
        If Not IsNull(rsrs!BEZEICH) Then
            ctmp = rsrs!BEZEICH
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(6).Text = ctmp
        
        iFehlerstufe = 51
        If Not IsNull(rsrs!LPZ) Then
            ctmp = rsrs!LPZ
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(8).Text = ctmp
        
        iFehlerstufe = 6
        If Not IsNull(rsrs!AGN) Then
            ctmp = rsrs!AGN
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(9).Text = ctmp
        
        iFehlerstufe = 6
        If Not IsNull(rsrs!PGN) Then
            ctmp = rsrs!PGN
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(39).Text = ctmp
        

        
        iFehlerstufe = 8
        If Not IsNull(rsrs!vkpr) Then
            dWert = rsrs!vkpr
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "#####0.00")
        Text1(12).Text = ctmp
        
        Dim dErrechneterZentraldrogpreis As Double
        Dim cErrechneterZentraldrogpreis As String
        If gbHauptg Then
            dErrechneterZentraldrogpreis = dWert * 80 / 100
            cErrechneterZentraldrogpreis = Runden(dErrechneterZentraldrogpreis)

            If CDbl(Format(cErrechneterZentraldrogpreis, "#####0.00")) <> CDbl(Format(rsrs!KVKPR1, "#####0.00")) Then
                Label4(39).Caption = Format(cErrechneterZentraldrogpreis, "#####0.00")
                Label4(39).Visible = True
                Label4(39).ForeColor = vbRed
            Else
                Label4(39).Caption = ""
                Label4(39).Visible = False
            End If
        End If
        
        iFehlerstufe = 9
        If Not IsNull(rsrs!BESTAND) Then
            dWert = rsrs!BESTAND
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "#####0")
        glBestandAlt = dWert
        Text1(13).Text = ctmp
        Label6.Caption = ctmp
        
        iFehlerstufe = 10
        If Not IsNull(rsrs!MWST) Then
            ctmp = rsrs!MWST
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
       
        gcMwSt = ctmp
        Text1(15).Text = ctmp
        
        iFehlerstufe = 11
        If Not IsNull(rsrs!NOTIZEN) Then
            ctmp = rsrs!NOTIZEN
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(16).Text = ctmp
        
        If Not IsNull(rsrs!GROESSE) Then
            ctmp = rsrs!GROESSE
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(43).Text = ctmp
        
        Text1(17).Text = ErmlastREK(cSuch)
        
        iFehlerstufe = 12
        If Not IsNull(rsrs!EAN) Then
            ctmp = rsrs!EAN
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(18).Text = ctmp
        
        iFehlerstufe = 13
        If Not IsNull(rsrs!EAN2) Then
            ctmp = rsrs!EAN2
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(20).Text = ctmp
        
        iFehlerstufe = 14
        If Not IsNull(rsrs!EAN3) Then
            ctmp = rsrs!EAN3
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(21).Text = ctmp
        
        iFehlerstufe = 141
        If Not IsNull(rsrs!INHALT) Then
            dWert = rsrs!INHALT
        Else
            dWert = 0
        End If
        
        iFehlerstufe = 15
        If dWert = Fix(dWert) Then
            ctmp = Format$(dWert, "#####0")
        Else
            ctmp = Format$(dWert, "#0.000")
        End If
        
        Text1(22).Text = ctmp
        
        iFehlerstufe = 16
        If Not IsNull(rsrs!INHALTBEZ) Then
            ctmp = rsrs!INHALTBEZ
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(23).Text = ctmp
        
        iFehlerstufe = 17
        If Not IsNull(rsrs!GRUNDPREIS) Then
            ctmp = rsrs!GRUNDPREIS
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(24).Text = ctmp
        
        iFehlerstufe = 18
        If Not IsNull(rsrs!MINBEST) Then
            dWert = rsrs!MINBEST
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "#####0")
        Text1(25).Text = ctmp
        
        iFehlerstufe = 19
        If Not IsNull(rsrs!RABATT_OK) Then
            ctmp = rsrs!RABATT_OK
        Else
            ctmp = "J"
        End If
        ctmp = Trim$(ctmp)
        Text1(26).Text = ctmp
        
        If Not IsNull(rsrs!UMS_OK) Then
            ctmp = rsrs!UMS_OK
        Else
            ctmp = "J"
        End If
        ctmp = Trim$(ctmp)
        Text1(38).Text = ctmp
        
        iFehlerstufe = 20
        If Not IsNull(rsrs!GEFUEHRT) Then
            ctmp = rsrs!GEFUEHRT
        Else
            ctmp = "N"
        End If
        ctmp = Trim$(ctmp)
        Text1(27).Text = ctmp
        
        iFehlerstufe = 21
        If Not IsNull(rsrs!ekpr) Then
            dWert = rsrs!ekpr
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "#####0.00")
        Text1(29).Text = ctmp
        
        iFehlerstufe = 22
        If Not IsNull(rsrs!KVKPR1) Then
            dWert = rsrs!KVKPR1
        Else
            dWert = 0
        End If
        ctmp = Format$(dWert, "#####0.00")
        Text1(30).Text = ctmp
        Label9.Caption = ctmp       'Vergleichswert für Änderungen
        
        If LeseSpezpreis(CLng(cSuch), 0) > 0 Then
            Command1(26).BackColor = vbRed
        Else
            Command1(26).BackColor = Command1(9).BackColor
        End If
        
        
        
        iFehlerstufe = 23
        If Not IsNull(rsrs!PREISSCHU) Then
            ctmp = rsrs!PREISSCHU
        Else
            ctmp = "N"
        End If
        ctmp = Trim$(ctmp)
        Text1(31).Text = ctmp
        
        iFehlerstufe = 24
        If Not IsNull(rsrs!BONUS_OK) Then
            ctmp = rsrs!BONUS_OK
        Else
            ctmp = "J"
        End If
        ctmp = Trim$(ctmp)
        Text1(33).Text = ctmp
        
        iFehlerstufe = 241
        If Not IsNull(rsrs!AUFDAT) Then
            If CLng(rsrs!AUFDAT) = 0 Then
                ctmp = ""
            Else
                ctmp = Trim$(rsrs!AUFDAT)
            End If
            
        Else
            ctmp = ""
        End If
        Label4(32).Caption = ctmp
        
        iFehlerstufe = 25
        
        If Not IsNull(rsrs!AWM) Then
            ctmp = rsrs!AWM
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        
        iFehlerstufe = 26
        
        If ctmp <> "" Then
            If ctmp = "98" Then
                Text1(34).Text = ctmp
                cmdfarbe.Caption = "Neu"
            ElseIf ctmp = "95" Then
                Text1(34).Text = ctmp
                cmdfarbe.Caption = "nicht lieferbar"
                cmdfarbe.BackColor = vbBlue
            ElseIf ctmp = "94" Then
                Text1(34).Text = ctmp
                cmdfarbe.Caption = "Preisaktion vorbereitet"
                cmdfarbe.BackColor = glfarbe(0)
            ElseIf ctmp = "93" Then
                Text1(34).Text = ctmp
                cmdfarbe.Caption = "Preisaktion jetzt"
                cmdfarbe.BackColor = vbWhite
            ElseIf ctmp = "92" Then
                Text1(34).Text = ctmp
                cmdfarbe.Caption = "lange nicht verkauft"
                cmdfarbe.BackColor = &H80000012
            Else
                If CByte(ctmp) < 10 Then
                    Text1(34).Text = ctmp
                    cmdfarbe.BackColor = glfarbe(ctmp)
                    cmdfarbe.Caption = ermFarbeBez(ctmp)
                ElseIf CByte(ctmp) > 10 And CByte(ctmp) < 20 Then
                    Text1(34).Text = ctmp
                    cmdfarbe.BackColor = glfarbe2(CInt(ctmp) - 10)
                    cmdfarbe.Caption = ermFarbeBez(CStr(CInt(ctmp)))
                Else
                    Text1(34).Text = "0"
                    cmdfarbe.BackColor = glfarbe(0)
                    cmdfarbe.Caption = ""
                End If
            End If
        Else
            Text1(34).Text = "0"
            cmdfarbe.BackColor = glfarbe(0)
            cmdfarbe.Caption = ""
        End If
        
        If Text1(5).Text <> "" Then
            If IsNumeric(Text1(5).Text) Then
                Label4(37).Caption = ErmlzZugangM(Text1(5).Text)
            End If
        End If

        If Text1(5).Text <> "" Then
            If IsNumeric(Text1(5).Text) Then
                Label4(35).Caption = ErmlzVK(Text1(5).Text)
            End If
        End If
        
        If LeseInterArt(Text1(5).Text) Then
            Check15.value = vbChecked
            Check15.ForeColor = vbRed
        Else
            Check15.value = vbUnchecked
            Check15.ForeColor = vbBlack
        End If
                
        iFehlerstufe = 27
        
        '**********aus ZUORDEAN
        cSQL = "Select gpean ,faktor from ZUORDEAN where ARTNR = " & cSuch & " "
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
        If Not rsrs1.EOF Then
            rsrs1.MoveFirst

            If Not IsNull(rsrs1!GPEAN) Then
                Text2(3).Text = rsrs1!GPEAN
            Else
                Text2(3).Text = ""
            End If

            If Not IsNull(rsrs1!Faktor) Then
                Text2(2).Text = rsrs1!Faktor
            Else
                Text2(2).Text = ""
            End If
        Else
            Text2(2).Text = ""
            Text2(3).Text = ""
        End If
        rsrs1.Close: Set rsrs1 = Nothing

        '**********aus ARTMERK
        Text1(37).Text = ""
        cSQL = "Select MERK from ARTMERK where ARTNR = " & cSuch & " "
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
        If Not rsrs1.EOF Then
            rsrs1.MoveFirst
            If Not IsNull(rsrs1!merk) Then
                Text1(37).Text = rsrs1!merk
            End If
        End If
        rsrs1.Close: Set rsrs1 = Nothing

        '**********aus Stornof
        Text1(40).Text = "J"
        cSQL = "Select MERK from STORNOF where ARTNR = " & cSuch & " "
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
        If Not rsrs1.EOF Then
            rsrs1.MoveFirst
            If Not IsNull(rsrs1!merk) Then
                Text1(40).Text = rsrs1!merk
            End If
        End If
        rsrs1.Close: Set rsrs1 = Nothing

        '**********aus Lagerplatz
        Text2(4).Text = ""
        cSQL = "Select LAGERP from LAGERPLATZ where ARTNR = " & cSuch & " "
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
        If Not rsrs1.EOF Then
            rsrs1.MoveFirst
            If Not IsNull(rsrs1!lagerp) Then
                Text2(4).Text = rsrs1!lagerp
            End If
        End If
        rsrs1.Close: Set rsrs1 = Nothing

        '**********aus KONDITIONEN
        Text2(6).Text = ""
        Text2(5).Text = ""
        cSQL = "Select * from KONDITIONEN where ARTNR = " & cSuch & " "
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
        If Not rsrs1.EOF Then
            rsrs1.MoveFirst

            If Not IsNull(rsrs1!kondi) Then
                Text2(5).Text = rsrs1!kondi
            End If

            If Not IsNull(rsrs1!Faktor) Then
                Text2(6).Text = rsrs1!Faktor
            End If
        End If
        rsrs1.Close: Set rsrs1 = Nothing

        '**********aus Textil
        Text2(7).Text = ""
        Text2(8).Text = ""
        Text2(9).Text = ""
        cSQL = "Select Modell,Material,Farbbez from TEXTIL where ARTNR = " & cSuch & " "
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
        If Not rsrs1.EOF Then
            rsrs1.MoveFirst
            If Not IsNull(rsrs1!Modell) Then
                Text2(7).Text = rsrs1!Modell
            End If

            If Not IsNull(rsrs1!Material) Then
                Text2(8).Text = rsrs1!Material
            End If

            If Not IsNull(rsrs1!Farbbez) Then
                Text2(9).Text = rsrs1!Farbbez
            End If
        End If
        rsrs1.Close: Set rsrs1 = Nothing

         '**********aus Gruppe
        Text2(10).Text = ""
        cSQL = "Select Gruppennr from Gruppe_Artikel where ARTNR = " & cSuch & " "
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
        If Not rsrs1.EOF Then
            rsrs1.MoveFirst
            If Not IsNull(rsrs1!Gruppennr) Then
                Text2(10).Text = rsrs1!Gruppennr
            End If
        End If
        rsrs1.Close: Set rsrs1 = Nothing

        fuellecombo1 cSuch


        iFehlerstufe = 29

        '**********aus INBEST

        Text1(32).Text = erminBestell(cSuch)


        iFehlerstufe = 31

        '*********aus Artlief

        cbo1fuellen cSuch
        Liefdetail cSuch

        '********ArtEAN

        If MehrEAN_vorhanden(CLng(cSuch)) Then
            Label4(49).ForeColor = glWarn
        Else
            Label4(49).ForeColor = glS1
        End If

        If StaffelKVK_vorhanden(CLng(cSuch)) Then
            Command1(34).ForeColor = glWarn
        Else
            Command1(34).ForeColor = glButtonForecolor
        End If

        iFehlerstufe = 39
               
        Label3(2).Caption = 5
        If gbFilNr And gcFilNr <> 0 Then
            Command8.Visible = True
        Else
            Command8.Visible = False
        End If
        
        iFehlerstufe = 40
        
        Frame3.Visible = True
        
        If gbBILDTAST = False Then
            Frame2.Visible = False
        Else
            Frame2.Visible = True
        End If
        
        Frame1.Visible = False
        If gbcomefromwoa = True Then
        
        Else
            Text1(5).SetFocus
        End If
    Else
        Frame3.Visible = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleDatenWKL10"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten. Fehlernummer: " & Trim$(Str$(iFehlerstufe))
    
    Fehlermeldung1
    
    Resume Next 'soll bleiben 01.03.04
    
End Sub
Private Sub Liefdetail(cSuch As String)
On Error GoTo LOKAL_ERROR

    Dim sSpanne         As String
    Dim se              As String
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim ctmp            As String
    Dim dWert           As Double
    Dim lJahr           As Long
    Dim sKalkulierbar   As String
    Dim cLinr           As String
    Dim lcount          As Long
    Dim bgefunden       As Boolean
    
    If cSuch = "" Then
        Exit Sub
    End If
    
    If IsNumeric(cSuch) = False Then
        Exit Sub
    End If
    
    Dim lLinrKleinerLek As Long
    
    If cbo1.Text = "" Then
    
        If glPrimLinr > 0 Then
            bgefunden = False
            For lcount = 0 To cbo1.ListCount - 1
                If cbo1.list(lcount) = glPrimLinr Then
                    bgefunden = True
                    cbo1.Text = cbo1.list(lcount)
                    Exit For
                End If
            Next lcount
       
            If bgefunden = False Then
                If cbo1.ListCount > 0 Then cbo1.Text = cbo1.list(0)
            End If
        
            
        Else
            'oder den, mit dem kleinsten LEK anzeigen
            cSQL = "Select min(LEKPR) as ek,linr  from ARTLIEF where ARTNR = " & cSuch & " and lekpr <> 0 group by linr order by min(lekpr) asc "
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!linr) Then
                   lLinrKleinerLek = rsrs!linr
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
            
            If lLinrKleinerLek > 0 Then
                bgefunden = False
                For lcount = 0 To cbo1.ListCount - 1
                    If cbo1.list(lcount) = lLinrKleinerLek Then
                        bgefunden = True
                        cbo1.Text = cbo1.list(lcount)
                        Exit For
                    End If
                Next lcount
            
                If bgefunden = False Then
                    If cbo1.ListCount > 0 Then cbo1.Text = cbo1.list(0)
                End If
            Else
        
        
        
        
                If cbo1.ListCount > 0 Then cbo1.Text = cbo1.list(0)
            End If
        End If
    End If
            
        cLinr = Trim(cbo1.Text)
        If cLinr <> "" Then
        
            cSQL = "Select * from lisrt where LINR = " & cLinr & " "
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                rsrs.MoveFirst
        
                If Not IsNull(rsrs!LIEFBEZ) Then     'LISRT LIEFBEZ
                    ctmp = rsrs!LIEFBEZ
                Else
                    ctmp = ctmp & Space(1)
                End If
                
                If Not IsNull(rsrs!STADT) Then     'LISRT STADT
                    ctmp = ctmp & Space(1) & rsrs!STADT
                Else
                    ctmp = ctmp & Space(1)
                End If
                
                If Not IsNull(rsrs!strasse) Then     'LISRT STRASSE
                    ctmp = ctmp & Space(1) & rsrs!strasse
                Else
                    ctmp = ctmp
                End If
                
                If Not IsNull(rsrs!Tel) Then     'LISRT telefon
                    ctmp = ctmp & Space(1) & "Tel.: " & rsrs!Tel
                Else
                    ctmp = ctmp
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
        
            lblLiefbez.Caption = ctmp
            cSQL = "Select * from ARTlief where (SYNSTATUS = 'E' or SYNSTATUS = 'A' or SYNSTATUS is null) and  ARTNR = " & cSuch & " and LINR = " & cLinr & " "
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                rsrs.MoveFirst
                
                If Not IsNull(rsrs!MINMEN) Then     'ARTLIEF MINMEN
                    dWert = rsrs!MINMEN
                Else
                    dWert = 0
                End If
                ctmp = Format$(dWert, "#####0")
                Text1(28).Text = ctmp
                
                If Not IsNull(rsrs!LIBESNR) Then    'ARTLIEF LIBESNR
                    ctmp = rsrs!LIBESNR
                Else
                    ctmp = ""
                End If
                ctmp = Trim$(ctmp)
                Text1(19).Text = ctmp
                
                        
                If Not IsNull(rsrs!RKZ) Then
                    ctmp = rsrs!RKZ
                Else
                    ctmp = ""
                End If
                ctmp = Trim$(ctmp)
                Text1(10).Text = ctmp

                
                If Not IsNull(rsrs!EXDAT) Then
                    If CLng(rsrs!EXDAT) = 0 Then
                        ctmp = ""
                    Else
                        ctmp = Trim$(rsrs!EXDAT)
                    End If
                Else
                    ctmp = ""
                End If
                Label4(34).Caption = ctmp
                
                If Not IsNull(rsrs!lekpr) Then  'ARTLIEF LEKPR
                    dWert = rsrs!lekpr
                Else
                    dWert = 0
                End If
                ctmp = Format$(dWert, "#####0.00")
                Text1(11).Text = ctmp
                
                If Not IsNull(rsrs!SPANNE) Then  'ARTLIEF SPANNE
                    dWert = rsrs!SPANNE
                Else
                    dWert = 0
                End If
                
                If dWert = 0 Then 'autoKalk = no
                
                    Label4(30).Visible = False
                    Label4(30).ForeColor = glS1
                    Label4(30).Refresh
                    
                    'Nettospanne errechnen
                    
                    Dim cKVKN   As String
                    Dim cek     As String
                    Dim cMwst   As String
                    
                    
                    If gsSpanne = "LEK" Then    'basierend auf LEK oder SEK
                    
                        cek = Text1(11).Text 'ermMaxLEKPR(cSuch)
                        
                    ElseIf gsSpanne = "SEK" Then
                        cek = Text1(29).Text
                    End If
                    
                    cMwst = Text1(15).Text
                    cKVKN = Text1(30).Text
                    Text1(14).Text = NettospanneInProzent(cKVKN, cek, cMwst)
                    
                    Handelsspanne_anzeigen cSuch, cMwst, cKVKN, cLinr
                
                Else 'autoKalk = yes
                    Label4(30).Visible = True
                    Label4(30).ForeColor = vbRed
                    Label4(30).Refresh
                    ctmp = Format$(dWert, "###0.00")
                    Text1(14).Text = ctmp
                End If
                
            End If
            rsrs.Close: Set rsrs = Nothing
            
        Else
            lblLiefbez.Caption = "Geben Sie bitte die Lieferantendaten für diesen Artikel ein!"
            lblLiefbez.Refresh
            
            Exit Sub
        End If
        
        If LeseStaffelpreis(CLng(cSuch), CLng(cbo1.Text)) Then
            Command1(27).BackColor = vbRed
        Else
            Command1(27).BackColor = Command1(9).BackColor
        End If
        
        
        
        Dim linBV As Long
            
        linBV = 0
        linBV = ermINBV(cSuch, cbo1.Text)
        
        If linBV > 0 Then
            Command5(5).Caption = "in BV(" & linBV & ")"
            Command5(5).ForeColor = vbRed
        Else
            Command5(5).Caption = "in BV"
            Command5(5).ForeColor = glS1
        End If
            
        
        
        
        
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Liefdetail"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub cbo1fuellen(cSuch As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cbo1.Clear
    
    cSQL = "Select * from ARTlief where (SYNSTATUS = 'E' or SYNSTATUS = 'A' or SYNSTATUS is null) and ARTNR = " & cSuch & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        If rsrs.RecordCount > 1 Then
            Label4(2).ForeColor = glWarn
        Else
            Label4(2).ForeColor = glS1
        End If
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!linr) Then
                cbo1.AddItem rsrs!linr
            Else
                cbo1.AddItem ""
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbo1fuellen"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeereArtikelDatenWKL10(bArtnrauch As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    If bArtnrauch Then
        Text1(5).Text = ""
    End If
    
    Text1(6).Text = ""
    
    For lcount = 8 To 34
        Text1(lcount).Text = ""
    Next lcount
    
    Text1(37).Text = ""
    Text1(38).Text = ""
    Text1(39).Text = ""
    Text1(40).Text = ""
    Text1(43).Text = ""
    
    cbo1.Text = ""
    cbo1.Clear
    
    cmdfarbe.BackColor = glfarbe(0)
    Label4(49).ForeColor = glS1
    Command1(34).ForeColor = glButtonForecolor
    
    Label6.Caption = ""
    Label9.Caption = ""
    Label4(32).Caption = ""
    Label4(34).Caption = ""
    lblLiefbez.Caption = ""
    
    Label4(40).Caption = ""
    Label4(41).Caption = ""
    Label4(42).Caption = ""
    Label10.Caption = ""
    
'    Label4(2).ForeColor = vbBlack
'    Label4(6).ForeColor = vbBlack
'    Label4(14).ForeColor = vbBlack
'    Label4(23).ForeColor = vbBlack
    
    Text2(3).Text = ""
    Text2(2).Text = ""
    
    Text2(4).Text = ""
    Text2(5).Text = ""
    Text2(6).Text = ""
    
    Text2(7).Text = ""
    Text2(8).Text = ""
    Text2(9).Text = ""
    
    Text2(10).Text = ""
    
    Label4(39).Caption = ""
    Label4(39).Visible = False
    
    giDlgZustand = giNEU
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereArtikelDatenWKL10"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeereDialogWKL10()
    On Error GoTo LOKAL_ERROR
    
    Text1(36).Text = ""
    Text1(1).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    Text1(0).Text = ""
    Text1(2).Text = ""
    Text1(7).Text = ""
    Text1(35).Text = ""
    Text1(41).Text = ""
    Text1(42).Text = ""
    Text1(44).Text = ""
    Text1(45).Text = ""
    Text1(46).Text = ""
    Text1(47).Text = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKL10"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SchreibeDatenWKL10()
    On Error GoTo LOKAL_ERROR
    
    Dim lBestandAlt As Long
    Dim lBewegung   As Long
    Dim lBestandneu As Long
    Dim lMinMen     As Long
    Dim lHeute      As Long
    Dim cSQL        As String
    Dim cKey        As String
    Dim ctmp        As String
    Dim cArtNr      As String
    Dim cBezeich    As String
    Dim cLinr       As String
    Dim cEAN        As String
    Dim cEAN2       As String
    Dim cEAN3       As String
    Dim cLiBesNr    As String
    Dim cJetzt      As String
    Dim dLEKPR      As Double
    Dim dWert       As Double
    Dim dBWert      As Double
    Dim bLeeren     As Boolean
    Dim bTrans      As Boolean
    Dim bNeu        As Boolean
    Dim bEtikett    As Boolean
    Dim bhatsichBestgeaendert       As Boolean
    Dim bEAN        As Boolean
    Dim bEAN2       As Boolean
    Dim bEAN3       As Boolean
    Dim bKVKPR1     As Boolean
    Dim rsrs        As Recordset
    Dim rsHis       As Recordset
    Dim rsArt       As Recordset
    Dim rsEti       As Recordset
    Dim rsZ         As Recordset
    Dim sSQLz       As String
    Dim i                       As Integer
    Dim iArtAnzahl              As Integer
    Dim bArtLiefInsert          As Boolean
    Dim bArtLiefUpdate          As Boolean
    Dim iZBestand(1 To 20)       As Integer
    
    bLeeren = True
    bTrans = False
    bNeu = False
    bEtikett = False
    bEAN = False
    bEAN2 = False
    bEAN3 = False
    bKVKPR1 = False
    bhatsichBestgeaendert = False
    
    cKey = Text1(5).Text
    cKey = Trim$(cKey)
    
    If cKey = "" Then
        MsgBox "Schreiben nicht möglich, da keine Artikelnummer vorhanden!", vbCritical, "STOP!"
        Text1(5).SetFocus
        Exit Sub
    End If
    
    cKey = String$(6 - Len(cKey), "0") & cKey
    Text1(5).Text = cKey
    
    
    ctmp = Trim$(Text1(13).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    dBWert = Val(ctmp)
    
    ctmp = Trim$(Text1(30).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    dWert = Val(ctmp)
    
    '** Änderung bei KVKPR1 **
    If Trim$(Label9.Caption) <> Trim$(Text1(30).Text) Then
        If (glBestandNeu > glBestandAlt) Or (glBestandNeu < glBestandAlt) Then
        '** BESTAND und KVKPR1 haben sich geändert **
            If gcFilNr = "1" Then
                For i = 1 To giAnzFil
                    If i <> 1 Then
                        sSQLz = "Select * from zbestand where filialnr = " & i
                        sSQLz = sSQLz & " and artnr = " & cKey
                        Set rsZ = gdBase.OpenRecordset(sSQLz)

                        If Not rsZ.EOF Then
                            If Not IsNull(rsZ!BESTAND) Then
                                iZBestand(i) = rsZ!BESTAND
                            Else
                                iZBestand(i) = 0
                            End If
                        Else
                            iZBestand(i) = 0
                        End If
                        rsZ.Close: Set rsZ = Nothing
                    End If
    
                    cSQL = "Select * from ETIDRU where ARTNR = " & cKey
                    cSQL = cSQL & " and FILNR = " & i
                    Set rsEti = gdBase.OpenRecordset(cSQL)
                    bEtikett = True
                
                    If Not rsEti.EOF Then
                        rsEti.Edit
                    Else
                        If iZBestand(i) <> 0 Or dWert <> 0 Then
                            rsEti.AddNew
                        Else
                            rsEti.Close: Set rsEti = Nothing
                            GoTo weiter2
                        End If
                    End If
                    rsEti!artnr = cKey
                    rsEti!BEZEICH = Trim$(Text1(6).Text)
                    rsEti!vkpr = dWert
                    dWert = 0
                    
                    If i = 1 Then
                        If Not IsNull(rsEti!BESTAND) Then
                            rsEti!BESTAND = dBWert + (rsEti!BESTAND)
                        Else
                            rsEti!BESTAND = dBWert
                        End If
                        If Not IsNull(rsEti!ANZAHL) Then
                            rsEti!ANZAHL = dBWert + (rsEti!ANZAHL)
                        Else
                            rsEti!ANZAHL = dBWert
                        End If
                    Else
                        If Not IsNull(rsEti!BESTAND) Then
                            If iZBestand(i) <> 0 Then rsEti!BESTAND = (rsEti!BESTAND) + iZBestand(i)
                        Else
                            If iZBestand(i) <> 0 Then rsEti!BESTAND = iZBestand(i)
                        End If
                        If Not IsNull(rsEti!ANZAHL) Then
                            If iZBestand(i) <> 0 Then rsEti!ANZAHL = (rsEti!ANZAHL) + iZBestand(i)
                        Else
                            If iZBestand(i) <> 0 Then rsEti!ANZAHL = iZBestand(i)
                        End If
                    End If
                    rsEti!LIBESNR = Trim$(Text1(19).Text)
                    rsEti!EAN = Trim$(Text1(18).Text)
                    rsEti!linr = Trim$(cbo1.Text)
                    rsEti!LPZ = Val(Trim$(Text1(8).Text))
                    rsEti!filnr = i
                    rsEti!Pcname = srechnertab
                    rsEti.Update
                    rsEti.Close: Set rsEti = Nothing
weiter2:
                Next i
            Else '** gcFilNr <> "1" **
                cSQL = "Select * from ETIDRU where ARTNR = " & cKey
                cSQL = cSQL & " and FILNR = " & gcFilNr
                Set rsEti = gdBase.OpenRecordset(cSQL)
                bEtikett = True
                If bEtikett = True Then
                    If Not rsEti.EOF Then
                        rsEti.Edit
                    Else
                        rsEti.AddNew
                    End If
                    rsEti!artnr = cKey
                    rsEti!BEZEICH = Trim$(Text1(6).Text)
                    
                    ctmp = Trim$(Text1(30).Text)
                    ctmp = fnMoveComma2Point$(ctmp)
                    dWert = Val(ctmp)
                    rsEti!vkpr = dWert
                    
                    ctmp = Trim$(Text1(13).Text)
                    ctmp = fnMoveComma2Point$(ctmp)
                    dWert = Val(ctmp)
                    rsEti!BESTAND = dWert
                    rsEti!ANZAHL = dWert
                    rsEti!LIBESNR = Trim$(Text1(19).Text)
                    rsEti!EAN = Trim$(Text1(18).Text)
                    rsEti!linr = Trim$(cbo1.Text)
                    rsEti!LPZ = Val(Trim$(Text1(8).Text))
                    rsEti!filnr = Val(gcFilNr)
                    rsEti!Pcname = srechnertab
                    rsEti.Update
                    rsEti.Close: Set rsEti = Nothing
                End If
            End If
        Else '**glBestandNeu = glBestandAlt, neue KVKPR1
            If gcFilNr = "1" Then
                For i = 1 To giAnzFil
                    If i <> 1 Then
                        iZBestand(i) = ermBestandfromZbestand(cKey, i)
                    End If
    
                    cSQL = "Select * from ETIDRU where ARTNR = " & cKey
                    cSQL = cSQL & " and FILNR = " & i
                    Set rsEti = gdBase.OpenRecordset(cSQL)
                    bEtikett = True
                
                    If Not rsEti.EOF Then
                        rsEti.Edit
                    Else
                        ctmp = Trim$(Text1(30).Text)
                        ctmp = fnMoveComma2Point$(ctmp)
                        dWert = Val(ctmp)
                        If iZBestand(i) <> 0 Or dWert <> 0 Then
                            rsEti.AddNew
                        Else
                            rsEti.Close: Set rsEti = Nothing
                            GoTo weiter3
                        End If
                    End If
                    rsEti!artnr = cKey
                    rsEti!BEZEICH = Trim$(Text1(6).Text)
                    rsEti!vkpr = dWert
                    dWert = 0
                    
                    If i = 1 Then
                        If Not IsNull(rsEti!BESTAND) Then
                            rsEti!BESTAND = dBWert + (rsEti!BESTAND)
                        Else
                            rsEti!BESTAND = dBWert
                        End If
                        If Not IsNull(rsEti!ANZAHL) Then
                            rsEti!ANZAHL = dBWert + (rsEti!ANZAHL)
                        Else
                            rsEti!ANZAHL = dBWert
                        End If
                    Else
                        If Not IsNull(rsEti!BESTAND) Then
                            If iZBestand(i) <> 0 Then rsEti!BESTAND = (rsEti!BESTAND) + iZBestand(i)
                        Else
                            If iZBestand(i) <> 0 Then rsEti!BESTAND = iZBestand(i)
                        End If
                        If Not IsNull(rsEti!ANZAHL) Then
                            If iZBestand(i) <> 0 Then rsEti!ANZAHL = (rsEti!ANZAHL) + iZBestand(i)
                        Else
                            If iZBestand(i) <> 0 Then rsEti!ANZAHL = iZBestand(i)
                        End If
                    End If
                    rsEti!LIBESNR = Trim$(Text1(19).Text)
                    rsEti!EAN = Trim$(Text1(18).Text)
                    rsEti!linr = Trim$(cbo1.Text)
                    rsEti!LPZ = Val(Trim$(Text1(8).Text))
                    rsEti!filnr = i
                    rsEti!Pcname = srechnertab
                    rsEti.Update
                    rsEti.Close: Set rsEti = Nothing
weiter3:
                Next i
            Else '** gcFilNr <> "1" **
                cSQL = "Select * from ETIDRU where ARTNR = " & cKey
                cSQL = cSQL & " and FILNR = " & gcFilNr
                Set rsEti = gdBase.OpenRecordset(cSQL)
                bEtikett = True
                If bEtikett = True Then
                    If Not rsEti.EOF Then
                        rsEti.Edit
                    Else
                        rsEti.AddNew
                    End If
                    rsEti!artnr = cKey
                    rsEti!BEZEICH = Trim$(Text1(6).Text)
                    
                    ctmp = Trim$(Text1(30).Text)
                    ctmp = fnMoveComma2Point$(ctmp)
                    dWert = Val(ctmp)
                    rsEti!vkpr = dWert
                    
                    ctmp = Trim$(Text1(13).Text)
                    ctmp = fnMoveComma2Point$(ctmp)
                    dWert = Val(ctmp)
                    If Not IsNull(rsEti!BESTAND) Then
                        rsEti!BESTAND = dWert
                    Else
                        rsEti!BESTAND = dWert
                    End If
                    If Not IsNull(rsEti!ANZAHL) Then
                        rsEti!ANZAHL = dWert
                    Else
                        rsEti!ANZAHL = dWert
                    End If
                    rsEti!LIBESNR = Trim$(Text1(19).Text)
                    rsEti!EAN = Trim$(Text1(18).Text)
                    rsEti!linr = Trim$(cbo1.Text)
                    rsEti!LPZ = Val(Trim$(Text1(8).Text))
                    rsEti!filnr = Val(gcFilNr)
                    rsEti!Pcname = srechnertab
                    rsEti.Update
                    rsEti.Close: Set rsEti = Nothing
                End If
            End If
        End If
        
    Else '** nur BESTAND hat sich geändert
        If glBestandNeu > glBestandAlt Then
            cSQL = "Select * from ETIDRU where ARTNR = " & cKey
            cSQL = cSQL & " and FILNR = " & gcFilNr
            Set rsEti = gdBase.OpenRecordset(cSQL)
            bEtikett = True
            If bEtikett = True Then
                If Not rsEti.EOF Then
                    rsEti.Edit
                Else
                    rsEti.AddNew
                End If
                rsEti!artnr = cKey
                rsEti!BEZEICH = Trim$(Text1(6).Text)
                
                ctmp = Trim$(Text1(30).Text)
                ctmp = fnMoveComma2Point$(ctmp)
                dWert = Val(ctmp)
                rsEti!vkpr = dWert
                
'                ctmp = Trim$(Text1(13).Text)
'                ctmp = fnMoveComma2Point$(ctmp)
'                dWert = Val(ctmp)
                
                dWert = glBestandNeu - glBestandAlt
                
                If Not IsNull(rsEti!BESTAND) Then
                    rsEti!BESTAND = rsEti!BESTAND + dWert
                Else
                    rsEti!BESTAND = dWert
                End If
                If Not IsNull(rsEti!ANZAHL) Then
                    rsEti!ANZAHL = rsEti!ANZAHL + dWert
                Else
                    rsEti!ANZAHL = dWert
                End If
                rsEti!LIBESNR = Trim$(Text1(19).Text)
                rsEti!EAN = Trim$(Text1(18).Text)
                rsEti!linr = Trim$(cbo1.Text)
                rsEti!LPZ = Val(Trim$(Text1(8).Text))
                rsEti!filnr = Val(gcFilNr)
                rsEti!Pcname = srechnertab
                rsEti.Update
                rsEti.Close: Set rsEti = Nothing
            End If
        Else '** nur BESTAND hat sich geändert aber BestandNeu < BestandAlt
        End If
    End If
    
    
    cSQL = "Select * from ARTIKEL where ARTNR = " & cKey & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If giDlgZustand = giNEU Then
            MsgBox "Artikelnummer bereits vorhanden! Bitte andere Artikelnummer eingeben.", vbCritical, "DOPPELTE ARTIKELNUMMER"
            Text1(5).SetFocus
            Exit Sub
        End If
        rsrs.Edit
        rsrs!SYNStatus = "E"
        bNeu = False
    Else
        rsrs.AddNew
        rsrs!SYNStatus = "A"
        rsrs!AUFDAT = DateValue(Now)
        bNeu = True
        bLeeren = False
    End If
    
    cArtNr = Trim$(Text1(5).Text)
    cBezeich = Trim$(Text1(6).Text)
    cLinr = Trim$(cbo1.Text)
    cEAN = Trim$(Text1(18).Text)
    
    
    
    If bNeu Then
        cSQL = "Select * from ARTLIEF where ARTNR = -1"
        Set rsArt = gdBase.OpenRecordset(cSQL)
        bArtLiefInsert = True
    Else
        cSQL = "Select * from ARTLIEF where ARTNR = " & cArtNr & " and LINR = " & cLinr
        Set rsArt = gdBase.OpenRecordset(cSQL)
        If Not rsArt.EOF Then
            bArtLiefInsert = False
        Else
            bArtLiefInsert = True
        End If
    End If
    
    If cKey = "" Then
        rsrs!artnr = Null
    Else
        rsrs!artnr = Trim$(Text1(5).Text)
    End If
    rsrs!BEZEICH = Trim$(Text1(6).Text)
    
    If gbNew Then
        ctmp = Trim$(cbo1.Text)
        If ctmp = "" Then
            rsrs!linr = Null
        Else
            rsrs!linr = ctmp
        End If
    Else
        If cbo1.ListCount = 0 And cLinr <> "" Then
            rsrs!linr = cLinr
        End If
    End If
    
    ctmp = Trim$(Text1(8).Text)
    If ctmp = "" Then
        rsrs!LPZ = 1
    Else
        rsrs!LPZ = Val(ctmp)
    End If
    
    ctmp = Trim$(Text1(9).Text)
    If ctmp = "" Then
        rsrs!AGN = Null
    Else
        rsrs!AGN = ctmp
    End If
    
    ctmp = Trim$(Text1(39).Text)
    If ctmp = "" Then
        rsrs!PGN = Null
    Else
        rsrs!PGN = ctmp
    End If
    

    
    ctmp = Trim$(Text1(11).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    dWert = Val(ctmp)
    rsrs!lekpr = dWert
    dLEKPR = dWert
    
    ctmp = Trim$(Text1(12).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    dWert = Val(ctmp)
    rsrs!vkpr = dWert
    
    If Text1(30).Text = "" Then Text1(30).Text = 0
    
    'Hat sich der KVKPR1 geändert
    If Not IsNull(rsrs!KVKPR1) Then
        If CDbl(Trim(CStr(rsrs!KVKPR1))) <> CDbl(Trim$(Text1(30).Text)) Then
            bKVKPR1 = True
        End If
    Else
        bKVKPR1 = True
    End If
    
    'Hat sich der Bestand geändert
    If Not IsNull(rsrs!BESTAND) Then
        If Trim(CStr(rsrs!BESTAND)) <> Trim$(Text1(13).Text) Then
            bhatsichBestgeaendert = True
        End If
    Else
        If Trim$(Text1(13).Text) <> "" Then
            bhatsichBestgeaendert = True
        End If
    End If
    
    rsrs!MWST = Trim$(Text1(15).Text)
    rsrs!NOTIZEN = Trim$(Text1(16).Text)
    rsrs!GROESSE = Trim$(Text1(43).Text)
    rsrs!LIBESNR = Trim$(Text1(19).Text)
    cLiBesNr = Trim$(Text1(19).Text)
    
    'Hat sich der EAN(1) geändert
    If Not IsNull(rsrs!EAN) Then
        If Trim(CStr(rsrs!EAN)) <> Trim$(Text1(18).Text) Then
            bEAN = True
        End If
    Else
        bEAN = True
    End If
    
    'Hat sich der EAN2 geändert
    If Not IsNull(rsrs!EAN2) Then
        If Trim(CStr(rsrs!EAN2)) <> Trim$(Text1(20).Text) Then
            bEAN2 = True
        End If
    Else
        bEAN2 = True
    End If
    
    'Hat sich der EAN3 geändert
    If Not IsNull(rsrs!EAN3) Then
        If Trim(CStr(rsrs!EAN3)) <> Trim$(Text1(21).Text) Then
            bEAN3 = True
        End If
    Else
        bEAN3 = True
    End If
 
    ctmp = Trim$(Text1(22).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    dWert = Val(ctmp)
    rsrs!INHALT = dWert
    rsrs!INHALTBEZ = Trim$(Text1(23).Text)
    rsrs!GRUNDPREIS = Trim$(Text1(24).Text)
    rsrs!MINBEST = Val(Trim$(Text1(25).Text))
    rsrs!RABATT_OK = Trim$(Text1(26).Text)
    rsrs!UMS_OK = Trim$(Text1(38).Text)
    
    
    If rsrs!GEFUEHRT <> Trim$(Text1(27).Text) Then
        
        If gbKL_LIVEGefSperr = True Then
                
            Dim bSperre As Boolean
            
            If Trim$(Text1(27).Text) = "N" Then
                bSperre = True
            ElseIf Trim$(Text1(27).Text) = "J" Then
                bSperre = False
            End If
            
            If live_ArtikelSperren_updaten(cArtNr, CInt(gcFilNr), bSperre) = True Then
            
                If Trim$(Text1(27).Text) = "N" Then
                    MsgBox "Das Sperrmerkmal wurde im Kisslive gesetzt.", vbInformation, "Winkiss Hinweis:"
                ElseIf Trim$(Text1(27).Text) = "J" Then
                    MsgBox "Das Sperrmerkmal wurde im Kisslive aufgehoben.", vbInformation, "Winkiss Hinweis:"
                End If
                
            End If
        End If

        schreibeProtokollAWMablauf " " & cArtNr & Space(8 - Len(cArtNr)) & ermBezeichausWGN(cArtNr) & " neue Farbe: " & Trim$(Text1(34).Text) & " Bediener: " & gcBedienerNr & " Artikel bea EinzelM"
    End If
    
    
    
    
    rsrs!GEFUEHRT = Trim$(Text1(27).Text)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    rsrs!MINMEN = Val(Trim$(Text1(28).Text))
    lMinMen = Val(Trim$(Text1(28).Text))
    
    ctmp = Trim$(Text1(29).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    dWert = Val(ctmp)
    rsrs!ekpr = dWert
    
    rsrs!PREISSCHU = Trim$(Text1(31).Text)
    rsrs!BONUS_OK = Trim$(Text1(33).Text)
    
    If IsNull(rsrs!UMS_OK) Then
        rsrs!UMS_OK = "J"
    End If
    
    If Text1(34).Text = "" Then
        rsrs!AWM = "0"
    Else
        If rsrs!AWM <> Trim$(Text1(34).Text) Then
        
            If gbKL_LIVEFarbe = True Then
                If live_Artikelfarbe_updaten(cArtNr, CInt(Trim$(Text1(34).Text))) = True Then
                    MsgBox "Diese Artikelfarbe wurde im Kisslive gespeichert.", vbInformation, "Winkiss Hinweis:"
                End If
            End If
    
            schreibeProtokollAWMablauf " " & cArtNr & Space(8 - Len(cArtNr)) & ermBezeichausWGN(cArtNr) & " neue Farbe: " & Trim$(Text1(34).Text) & " Bediener: " & gcBedienerNr & " Artikel bea EinzelM"
        End If
    
    
        rsrs!AWM = Trim$(Text1(34).Text)
    End If
    
    
    cSQL = "Select * from ZUGANG where ARTNR = -1"
    Set rsHis = gdBase.OpenRecordset(cSQL)
    
    lHeute = Fix(Now)
    cJetzt = Format$(Now, "HH:MM")
    
    If bhatsichBestgeaendert Then
        
        rsHis.AddNew
        rsHis!artnr = Val(cArtNr)
        rsHis!BEZEICH = cBezeich
        rsHis!linr = Val(cLinr)
        rsHis!EAN = cEAN
        rsHis!ADATE = lHeute
        rsHis!Uhrzeit = cJetzt
        If gcBedienerNr = "" Then gcBedienerNr = "99"
        rsHis!BEDNU = gcBedienerNr
        rsHis!bedname = gcUserName
        rsHis!FILIALNR = 1
        rsHis!bestandalt = glBestandAlt
        rsHis!BEWEGUNG = glBestandNeu - glBestandAlt
        rsHis!BESTANDneu = glBestandNeu
        
        rsHis!rek = dLEKPR
        rsHis.Update
    End If
    
    ctmp = Trim$(Text1(17).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    dWert = Val(ctmp)
    
    UpdateLastREK Trim$(Text1(5).Text), dWert

    If bArtLiefInsert Then
        rsArt.AddNew
        rsArt!SYNStatus = "A"
    Else
        rsArt.Edit
        rsArt!SYNStatus = "E"
    End If
    rsArt!artnr = Val(cArtNr)
    rsArt!linr = Val(cLinr)
    rsArt!lekpr = dLEKPR
    rsArt!LIBESNR = cLiBesNr
    rsArt!MINMEN = lMinMen
    
    If Label4(30).Visible = True Then
        ctmp = Trim$(Text1(14).Text)
        ctmp = fnMoveComma2Point$(ctmp)
        dWert = Val(ctmp)
        rsArt!SPANNE = dWert
    Else
        dWert = 0
        rsArt!SPANNE = dWert
    End If
    
    
    
    'RKZ check
    
    rsArt!RKZ = Trim$(Text1(10).Text)
    If rsArt!RKZ = "J" Then
        If Not IsNull(rsArt!EXDAT) Then
        
            If CLng(rsArt!EXDAT) = 0 Then
                rsArt!EXDAT = DateValue(Now)
            Else
                
            End If
        Else
            rsArt!EXDAT = DateValue(Now)
        End If
    Else
        rsArt!EXDAT = 0
    End If
    
    
    
    rsArt.Update
    rsArt.Close: Set rsArt = Nothing
    
    IstdieLinrinUeberli cLinr, cArtNr, cLiBesNr, dLEKPR, lMinMen
    
    bNeu = False
    bArtLiefInsert = False
    
    BeginTrans
    bTrans = True
    
    rsrs!LASTDATE = DateValue(Now)
    rsrs!LASTTIME = TimeValue(Now)
    rsrs.Update

    
    
    CommitTrans
    rsHis.Close: Set rsHis = Nothing
    rsrs.Close: Set rsrs = Nothing
    
    If bhatsichBestgeaendert Then
        ctmp = Trim$(Text1(13).Text)
        ctmp = fnMoveComma2Point$(ctmp)
        dWert = Val(ctmp)
        Bestandsveraenderung cArtNr, CLng(dWert), "Artikelbearbeitung"
    End If
    
    If bEAN Then
        ctmp = Trim$(Text1(18).Text)
        Artikelveraenderung cArtNr, ctmp, "Artikelbearbeitung", "EAN"
    End If
    
    If bEAN2 Then
        ctmp = Trim$(Text1(20).Text)
        Artikelveraenderung cArtNr, ctmp, "Artikelbearbeitung", "EAN2"
    End If
    
    If bEAN3 Then
        ctmp = Trim$(Text1(21).Text)
        Artikelveraenderung cArtNr, ctmp, "Artikelbearbeitung", "EAN3"
    End If
    
    If bKVKPR1 Then
        ctmp = Trim$(Text1(30).Text)
        Artikelveraenderung cArtNr, ctmp, "Artikelbearbeitung", "KVKPR1"
    End If
    
    speichernGP1
    speichernLAGERP
    speichernTEXTIL
    speichernGruppe
    
    speichernMerkmal cArtNr, Text1(37).Text
    
    If UCase(Trim(Text1(40).Text)) = "N" Then
        speichernStornof cArtNr, Text1(40).Text
    End If
    
    delKONDITIONEN Text1(5).Text, Text2(5).Text
    speichernKONDITIONEN
    

    
    
    
    
    
    If bLeeren Then
        LeereArtikelDatenWKL10 True
    Else
        Label5(0).Visible = True
        anzeige "LASER", "Der Artikel ist gespeichert", Label5(0)
        Pause (1)
       
'        MsgBox "Artikel gespeichert!", vbInformation, "INFO"
    End If
    
    Text1(5).SetFocus
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWKL10"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Sub
Private Function SucheArtikelWKL10(cTeil As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim cFeld       As String
    Dim cLinr       As String
    Dim cwhere      As String
    Dim iRet        As Integer
    Dim cEAN        As String
    Dim cArtNr      As String
    Dim cEigNr      As String
    Dim bytePGN     As Byte
    Dim iStufe      As Integer
    Dim cJoin       As String
    Dim cMarke      As String
    Dim lcount      As Long
    Dim llpzvon     As Double
    Dim llpzbis     As Double
    Dim siAnzeige   As Single
    Dim i           As Integer
    
    Dim dLagerwertzumSEK As Double
    Dim dPennerwertzumSEK As Double
    Dim dPennerAnteilSEK As Double
    Dim dPennerAnteilST As Double
    Dim lLagerST As Long
    Dim lPennerST As Long
    Dim dEINKaufswert As Double
    Dim dEINKaufswertvj As Double
    Dim dUmsBraktJahr As Double
    Dim dUmsBrvorJahr As Double
    Dim dUmsSEKaktJahr As Double
    Dim dUmsSEKvorJahr As Double
    Dim dUms12M As Double
    Dim dUms12MVJZR As Double
    Dim dUms12MDIFFabs As Double
    Dim dUms12MDIFFrela As Double
    Dim dUmsSEK12M As Double
    Dim dUmsSEK12MVJZR As Double
    Dim dUmsSEK12MDIFFabs As Double
    Dim dUmsSEK12MDIFFrela As Double
    
    Dim l_VKM_aktJahr As Long
    Dim l_VKM_vorJahr As Long
    
    Dim l_VKM_12M As Long
    Dim l_VKM_12MVJZR As Long
    
    Dim bymonat As Byte
    Dim iJahr As Integer
    Dim lMax As Long
    Dim j As Integer
    
    cMarke = Trim$(Text1(2).Text)
    
    llpzvon = Val(Text1(41).Text)
    llpzbis = Val(Text1(42).Text)
    
    If llpzvon > 0 And llpzbis = 0 Then
        llpzbis = llpzvon
    End If
    
    SucheArtikelWKL10 = False

    loeschNEW "TOP" & srechnertab, gdBase
    
    iRet = fnPruefeEingabeWKL10()
    If iRet <> 0 Then
        anzeige "rot", "Suchkriterium!", Label0(4)
        Text1(1).SetFocus
        Exit Function
    End If

    Screen.MousePointer = 11
    
    anzeige "normal", "", Label1(59)
    
    picprogress.Visible = True
    txtStatus.Text = 20
    
    Frame1.Visible = False
    
    cSQL = "Select distinct A.ARTNR "
    cSQL = cSQL & ", A.Bezeich"
    cSQL = cSQL & ", A.Bestand"
    cSQL = cSQL & ", A.KVKPR1"
    cSQL = cSQL & ", A.VKPR"
    cSQL = cSQL & ", A.PGN"
    cSQL = cSQL & ", A.EKPR"
    cSQL = cSQL & ", A.EAN"
    cSQL = cSQL & ", A.EAN2"
    cSQL = cSQL & ", A.EAN3"
    cSQL = cSQL & ", A.LPZ"
    cSQL = cSQL & ", A.gefuehrt"
    cSQL = cSQL & ", A.MINBEST"
    cSQL = cSQL & ", A.NOTIZEN"
    cSQL = cSQL & ", A.AGN"
    cSQL = cSQL & ", A.AWM"
    cSQL = cSQL & ", A.GROESSE"
    
    cSQL = cSQL & ", b.EXDAT"
    cSQL = cSQL & ", b.RKZ"
    cSQL = cSQL & ", b.MINMEN"
    cSQL = cSQL & ", b.LINR"
    cSQL = cSQL & ", b.LEKPR "
    cSQL = cSQL & ", b.LIBESNR"
    
    cSQL = cSQL & ", A.Preisschu"
    cSQL = cSQL & ", A.Rabatt_OK"
    cSQL = cSQL & ", A.BONUS_OK"
    cSQL = cSQL & ", '' as Merk "
    cSQL = cSQL & ", 0 as SpezKVK "
    cSQL = cSQL & ", VKPR as LUG "
    cSQL = cSQL & ", 0.0 as LUGD "
    cSQL = cSQL & ", 0.0 as LRW "
    cSQL = cSQL & ", 'N' as SHOP "
    cSQL = cSQL & ", '' as MODELL "
    cSQL = cSQL & ", '' as MATERIAL "
    cSQL = cSQL & ", '' as FARBBEZ "
    cSQL = cSQL & ", 0 as GRUPPENNR "
    cSQL = cSQL & ", VKPR as HS "
    cSQL = cSQL & ", A.MWST "

    If llpzvon > 0 Then
        cSQL = cSQL & ", c.Lagerp"
    Else
        cSQL = cSQL & ", VKPR as Lagerp"
    End If
    
    cLinr = Trim(Text1(7).Text)
    
    If cLinr <> "" Then
        If IsNumeric(cLinr) Then
            cJoin = " and A.ARTNR = B.ARTNR "
        Else
            cJoin = " and A.ARTNR = B.ARTNR " 'and A.LINR = B.LINR"
        End If
    Else
        cJoin = " and A.ARTNR = B.ARTNR " 'and A.LINR = B.LINR "
    End If
    
    If cMarke <> "" Then
        If Datendrin("MA" & srechnertab, gdBase) Then
            cSQL = cSQL & " into Top" & srechnertab & " from ARTIKEL A, ARTLIEF B " ', MA" & srechnertab & " D"
            cJoin = cJoin & " and a.artnr in (Select artnr from MA" & srechnertab & ") "
        Else
            cSQL = cSQL & " into Top" & srechnertab & " from ARTIKEL A, ARTLIEF B "
        End If
    Else
        cSQL = cSQL & " into Top" & srechnertab & " from ARTIKEL A, ARTLIEF B "
        
    End If
    
    If llpzvon > 0 Then
        cSQL = cSQL & " , lagerplatz c where A.artnr = c.artnr"
        cwhere = " and "
    Else
        cwhere = " Where "
    End If
    
    cwhere = cwhere & " ( A.SYNSTATUS = 'E' or A.SYNSTATUS = 'A' or A.SYNSTATUS is null )"
    cwhere = cwhere & " and ( B.SYNSTATUS = 'E' or B.SYNSTATUS = 'A' or B.SYNSTATUS is null ) "
    
    Dim lVon As Long
    Dim lBis As Long
    
    lVon = datumwandlung(DateValue(Now) - 365)
    lBis = datumwandlung(DateValue(Now))
    
    If Check5.value = vbChecked Then
        cwhere = cwhere & " and a.artnr not in (Select artnr from kassjour where ADATE between " & lVon & " and " & lBis & "   ) "
    End If
    
    If llpzvon > 0 Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & " c.lagerp between " & llpzvon & " and " & llpzbis & " "
    End If
    
    If Text1(36).Text <> "" Then
        cFeld = LTrim(Text1(36).Text)
    
'        cFeld = SwapStr(cFeld, "     ", "*")
'        cFeld = SwapStr(cFeld, "    ", "*")
'        cFeld = SwapStr(cFeld, "   ", "*")
'        cFeld = SwapStr(cFeld, "  ", "*")
'        cFeld = SwapStr(cFeld, " ", "*")

        cFeld = SwapStr(cFeld, "???", "[???]*")
        
        
        If cFeld <> "" Then
        
            If cTeil = "1" Then '1.Durchgang
            
                If cwhere = "" Then
                    cwhere = " where "
                Else
                    cwhere = cwhere & " and "
                End If
                cwhere = cwhere & " A.BEZEICH like '" & cFeld & "*' "
    
            Else
                Dim sArray() As String
                sArray = Split(cFeld, " ")
                
                For i = 0 To UBound(sArray)
                
                    cFeld = sArray(i)
                    If cwhere = "" Then
                        cwhere = " where "
                    Else
                        cwhere = cwhere & " and "
                    End If
                    If cTeil = "2" Then
                        cwhere = cwhere & " A.BEZEICH like '*" & cFeld & "*' "
                    Else
                        cwhere = cwhere & " A.BEZEICH like '" & cFeld & "*' "
                    End If
                    
                
                Next i
            End If
        End If
        
    End If
    
    iStufe = 4
    cFeld = Text1(1).Text
    cFeld = Trim$(cFeld)
    
    If cFeld <> "" Then
        If IsNumeric(cFeld) = True Then
            If cwhere = "" Then
                cwhere = "where "
            Else
                cwhere = cwhere & "and "
            End If
            
            
            cEAN = cFeld
            
            If Len(cFeld) <= 6 Then
                cArtNr = cFeld
                cEAN = "1111111111111"
                
            Else
                
                cArtNr = ""
            End If
            
            If Left(cFeld, 1) = "2" Or Left(cFeld, 1) = "0" And Len(cFeld) = 8 Then
                cEigNr = Mid(cFeld, 2, 6)
            Else
                cEigNr = ""
            End If
            
            cwhere = cwhere & "("
            
            If cEAN <> "" Then
                cwhere = cwhere & "A.EAN like '" & cEAN & "*' "
                cwhere = cwhere & "or A.EAN2 like '" & cEAN & "*' "
                cwhere = cwhere & "or A.EAN3 like '" & cEAN & "*' "
            End If
            
            If cArtNr <> "" Then
                If InStr(cArtNr, "*") > 0 Then
                    cwhere = cwhere & " or A.ARTNR like '" & cArtNr & "' "
                Else
                    cwhere = cwhere & " or A.ARTNR = " & cArtNr & " "
                End If
            End If
            If cEigNr <> "" Then
                cwhere = cwhere & " or A.ARTNR = " & cEigNr & " "
            End If
            cwhere = cwhere & ") "
        Else
            Screen.MousePointer = 0
            Text1(1).SetFocus
            MsgBox "Artikelnummer oder EAN - Code ?", vbCritical, "SUCHEN"
        
            Exit Function
        End If
    End If
    iStufe = 5
    
    If cLinr <> "" Then
        If IsNumeric(cLinr) Then
            If cwhere = "" Then
                cwhere = "where "
            Else
                cwhere = cwhere & "and "
            End If
            cwhere = cwhere & " B.LINR = " & cLinr & " "
        Else
            Screen.MousePointer = 0
            MsgBox "Berichtigen Sie Ihre Lieferanteneingabe!", vbInformation, "Winkiss Hinweis:"
            Exit Function
        End If
    End If
    
    If List3.ListCount <> 0 Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
    
        cwhere = cwhere & "  (A.LPZ = " & Mid(List3.list(0), 1, InStr(1, List3.list(0), " ")) & " "
        For lcount = 1 To List3.ListCount - 1
            cwhere = cwhere & " or A.LPZ = " & Mid(List3.list(lcount), 1, InStr(1, List3.list(lcount), " ")) & " "
        Next lcount
        cwhere = cwhere & ")"
    Else
        iStufe = 6
        cFeld = Trim(Left(Text1(35).Text, 3))
        If cFeld <> "" Then
            If cwhere = "" Then
                cwhere = "where "
            Else
                cwhere = cwhere & "and "
            End If
            cwhere = cwhere & " A.LPZ = " & cFeld & " "
        End If
    End If
    
    'AGN
    
    If List4.Visible = True And List4.ListCount > 0 Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
    
        cwhere = cwhere & "(A.AGN = " & Mid(List4.list(0), 1, InStr(1, List4.list(0), " ")) & " "
        For lcount = 1 To List4.ListCount - 1
            cwhere = cwhere & " or A.agn= " & Mid(List4.list(lcount), 1, InStr(1, List4.list(lcount), " ")) & " "
        Next lcount
        cwhere = cwhere & " ) "
       
    Else
        'agn
        cFeld = Trim$(Text1(3).Text)
        If cFeld <> "" Then
            If cwhere = "" Then
                cwhere = "where "
            Else
                cwhere = cwhere & "and "
            End If
            cwhere = cwhere & " A.AGN = " & cFeld & " "
        End If
        
    End If

    
    cFeld = Text1(0).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.PGN = " & cFeld & " "
    End If
    
    If Check12.value = vbChecked Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.Bestand < 0 "
    End If
    
    cFeld = Text1(44).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.Groesse like  '" & cFeld & "*' "
    End If
    
    iStufe = 8
    
    cFeld = Label1(2).Tag
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.AWM = '" & cFeld & "' "
    End If
    
    iStufe = 9
    cFeld = Text1(4).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        
        If Right(cFeld, 1) = "*" Then
            cwhere = cwhere & "B.LIBESNR like '" & cFeld & "' "
        Else
            cwhere = cwhere & "B.LIBESNR = '" & cFeld & "' " 'dass machte die Sache für Topin schneller
        End If
    End If
    iStufe = 10
    
    
    
    cFeld = Text1(45).Text
    cFeld = Trim$(cFeld)
    cFeld = SwapStr(cFeld, ",", ".")
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.KVKPR1 >= " & cFeld & " "
    End If
    
    
    cFeld = Text1(46).Text
    cFeld = Trim$(cFeld)
    cFeld = SwapStr(cFeld, ",", ".")
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.KVKPR1 <= " & cFeld & " "
    End If
    
    

    
    cFeld = Text1(47).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & " A.aufdat = " & CLng(DateValue(cFeld)) & " "
    End If
    
    
    
    
    
    If Check14.value = vbChecked Then
    
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & " A.artnr in (Select artnr from interart) "
    
    
    End If
    
    
    
    
    
    
    cSQL = cSQL & cwhere & cJoin

    txtStatus.Text = 72
    
    anzeige "normal", "Artikel suchen...", Label0(4)
    
'    MsgBox cSQL





    gdBase.Execute cSQL, dbFailOnError
    
    CheckIndex "Top" & srechnertab, "Artnr", "", gdBase
    
    'duplikate löschen
    
    txtStatus.Text = 74
    
    DuplikateDelTabelle "Top" & srechnertab, gdBase, ""
    
    'Ende Duplikate
    
    txtStatus.Text = 76
    
    cSQL = "Update Top" & srechnertab & " set lagerp = 0"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Top" & srechnertab & " inner join Lagerplatz on "
    cSQL = cSQL & " Top" & srechnertab & ".artnr = Lagerplatz.artnr "
    cSQL = cSQL & " set Top" & srechnertab & ".lagerp = Lagerplatz.lagerp "
    gdBase.Execute cSQL, dbFailOnError
    
    txtStatus.Text = 77
    
    'such mal alle ArtliefArtikel, die kein RKZ haben
    'das Ganze machst du nur, wenn keine Lieferant gewählt wurde
    
    If Trim(Text1(4).Text) = "" And Trim(Text1(7).Text) = "" Then
    
        cSQL = "Update Top" & srechnertab & " inner join artlief on "
        cSQL = cSQL & " Top" & srechnertab & ".artnr = artlief.artnr "
        cSQL = cSQL & " set Top" & srechnertab & ".linr = artlief.linr "
        cSQL = cSQL & " where artlief.RKZ = 'N' "
        gdBase.Execute cSQL, dbFailOnError
        
    End If
    
    cSQL = "Update Top" & srechnertab & " inner join artlief on "
    cSQL = cSQL & " Top" & srechnertab & ".artnr = artlief.artnr "
    cSQL = cSQL & " and Top" & srechnertab & ".linr = artlief.linr "
    cSQL = cSQL & " set Top" & srechnertab & ".LEKPR = artlief.LEKPR "
    cSQL = cSQL & " , Top" & srechnertab & ".LIBESNR = artlief.LIBESNR "
    cSQL = cSQL & " , Top" & srechnertab & ".MINMEN = artlief.MINMEN "
    cSQL = cSQL & " , Top" & srechnertab & ".RKZ = artlief.RKZ "
    cSQL = cSQL & " , Top" & srechnertab & ".EXDAT = artlief.EXDAT "
    gdBase.Execute cSQL, dbFailOnError
    
    
    If Check13.value = vbChecked Then
        cSQL = "delete from Top" & srechnertab & " "
        cSQL = cSQL & " where RKZ = 'N' "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    txtStatus.Text = 78
    
    cSQL = "Update Top" & srechnertab & " inner join PREISE on "
    cSQL = cSQL & " Top" & srechnertab & ".artnr = PREISE.artnr "
    cSQL = cSQL & " set Top" & srechnertab & ".SPEZKVK = PREISE.PREISWERT "
    cSQL = cSQL & " where Preise.preistyp = 0"
    gdBase.Execute cSQL, dbFailOnError
    
    txtStatus.Text = 79
    
    cSQL = "Update Top" & srechnertab & " inner join ARTMERK on "
    cSQL = cSQL & " Top" & srechnertab & ".artnr = ARTMERK.artnr "
    cSQL = cSQL & " set Top" & srechnertab & ".MERK = ARTMERK.MERK "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Top" & srechnertab & " inner join INTERART on "
    cSQL = cSQL & " Top" & srechnertab & ".artnr = INTERART.artnr "
    cSQL = cSQL & " set Top" & srechnertab & ".SHOP = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Top" & srechnertab & " inner join TEXTIL on "
    cSQL = cSQL & " Top" & srechnertab & ".artnr = TEXTIL.artnr "
    cSQL = cSQL & " set Top" & srechnertab & ".MODELL = TEXTIL.MODELL "
    cSQL = cSQL & " , Top" & srechnertab & ".MATERIAL = TEXTIL.MATERIAL "
    cSQL = cSQL & " , Top" & srechnertab & ".FARBBEZ = TEXTIL.FARBBEZ "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Top" & srechnertab & " inner join GRUPPE_ARTIKEL on "
    cSQL = cSQL & " Top" & srechnertab & ".artnr = GRUPPE_ARTIKEL.artnr "
    cSQL = cSQL & " set Top" & srechnertab & ".GRUPPENNR = GRUPPE_ARTIKEL.GRUPPENNR "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Top" & srechnertab & " set HS = 0"
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    
    'ist ein Lieferant gewählt?
    If cLinr <> "" Then
        'dann Handelsspanne nach LEKPR vom Lieferanten
        
        cSQL = "Update Top" & srechnertab & " set HS = "
        cSQL = cSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStV & ")) - lekpr) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStV & ")) "
        cSQL = cSQL & "  "
        cSQL = cSQL & "  where MWST = 'V' and lekpr > 0 and KVKPR1 > 0"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update Top" & srechnertab & " set HS = "
        cSQL = cSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStE & ")) - lekpr) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStE & ")) "
        cSQL = cSQL & "  "
        cSQL = cSQL & "  where MWST = 'E' and lekpr > 0 and KVKPR1 > 0"
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update Top" & srechnertab & " set HS = "
        cSQL = cSQL & "  ((((KVKPR1 * 100) / (100)) - lekpr) *100) / ((KVKPR1 * 100) / (100)) "
        cSQL = cSQL & "  "
        cSQL = cSQL & "  where MWST = 'O' and lekpr > 0 and KVKPR1 > 0"
        gdBase.Execute cSQL, dbFailOnError
        
        
        
        
        
        
    Else
    
    
        
    
    
    
    
    
    
    
    
    
    
    
    
        If gsSpanne = "SEK" Then
            'Handelsspanne HS über Schnittek
            cSQL = "Update Top" & srechnertab & " set HS = "
            cSQL = cSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStV & ")) - ekpr) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStV & ")) "
            cSQL = cSQL & "  "
            cSQL = cSQL & "  where MWST = 'V' and ekpr > 0 and KVKPR1 > 0"
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update Top" & srechnertab & " set HS = "
            cSQL = cSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStE & ")) - ekpr) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStE & ")) "
            cSQL = cSQL & "  "
            cSQL = cSQL & "  where MWST = 'E' and ekpr > 0 and KVKPR1 > 0"
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update Top" & srechnertab & " set HS = "
            cSQL = cSQL & "  ((((KVKPR1 * 100) / (100)) - ekpr) *100) / ((KVKPR1 * 100) / (100)) "
            cSQL = cSQL & "  "
            cSQL = cSQL & "  where MWST = 'O' and ekpr > 0 and KVKPR1 > 0"
            gdBase.Execute cSQL, dbFailOnError
            
        Else
            'Handelsspanne HS über Listenek - größter
            
            
            
            If gbEKMAX = True Then
        
                cSQL = "Update Top" & srechnertab & " set lekpr = 0"
                gdBase.Execute cSQL, dbFailOnError
                
                loeschNEW "MAX_LEK_" & srechnertab, gdBase
                
                cSQL = "select artnr, max(lekpr) as Maxlekpr into MAX_LEK_" & srechnertab & " from Artlief "
                cSQL = cSQL & " where artnr in(Select artnr from Top" & srechnertab & ")"
                cSQL = cSQL & " group by artnr "
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Update Top" & srechnertab & " inner join MAX_LEK_" & srechnertab & " on "
                cSQL = cSQL & " Top" & srechnertab & ".artnr = MAX_LEK_" & srechnertab & ".artnr "
                cSQL = cSQL & " set Top" & srechnertab & ".lekpr = MAX_LEK_" & srechnertab & ".maxlekpr "
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Update Top" & srechnertab & " set HS = "
                cSQL = cSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStV & ")) - lekpr) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStV & ")) "
                cSQL = cSQL & "  "
                cSQL = cSQL & "  where MWST = 'V' and lekpr > 0 and KVKPR1 > 0"
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Update Top" & srechnertab & " set HS = "
                cSQL = cSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStE & ")) - lekpr) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStE & ")) "
                cSQL = cSQL & "  "
                cSQL = cSQL & "  where MWST = 'E' and lekpr > 0 and KVKPR1 > 0"
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Update Top" & srechnertab & " set HS = "
                cSQL = cSQL & "  ((((KVKPR1 * 100) / (100)) - lekpr) *100) / ((KVKPR1 * 100) / (100)) "
                cSQL = cSQL & "  "
                cSQL = cSQL & "  where MWST = 'O' and lekpr > 0 and KVKPR1 > 0"
                gdBase.Execute cSQL, dbFailOnError
            
            Else
            
                'Handelsspanne HS über Listenek - kleinster
            
                cSQL = "Update Top" & srechnertab & " set lekpr = 0"
                gdBase.Execute cSQL, dbFailOnError
                
                loeschNEW "MIN_LEK_" & srechnertab, gdBase
                
                cSQL = "select artnr, min(lekpr) as Minlekpr into MIN_LEK_" & srechnertab & " from Artlief "
                cSQL = cSQL & " where artnr in(Select artnr from Top" & srechnertab & ")"
                cSQL = cSQL & " group by artnr "
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Update Top" & srechnertab & " inner join MIN_LEK_" & srechnertab & " on "
                cSQL = cSQL & " Top" & srechnertab & ".artnr = MIN_LEK_" & srechnertab & ".artnr "
                cSQL = cSQL & " set Top" & srechnertab & ".lekpr = MIN_LEK_" & srechnertab & ".minlekpr "
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Update Top" & srechnertab & " set HS = "
                cSQL = cSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStV & ")) - lekpr) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStV & ")) "
                cSQL = cSQL & "  "
                cSQL = cSQL & "  where MWST = 'V' and lekpr > 0 and KVKPR1 > 0"
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Update Top" & srechnertab & " set HS = "
                cSQL = cSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStE & ")) - lekpr) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStE & ")) "
                cSQL = cSQL & "  "
                cSQL = cSQL & "  where MWST = 'E' and lekpr > 0 and KVKPR1 > 0"
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Update Top" & srechnertab & " set HS = "
                cSQL = cSQL & "  ((((KVKPR1 * 100) / (100)) - lekpr) *100) / ((KVKPR1 * 100) / (100)) "
                cSQL = cSQL & "  "
                cSQL = cSQL & "  where MWST = 'O' and lekpr > 0 and KVKPR1 > 0"
                gdBase.Execute cSQL, dbFailOnError
            
            End If
        End If
        
        loeschNEW "MIN_LEK_" & srechnertab, gdBase
        loeschNEW "MAX_LEK_" & srechnertab, gdBase
        
        
        'bevor es über SEK oder LEK geht
        'die Frage nach PrimLinr
        
        If glPrimLinr > 0 Then
        
'            cSQL = "Update Top" & srechnertab & " set lekpr = 0"
'            gdBase.Execute cSQL, dbFailOnError
            
            loeschNEW "PRIM_" & srechnertab, gdBase
            
            cSQL = "select artnr, lekpr into PRIM_" & srechnertab & " from Artlief "
            cSQL = cSQL & " where artnr in(Select artnr from Top" & srechnertab & ")"
            cSQL = cSQL & " and Linr = " & glPrimLinr
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update Top" & srechnertab & " inner join PRIM_" & srechnertab & " on "
            cSQL = cSQL & " Top" & srechnertab & ".artnr = PRIM_" & srechnertab & ".artnr "
            cSQL = cSQL & " set Top" & srechnertab & ".lekpr = PRIM_" & srechnertab & ".lekpr "
            cSQL = cSQL & " , Top" & srechnertab & ".linr = " & glPrimLinr
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update Top" & srechnertab & " inner join artlief on "
            cSQL = cSQL & " Top" & srechnertab & ".artnr = artlief.artnr "
            cSQL = cSQL & " and Top" & srechnertab & ".linr = artlief.linr "
            
            cSQL = cSQL & " set Top" & srechnertab & ".LIBESNR = artlief.LIBESNR "
            cSQL = cSQL & " , Top" & srechnertab & ".MINMEN = artlief.MINMEN "
            cSQL = cSQL & " , Top" & srechnertab & ".RKZ = artlief.RKZ "
            cSQL = cSQL & " , Top" & srechnertab & ".EXDAT = artlief.EXDAT "
            gdBase.Execute cSQL, dbFailOnError
        
            cSQL = "Update Top" & srechnertab & " set HS = "
            cSQL = cSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStV & ")) - lekpr) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStV & ")) "
            cSQL = cSQL & "  "
            cSQL = cSQL & "  where MWST = 'V' and lekpr > 0 and KVKPR1 > 0"
            cSQL = cSQL & " and Linr = " & glPrimLinr
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update Top" & srechnertab & " set HS = "
            cSQL = cSQL & "  ((((KVKPR1 * 100) / (100 + " & gdMWStE & ")) - lekpr) *100) / ((KVKPR1 * 100) / (100 + " & gdMWStE & ")) "
            cSQL = cSQL & "  "
            cSQL = cSQL & "  where MWST = 'E' and lekpr > 0 and KVKPR1 > 0"
            cSQL = cSQL & " and Linr = " & glPrimLinr
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update Top" & srechnertab & " set HS = "
            cSQL = cSQL & "  ((((KVKPR1 * 100) / (100)) - lekpr) *100) / ((KVKPR1 * 100) / (100)) "
            cSQL = cSQL & "  "
            cSQL = cSQL & "  where MWST = 'O' and lekpr > 0 and KVKPR1 > 0"
            cSQL = cSQL & " and Linr = " & glPrimLinr
            gdBase.Execute cSQL, dbFailOnError
            
            loeschNEW "PRIM_" & srechnertab, gdBase
            
        End If
        
        
        
        
        
    End If
    
    
    
    
    
    
    
    
    
    Dim rsrs As Recordset
    Dim rsArt As Recordset
    Dim zbrestyes As Boolean
    
    zbrestyes = False
    
    If Trim$(gcFilNr) = "0" Then
        
    Else
        If NewTableSuchenDBKombi("ZBREST", gdBase) Then
            zbrestyes = True
        End If
    End If
    
    txtStatus.Text = 80
    
    cSQL = "Update Top" & srechnertab
    cSQL = cSQL & " set MINMEN = 0"
    gdBase.Execute cSQL, dbFailOnError
    
    txtStatus.Text = 81
    
    cSQL = "Update Top" & srechnertab
    cSQL = cSQL & " set LUG = 0"
    gdBase.Execute cSQL, dbFailOnError
    
    txtStatus.Text = 0
    
    siAnzeige = 0
    Set rsArt = gdBase.OpenRecordset("Top" & srechnertab)
    If Not rsArt.EOF Then
    
        rsArt.MoveLast
        lcount = rsArt.RecordCount
        rsArt.MoveFirst
        Do While Not rsArt.EOF
        
            siAnzeige = siAnzeige + 1
            txtStatus.Text = CStr((100 * siAnzeige) / lcount)
        
            If Not IsNull(rsArt!artnr) Then
                cArtNr = rsArt!artnr
                
                cSQL = "Select SUM(BESTVOR) as BESTELLT from "
                If Trim$(gcFilNr) = "0" Then
                    cSQL = cSQL & " bestrest "
                Else
                    If zbrestyes Then
                        cSQL = cSQL & " zbrest "
                    Else
                        cSQL = cSQL & " bestrest "
                    End If
                End If
                
                cSQL = cSQL & "  where ARTNR = " & cArtNr
                Set rsrs = gdBase.OpenRecordset(cSQL)
                If Not rsrs.EOF Then
                    rsrs.MoveFirst
                    If Not IsNull(rsrs!BESTELLT) Then
                        rsArt.Edit
                        rsArt!MINMEN = rsrs!BESTELLT
                        rsArt.Update
                    End If
                End If
                rsrs.Close: Set rsrs = Nothing
            End If
        
        rsArt.MoveNext
        Loop
    End If
    rsArt.Close: Set rsArt = Nothing
    
    cSQL = " Alter table TOP" & srechnertab & " add LAGERWSEK double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add EKaktJahr double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add EKvorJahr double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add UmsBraktJahr double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add UmsBrvorJahr double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add UmsSEKaktJahr double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add UmsSEKvorJahr double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add UmsBrakt12M double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add UmsBrvor12M double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add UmsSEKakt12 double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add UmsSEKvor12 double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add UmsBr12MDIFFabs double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add UmsBr12MDIFFrela double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add UmsSEK12MDIFFabs double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add UmsSEK12MDIFFrela double "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add LWE DATETIME "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add LVK DATETIME "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = " Alter table TOP" & srechnertab & " add VKMaktJahr long "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add VKMvorJahr long "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add VKMakt12M long "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table TOP" & srechnertab & " add VKMvor12M long "
    gdBase.Execute cSQL, dbFailOnError
    
    
   
    If Check10.value = vbChecked Then
    
        anzeige "normal", "mit Detailzahlen: es werden Hintergrunddaten zusammengefasst...", Label1(59)

'        If UMS_ARTNRaktuell = False Then
            ErzeugeArtnrUmsatz
'        End If

        siAnzeige = 0
        Set rsArt = gdBase.OpenRecordset("Top" & srechnertab)
        If Not rsArt.EOF Then
        
            rsArt.MoveLast
            lcount = rsArt.RecordCount
            lMax = rsArt.RecordCount
            rsArt.MoveFirst
            Do While Not rsArt.EOF
            
                siAnzeige = siAnzeige + 1
                txtStatus.Text = CStr((100 * siAnzeige) / lMax)

                If Not IsNull(rsArt!artnr) Then
                    cArtNr = rsArt!artnr
                    
                    'Lagerumschlag
                    Dim dlug As Double
                    Dim dlugd As Double
                    Dim dLRW As Double
                    
                    dlug = HoleLagerumschlag1(cArtNr)
                    dlugd = 0
                    If dlug > 0 Then
                        dlugd = 360 / dlug
                    End If
                    
                    dLRW = Val(wievieleTage(cArtNr))
                    
                    'Lagerumschlag ende

                    lcount = lcount - 1
                    anzeige "normal", "noch " & CStr(lcount) & " Artikel: " & rsArt!BEZEICH, Label1(59)

                    dEINKaufswert = CDbl(EinkaufsumsatzermittlungArtikel(cArtNr, gdBase, CInt(Year(Now))))
                    dEINKaufswertvj = CDbl(EinkaufsumsatzermittlungArtikel(cArtNr, gdBase, CInt(Year(Now) - 1)))

                    dUmsBraktJahr = ermgesUmsatzARTnr(0, CInt(Year(Now)), cArtNr)
                    dUmsBrvorJahr = ermgesUmsatzARTnr(0, CInt(Year(Now) - 1), cArtNr)
                    
                    l_VKM_aktJahr = ermgesVKMengeARTnr(0, CInt(Year(Now)), cArtNr)
                    l_VKM_vorJahr = ermgesVKMengeARTnr(0, CInt(Year(Now) - 1), cArtNr)

                    dUmsSEKaktJahr = ermgesEKUmsatzARTNR(0, CInt(Year(Now)), cArtNr)
                    dUmsSEKvorJahr = ermgesEKUmsatzARTNR(0, CInt(Year(Now) - 1), cArtNr)

                    dUms12M = 0
                    dUms12MVJZR = 0
                    dUmsSEK12M = 0
                    dUmsSEK12MVJZR = 0
                    
                    l_VKM_12M = 0
                    l_VKM_12MVJZR = 0

                    bymonat = Month(DateValue(Now))
                    iJahr = Year(DateValue(Now))

                    For j = 1 To 12

                        If bymonat = 1 Then
                            bymonat = 12
                            iJahr = iJahr - 1
                        Else
                            bymonat = bymonat - 1
                            iJahr = iJahr
                        End If

                        dUms12M = dUms12M + ermgesUmsatzARTnr(bymonat, iJahr, cArtNr)
                        dUms12MVJZR = dUms12MVJZR + ermgesUmsatzARTnr(bymonat, iJahr - 1, cArtNr)
                        
                        l_VKM_12M = l_VKM_12M + ermgesVKMengeARTnr(bymonat, iJahr, cArtNr)
                        l_VKM_12MVJZR = l_VKM_12MVJZR + ermgesVKMengeARTnr(bymonat, iJahr - 1, cArtNr)

                        dUmsSEK12M = dUmsSEK12M + ermgesEKUmsatzARTNR(bymonat, iJahr, cArtNr)
                        dUmsSEK12MVJZR = dUmsSEK12MVJZR + ermgesEKUmsatzARTNR(bymonat, iJahr - 1, cArtNr)

                    Next j

'                    dLagerwertzumSEK = LAGEREKermittlungJetztARTNR(cArtNr)

                Else
                
                    l_VKM_aktJahr = 0
                    l_VKM_vorJahr = 0
    
                    l_VKM_12M = 0
                    l_VKM_12MVJZR = 0
                
                    dEINKaufswert = 0
                    dEINKaufswertvj = 0

                    dUmsBraktJahr = 0
                    dUmsBrvorJahr = 0

                    dUmsSEKaktJahr = 0
                    dUmsSEKvorJahr = 0

                    dLagerwertzumSEK = 0
                    dlug = 0

                    dlugd = 0
                    dLRW = 0

                    dUms12M = 0
                    dUms12MVJZR = 0
                    dUmsSEK12M = 0
                    dUmsSEK12MVJZR = 0

                End If

                rsArt.Edit
                rsArt!LUG = dlug
                rsArt!LUGD = dlugd
                rsArt!LRW = dLRW
'                rsArt!LWE = ErmlzZugang(cArtNr)
                
'                rsArt!LVK = ErmlzVK(cArtNr)
                   
'                rsArt!LAGERWSEK = dLagerwertzumSEK
                rsArt!EKaktJahr = dEINKaufswert
                rsArt!EKvorJahr = dEINKaufswertvj

                rsArt!UmsBraktJahr = dUmsBraktJahr
                rsArt!UmsBrvorJahr = dUmsBrvorJahr
                
                
                
                
                
                
                
                
                
                
                rsArt!VKMaktJahr = l_VKM_aktJahr
                rsArt!VKMvorJahr = l_VKM_vorJahr
                
                rsArt!VKMakt12M = l_VKM_12M
                rsArt!VKMvor12M = l_VKM_12MVJZR

                rsArt!UmsSEKaktJahr = dUmsSEKaktJahr
                rsArt!UmsSEKvorJahr = dUmsSEKvorJahr

                rsArt!UmsBrakt12M = dUms12M
                rsArt!UmsBrvor12M = dUms12MVJZR

                dUms12MDIFFabs = 0
                dUms12MDIFFabs = dUms12M - dUms12MVJZR

                dUms12MDIFFrela = 0
                If dUms12M <> 0 Then
                    dUms12MDIFFrela = 100 * dUms12MDIFFabs / dUms12M
                End If

                rsArt!UmsSEKakt12 = dUmsSEK12M
                rsArt!UmsSEKvor12 = dUmsSEK12MVJZR

                dUmsSEK12MDIFFabs = 0
                dUmsSEK12MDIFFabs = dUmsSEK12M - dUmsSEK12MVJZR

                dUmsSEK12MDIFFrela = 0
                If dUmsSEK12M <> 0 Then
                    dUmsSEK12MDIFFrela = 100 * dUmsSEK12MDIFFabs / dUmsSEK12M
                End If

                rsArt!UmsBr12MDIFFabs = dUms12MDIFFabs
                rsArt!UmsSEK12MDIFFabs = dUmsSEK12MDIFFabs

                rsArt!UmsBr12MDIFFrela = dUms12MDIFFrela
                rsArt!UmsSEK12MDIFFrela = dUmsSEK12MDIFFrela

                rsArt.Update

            rsArt.MoveNext
            Loop
        End If
        rsArt.Close
    End If
    
    anzeige "normal", "", Label1(59)
  
    picprogress.Visible = False
    
    SucheArtikelWKL10 = True
    
    Screen.MousePointer = 0

Exit Function
LOKAL_ERROR:
    If err.Number = 5 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SucheArtikelWKL10"
        Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten." & iStufe
        
        Fehlermeldung1
    End If
    Resume Next
   
End Function

Private Sub cbo1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim cSQL As String
    Dim cLinr As String
    Dim cArtNr As String
    Dim dWert As Double
    Dim ctmp As String
    
    cArtNr = Trim(Text1(5).Text)
    cLinr = Trim(cbo1.Text)
    
    If cLinr = "" And cArtNr <> "" Then
        'Lösche Artliefeintrag
        cSQL = "Delete from artlief where artnr = " & cArtNr
        cSQL = cSQL & " and linr is null"
        gdBase.Execute cSQL, dbFailOnError
        
        cbo1fuellen cArtNr
        Liefdetail cArtNr
        Exit Sub
    End If
    Liefdetail cArtNr

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbo1_Click"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cbo1_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    cbo1.BackColor = glSelBack1
    Label3(1).Caption = Label4(2).Caption
    Label3(2).Caption = "-3"
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbo1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cboGp_GotFocus()
    On Error GoTo LOKAL_ERROR

    cboGp.SelStart = 0
    cboGp.SelLength = Len(cboGp.Text)
    cboGp.BackColor = glSelBack1

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbogp_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cbo1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        gF2Prompt.cFeld = "LINR"
            
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
            If gF2Prompt.cWahl <> "" Then
                cbo1.Text = gF2Prompt.cWahl
                cbo1_Click
            End If
        End If
        
    End If
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbo1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cbo1_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    cbo1.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbo1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cboGp_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        
        Text2(1).SetFocus
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboGp_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub cboGp_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    cboGp.BackColor = vbWhite
    
    If Trim(cboGp.Text) = "" Then
        Label7(7).Caption = "Bitte geben Sie die Packungs - EAN an!"
        Label7(7).Refresh
        
        Exit Sub
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboGp_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check11_Click()
On Error GoTo LOKAL_ERROR
    
    If Check11.value = vbChecked Then
        gbBestandsgrund = False
    Else
        gbBestandsgrund = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check11_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check12_Click()
On Error GoTo LOKAL_ERROR
    
    If Check12.value = vbChecked Then
        Check7.Visible = False
        Check7.value = False
    Else
        Check7.Visible = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check12_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check15_Click()
On Error GoTo LOKAL_ERROR
    
    If Check15.value = vbChecked Then
        Internet_Shop Text1(5).Text, "J"
            
        If LeseInterArt(Text1(5).Text) Then
            Check15.ForeColor = vbRed
        Else
            Check15.ForeColor = vbBlack
        End If
    Else
        Internet_Shop Text1(5).Text, "N"
            
        If LeseInterArt(Text1(5).Text) Then
            Check15.ForeColor = vbRed
        Else
            Check15.ForeColor = vbBlack
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check15_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check2_Click()
On Error GoTo LOKAL_ERROR
    
    CheckandZeig
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check2_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check1_Click()
On Error GoTo LOKAL_ERROR
    
    CheckandZeig
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check3_Click()
On Error GoTo LOKAL_ERROR
    
    CheckandZeig
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check3_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub CheckandZeig()
On Error GoTo LOKAL_ERROR
    

    If Check3.value = vbChecked Then
        bAusblenden = True
    Else
        bAusblenden = False
    End If
    
    If Check2.value = vbChecked Then
        bgef = True
    Else
        bgef = False
    End If
    
    If Check1.value = vbChecked Then
        bBest = True
    Else
        bBest = False
    End If
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "CheckandZeig"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub CheckandZeig1()
On Error GoTo LOKAL_ERROR
    

    If Check8.value = vbChecked Then
        bAusblenden = True
    Else
        bAusblenden = False
    End If
    
    If Check6.value = vbChecked Then
        bgef = True
    Else
        bgef = False
    End If
    
    If Check7.value = vbChecked Then
        bBest = True
    Else
        bBest = False
    End If
    
    If Check9.value = vbChecked Then
        binbest = True
    Else
        binbest = False
    End If
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "CheckandZeig1"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check5_Click()
On Error GoTo LOKAL_ERROR

    If Check5.value = vbChecked Then
        Check5.ForeColor = vbRed
    Else
        Check5.ForeColor = glS1
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check5_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check8_Click()
On Error GoTo LOKAL_ERROR

    If Check8.value = vbChecked Then
        Check8.ForeColor = vbRed
        Check13.Visible = False
        Check13.value = False
    Else
        
        Check8.ForeColor = glS1
        Check13.Visible = True
    End If
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check8_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check13_Click()
On Error GoTo LOKAL_ERROR

    If Check13.value = vbChecked Then
        
        Check8.Visible = False
        Check8.value = False
    Else
        
        
        Check8.Visible = True
    End If
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check13_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check7_Click()
On Error GoTo LOKAL_ERROR

    If Check7.value = vbChecked Then
        Check7.ForeColor = vbRed
        
        
        
    Else
        Check7.ForeColor = glS1
        
    End If
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check7_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check9_Click()
On Error GoTo LOKAL_ERROR

    If Check9.value = vbChecked Then
        Check9.ForeColor = vbRed
    Else
        Check9.ForeColor = glS1
    End If
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check9_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check6_Click()
On Error GoTo LOKAL_ERROR

    If Check6.value = vbChecked Then
        Check6.ForeColor = vbRed
    Else
        Check6.ForeColor = glS1
    End If
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check6_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdfarbe_Click()
    On Error GoTo LOKAL_ERROR

    Frame5.Top = 0
    Frame5.Top = cmdfarbe.Top + Frame5.Top + cmdfarbe.Height + Frame3.Top + 50
    Frame5.Left = (cmdfarbe.Left + cmdfarbe.Width) - Frame5.Width
    Frame5.Visible = True
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdfarbe_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdfarbe_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyF2 Then
        frmWKL49.Show 1
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdfarbe_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo1_Click()
On Error GoTo LOKAL_ERROR

Dim cSQL As String
Dim rsrs As Recordset

    cSQL = "Select * from KONDITIONEN where ARTNR = " & Text1(5) & " "
    cSQL = cSQL & " and KONDI = " & Combo1.Text
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!kondi) Then
            Text2(5).Text = rsrs!kondi
        Else
            Text2(5).Text = ""
        End If

        If Not IsNull(rsrs!Faktor) Then
            Text2(6).Text = rsrs!Faktor
        Else
            Text2(6).Text = ""
        End If
    Else
        Text2(6).Text = ""
        Text2(5).Text = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_Click"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    
    SpaltennummerHS = 255

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "BESTAND"
                SpaltennummerBESTAND = i
            Case Is = "AWM"
                SpaltennummerAWM = i
            Case Is = "ARTNR"
                SpaltennummerArtnr = i
            Case Is = "BEZEICH"
                SpaltennummerBEZEICH = i
            Case Is = "KVKPR1"
                SpaltennummerKVKPR1 = i
            Case Is = "HS"
                SpaltennummerHS = i
            Case Is = "GEFUEHRT"
                SpaltennummerGEFUEHRT = i
            Case Is = "RABATT_OK"
                SpaltennummerRABATT_OK = i
            Case Is = "BONUS_OK"
                SpaltennummerBONUS_OK = i
            Case Is = "PREISSCHU"
                SpaltennummerPREISSCHU = i
            Case Is = "RKZ"
                SpaltennummerRKZ = i
            Case Is = "PGN"
                SpaltennummerPGN = i
            Case Is = "EAN"
                SpaltennummerEAN = i
            Case Is = "EAN2"
                SpaltennummerEAN2 = i
            Case Is = "EAN3"
                SpaltennummerEAN3 = i
            Case Is = "LPZ"
                SpaltennummerLPZ = i
            Case Is = "NOTIZEN"
                SpaltennummerNOTIZEN = i
            Case Is = "AGN"
                SpaltennummerAGN = i
            Case Is = "LINR"
                SpaltennummerLINR = i
            Case Is = "LEKPR"
                SpaltennummerLEKPR = i
            Case Is = "VKPR"
                SpaltennummerLVKPR = i
            Case Is = "LIBESNR"
                SpaltennummerLIBESNR = i
            Case Is = "LAGERP"
                SpaltennummerLagerP = i
            Case Is = "GROESSE"
                SpaltennummerGROESSE = i
            Case Is = "MINBEST"
                SpaltennummerMB = i
            Case Is = "SHOP"
                SpaltennummerSHOP = i
            Case Is = "MODELL"
                SpaltennummerModell = i
            Case Is = "MATERIAL"
                SpaltennummerMaterial = i
            Case Is = "FARBBEZ"
                SpaltennummerFarbbez = i
            Case Is = "GRUPPENNR"
                SpaltennummerGRUPPE = i
            Case Is = "MWST"
                SpaltennummerMWST = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function zeige_Grid(bAusblenden As Boolean, bBest As Boolean, bgef As Boolean, binbest As Boolean, sOrderby As String) As Boolean
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
    Dim cSQL As String
    
    zeige_Grid = True
    
    If Not NewTableSuchenDBKombi("TOP" & srechnertab, gdBase) Then
        anzeige "rot2", "Keine Artikel", Label0(4)
        Text1(1).SetFocus
        Exit Function
    End If
    
    
    Dim bAnd        As Boolean
    bAnd = False
    
    cSQL = "Select * from TOP" & srechnertab
    If bAusblenden Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " RKZ = 'N' "

        bAnd = True
    End If
    
    If bBest Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " bestand > 0 "
        bAnd = True
    End If
    
    If binbest Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " MINMEN > 0 "
        bAnd = True
    End If
    
    If bgef Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " GEFUEHRT = 'J' "
        bAnd = True
    End If
    
    
    Set recAnz = gdBase.OpenRecordset(cSQL)
    
    If recAnz.EOF Then
        
        MSFlexGrid1.Visible = False
        MSFlexGrid1.Clear
        Frame1.Visible = False
        Frame0.Visible = True
        anzeige "rot2", "Keine Artikel", Label0(4)
        Text1(1).SetFocus
        recAnz.Close: Set recAnz = Nothing
        Exit Function
    Else
        recAnz.MoveLast
        
        If recAnz.RecordCount > 8000 Then
            iRet = MsgBox("Es wurden mehr als 8000 (" & recAnz.RecordCount & ") Datensätze gefunden. Bitte schränken Sie Ihre Suche weiter ein.", vbInformation, "Winkiss Hinweis:")
            recAnz.Close: Set recAnz = Nothing
            zeige_Grid = False
            Exit Function
        End If
        
        If recAnz.RecordCount > 2000 Then
            iRet = MsgBox("Es wurden mehr als 2000 (" & recAnz.RecordCount & ") Datensätze gefunden." & vbCrLf & "Wirklich anzeigen?", vbQuestion + vbYesNo, "Winkiss Frage:")
            If iRet = vbNo Then
                recAnz.Close: Set recAnz = Nothing
                zeige_Grid = False
                Exit Function
            End If
        End If
        
        anzeige "normal", recAnz.RecordCount & " Artikel werden angezeigt", Label8(5)
        
    End If
    recAnz.Close: Set recAnz = Nothing
    
    
    Screen.MousePointer = 11

    Tabcheck "BEAART"
    
    FormatGridOverTablay "BEAART"

    With MSFlexGrid1
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 25
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
        Next j
        
        
    
        
        FuellenMSFlex10 bAusblenden, bBest, bgef, binbest, sOrderby
        
        ermittlespalten
        
        Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
        
        FaerbenGrid MSFlexGrid1, CInt(SpaltennummerAWM), CInt(SpaltennummerArtnr)
        
        FaerbeBestvor MSFlexGrid1, CByte(SpaltennummerBEZEICH), CByte(SpaltennummerArtnr)
        
        FaerbeKVKweilSpezKvk MSFlexGrid1, CByte(SpaltennummerKVKPR1), CByte(SpaltennummerArtnr)
        
        If FaerbenGridaufGrundmehrLINR(MSFlexGrid1, CInt(SpaltennummerLINR), CInt(SpaltennummerArtnr)) Then
'            anzeige "rot2", "Rot unterlegte Lieferantennummern sind Zweitlieferanten pro Artikel. Sollten Sie diese überschreiben, so werden alle Zweitlieferanten des Artikels gelöscht.", Label8(5)
            anzeige "normal", MSFlexGrid1.Rows - 2 & " Artikel werden angezeigt", Label8(5)
        Else
            anzeige "normal", MSFlexGrid1.Rows - 2 & " Artikel werden angezeigt", Label8(5)
        End If
        
        Frame1.Visible = True
        Frame0.Visible = False
        
        
        Frame2.Visible = False
        
        .Visible = True
        .Redraw = True
        .RowHeight(2) = MSFlexGrid1.RowHeight(0)

        .Row = 1
        .SetFocus
    
    End With
  
    Screen.MousePointer = 0
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_Grid"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
'    Resume Next

End Function
Private Sub BereiteExportDaten(bAusblenden As Boolean, bBest As Boolean, bgef As Boolean, binbest As Boolean)
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
    Dim cSQL As String
    
    If Not NewTableSuchenDBKombi("TOP" & srechnertab, gdBase) Then
        Exit Sub
    End If
    
    loeschNEW "TOPZ" & srechnertab, gdBase
    
    Dim bAnd        As Boolean
    bAnd = False
    
    cSQL = "Select * into TOPZ" & srechnertab & " from TOP" & srechnertab
    If bAusblenden Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " RKZ = 'N' "
        bAnd = True
    End If
    
    If bBest Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " bestand > 0 "
        bAnd = True
    End If
    
    If binbest Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " MINMEN > 0 "
        bAnd = True
    End If
    
    If bgef Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " GEFUEHRT = 'J' "
        bAnd = True
    End If
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BereiteExportDaten"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub FaerbeBestvor(gridx As MSFlexGrid, spaltebestvor As Byte, spalteartnr As Byte)
    On Error GoTo LOKAL_ERROR

    Dim j           As Integer
    Dim lfakt       As Long
    Dim sierg       As Single
    Dim sArtnr      As String
    
    
    With gridx
        .Redraw = False
        For j = 1 To .Rows - 1
            .Row = j
            .Col = spalteartnr
            sArtnr = .Text
            lfakt = ermfak(sArtnr)
            If lfakt > 0 Then
                .Col = spaltebestvor
                .CellBackColor = vbYellow
            End If
            
        Next j
        .Redraw = True
    
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "FaerbeBestvor"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FaerbeKVKweilSpezKvk(gridx As MSFlexGrid, spaltezufaerb As Byte, spalteartnr As Byte)
    On Error GoTo LOKAL_ERROR

    Dim j           As Integer
    Dim sArtnr      As String
    
    With gridx
        .Redraw = False
        For j = 1 To .Rows - 1
            .Row = j
            .Col = spalteartnr
            sArtnr = .Text
            If LeseSpezpreis(CLng(Val(sArtnr)), 0) > 0 Then
                .Col = spaltezufaerb
                .CellForeColor = vbRed
            End If
        Next j
        .Redraw = True
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "FaerbeKVKweilSpezKvk"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSFlexGrid, siEigFak As Single)
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
Private Sub FuellenMSFlex10(bAusblenden As Boolean, bBest As Boolean, bgef As Boolean, binbest As Boolean, sOrderby As String)
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
    Dim corder      As String
    Dim bAnd        As Boolean
    bAnd = False
    
    
    If sOrderby <> "" Then
        corder = sOrderby
    Else
        If Option1(0).value Then
             corder = " order by  LPZ, BEZEICH "
             
        ElseIf Option1(1).value Then
             corder = " order by BEZEICH "
            
        ElseIf Option1(2).value Then
            corder = " order by AGN,BEZEICH "
            
        ElseIf Option1(3).value Then
            corder = " order by AWM desc, LPZ,BEZEICH "
            
        ElseIf Option1(4).value Then
            corder = " order by BESTAND desc "
            
        ElseIf Option1(5).value Then
            corder = " order by PGN "
        ElseIf Option1(6).value Then
            corder = " order by LIBESNR "
            
        End If
    End If

    picprogress.Visible = True
    
    cSQL = "Select * from TOP" & srechnertab
    
    If bAusblenden Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " RKZ = 'N' "
        bAnd = True
    End If
    
    If bBest Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " bestand > 0 "
        bAnd = True
    End If
    
    If binbest Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " MINMEN > 0 "
        bAnd = True
    End If
    
    If bgef Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " GEFUEHRT = 'J' "
        bAnd = True
    End If
    
    cSQL = cSQL & corder
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSFlexGrid1
        
        counter = rsrs.RecordCount
        lrow = 1
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                txtStatus.Text = (lrow * 100) / counter
                lrow = lrow + 1
                .Rows = lrow + 1
                .Col = 0
                
                For i = 0 To byAnzahlSpalten - 1
                    .Row = 0
                    .Col = i
                    
                    If sSpaltenname(i) = .Text Then
                        
                        Select Case sSpaltenname(i)
                            Case Is = "L EKPR", "L VKPR", "S EKPR", "K VKPR", "spez KVK", "LUG"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "####0.00")
                            
                            Case Is = "UMSATZ Br akt Jahr", "UMSATZ Br vor Jahr", "UMSATZ SEK akt Jahr", "UMSATZ SEK vor Jahr"
            
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "#######0.00")
                                
                            Case Is = "UMS Br l. 12M", "UMS Br l. 12M VJZR", "UMS SEK l. 12M", "UMS SEK l. 12M VJZR"
                                
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "#######0.00")
                                
                            Case Is = "DIFF UMS BR 12M €", "DIFF UMS BR 12M %", "DIFF UMS SEK 12M €", "DIFF UMS SEK 12M %"
                                
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "#######0.00")
                                If CDbl(sWert) < 0 Then
                                    .CellForeColor = vbRed
                                Else
                                    .CellForeColor = vbBlack
                                End If
                                
                            Case Is = "LAGER(SEK)", "EINKAUF akt Jahr", "EINKAUF vor Jahr"
                                
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "#######0.00")
                            
                            Case Is = "Shop"
                                
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "N"
                                End If
                                .Row = lrow
                                .Text = sWert
                                
                                If sWert = "J" Then
                                    .CellForeColor = vbRed
                                Else
                                    .CellForeColor = vbBlack
                                End If
                                
                            Case Is = "HS"
                                
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = Fix(rsrs(sSpaltenbez(i)))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = sWert
                                
                                If CDbl(sWert) >= 100 Then
                                    .CellBackColor = &H8000&
                                ElseIf CDbl(sWert) > 79.99 Then
                                    .CellBackColor = &HC000&
                                ElseIf CDbl(sWert) > 59.99 Then
                                    .CellBackColor = &HFF00&
                                ElseIf CDbl(sWert) > 39.99 Then
                                    .CellBackColor = &HFFFF&
                                ElseIf CDbl(sWert) > 19.99 Then
                                    .CellBackColor = &H80C0FF
                                ElseIf CDbl(sWert) > 0 Then
                                    .CellBackColor = &H80FF&
                                ElseIf CDbl(sWert) <= 0 Then
                                    .CellBackColor = &HFF&
                                End If
                                

    
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
        .Visible = True
    End With
    
    picprogress.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMSFlex10"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
        
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub Command1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    Screen.MousePointer = 11
    
    Select Case index
    
        Case 35 'Lagerplatz
            frmWKL206.Show 1
        Case 34
            Screen.MousePointer = 0
            gsARTNR = Text1(5).Text
            frmWKL205.Show 1
            
            If StaffelKVK_vorhanden(CLng(gsARTNR)) Then
                Command1(34).ForeColor = glWarn
            Else
                Command1(34).ForeColor = glButtonForecolor
            End If
            gsARTNR = ""
    
        Case 33
            ctmp = Text1(20).Text
            Text1(20).Text = Text1(21).Text
            Text1(21).Text = ctmp
        Case 32
            ctmp = Text1(18).Text
            Text1(18).Text = Text1(20).Text
            Text1(20).Text = ctmp
        Case 31 'massen einfügen
            If Text3.Text = "" Then
                Exit Sub
            End If
        
            If MSFlexGrid1.RowSel > 1 Then

                FlexGrid_Update MSFlexGrid1
                
                MSFlexGrid1.Row = 1
                MSFlexGrid1.SetFocus
            Else
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
            End If
        Case 30
            frmWKL180.Show 1
            If Val(gsGRUPPENNR) > 0 Then
                Text2(10).Text = Val(gsGRUPPENNR)
            End If
            gsGRUPPENNR = ""
        Case 29
            Screen.MousePointer = 0
            gsARTNR = Text1(5).Text
            frmWKL170.Show 1
            
            If LeseGeschwisterArt(gsARTNR) Then
                Command1(29).BackColor = vbRed
            Else
                Command1(29).BackColor = Command1(9).BackColor
            End If
            gsARTNR = ""
    
        Case 28
            Screen.MousePointer = 0
            gsARTNR = Text1(5).Text
            frmWKL163.Show 1
            
            If LeseInterArt(gsARTNR) Then
                Command1(28).BackColor = vbRed
            Else
                Command1(28).BackColor = Command1(9).BackColor
            End If
            gsARTNR = ""
        Case 27
            Screen.MousePointer = 0
            Text1_KeyUp 11, vbKeyF2, 0
        Case 26
            Screen.MousePointer = 0
            Text1_KeyUp 30, vbKeyF2, 0
        Case 25
            Screen.MousePointer = 0
            frmWKL143.Show 1
        Case 24
            Screen.MousePointer = 0
            frmWKL127.Show 1
    
        Case 23
            If SucheArtikelWKL10("2") Then

                If NewTableSuchenDBKombi("E10", gdApp) Then
                    voreinstellungladen10
                End If
                Me.Refresh

                If MSFlexGrid1.Visible = False Then
                    CheckandZeig
                    zeige_Grid bAusblenden, bBest, bgef, binbest, ""
                End If

                If MSFlexGrid1.Visible = True Then
                    If MSFlexGrid1.Cols > 1 Then
                        MSFlexGrid1.Col = 1
                        MSFlexGrid1.Row = 2
                        MSFlexGrid1.SetFocus
                    End If
                End If
            End If
        
        Case 22
            zeige_Grid bAusblenden, bBest, bgef, binbest, ""
        Case 0     '** Suchen **
        
            Command8.Visible = True
            Command9.Visible = True
            voreinstellungspeichern10A
            voreinstellungspeichern10
            anzeige "normal", "Artikeldaten bearbeiten", Label0(4)
            
            If SucheArtikelWKL10("1") Then
                If NewTableSuchenDBKombi("E10", gdApp) Then
                    voreinstellungladen10
                End If
                Me.Refresh
                
                If MSFlexGrid1.Visible = False Then
                    CheckandZeig1
                    If zeige_Grid(bAusblenden, bBest, bgef, binbest, "") = False Then
                        Exit Sub
                    End If
                End If
               
                If MSFlexGrid1.Visible = True Then
                    If MSFlexGrid1.Cols > 1 Then
                        MSFlexGrid1.Col = 1
                        MSFlexGrid1.Row = 2
                        MSFlexGrid1.SetFocus
                        
                        If gbArtEindeut Then
                            If MSFlexGrid1.Rows = 3 Then
                                Command2_Click 0
                            End If
                        End If
                        
                    End If
                Else
                    If SucheArtikelWKL10("2") Then
                        If NewTableSuchenDBKombi("E10", gdApp) Then
                            voreinstellungladen10
                        End If
                        Me.Refresh

                        If MSFlexGrid1.Visible = False Then
                            CheckandZeig1
                            zeige_Grid bAusblenden, bBest, bgef, binbest, ""
                        End If

                        If MSFlexGrid1.Visible = True Then
                            If MSFlexGrid1.Cols > 1 Then
                                MSFlexGrid1.Col = 1
                                MSFlexGrid1.Row = 2
                                MSFlexGrid1.SetFocus
                            End If
                            
                            If gbArtEindeut Then
                                If MSFlexGrid1.Rows = 3 Then
                                    Command2_Click 0
                                End If
                            End If
                            
                        Else
                        
                            Dim cSuchU_EAN As String
                            cSuchU_EAN = Text1(1).Text
                            
                            If Ist_in_ARTIKEL(cSuchU_EAN) = True Then
    
                                Screen.MousePointer = 0
                                MsgBox "Der gesuchte Artikel kann leider nicht angezeigt werden, bitte überprüfen Sie ihre Filteroptionen.", vbInformation, "Winkiss Hinweis:"
                                Exit Sub
                            
                            End If
        
                            Text1(1).Text = unbekanntenEAN_Suchen_und_Anlegen(Text1(1).Text)
                            
                            If Text1(1).Text = "" Then
                                Text1(1).Text = unbekanntenEAN_Suchen_und_Anlegen_DrogAlles(cSuchU_EAN) 'Über Drogerie und Spielwaren Schalter
                            End If
                            
                            If Text1(1).Text <> cSuchU_EAN Then
                                Command1_Click 0 'also nochmal suchen
                                Exit Sub
                            End If
                        
                        End If
                    End If
                End If
            Else
                
            End If
        
        Case Is = 1     '** Neu **

            Command8.Visible = False
            Command9.Visible = False
            LeereArtikelDatenWKL10 True
            
            Text1(5).Text = HoleFreieArtikelNrWKL10
                
            
            If Text1(5).Text <> "" Then
                Sicherheitslöschen Text1(5).Text
            End If
            
            Text1(15).Text = gsMWST
            Frame3.Visible = True
            Frame0.Visible = False
            Text1(5).SetFocus
            giDlgZustand = giNEU
            gbNew = True

            
        Case Is = 2     'Beenden
            
            Unload frmWKL10
        Case Is = 3     'Umverpackung EAN
            
            Frame6.Visible = True
            Frame0.Visible = False
            startframe6
                
        Case Is = 4     'Umverpackung EAN
            Frame0.Visible = True
            Frame6.Visible = False
        Case Is = 5     'Umverpackung speichern
            speichernGP
        Case Is = 6     'F2 Lieferant
            Screen.MousePointer = 0
            Text1_KeyUp 7, vbKeyF2, 0
        Case Is = 7     'F2 AGN
            Screen.MousePointer = 0
            Text1_KeyUp 3, vbKeyF2, 0
            
        Case Is = 10     'F2 PGN
            Screen.MousePointer = 0
            Text1_KeyUp 0, vbKeyF2, 0
            
        Case Is = 16     'F2 Marke
            Screen.MousePointer = 0
            Text1_KeyUp 2, vbKeyF2, 0
            
        Case Is = 8    'F2 linr
            Screen.MousePointer = 0
            cbo1_KeyUp vbKeyF2, 0
        Case Is = 9     'F2 Linie
            Screen.MousePointer = 0
            Text1_KeyUp 9, vbKeyF2, 0
        Case Is = 55     'F2 Linie
            Screen.MousePointer = 0
            Text1_KeyUp 35, vbKeyF2, 0
        Case Is = 11    'F2 farbe
            Screen.MousePointer = 0
            
            gsBackcolor = Label1(2).BackColor
            gsForecolor = Label1(2).ForeColor
            gsArtikelFarbe = Label1(2).Tag
            
            frmWKL49.Show 1
            
            Label1(2).BackColor = gsBackcolor
            Label1(2).ForeColor = gsForecolor
            Label1(2).Tag = gsArtikelFarbe
            
            If gsArtikelFarbe <> "" Then
                Label1(2).Caption = "Farbauswahl"
            Else
                Label1(2).Caption = "alle Farben"
            End If
        Case Is = 12    'F2 farbe
            Screen.MousePointer = 0
            cmdfarbe_KeyUp vbKeyF2, 0
        Case Is = 13   'lug
            Screen.MousePointer = 11
            gsARTNR = Text1(5).Text
            If Text1(29).Text = "" Then Text1(29).Text = "0"
            gsSEK = Text1(29).Text
            frmWKL62.Show 1
            gsARTNR = ""
        Case Is = 14  'lV
            Screen.MousePointer = 0
            gsARTNR = Text1(5).Text
            frmWKL63.Show 1
            gsARTNR = ""
            
        Case Is = 15   'lzu
            Screen.MousePointer = 0
            
            gsARTNR = Text1(5).Text
            frmWKL64.Show 1
            gsARTNR = ""
        Case Is = 10
            
            If Frame5.Visible = False Then
                
                Frame5.Top = Command1(10).Top + Frame5.Height
                Frame5.Left = Command1(10).Left
                Frame5.Visible = True
                
            Else
                Frame5.Visible = False
                Text1(2).Text = ""
                Command1(10).Caption = "alle"
                Command1(10).BackColor = glfarbe(0)
            End If
        Case 17

            Screen.MousePointer = 0
            frmWKL201.Show 1
            
        Case 18
            gsARTNR = Trim(Text1(18).Text)
            frmWKL84.Show 1
            gsARTNR = ""
        Case 19
            zeigeHilfeDabapfad "LPROTOK", "KVKPR1.txt"
        
        Case 20
            gsARTNR = Trim(Text1(20).Text)
            frmWKL84.Show 1
            gsARTNR = ""
        Case 21
            gsARTNR = Trim(Text1(21).Text)
            frmWKL84.Show 1
            gsARTNR = ""
            
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub startframe6()
    On Error GoTo LOKAL_ERROR
    Dim i As Integer
    
    For i = 8 To 11
        Label7(i).Visible = False
    Next i
    
    Text2(0).Text = ""
    Text2(1).Text = ""
    cboGp.Text = ""
    cboGp.Clear
    Text2(0).SetFocus
    
    Label7(7).Caption = "Ihre Eingabe bitte..."
    Label7(7).Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "startframe6"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speichernGP()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsEAN As Recordset
    
    
    
    If Trim(Text2(0).Text) = "" Then
        Label7(7).Caption = "Sie müssen erst die EAN des Einzelprodukts angeben!"
        Label7(7).Refresh
        Text2(0).SetFocus
        Exit Sub
    End If
    
    If Trim(Text2(0).Text) = Trim(cboGp.Text) Then
        Label7(7).Caption = "Achtung gleiche EAN - Codes werden nicht abgespeichert."
        Label7(7).Refresh
        Exit Sub
    End If
    
    If Trim(cboGp.Text) = "" Then
        Label7(7).Caption = "Bitte geben Sie die Packungs - EAN an!"
        Label7(7).Refresh
        cboGp.SetFocus
        Exit Sub
    End If
    
    If Trim(Text2(1).Text) = "" Then
        Label7(7).Caption = "Bitte geben Sie den Packungsinhalt an!"
        Label7(7).Refresh
        Text2(1).SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(Text2(1).Text) Then
        Label7(7).Caption = "Bitte geben Sie eine Zahl als Packungsinhalt an!"
        Label7(7).Refresh
        Text2(1).SetFocus
        Exit Sub
    End If
    
    sSQL = "Select * from ZUORDEAN where GPEAN = '" & Trim(cboGp.Text) & "'"
    Set rsEAN = gdBase.OpenRecordset(sSQL)
    If rsEAN.EOF Then
        rsEAN.Close
        sSQL = "Insert into ZUORDEAN (EAN,FAKTOR,GPEAN) Values (" & Trim(Text2(0).Text) & ", " & Trim(Text2(1).Text) & " , '" & Trim(cboGp.Text) & "') "
        gdBase.Execute sSQL, dbFailOnError
        
        
    Else
        rsEAN.Close
        sSQL = "Delete from ZUORDEAN where GPEAN = '" & Trim(cboGp.Text) & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into ZUORDEAN (EAN,FAKTOR,GPEAN) Values (" & Trim(Text2(0).Text) & ", " & Trim(Text2(1).Text) & " , '" & Trim(cboGp.Text) & "') "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    startframe6
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichernGP"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speichernGP1()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Delete from ZUORDEAN where ARTNR = " & Trim(Text1(5).Text)
    gdBase.Execute sSQL, dbFailOnError
    
    If Trim(Text2(3).Text) <> "" Then
    
        If Trim(Text2(2).Text) = "" Then Text2(2).Text = "0"

        sSQL = "Insert into ZUORDEAN (ARTNR,EAN,FAKTOR,GPEAN) Values (" & Trim(Text1(5).Text) & ", '" & Trim(Text1(18).Text) & "' ," & Trim(Text2(2).Text) & ", '" & Trim(Text2(3).Text) & "') "
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichernGP1"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speichernLAGERP()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Delete from LAGERPLATZ where ARTNR = " & Trim(Text1(5).Text)
    gdBase.Execute sSQL, dbFailOnError
    
    If Trim(Text2(4).Text) <> "" Then
        sSQL = "Insert into LAGERPLATZ (ARTNR,LAGERP) Values (" & Trim(Text1(5).Text) & ", " & Trim(Text2(4).Text) & ") "
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichernLAGERP"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speichernGruppe()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Delete from Gruppe_Artikel where ARTNR = " & Trim(Text1(5).Text)
    gdBase.Execute sSQL, dbFailOnError
    
    If Trim(Text2(10).Text) <> "" Then
        sSQL = "Insert into Gruppe_Artikel (ARTNR,Gruppennr) Values (" & Trim(Text1(5).Text) & ", " & Trim(Text2(10).Text) & ") "
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichernGruppe"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speichernTEXTIL()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Delete from TEXTIL where ARTNR = " & Trim(Text1(5).Text)
    gdBase.Execute sSQL, dbFailOnError
    
    If Trim(Text2(7).Text) <> "" Or Trim(Text2(8).Text) <> "" Or Trim(Text2(9).Text) <> "" Then
        sSQL = "Insert into TEXTIL (ARTNR,MODELL,MATERIAL,FARBBEZ) Values (" & Trim(Text1(5).Text) & ", '" & Trim(Text2(7).Text) & "', '" & Trim(Text2(8).Text) & "', '" & Trim(Text2(9).Text) & "') "
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichernTEXTIL"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub speichernKONDITIONEN()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    If Trim(Text2(5).Text) <> "" Then
        If Trim(Text2(6).Text) = "" Then Text2(6).Text = "1"
        
        sSQL = "Insert into KONDITIONEN (ARTNR,KONDI,Faktor) Values (" & Trim(Text1(5).Text) & ", " & Trim(Text2(5).Text) & ", " & Trim(Text2(6).Text) & ")"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speichernKONDITIONEN"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub delKONDITIONEN(sArt As String, sKondi As String)
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    
    sSQL = "Delete from KONDITIONEN where ARTNR = " & Trim(sArt)
    If sKondi <> "" Then
        sSQL = sSQL & " and Kondi = " & sKondi
    End If
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "delKONDITIONEN"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub


Private Function fnPruefeEingabeWKL10()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim sSQL As String
    
    fnPruefeEingabeWKL10 = 1
    
    If Trim$(Text1(3).Text) <> "" Then
        If IsNumeric(Trim$(Text1(3).Text)) = False Then
            Text1(3).Text = ""
        End If
    End If
    
    If Trim$(Text1(0).Text) <> "" Then
        If IsNumeric(Trim$(Text1(0).Text)) = False Then
            Text1(0).Text = ""
        End If
    End If
    
    If Trim$(Text1(7).Text) <> "" Then
        If IsNumeric(Trim$(Text1(7).Text)) = False Then
            Text1(7).Text = ""
        End If
    End If
    
    If Trim$(Text1(35).Text) <> "" Then
        If IsNumeric(Trim$(Text1(35).Text)) = False Then
            Text1(35).Text = ""
        End If
    End If
    
    If Trim$(Text1(2).Text) <> "" Then
        If LoeseMarkenstringinLPZ12(Trim$(Text1(2).Text)) = True Then
            fnPruefeEingabeWKL10 = 0
        Else
            Text1(2).Text = ""
        End If
    Else
    
        sSQL = "Delete from  MA" & srechnertab
        SQL_Befehl_ausführen sSQL
        
        
    End If
    
    If Trim$(Text1(7).Text) <> "" Then 'Liefnr
        If IsNumeric(Text1(7).Text) Then
            fnPruefeEingabeWKL10 = 0
            Exit Function
        End If
    End If
    
    If Trim$(Text1(36).Text) <> "" Then
        fnPruefeEingabeWKL10 = 0
        Exit Function
    End If
    
    If Trim$(Text1(44).Text) <> "" Then
        fnPruefeEingabeWKL10 = 0
        Exit Function
    End If
    
    If Trim$(Label1(2).Tag) <> "" Then
        fnPruefeEingabeWKL10 = 0
        Exit Function
    End If
    
    For lcount = 0 To 4
        If Trim$(Text1(lcount).Text) <> "" Then
            fnPruefeEingabeWKL10 = 0
            Exit Function
        End If
    Next lcount
    
    If Trim$(Text1(41).Text) <> "" Then
        fnPruefeEingabeWKL10 = 0
        Exit Function
    End If
    
    If List4.Visible = True And List4.ListCount > 0 Then
        fnPruefeEingabeWKL10 = 0
        Exit Function
    End If
    
    If Trim$(Text1(45).Text) <> "" Then
        fnPruefeEingabeWKL10 = 0
        Exit Function
    End If
    
    If Check12.value = vbChecked Then
        fnPruefeEingabeWKL10 = 0
        Exit Function
    End If
    
    If Check14.value = vbChecked Then
        fnPruefeEingabeWKL10 = 0
        Exit Function
    End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeWKL10"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Private Sub Command1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    Select Case index
    
        Case 10
            If KeyCode = vbKeyF2 Then
                frmWKL49.Show 1
            End If
        
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command10_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    
        cmdfarbe.BackColor = Command10(index).BackColor
        cmdfarbe.Caption = ""
        Text1(34).Text = index
        Frame5.Visible = False
    
    
    
    
    
    
   
       
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command10_Click"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command11_Click()
On Error GoTo LOKAL_ERROR

    'Mail schicken

    gcBestellEmail.Attachment1 = ""
    gcBestellEmail.Attachment2 = ""
    gcBestellEmail.Attachment3 = ""
    gcBestellEmail.Attachment4 = ""
    gcBestellEmail.Attachment5 = ""

    Dim sTemp As String
    
    gcBestellEmail.Subject = "Frage/Stammdatenpflege"
    sTemp = "Die Angaben des Artikels sind falsch oder unvollständig:" & vbCrLf & vbCrLf
    sTemp = sTemp & "Artikelbezeichnung: " & Text1(6).Text & vbCrLf
    sTemp = sTemp & "Lieferant: " & ermLiefBez(CLng(cbo1.Text)) & vbCrLf
    sTemp = sTemp & "BestellNr: " & Text1(19).Text & vbCrLf
    sTemp = sTemp & "VPE: " & Text1(28).Text & vbCrLf
    sTemp = sTemp & "Listen EK: " & Text1(11).Text & vbCrLf
    sTemp = sTemp & "Listen VK: " & Text1(12).Text & vbCrLf
    If Text1(18).Text <> "" Then
        sTemp = sTemp & "EAN: " & Text1(18).Text & vbCrLf
    End If
    
    If Text1(20).Text <> "" Then
        sTemp = sTemp & "EAN: " & Text1(20).Text & vbCrLf
    End If
    
    If Text1(21).Text <> "" Then
        sTemp = sTemp & "EAN: " & Text1(21).Text & vbCrLf
    End If
    
    If Text1(10).Text = "J" Then
        sTemp = sTemp & "EX: Ja " & vbCrLf
    End If
    sTemp = sTemp & vbCrLf
    
    sTemp = sTemp & "meine Anmerkung:"
    
    
    gcBestellEmail.Message = sTemp
    gcBestellEmail.Recipient = ermEmailAdress("Stammdatenpflege")
            
    frmWKL129.Show 1
            
    gcBestellEmail.Attachment1 = ""
    gcBestellEmail.Subject = ""
    gcBestellEmail.Message = ""
    gcBestellEmail.Recipient = ""

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cSuch As String
    Dim iRet As Integer
    Dim cMeld As String
    
    MSFlexGrid1_SelChange
    
    List3.Clear
    List3.Visible = False
    
    List4.Clear
    List4.Visible = False
    
    If index = 4 Then       'Schließen
    
        voreinstellungspeichern10
        
        If MSFlexGrid1.Visible = True Then
            MSFlexGrid1.Visible = False
        End If
        
        Frame0.Visible = True
        Frame1.Visible = False
        If gbBILDTAST = False Then
            Frame2.Visible = False
        Else
            Frame2.Visible = True
        End If

        If NewTableSuchenDBKombi("E10A", gdApp) Then
            voreinstellungladen10A
        End If

        Select Case iFocus
            Case 0, 1, 2, 3, 4, 7, 36, 41, 42
                Text1(iFocus).SetFocus
            Case Else
                Text1(36).SetFocus
        End Select
        
        Exit Sub
    ElseIf index = 1 Then   'Listen
    
        CheckandZeig
        BereiteExportDaten bAusblenden, bBest, bgef, binbest
    
        frmWKL76.Show 1
        Exit Sub
    
    ElseIf index = 5 Then   'temporäre VK-Preise
    
        frmWK10a.Show 1
        Exit Sub
    
    ElseIf index = 6 Then   'Extras Sonderkontitionen
    
        If MSFlexGrid1.Row < 1 Then
            Screen.MousePointer = 0
            MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
            Exit Sub
        End If
        gsARTNR = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr))
        
        frmWKL122.Show 1
        gsARTNR = ""
        Exit Sub
    ElseIf index = 9 Then   'in BV
    
        If MSFlexGrid1.Row < 1 Then
            Screen.MousePointer = 0
            MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
            Exit Sub
        End If
        
        insert_BestVor Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr)), Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR))
    ElseIf index = 21 Then
    
        If Frame2.Visible Then
            Frame2.Visible = False
        Else
            Frame2.Visible = True
        End If
        
        Exit Sub
    
    
    End If
    
    Screen.MousePointer = 11
    
    If MSFlexGrid1.Row < 1 Then
        Screen.MousePointer = 0
        MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr)
    cSuch = Trim$(cSuch)
    
    If IsNumeric(cSuch) Then
    
    Else
        Screen.MousePointer = 0
        MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    Select Case index
        Case 0      'Auswählen
            gbBestandsgrund = True
            Check11.value = vbUnchecked
            HoleDatenWKL10 cSuch
        Case 2      'Löschen
        
            If MSFlexGrid1.RowSel > 1 Then
                cMeld = "ACHTUNG!" & vbCrLf & vbCrLf
                cMeld = cMeld & "Das Löschen eines Artikels kann zu Unstimmigkeiten" & vbCrLf
                cMeld = cMeld & "in der Datenbank führen, wenn der Artikel bereits" & vbCrLf
                cMeld = cMeld & "verkauft wurde!" & vbCrLf & vbCrLf
                cMeld = cMeld & "Wollen Sie den/die Artikel trotzdem löschen?"
                iRet = MsgBox(cMeld, vbYesNo + vbQuestion, "Winkiss Hinweis:")
                If iRet = vbYes Then
                    FlexGrid_Delete MSFlexGrid1
                    
                    MSFlexGrid1.Row = 1
                    MSFlexGrid1.SetFocus
                End If
            Else
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
            End If
            
        Case 10 'Etiketten
            If MSFlexGrid1.RowSel > 1 Then
            
                loeschNEW "LSTEETI", gdBase
                CreateTableT2 "LSTEETI", gdBase
                
                FlexGrid_Etiketten MSFlexGrid1
                
                Dim rsrs As DAO.Recordset
                Dim sSQL As String
                
                Set rsrs = gdBase.OpenRecordset("LSTEETI")
                If Not rsrs.EOF Then
                    rsrs.MoveFirst
                    Do While Not rsrs.EOF
                    
                    rsrs.Edit
                    rsrs!linr = ermLiefLinrmitkleinstenLEKPR(rsrs!artnr, gdBase)
                    rsrs.Update
                    
                    rsrs.MoveNext
                    Loop
                End If
                rsrs.Close: Set rsrs = Nothing
                
                sSQL = "Update LSTEETI inner join Artlief on LSTEETI.Artnr = Artlief.artnr and LSTEETI.linr = Artlief.linr "
                sSQL = sSQL & " set LSTEETI.LIBESNR = Artlief.LIBESNR"
                gdBase.Execute sSQL, dbFailOnError
                
                gsETILS = "aus Lieferschein"
                frmWKL30.Show 1
                
                
                
                
                
                
                
                
                
                
                
                
                MSFlexGrid1.Row = 1
                MSFlexGrid1.SetFocus
                
            Else
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
            End If
        
        Case 3      'Kopieren
            HoleDatenWKL10 cSuch
            Text1(5).Text = HoleFreieArtikelNrWKL10
            
            If Text1(5).Text <> "" Then
                Sicherheitslöschen Text1(5).Text
            End If
            
            Text1(13).Text = ""
            Text1(17).Text = ""
            Text1(18).Text = ""
            Text1(19).Text = ""
            Text1(20).Text = ""
            Text1(21).Text = ""
            Text1(29).Text = "0" 'schnittek auf o
            Label5(0).ForeColor = vbRed
            Label5(0).Visible = True
            Label5(0).Caption = "Kopie eines Artikels!     ArtNr wurde neu ermittelt!"
            gbNew = True
        Case 7      'Shop
        
            Screen.MousePointer = 0
            gsARTNR = cSuch
            frmWKL163.Show 1
            
            If SpaltennummerSHOP > 0 Then
                If LeseInterArt(gsARTNR) Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerSHOP) = "J"
                    MSFlexGrid1.Col = SpaltennummerSHOP
                    MSFlexGrid1.CellForeColor = vbRed
                Else
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerSHOP) = "N"
                    MSFlexGrid1.Col = SpaltennummerSHOP
                    MSFlexGrid1.CellForeColor = vbBlack
                End If
            End If
            gsARTNR = ""
        Case 8      'Gruppieren
        
            If MSFlexGrid1.RowSel > 1 Then
                
                frmWKL180.Show 1
                If Val(gsGRUPPENNR) > 0 Then
                    FlexGrid_Gruppieren MSFlexGrid1, Val(gsGRUPPENNR)
                End If
                gsGRUPPENNR = ""
                
                MSFlexGrid1.Row = 1
                MSFlexGrid1.SetFocus
                
            Else
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
            End If
            
            
            
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LoescheArtikelWKL10(cArtNr As String)
    On Error GoTo LOKAL_ERROR
        
    Dim cSQL As String
    
    SicherInArtikelsic CLng(cArtNr)
    
    cSQL = "Delete from ARTIKEL where ARTNR = " & cArtNr & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from Artlief where ARTNR = " & cArtNr & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from ARTEAN_K where ARTNR = " & cArtNr & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from TOP" & srechnertab & " where ARTNR = " & cArtNr & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from TOP" & srechnertab & " where ARTNR = " & cArtNr & " "
    gdBase.Execute cSQL, dbFailOnError
    
    schreibeProtokollgArtikel "Artikel: " & cArtNr & " " & ErmittleDetails(cArtNr) & " wurde gelöscht."
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheArtikelWKL10"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function UpdateArtikelWKL10(cArtNr As String, nCol As Long, sWert As String) As Boolean
    On Error GoTo LOKAL_ERROR
        
    Dim cSQL As String
    Dim sSpalte As String
    
    UpdateArtikelWKL10 = False
        
    sSpalte = sSpaltenbez(nCol)
    
    Select Case UCase(sSpalte)
    
        Case "AGN"
        
            If Val(sWert) > 99999 Then
                Exit Function
            End If
            
            UpdateArtikelWKL10 = True
            
            cSQL = "Update ARTIKEL set " & sSpalte & " = " & Val(sWert) & " where ARTNR = " & cArtNr & " "
            gdBase.Execute cSQL, dbFailOnError
            
        Case "RABATT_OK"
        
            If UCase(sWert) <> "J" And UCase(sWert) <> "N" Then
                Exit Function
            End If
        
            UpdateArtikelWKL10 = True
            
            cSQL = "Update ARTIKEL set " & sSpalte & " = '" & UCase(sWert) & "' where ARTNR = " & cArtNr & " "
            gdBase.Execute cSQL, dbFailOnError
            
        Case "GEFUEHRT"
        
            If UCase(sWert) <> "J" And UCase(sWert) <> "N" Then
                Exit Function
            End If
        
            UpdateArtikelWKL10 = True
            
            cSQL = "Update ARTIKEL set " & sSpalte & " = '" & UCase(sWert) & "' where ARTNR = " & cArtNr & " "
            gdBase.Execute cSQL, dbFailOnError
            
        Case "LAGERP"
        
            If Val(sWert) = 0 Then
                cSQL = "Delete * from Lagerplatz where artnr = " & cArtNr
                gdBase.Execute cSQL, dbFailOnError
            End If
            
            If Val(sWert) > 0 Then
            
            
                cSQL = "Delete * from Lagerplatz where artnr = " & cArtNr
                gdBase.Execute cSQL, dbFailOnError
            
                cSQL = "Insert into LAGERPLATZ (ARTNR,LAGERP) Values (" & cArtNr & ", " & Val(sWert) & ") "
                gdBase.Execute cSQL, dbFailOnError
            End If
            
            UpdateArtikelWKL10 = True
            
    End Select
    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UpdateArtikelWKL10"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Gruppiere_ArtikelWKL10(cArtNr As String, lGruppArtikel As Long)
    On Error GoTo LOKAL_ERROR
        
    Dim cSQL As String
    
    cSQL = "Delete from GRUPPE_ARTIKEL where GRUPPENNR = " & lGruppArtikel & " "
    cSQL = cSQL & " and ARTNR = " & cArtNr & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into GRUPPE_ARTIKEL (GRUPPENNR,ARTNR) values (" & lGruppArtikel & "," & cArtNr & ")"
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Gruppiere_ArtikelWKL10"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    
    iZielIndex = Label3(2).Caption
    
    Select Case iZielIndex
        Case Is = -3
            cbo1.Text = cbo1.Text & Command3(index).Caption
            cbo1.SetFocus
        Case 11, 29, 12, 30, 14, 45, 46 'Preise 'Hier die doppelten Kommas prüfen
            If Command3(index).Caption = "," Then
                If InStr(Text1(iZielIndex).Text, ",") > 0 Then
                    Text1(iZielIndex).SetFocus
                Else
                    Text1(iZielIndex).Text = Text1(iZielIndex).Text & Command3(index).Caption
                    Text1(iZielIndex).SetFocus
                End If
            Else
                Text1(iZielIndex).Text = Text1(iZielIndex).Text & Command3(index).Caption
                Text1(iZielIndex).SetFocus
            End If
        Case Else
            Text1(iZielIndex).Text = Text1(iZielIndex).Text & Command3(index).Caption
            Text1(iZielIndex).SetFocus
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command3_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    
    iZielIndex = Label3(2).Caption
    If Frame0.Visible Then
        Select Case iZielIndex
            Case Is = -3
            Case Is = -2
            Case Is = 2

            Case Else
                Text1(iZielIndex).BackColor = glSelBack1
        End Select
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command30_Click()
    On Error GoTo LOKAL_ERROR
    
    Text1(47).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
    Text1(47).SetFocus
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command30_Click"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command4_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    iZielIndex = Label3(2).Caption
    
    Set WshShell = CreateObject("WScript.Shell")
    'WshShell.SendKeys "+{Tab}", True
        
    Select Case index
        Case Is = 0     'CLEAR
            Select Case iZielIndex
                Case Is = -3
                    cbo1.Text = ""
                Case Else
                    Text1(iZielIndex).Text = ""
            End Select
            
        Case Is = 1     'ENTER
            Select Case iZielIndex
                Case Is = -3
                    If Len(cbo1.Text) > 0 Then
                        cbo1.Text = Left(cbo1.Text, Len(cbo1.Text) - 1)
                    End If
                Case Else
                    If Len(Text1(iZielIndex).Text) > 0 Then
                        Text1(iZielIndex).Text = Left(Text1(iZielIndex).Text, Len(Text1(iZielIndex).Text) - 1)
                    End If
            End Select
        Case Is = 2     'BEFORE
            WshShell.SendKeys "+{Tab}", True
        Case Is = 3     'NEXT
            WshShell.SendKeys "{Tab}", True
        Case Is = 4     'Switch UPPER / lower
            SwitchUpperLowerCaseWKL10
    End Select
    
    
    
    Select Case iZielIndex
        Case Is = -3
            cbo1.SetFocus
        Case Else
            Text1(iZielIndex).SetFocus
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    

End Sub
Private Sub SwitchUpperLowerCaseWKL10()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    If Left(Command4(4).Caption, 1) = "A" Then
        For lcount = 22 To 32
            Command3(lcount).Caption = LCase(Command3(lcount).Caption)
        Next lcount
        For lcount = 35 To 45
            Command3(lcount).Caption = LCase(Command3(lcount).Caption)
        Next lcount
        For lcount = 47 To 55
            Command3(lcount).Caption = LCase(Command3(lcount).Caption)
        Next lcount
        Command3(54).Caption = ","
        Command3(55).Caption = "."
        Command3(56).Caption = "-"
        
        Command4(4).Caption = "a -> A"
    Else
        For lcount = 22 To 32
            Command3(lcount).Caption = UCase(Command3(lcount).Caption)
        Next lcount
        For lcount = 35 To 45
            Command3(lcount).Caption = UCase(Command3(lcount).Caption)
        Next lcount
        For lcount = 47 To 55
            Command3(lcount).Caption = UCase(Command3(lcount).Caption)
        Next lcount
        
        Command3(54).Caption = ";"
        Command3(55).Caption = ":"
        Command3(56).Caption = "_"
        
        Command4(4).Caption = "A -> a"
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SwitchUpperLowerCaseWKL10"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    
    iZielIndex = Label3(2).Caption
    Select Case iZielIndex
        Case Is = -3
            cbo1.BackColor = glSelBack1
        Case Else
            Text1(iZielIndex).BackColor = glSelBack1
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command5_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    Dim bGespeichert    As Boolean
    Dim lRet            As Long
    Dim cPos(1 To 4)    As Long
    Dim i               As Integer
    Dim cTxT            As String
    Dim sArtnr          As String
    Dim sLifnr          As String
    Dim sDelLifnr       As String
    Dim sSQL            As String
    Dim iRet            As Integer
    Dim ctmp            As String
    
    bGespeichert = False
    
    Select Case index
    
        Case 6 'Größenansicht
            Screen.MousePointer = 0
            frmWKL181.Show 1
        Case 5 'ab in einen Bestellvorschlag
            insert_BestVor Text1(5).Text, cbo1.Text
            
            Dim linBV As Long
            
            linBV = 0
            linBV = ermINBV(Text1(5).Text, cbo1.Text)
            
            If linBV > 0 Then
                Command5(5).Caption = "in BV(" & linBV & ")"
                Command5(5).ForeColor = vbRed
            Else
                Command5(5).Caption = "in BV"
                Command5(5).ForeColor = glS1
            End If
            
            
            
        Case 0    '** Speichern **

            Frame5.Visible = False
            
            If Text1(18).Text = "0" Then Text1(18).Text = ""
            If Text1(20).Text = "0" Then Text1(20).Text = ""
            If Text1(21).Text = "0" Then Text1(21).Text = ""
            Text1(6).Text = SwapStr(Text1(6).Text, "'", " ")
            Text1(6).Text = SwapStr(Text1(6).Text, ";", " ")
            Text1(6).Text = SwapStr(Text1(6).Text, ",", " ")
            Text1(6).Text = SwapStr(Text1(6).Text, "*", " ")
            
            lRet = fnPruefeLINRWKL10()
            If lRet <> 0 Then
                MsgBox "Die eingegebene Lieferantennummer ist unbekannt!" & vbCrLf & "Bitte unter STAMMDATEN -> Lieferanten bearbeiten!", vbInformation, "Winkiss Hinweis:"
                cbo1.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            If Val(Text1(30).Text) > 99999 Then
                MsgBox "Bitte überprüfen Sie den Kassenverkaufspreis!", vbOKOnly + vbInformation, "Winkiss Hinweis:"
                Text1(30).SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            lRet = fnPruefeDialogEingabenWKL10()
            If lRet = 0 Then
                lRet = fnPruefeAGNWKL10()
                If lRet <> 0 Then
                    MsgBox "Die eingegebene AGN ist unbekannt!" & vbCrLf & "Bitte unter STAMMDATEN -> ARTIKELGRUPPEN nachdefinieren!", vbInformation, "UNBEKANNTER WERT"
                End If
                glBestandNeu = Val(Text1(13).Text)
                If glBestandNeu < glBestandAlt Then
                    If glLevel < 7 Then
                        MsgBox "Mengen-Reduzierung nicht möglich!" & vbCrLf & vbCrLf & "Bestandsminderungen sind nur mit Zugriffs-Level 7 oder höher erlaubt!", vbInformation, "INFO"
                        Exit Sub
                    End If
                End If
                
                If checkthisean(Trim$(Text1(18).Text), Trim$(Text1(5).Text)) = True Then

                Else
                    Text1(18).SetFocus
                    Screen.MousePointer = 0
                    
                    ctmp = "Diese EAN - Nummer ist schon vergeben. Diese kann kein zweites Mal gespeichert werden." & vbCrLf & vbCrLf
                    ctmp = ctmp & "Möchten Sie diese EAN trotzdem vergeben?"
                    iRet = MsgBox(ctmp, vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
                    
                    If iRet = vbYes Then
                        sSQL = "Update Artikel set ean = '' where ean = '" & Trim$(Text1(18).Text) & "'": gdBase.Execute sSQL, dbFailOnError
                        sSQL = "Update Artikel set ean2 = '' where ean2 = '" & Trim$(Text1(18).Text) & "'": gdBase.Execute sSQL, dbFailOnError
                        sSQL = "Update Artikel set ean3 = '' where ean3 = '" & Trim$(Text1(18).Text) & "'": gdBase.Execute sSQL, dbFailOnError
                    Else
                        Exit Sub
                    End If
                    
                    
                End If
                
                If checkthisean(Trim$(Text1(20).Text), Trim$(Text1(5).Text)) = True Then

                Else
                    Text1(20).SetFocus
                    Screen.MousePointer = 0
                    ctmp = "Diese EAN - Nummer ist schon vergeben. Diese kann kein zweites Mal gespeichert werden." & vbCrLf & vbCrLf
                    ctmp = ctmp & "Möchten Sie diese EAN trotzdem vergeben?"
                    iRet = MsgBox(ctmp, vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
                    
                    If iRet = vbYes Then
                        sSQL = "Update Artikel set ean = '' where ean = '" & Trim$(Text1(20).Text) & "'": gdBase.Execute sSQL, dbFailOnError
                        sSQL = "Update Artikel set ean2 = '' where ean2 = '" & Trim$(Text1(20).Text) & "'": gdBase.Execute sSQL, dbFailOnError
                        sSQL = "Update Artikel set ean3 = '' where ean3 = '" & Trim$(Text1(20).Text) & "'": gdBase.Execute sSQL, dbFailOnError
                    Else
                        Exit Sub
                    End If
                End If
                
                If checkthisean(Trim$(Text1(21).Text), Trim$(Text1(5).Text)) = True Then

                Else
                    Text1(21).SetFocus
                    Screen.MousePointer = 0
                    ctmp = "Diese EAN - Nummer ist schon vergeben. Diese kann kein zweites Mal gespeichert werden." & vbCrLf & vbCrLf
                    ctmp = ctmp & "Möchten Sie diese EAN trotzdem vergeben?"
                    iRet = MsgBox(ctmp, vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
                    
                    If iRet = vbYes Then
                        sSQL = "Update Artikel set ean = '' where ean = '" & Trim$(Text1(21).Text) & "'": gdBase.Execute sSQL, dbFailOnError
                        sSQL = "Update Artikel set ean2 = '' where ean2 = '" & Trim$(Text1(21).Text) & "'": gdBase.Execute sSQL, dbFailOnError
                        sSQL = "Update Artikel set ean3 = '' where ean3 = '" & Trim$(Text1(21).Text) & "'": gdBase.Execute sSQL, dbFailOnError
                    Else
                        Exit Sub
                    End If
                End If
                
                If gbNewArt Then
                    If gbNew = True Then
                        Dim ierg As Integer
                        ierg = Artfrei(Text1(5).Text)
                        If ierg = 0 Then
                             
                        ElseIf ierg = 1 Then
                            Screen.MousePointer = 0
                            Text1(5).SetFocus
                            MsgBox "Diese Artikelnummer ist schon vergeben. Diese kann kein zweites Mal gespeichert werden.", vbInformation, "Winkiss Hinweis:"
                        
                            Exit Sub
                            
                        ElseIf ierg = 2 Then
                            Screen.MousePointer = 0
                            Text1(5).SetFocus
                            ctmp = "Diese Artikelnummer wurde bereits verwendet und besitzt Vergangenheitsdaten. Diese Artikelnummer kann kein weiteres Mal verwendet werden." & vbCrLf & vbCrLf
                            ctmp = ctmp & "Nachfolgend erhalten Sie Tipps wie Sie diesen Fehler abstellen können."
                            MsgBox ctmp, vbInformation, "Winkiss Hinweis:"
                        
                            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/hilfe-bei-problemen/44-software-probleme-winkiss/242-artikelnummer-wird-bereits-verwendet.html"
                            
                            
                            Exit Sub
                        End If
                    End If
                Else
                    Screen.MousePointer = 0
                    MsgBox "Keine Artikelanlage möglich (Voreinstellung)", vbInformation, "Winkiss Hinweis:"
                    Exit Sub
                End If
                bBeimSpeichern = True
                SchreibeDatenWKL10
                bBeimSpeichern = False
                
                If gbcomefromwoa = True Then
                    Unload frmWKL10
                    Exit Sub
                End If
                bGespeichert = True
                DoEvents
                
                If gbNew = True Then
                    
                    If gbNewArtNrVorschlag Then
                        Text1(5).Text = HoleFreieArtikelNrWKL10
                    Else
                        Text1(5).Text = ""
                    End If
                Else
                    LeereArtikelDatenWKL10 True

                    If gbNewArtNrVorschlag Then
                        Text1(5).Text = HoleFreieArtikelNrWKL10
                    End If

                    gbNew = True
                End If
                Label5(0).Visible = False
            Else
                If lRet = 18 Or lRet = 20 Or lRet = 21 Then
                    MsgBox "Die eingegebene EAN ist ungültig!", vbCritical, "STOP!"
                    Text1(lRet).SetFocus
                Else
                    If lRet = 99 Then
                        MsgBox "Bitte ArtNr und ArtBez. angeben!", vbCritical, "STOP!"
                        Text1(5).SetFocus
                    Else
                        Text1(lRet).SetFocus
                    End If
                End If
            End If
            
        Case Is = 1
        
            If gbcomefromwoa = True Then
'                LogtoEnd Me
                Unload frmWKL10
                Exit Sub
            End If
            
            If gbNew = False And Text1(5).Text <> "" Then
                lRet = fnPruefeLINRWKL10()
                If lRet <> 0 Then
                    MsgBox "Die eingegebene Lieferantennummer ist unbekannt!" & vbCrLf & "Bitte unter STAMMDATEN -> Lieferanten bearbeiten!", vbInformation, "Winkiss Hinweis:"
                    cbo1.SetFocus
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
            
            lblLiefbez.Caption = ""
            lblLiefbez.Refresh
            
            Frame3.Visible = False
            Frame5.Visible = False
            
            Label5(0).Visible = False
            gbNew = False
            If MSFlexGrid1.Rows = 2 Then
                Frame1.Visible = False
                Frame0.Visible = True
                Text1(1).SetFocus
            Else
                Frame2.Visible = False
                Frame1.Visible = True

                If bGespeichert Then
                    Command1_Click 0
                Else
                    If MSFlexGrid1.Visible = True Then
                        MSFlexGrid1.SetFocus
                    Else
                        Frame1.Visible = False
                        Frame0.Visible = True
                        Text1(1).SetFocus
                    End If
                End If
                giDlgZustand = giUPD
            End If
            
            LeereArtikelDatenWKL10 True
            
        Case Is = 2  '//neue Lieferantenzuordnung
        
            cbo1.Text = ""
            Text1(11).Text = ""
            Text1(28).Text = ""
            Text1(19).Text = ""
            Text1(10).Text = ""
            Label4(34).Caption = ""
            cbo1.SetFocus
            lblLiefbez.Caption = "Geben Sie bitte die Lieferantendaten für diesen Artikel ein!"
            lblLiefbez.Refresh
            
            Line1(0).BorderColor = glWarn
            Line1(2).BorderColor = glWarn
            Line1(3).BorderColor = glWarn
            Line1(1).BorderColor = glWarn
            
            

        Case Is = 3   'LIEFERANTENZUORDNUNG LÖSCHEN
        
            sArtnr = Trim(Text1(5).Text)
            sDelLifnr = Trim(cbo1.Text)
            
            If sArtnr <> "" And sDelLifnr <> "" Then
                sSQL = "Delete from artlief where artnr = " & sArtnr
                sSQL = sSQL & " and linr = " & Val(sDelLifnr)
                gdBase.Execute sSQL, dbFailOnError
                
                cbo1fuellen sArtnr
                Liefdetail sArtnr
            End If
            
            If cbo1.Text = "" Then
                cbo1.Text = ""
                Text1(11).Text = ""
                Text1(28).Text = ""
                Text1(19).Text = ""
                Text1(10).Text = ""
                Label4(34).Caption = ""
                cbo1.SetFocus
                lblLiefbez.Caption = "Geben Sie bitte die Lieferantendaten für diesen Artikel ein!"
                lblLiefbez.Refresh
                
                Line1(0).BorderColor = glWarn
                Line1(2).BorderColor = glWarn
                Line1(3).BorderColor = glWarn
                Line1(1).BorderColor = glWarn
                
            Else
                sLifnr = Trim(cbo1.Text)
                sSQL = "Update artikel set linr = " & Val(sLifnr) & " where artnr = " & sArtnr & " And linr = " & Val(sDelLifnr) & ""
                gdBase.Execute sSQL, dbFailOnError
            End If
        Case 4 ' Etikett in den Pool
        
            schreibe_Etikett_einzeln Text1(5).Text
        
        Case 11
            gsHelpstring = "Artikel bearbeiten"
            frmWKL110.Show 1
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub schreibe_Etikett_einzeln(sArtnr As String)
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsEti As DAO.Recordset
    Dim ctmp As String
    Dim dWert As Double
    
    If Val(sArtnr) = 0 Then
        Exit Sub
    End If

    sSQL = "Select * from ETIDRU where ARTNR = " & sArtnr
    sSQL = sSQL & " and FILNR = " & gcFilNr
    Set rsEti = gdBase.OpenRecordset(sSQL)
    
    If Not rsEti.EOF Then
        rsEti.Edit
    Else
        rsEti.AddNew
    End If
    rsEti!artnr = sArtnr
    rsEti!BEZEICH = Trim$(Text1(6).Text)
    
    ctmp = Trim$(Text1(30).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    dWert = Val(ctmp)
    rsEti!vkpr = dWert
    
    ctmp = Trim$(Text1(13).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    dWert = Val(ctmp)
    rsEti!BESTAND = dWert
    rsEti!ANZAHL = dWert
    rsEti!LIBESNR = Trim$(Text1(19).Text)
    rsEti!EAN = Trim$(Text1(18).Text)
    rsEti!linr = Trim$(cbo1.Text)
    rsEti!LPZ = Val(Trim$(Text1(8).Text))
    rsEti!filnr = Val(gcFilNr)
    rsEti!Pcname = srechnertab
    rsEti.Update
    rsEti.Close: Set rsEti = Nothing
   
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "schreibe_Etikett_einzeln"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command6_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    Dim cLBSatz As String
    Dim cLinr As String
    Dim cEkPr As String
    Dim cMinMen As String
    
    Dim iRet As Integer
    
    Select Case index
        Case 0
            gsARTNR = Text1(5).Text
            frmWKL78.Show 1
            gsARTNR = ""
        Case 1
            gsARTNR = Text1(5).Text
            frmWKL79.Show 1
            gsARTNR = ""
        Case 2
            gsARTNR = Text1(5).Text
            frmWKL80.Show 1
            gsARTNR = ""
        Case Is = 3     'Schließen
            Frame4.Visible = False
            Frame3.Enabled = True
            Text1(5).SetFocus
        Case Is = 4
            delKONDITIONEN Text1(5).Text, Text2(5).Text
            speichernKONDITIONEN
            fuellecombo1 Text1(5).Text
        Case Is = 5
            delKONDITIONEN Text1(5).Text, Text2(5).Text
            
            Text2(5).Text = ""
            Text2(6).Text = ""
            fuellecombo1 Text1(5).Text
            
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command7_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim cZeichen As String
    Dim cZiel As String
    Dim cValid As String
    Dim lcount As Long
    Dim se As String
    
    Screen.MousePointer = 11
    cValid = "1234567890,"
    
    'Anzeigesteuerung:
    'ist gdRechner(0) = 0, dann wird die Umrechnung
    'aus frmWKL02 nicht übernommen
    'ist gdRechner(0) = 1, dann wird die Umrechnung
    'aus frmWKL02 übernommen
    
    gdRechner(0) = 0
    
    If gsSpanne = "LEK" Then
        se = Text1(11).Text
    ElseIf gsSpanne = "SEK" Then
        se = Text1(29).Text
    End If
    
    'LEK oder SEK
    ctmp = se
    cZiel = ""
    For lcount = 1 To Len(ctmp)
        cZeichen = Mid(ctmp, lcount, 1)
        If InStr(cValid, cZeichen) > 0 Then
            cZiel = cZiel & cZeichen
        End If
    Next lcount
    ctmp = cZiel
    ctmp = fnMoveComma2Point(ctmp)
    gdRechner(1) = Val(ctmp)
    
    'VKPR
    ctmp = Text1(12).Text
    cZiel = ""
    For lcount = 1 To Len(ctmp)
        cZeichen = Mid(ctmp, lcount, 1)
        If InStr(cValid, cZeichen) > 0 Then
            cZiel = cZiel & cZeichen
        End If
    Next lcount
    ctmp = cZiel
    ctmp = fnMoveComma2Point(ctmp)
    gdRechner(2) = Val(ctmp)
    
    'KVKPR
    ctmp = Text1(30).Text
    cZiel = ""
    For lcount = 1 To Len(ctmp)
        cZeichen = Mid(ctmp, lcount, 1)
        If InStr(cValid, cZeichen) > 0 Then
            cZiel = cZiel & cZeichen
        End If
    Next lcount
    ctmp = cZiel
    ctmp = fnMoveComma2Point(ctmp)
    gdRechner(3) = Val(ctmp)

    frmWKL02.Show 1
    
    If gdRechner(0) = 1 Then
        Text1(12).Text = Format$(gdRechner(2), "#####0.00")
        Text1(30).Text = Format$(gdRechner(3), "#####0.00")
        Text1(14).Text = Format$(gdRechner(4), "#####0.00")
        If gdRechner(4) = 0 Then
            Label4(30).Visible = False
            Label4(30).ForeColor = glS1
            Label4(30).Refresh
        
        Else
            Label4(30).Visible = True
            Label4(30).ForeColor = vbRed
            Label4(30).Refresh
        End If
        
    End If
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command8_Click()
    On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11
    
    gcArtNrFiliale = Trim$(Str$(Val(Text1(5).Text)))
    frmWKLae.Show 1
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command9_Click()
    On Error GoTo LOKAL_ERROR
    
    gcArtNrFiliale = Text1(5).Text
    frmWKLam.Show 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command9_Click"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command98_Click()
On Error GoTo LOKAL_ERROR
    
    gsZSpalte = "Artnr"
    gsZSpalte1 = "linr"
    gstab = "BEAART"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command98_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub
Private Sub voreinstellungladen10()
    On Error GoTo LOKAL_ERROR

    Dim rs As Recordset

    Set rs = gdApp.OpenRecordset("E10")
    If Not rs.EOF Then
        Option1(0).value = rs!bo0
        Option1(1).value = rs!bo1
        Option1(2).value = rs!bo2
        Option1(3).value = rs!bo3
        Option1(4).value = rs!bo4
        Option1(5).value = rs!bo5
        Option1(6).value = rs!bo6
        
        If rs!bo7 = True Then
            Check2.value = vbUnchecked
            Check6.value = vbUnchecked
            Check6.ForeColor = glS1
        Else
            Check2.value = vbChecked
            Check6.value = vbChecked
            Check6.ForeColor = vbRed
        End If

        If rs!bo8 = True Then
            Check1.value = vbUnchecked
            Check7.value = vbUnchecked
            Check7.ForeColor = glS1
        Else
            Check1.value = vbChecked
            Check7.value = vbChecked
            Check7.ForeColor = vbRed
        End If
        
        If rs!bo9 = True Then
            Check3.value = vbUnchecked
            Check8.value = vbUnchecked
            Check8.ForeColor = glS1
        Else
            Check3.value = vbChecked
            Check8.value = vbChecked
            Check8.ForeColor = vbRed
        End If
        
        If rs!bo10 = True Then
            Check4.value = vbUnchecked
        Else
            Check4.value = vbChecked
        End If
        
        If rs!bo11 = True Then
            Check11.value = vbUnchecked
        Else
            Check11.value = vbChecked
        End If
        
        
        
    End If
    rs.Close: Set rs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen10"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungladen10A()
    On Error GoTo LOKAL_ERROR

    Dim rs As Recordset

    Set rs = gdApp.OpenRecordset("E10A")
    If Not rs.EOF Then
        If Not IsNull(rs!bo10) Then
            iFocus = rs!bo10
        Else
            iFocus = 36
        End If
    End If
    rs.Close: Set rs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen10A"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern10()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    Dim bo0 As Integer
    Dim bo1 As Integer
    Dim bo2 As Integer
    Dim bo3 As Integer
    Dim bo4 As Integer
    Dim bo5 As Integer
    Dim bo6 As Integer
    Dim bo7 As Integer
    Dim bo8 As Integer
    Dim bo9 As Integer
    
    Dim bo10 As Integer
    Dim bo11 As Integer

    loeschNEW "E10", gdApp
    CreateTable "E10", gdApp

    bo0 = Option1(0).value
    bo1 = Option1(1).value
    bo2 = Option1(2).value
    bo3 = Option1(3).value
    bo4 = Option1(4).value
    bo5 = Option1(5).value
    bo6 = Option1(6).value

    If Check6.value = vbChecked Then
        bo7 = 0
    Else
        bo7 = -1
    End If
    
    If Check7.value = vbChecked Then
        bo8 = 0
    Else
        bo8 = -1
    End If
    
    If Check8.value = vbChecked Then
        bo9 = 0
    Else
        bo9 = -1
    End If
    
    If Check4.value = vbChecked Then
        bo10 = 0
    Else
        bo10 = -1
    End If
    
    If Check11.value = vbChecked Then
        bo11 = 0
    Else
        bo11 = -1
    End If

    sSQL = "Insert into E10 ( bo0,bo1,bo2,bo3,bo4,bo5,bo6,bo7,bo8,bo9,bo10,bo11) "
    sSQL = sSQL & " values (" & bo0 & "," & bo1 & "," & bo2 & "," & bo3 & "," & bo4
    sSQL = sSQL & " ," & bo5 & "," & bo6 & "," & bo7 & "," & bo8 & "," & bo9 & "," & bo10 & "," & bo11
    sSQL = sSQL & " )"
    gdApp.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern10"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Handelsspanne_anzeigen(sArt As String, cMwst As String, sKVK As String, cLinr As String)
On Error GoTo LOKAL_ERROR

    'Handelsspanne anzeigen
    
    
    Dim sNettospanne    As String
    Dim dNettospanne    As Double
    Dim sKVKPR          As String
    Dim sEKpr           As String
    Dim sEKsql          As String
    
    
    Dim sSQL            As String
    Dim rsrs            As DAO.Recordset
    Dim sLiefBez        As String
    
    Shape1.Visible = False
    Label39(3).Caption = ""
    
    
    If gsSpanne = "LEK" Then

        'größten LEK suchen
        sSQL = "Select lekpr from artlief where artnr = " & sArt & " and linr = " & cLinr
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            sEKpr = "0"
            If Not IsNull(rsrs!lekpr) Then
                sEKpr = rsrs!lekpr
            End If
        End If
        rsrs.Close
        
        
        sLiefBez = ""
        sLiefBez = ermLiefBez(CLng(cLinr))
            
        
        
        sNettospanne = NettospanneInProzent(sKVK, sEKpr, cMwst)
        dNettospanne = CDbl(sNettospanne)
        
        If dNettospanne >= 100 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &H8000&
        ElseIf dNettospanne > 79.99 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &HC000&
        ElseIf dNettospanne > 59.99 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &HFF00&
        ElseIf dNettospanne > 39.99 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &HFFFF&
        ElseIf dNettospanne > 19.99 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &H80C0FF
        ElseIf dNettospanne > 0 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &H80FF&
        ElseIf dNettospanne <= 0 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &HFF&
        End If
        
        Shape1.Visible = True
        
        
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Handelsspanne_anzeigen"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub dyn_Handelsspanne_anzeigen(cMwst As String, sKVK As String, sLiefBez As String, sLEK As String, sSEK As String)
On Error GoTo LOKAL_ERROR

    'Handelsspanne anzeigen
    
    If sLiefBez = "" Then
        Exit Sub
    End If
    
    Dim sNettospanne    As String
    Dim dNettospanne    As Double
    Dim sKVKPR          As String
    Dim sEKpr           As String
    
    Shape1.Visible = False
    Label39(3).Caption = ""
    
        sEKpr = sLEK
        
        sNettospanne = NettospanneInProzent(sKVK, sEKpr, cMwst)
        dNettospanne = CDbl(sNettospanne)
        
        If dNettospanne >= 100 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &H8000&
        ElseIf dNettospanne > 79.99 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &HC000&
        ElseIf dNettospanne > 59.99 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &HFF00&
        ElseIf dNettospanne > 39.99 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &HFFFF&
        ElseIf dNettospanne > 19.99 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &H80C0FF
        ElseIf dNettospanne > 0 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &H80FF&
        ElseIf dNettospanne <= 0 Then
            Label39(3).Caption = Fix(dNettospanne)
            Label39(3).ToolTipText = "Nettospanne: " & sNettospanne & "% (Basis: Listen-EK: " & Format(sEKpr, "###,##0.00") & "€ " & sLiefBez & ")"
            Shape1.BackColor = &HFF&
        End If
        
        Shape1.Visible = True
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "dyn_Handelsspanne_anzeigen"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern10A()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim bo10 As Integer

    loeschNEW "E10A", gdApp
    CreateTable "E10A", gdApp
    
    Select Case iFocus
        Case 0, 1, 2, 3, 4, 7, 36, 41, 42
            bo10 = iFocus
        Case Else
            bo10 = 36
    End Select

    sSQL = "Insert into E10A ( bo10) "
    sSQL = sSQL & " values (" & bo10
    sSQL = sSQL & " )"
    gdApp.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern10A"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    
    Screen.MousePointer = 11
    
    gbcomefromwoa = False
    bBeimSpeichern = False
    
    WKL10Positionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Label0(4)
    
    Me.WindowState = 2
    
    Label1(13).BackColor = vbRed
    
    For i = 0 To 9
        Command10(i).BackColor = glfarbe(i)
    Next i
    
    For i = 1 To 9
        Command10(i + 10).BackColor = glfarbe2(i)
    Next i
    
    Command1(32).BackColorFrom = vbWhite
    Command1(32).BackColorTo = vbWhite
    
    Command1(33).BackColorFrom = vbWhite
    Command1(33).BackColorTo = vbWhite
    
    Command1(32).HoverColorFrom = vbWhite
    Command1(33).HoverColorFrom = vbWhite
    Command1(32).HoverColorTo = vbWhite
    Command1(33).HoverColorTo = vbWhite
    
    Command1(32).DownColorFrom = vbWhite
    Command1(32).DownColorTo = vbWhite
    
    Command1(33).DownColorFrom = vbWhite
    Command1(33).DownColorTo = vbWhite


    Me.Refresh
    
    anzeige "normal", "", Label1(59)
    
    gbNew = False
    
    If gbBILDTAST = False Then
        Frame2.Visible = False
    Else
        Frame2.Visible = True
    End If
    
    Frame0.Visible = True
    LeereDialogWKL10
    
    VorBereitLagerumschlag
    VorBereitLagerumschlagVJ
    VorBereitLagerumschlagVVJ
    
    If NewTableSuchenDBKombi("E10", gdApp) Then
    
        If SpalteInTabellegefundenNEW("E10", "bo10", gdApp) = False Then
            SpalteAnfuegenNEW "E10", "bo10", "BIT", gdApp
        End If
        
        If SpalteInTabellegefundenNEW("E10", "bo11", gdApp) = False Then
            SpalteAnfuegenNEW "E10", "bo11", "BIT", gdApp
        End If
    
        voreinstellungladen10
    End If
    
    If Trim(gsARTNR) <> "" Then
        gbcomefromwoa = True
        Frame0.Visible = False
        HoleDatenWKL10 gsARTNR
        Me.Refresh
        gsARTNR = ""
    End If
    
    If NewTableSuchenDBKombi("E10A", gdApp) Then
        voreinstellungladen10A
    End If
    
    If gbNewArt = True Then
        Command1(1).Enabled = True
        Command2(3).Enabled = True
    Else
        Command1(1).Enabled = False
        Command2(3).Enabled = False
    End If
    
    Text1(iFocus).TabIndex = 0
    
    Shape1.BorderColor = Line1(0).BorderColor
    Label39(3).ForeColor = vbBlack
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "TOP" & srechnertab, gdBase
    loeschNEW "UMS_ARTNR" & srechnertab, gdBase
    
    gbBestandsgrund = True
    Check11.value = vbUnchecked
    Erase gLayout
    LogtoEnd Me
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label4(30).Caption = "auto Kalkulation = ja"

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame3_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_DblClick(index As Integer)
On Error GoTo LOKAL_ERROR

If index = 2 Then
    Label1(index).Caption = "alle Farben"
    Label1(index).Tag = ""
    Label1(index).BackColor = Label1(4).BackColor
    Label1(index).ForeColor = Label1(4).ForeColor
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label3_Change(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    If index = 2 Then
        
        Text1(1).BackColor = vbWhite

        
        For lcount = 3 To 6
            Text1(lcount).BackColor = vbWhite
        Next lcount
        
        For lcount = 8 To 19
            Text1(lcount).BackColor = vbWhite
        Next lcount

    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label3_Change"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label4_Click(index As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim sSQL As String
    
    Select Case index
    
        Case 30
            iRet = MsgBox("Möchten Sie die automatische Kalkulation aufheben?", vbQuestion + vbYesNo, "Winkiss Frage:")
            
            If iRet = vbYes Then
                If Text1(5).Text <> "" Then
                    If IsNumeric(Text1(5).Text) Then
                        sSQL = "Update artlief set spanne = 0 where artnr = " & Text1(5).Text
                        gdBase.Execute sSQL, dbFailOnError
                        
                        Label4(30).Visible = False
                    End If
                End If
            Else
            
            End If
        Case 9
            frmWKL165.Show 1
        Case 49
        
            gsARTNR = Text1(5).Text
            frmWKL204.Show 1
                
            If MehrEAN_vorhanden(CLng(gsARTNR)) Then
                Label4(49).ForeColor = glWarn
            Else
                Label4(49).ForeColor = glS1
            End If
                
            gsARTNR = ""
            
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label4_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim sSQL As String
    
    Select Case index
    
        Case 30
            Label4(30).Caption = "auto. Kalk. = aus?"
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
'Private Sub MSFlexGrid1_Click()
'On Error GoTo LOKAL_ERROR
'
'    Bildzeigen MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr), Image1, Picture3, 80
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "MSFlexGrid1_Click"
'    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub

Private Sub MSFlexGrid1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim lcol As Long
    Dim lrow As Long
    lcol = MSFlexGrid1.Col
    lrow = MSFlexGrid1.Row
    
    If MSFlexGrid1.Row > 1 Then
        Command2_Click 0
    Else
    
        If MSFlexGrid1.TextMatrix(0, MSFlexGrid1.Col) = "le. WE" Then
            'sortier anders
            If byteSortReihen = 1 Then
        
                zeige_Grid bAusblenden, bBest, bgef, binbest, " order by lwe asc "
                byteSortReihen = 2
                
            ElseIf byteSortReihen = 2 Then
                zeige_Grid bAusblenden, bBest, bgef, binbest, " order by lwe desc "
                byteSortReihen = 1
                
            End If
        ElseIf MSFlexGrid1.TextMatrix(0, MSFlexGrid1.Col) = "le. VK" Then
            'sortier anders
            If byteSortReihen = 1 Then
        
                zeige_Grid bAusblenden, bBest, bgef, binbest, " order by lVK asc "
                byteSortReihen = 2
                
            ElseIf byteSortReihen = 2 Then
                zeige_Grid bAusblenden, bBest, bgef, binbest, " order by lVK desc "
                byteSortReihen = 1
                
            End If
        ElseIf MSFlexGrid1.TextMatrix(0, MSFlexGrid1.Col) = "LiefbestNr" Then
            'sortier anders
            
            cSQL = "Update TOP" & srechnertab & " set LIBESNR = '' where LIBESNR is null"
            gdBase.Execute cSQL, dbFailOnError
                
            If byteSortReihen = 1 Then
                zeige_Grid bAusblenden, bBest, bgef, binbest, " order by val(LIBESNR) asc "
                byteSortReihen = 2
                
            ElseIf byteSortReihen = 2 Then
                zeige_Grid bAusblenden, bBest, bgef, binbest, " order by val(LIBESNR) desc "
                byteSortReihen = 1
                
            End If
        Else
            sortierenGrid MSFlexGrid1
        End If
    End If
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Bildzeigen(sArt As String, imgx As Image, PicX As PictureBox, iSize As Integer)
On Error GoTo LOKAL_ERROR

    Dim sPfad As String

    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "PICTURE\ARTIKEL"
    

    If FileExists(sPfad & "\" & sArt & ".jpg") Then
        imgx.Picture = LoadPicture(sPfad & "\" & sArt & ".jpg")
    Else
        If FileExists(sPfad & "\" & "keinBild.jpg") Then
            imgx.Picture = LoadPicture(sPfad & "\" & "keinBild.jpg")
        Else
            PicX.Visible = False
            Exit Sub
        End If
    End If
    
    zeigImage_In_Picture_Kasse imgx, PicX, iSize
    PicX.Tag = sArt
    PicX.Visible = True
    
Exit Sub
LOKAL_ERROR:

    If err.Number = 481 Then
'        MsgBox "Diese Bild kann nicht gespeichert werden, ungültiges Dateiformat", vbInformation, "Winkiss Hinweis:"
        Kill sPfad & "\" & sArt & ".jpg"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Bildzeigen"
        Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Sub
Private Sub MSFlexGrid1_SelChange()
    On Error GoTo LOKAL_ERROR

    Dim lColmerker  As Long
    Dim lRowmerker  As Long
    Dim cJahrNow    As String
    Dim cArtNr      As String
    Dim cartnrzuSpeichern As String
    Dim dat         As Date
    Dim lBest     As Long
    Dim cPreis       As String
    
    
    If MSFlexGrid1.Row > 1 Then
    
        MSFlexGrid1.Redraw = False
        
        If gbAender Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            lBest = Val(MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerBESTAND))
            If lBest > 1000 Then
                MsgBox MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerBESTAND) & " Dieser Wert wird nicht gespeichert.", vbInformation, "Winkiss Hinweis:"
                gbAender = False
                MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerBESTAND) = 0
                MSFlexGrid1.Redraw = True
                Exit Sub
            ElseIf lBest < -1000 Then
                MsgBox MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerBESTAND) & " Dieser Wert wird nicht gespeichert.", vbInformation, "Winkiss Hinweis:"
                gbAender = False
                MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerBESTAND) = 0
                MSFlexGrid1.Redraw = True
                Exit Sub
            End If
            
'            lbest = Val(MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerBESTAND))
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Bestandsveraenderung cartnrzuSpeichern, lBest, "Artikel bea Tabelle"
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAender = False
            
        End If
        
        If gbAenderRKZ Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerRKZ)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "RKZ"
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderRKZ = False
            
        End If
        
        If gbAenderMWST Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerMWST)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "MWST"
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderMWST = False
            
        End If
        
        If gbAenderSHOP Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerSHOP)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
'            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "RKZ"
            
            Internet_Shop cartnrzuSpeichern, cPreis
            
            If LeseInterArt(cartnrzuSpeichern) Then

                MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerSHOP) = "J"
                MSFlexGrid1.Row = Val(lbl6(0).Caption)
                MSFlexGrid1.Col = SpaltennummerSHOP
                MSFlexGrid1.CellForeColor = vbRed
            Else
                MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerSHOP) = "N"
                MSFlexGrid1.Row = Val(lbl6(0).Caption)
                MSFlexGrid1.Col = SpaltennummerSHOP
                MSFlexGrid1.CellForeColor = vbBlack
            End If
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderSHOP = False
            
        End If
        
        If gbAendergefuehrt Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerGEFUEHRT)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "GEFUEHRT"
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAendergefuehrt = False
            
        End If
        
        If gbAenderpreisSchu Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerPREISSCHU)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "PREISSCHU"
           
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderpreisSchu = False
            
        End If
        
        If gbAenderRABATT_OK Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerRABATT_OK)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "RABATT_OK"
           
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderRABATT_OK = False
            
        End If
        
        If gbAenderBONUS_OK Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerBONUS_OK)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "BONUS_OK"
           
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderBONUS_OK = False
            
        End If
        
        If gbAenderKVK Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerKVKPR1)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "KVKPR1"
                
                schreibeWKEtidru cartnrzuSpeichern, ermBESTAND(cartnrzuSpeichern), CLng(gcFilNr)
                
            End If
            
            If SpaltennummerHS < 255 Then
            
                Dim dHS As Double
                
                dHS = Fix(NettospanneInProzent_neu(cartnrzuSpeichern))
                
                MSFlexGrid1.Col = SpaltennummerHS
                MSFlexGrid1.Row = Val(lbl6(0).Caption)
                
                If dHS >= 100 Then
                    MSFlexGrid1.CellBackColor = &H8000&
                ElseIf dHS > 79.99 Then
                    MSFlexGrid1.CellBackColor = &HC000&
                ElseIf dHS > 59.99 Then
                    MSFlexGrid1.CellBackColor = &HFF00&
                ElseIf dHS > 39.99 Then
                    MSFlexGrid1.CellBackColor = &HFFFF&
                ElseIf dHS > 19.99 Then
                    MSFlexGrid1.CellBackColor = &H80C0FF
                ElseIf dHS > 0 Then
                    MSFlexGrid1.CellBackColor = &H80FF&
                ElseIf dHS <= 0 Then
                    MSFlexGrid1.CellBackColor = &HFF&
                End If
                
                
                MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerHS) = dHS
            End If
            
            
                    
        
 
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderKVK = False
            
        End If
        
        If gbAenderLEKPR Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerLEKPR)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "LEKPR"
            End If
 
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderLEKPR = False
        End If
        
        If gbAenderLVKPR Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerLVKPR)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "VKPR"
            End If
 
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderLVKPR = False
        End If
        
        If gbAenderBEZEICH Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerBEZEICH)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
        
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "BEZEICH"

            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderBEZEICH = False
            
        End If
        
        If gbAenderEAN Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerEAN)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "EAN"
            End If
 
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderEAN = False
            
        End If
        
        If gbAenderEAN3 Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerEAN3)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "EAN3"
            End If
 
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderEAN3 = False
            
        End If
        
        
        If gbAenderEAN2 Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerEAN2)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "EAN2"
            End If
 
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderEAN2 = False
        End If
        
        
        If gbAenderLPZ Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerLPZ)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "LPZ"
            End If
 
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderLPZ = False
        End If
        
        If gbAenderMB Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerMB)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "MINBEST"
            End If
 
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderMB = False
            
        End If
        
        If gbAenderGROESSE Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerGROESSE)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "GROESSE"
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderGROESSE = False
            
        End If
        
        
        If gbAenderLAGERP Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerLagerP)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "LAGERP"
            End If
 
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderLAGERP = False
        End If
        
        If gbAenderMODELL Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerModell)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "MODELL"
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderMODELL = False
        End If
        
        If gbAenderMATERIAL Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerMaterial)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "MATERIAL"
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderMATERIAL = False
        End If
        
        If gbAenderGRUPPE Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerGRUPPE)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "GRUPPE"
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderGRUPPE = False
        End If
        
        If gbAenderFARBBEZ Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerFarbbez)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "FARBBEZ"
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderFARBBEZ = False
        End If
        
        If gbAenderPGN Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerPGN)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "PGN"
            End If
 
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderPGN = False
            
        End If
        
        If gbAenderAWM Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerAWM)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "AWM"
            End If
 
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderAWM = False
            
        End If
        
        
        If gbAenderLINR Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerLINR)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "LINR"
            End If
 
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderLINR = False
            
        End If
        
        If gbAenderLIBESNR Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerLIBESNR)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "LIBESNR"
            
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderLIBESNR = False
            
        End If
    
        MSFlexGrid1.Redraw = True
        
        If gbAenderAGN Then
            lColmerker = MSFlexGrid1.Col
            lRowmerker = MSFlexGrid1.Row
            
            cPreis = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerAGN)
            cartnrzuSpeichern = MSFlexGrid1.TextMatrix(Val(lbl6(0).Caption), SpaltennummerArtnr)
            
            If IsNumeric(cPreis) Then
                Artikelveraenderung cartnrzuSpeichern, cPreis, "Artikel bea Tabelle", "AGN"
            End If
 
            MSFlexGrid1.Col = lColmerker
            MSFlexGrid1.Row = lRowmerker
            
            gbAenderAGN = False
            
        End If
    
        MSFlexGrid1.Redraw = True
        
        Bildzeigen MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr), Image1, Picture3, 80
    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFLexGrid1_SelChange"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
    
End Sub
Private Sub Internet_Shop(cArtNr As String, sMerk As String)
On Error GoTo LOKAL_ERROR

    If Len(sMerk) > 1 Then
        MsgBox sMerk & " Dieser Wert wird nicht gespeichert.(J oder N sind zulässig)", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    ElseIf Trim(sMerk) = "" Then
        Exit Sub
    End If

    Select Case UCase(sMerk)
        Case "J"
            Insert_Shop cArtNr, ermdisKat(cArtNr)
        Case "N"
            delInterart cArtNr
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Internet_Shop"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermdisKat(cART As String) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lAnz As Long
    
    ermdisKat = ""
    
    
    
    If Not NewTableSuchenDBKombi("TOP" & srechnertab, gdBase) Then
    
        sSQL = "Select distinct(kategorie) as kati from INTERART where kategorie <> '' "
    Else
    
        sSQL = "Select distinct(kategorie) as kati from INTERART where artnr in (Select artnr from TOP" & srechnertab & ")"
        sSQL = sSQL & " and kategorie <> '' "
    
    End If

    

    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnz = rsrs.RecordCount

        If lAnz = 1 Then
        
        
            rsrs.MoveFirst
            
'            Do While Not rsrs.EOF
        
        
        
            If Not IsNull(rsrs!kati) Then
                ermdisKat = rsrs!kati
            End If
            
'            rsrs.MoveNext
'
'            Loop
        Else
        
            
                        
        End If

    End If
    rsrs.Close: Set rsrs = Nothing
    
    If ermdisKat = "" Then
'        MsgBox "Es kann keine eindeutige Kategorie zugewiesen werden.", vbInformation, "Winkiss Hinweis:"
    End If


Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermdisKat"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Insert_Shop(cArtNr As String, cKat As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    Dim cKURZBESCHREIB As String
    Dim cBESCHREIB As String
    Dim cArtBez As String
    Dim dSHOPKVK As Double
    Dim sPfad As String

    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "PICTURE\ARTIKEL"
    
    cArtBez = ermBezeichausWGN(cArtNr)
    cKURZBESCHREIB = cArtBez
    cBESCHREIB = cArtBez
    dSHOPKVK = 0
    dSHOPKVK = CDbl(ermKVKPR1(cArtNr))

    delInterart cArtNr
    
    sSQL = "INSERT into INTERART (ARTNR,ARTBEZ,INTERBEZ,BESCHREIB,LASTDATE,SHOPKVK,KATEGORIE,BILDgr) values "
    sSQL = sSQL & " ('" & cArtNr & "','" & cArtBez & "','" & cKURZBESCHREIB & "','" & cBESCHREIB & "', '" & DateValue(Now) & "'"
    
    sSQL = sSQL & ", '" & dSHOPKVK & "'  "
    sSQL = sSQL & ", '" & cKat & "'  "
    
    If FileExists(sPfad & "\" & cArtNr & ".jpg") Then
        sSQL = sSQL & ", '" & cArtNr & "' & '.jpg' "
    Else
        sSQL = sSQL & ", 'keinBild.jpg'  "
    End If
    
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Insert_Shop"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long
    
    lrow = MSFlexGrid1.Row
    lcol = MSFlexGrid1.Col
    
    If KeyCode <> vbKeyDown And KeyCode <> vbKeyUp And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft Then  'And KeyCode <> vbKeyReturn
    
        Select Case lcol
            Case Is = SpaltennummerBESTAND, SpaltennummerKVKPR1, SpaltennummerLagerP, SpaltennummerLPZ, SpaltennummerPGN, SpaltennummerAGN, _
            SpaltennummerEAN, SpaltennummerEAN2, SpaltennummerEAN3, SpaltennummerBEZEICH, SpaltennummerMWST, _
            SpaltennummerGEFUEHRT, SpaltennummerRABATT_OK, SpaltennummerPREISSCHU, SpaltennummerRKZ, _
            SpaltennummerLEKPR, SpaltennummerLVKPR, SpaltennummerLINR, SpaltennummerLIBESNR, SpaltennummerAWM, SpaltennummerGROESSE, SpaltennummerMB, _
            SpaltennummerSHOP, SpaltennummerBONUS_OK, SpaltennummerModell, SpaltennummerMaterial, SpaltennummerFarbbez, SpaltennummerGRUPPE, _
            SpaltennummerHS
        
                If iKeypress = 0 And KeyCode <> vbKeyBack And KeyCode <> vbKeyF2 And KeyCode <> vbKeyReturn Then
                
                    If Check4.value = vbChecked Then
                        MSFlexGrid1.Row = lrow
                        MSFlexGrid1.Col = lcol
                        MSFlexGrid1.Text = ""
                    End If

                ElseIf iKeypress > 0 And KeyCode = 46 Or KeyCode = 8 Then

                    MSFlexGrid1.Row = lrow
                    MSFlexGrid1.Col = lcol
                    MSFlexGrid1.Text = ""

                End If
                iKeypress = iKeypress + 1
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    Dim lcol As Long
    Dim lrow As Long
    Dim cArtNr As String
    lcol = MSFlexGrid1.Col
    lrow = MSFlexGrid1.Row
    
    lbl6(0).Caption = lrow

    cZeichen = Chr$(KeyAscii)
    
    Select Case lcol
        Case Is = SpaltennummerLIBESNR
            gbAenderLIBESNR = True
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerGROESSE
            gbAenderGROESSE = True
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerBEZEICH
            gbAenderBEZEICH = True
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
            
        Case Is = SpaltennummerModell
            gbAenderMODELL = True
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerMaterial
            gbAenderMATERIAL = True
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerFarbbez
            gbAenderFARBBEZ = True
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerKVKPR1
            gbAenderKVK = True
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    
                    
                    
                    MSFlexGrid1.Text = cValid
                    
'
                    
                    
                    
                    
                End If
            End If
        Case Is = SpaltennummerLEKPR
            gbAenderLEKPR = True
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerLVKPR
            gbAenderLVKPR = True
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerRKZ
            gbAenderRKZ = True
            cValid = "jnJN"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerMWST
            gbAenderMWST = True
            cValid = "veoVEO"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerSHOP
            gbAenderSHOP = True
            cValid = "jnJN" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerRABATT_OK
            gbAenderRABATT_OK = True
            cValid = "jnJN"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerBONUS_OK
            gbAenderBONUS_OK = True
            cValid = "jnJN"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerPREISSCHU
            gbAenderpreisSchu = True
            cValid = "jnJN"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerGEFUEHRT
            gbAendergefuehrt = True
            cValid = "jnJN"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
            
        Case Is = SpaltennummerBESTAND
            gbAender = True
            cValid = "1234567890-" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerGRUPPE
            gbAenderGRUPPE = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerMB
            gbAenderMB = True
            cValid = "1234567890-" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerAWM
            gbAenderAWM = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerEAN3
            gbAenderEAN3 = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    
                    If checkthisean(cValid, MSFlexGrid1.TextMatrix(lrow, CLng(SpaltennummerArtnr))) = True Then
                        MSFlexGrid1.Text = cValid
                    Else
                        
                        Screen.MousePointer = 0
                        MsgBox "Diese EAN - Nummer ist schon vergeben. Diese kann kein zweites Mal gespeichert werden.", vbInformation, "Winkiss Hinweis:"
                        cValid = ""
                        MSFlexGrid1.Text = ""
                        Exit Sub
                    End If
                End If
            End If
        Case Is = SpaltennummerEAN2
            gbAenderEAN2 = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    
                    If checkthisean(cValid, MSFlexGrid1.TextMatrix(lrow, CLng(SpaltennummerArtnr))) = True Then
                        MSFlexGrid1.Text = cValid
                    Else
                        
                        Screen.MousePointer = 0
                        MsgBox "Diese EAN - Nummer ist schon vergeben. Diese kann kein zweites Mal gespeichert werden.", vbInformation, "Winkiss Hinweis:"
                        cValid = ""
                        MSFlexGrid1.Text = ""
                        Exit Sub
                    End If
                End If
            End If
        Case Is = SpaltennummerEAN
            gbAenderEAN = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    
                    
                    If checkthisean(cValid, MSFlexGrid1.TextMatrix(lrow, CLng(SpaltennummerArtnr))) = True Then
                        MSFlexGrid1.Text = cValid
                    Else
                        
                        Screen.MousePointer = 0
                        MsgBox "Diese EAN - Nummer ist schon vergeben. Diese kann kein zweites Mal gespeichert werden.", vbInformation, "Winkiss Hinweis:"
                        cValid = ""
                        MSFlexGrid1.Text = ""
                        Exit Sub
                    End If
                    
                    
                    
                End If
            End If
        Case Is = SpaltennummerLPZ
            gbAenderLPZ = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerLagerP
            gbAenderLAGERP = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerPGN
            gbAenderPGN = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerLINR
            gbAenderLINR = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
        Case Is = SpaltennummerAGN
            gbAenderAGN = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.Col = lcol
                cValid = MSFlexGrid1.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid1.Text = cValid
                End If
            End If
    
     End Select
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case KeyCode
        Case Is = vbKeyReturn
            Command2_Click 0
        Case Is = vbKeyEscape
            Command2_Click 4
    End Select
    
    If KeyCode = vbKeyF4 Then
    
        gsARTNR = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr)
        If gsARTNR <> "" Then
            frmWKL63.Show 1
            Me.Refresh
        End If
        gsARTNR = ""
    End If
    
    
    
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_LeaveCell()
On Error GoTo LOKAL_ERROR
    
    iKeypress = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_LeaveCell"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Picture3_Click()
On Error GoTo LOKAL_ERROR
    
    gsARTNR = Picture3.Tag
    frmWKL163.Show 1

    gsARTNR = ""
   
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Picture3_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Sub
Private Sub Text1_Change(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If index = 5 Then
        If Len(Text1(5).Text) = 6 Then
            If IsNumeric(Text1(5).Text) Then
                Dim ierg As Integer
                ierg = Artfrei(Text1(5).Text)
                If ierg = 1 Or ierg = 2 Then
                
                    If bBeimSpeichern = False Then
                        HoleDatenWKL10 Text1(5).Text
                    End If
                    giDlgZustand = giUPD
                    gbNew = False
                    

                Else
                    giDlgZustand = giNEU
                    gbNew = True
                End If
            Else
                Text1(5).Text = ""
            End If
        Else
            LeereArtikelDatenWKL10 False
'            gbNew = True
        End If
    End If


    If index = 7 Then
        LiefKuerzelAufloesung Label1(10), Text1(7)
    End If
    
    If index = 0 Then
        If Len(Text1(0).Text) = 0 Then
            Label1(8).Caption = "keine Auswahl"
        End If
    End If
    
    If index = 9 Then
        If Len(Text1(9).Text) >= 3 Then
            Label10.Caption = Ermittleagntext(Text1(9).Text)
        End If
    End If
    
    
    
    
    
    
    Dim sMWST As String
    Dim sKVK As String
    Dim sLEK As String
    Dim sSEK As String
    Dim sLiefBez As String
    
    sMWST = Text1(15).Text
    sKVK = Text1(30).Text
    sLEK = Text1(11).Text
    sSEK = Text1(29).Text
    sLiefBez = lblLiefbez.Caption
    
    If sKVK = "" Then sKVK = "0"
    If sLEK = "" Then sLEK = "0"
    If sSEK = "" Then sSEK = "0"
    If sMWST = "" Then sMWST = "V"
    
    
    If index = 30 Then 'KVKPR
        'Nettospanne anzeigen
        dyn_Handelsspanne_anzeigen sMWST, sKVK, sLiefBez, sLEK, sSEK
    End If
    
    If index = 11 Then 'ListenEk
        'Nettospanne anzeigen
        dyn_Handelsspanne_anzeigen sMWST, sKVK, sLiefBez, sLEK, sSEK
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Artfrei(sZiff As String) As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsArt As Recordset
    
    Artfrei = 0
    sSQL = "Select * from artikel where artnr = " & Val(sZiff)
    
    Set rsArt = gdBase.OpenRecordset(sSQL)
    If rsArt.EOF Then
        Artfrei = 0
    Else
        Artfrei = 1
    End If
    
    rsArt.Close: Set rsArt = Nothing
    
    If Artfrei = 0 Then
    
        sSQL = "Select * from kassjour where artnr = " & Val(sZiff)
        
    
        Set rsArt = gdBase.OpenRecordset(sSQL)
        If rsArt.EOF Then
            Artfrei = 0
        Else
            Artfrei = 2
        End If
    
        rsArt.Close: Set rsArt = Nothing
        
   
    
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Artfrei"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function


Private Sub Text1_DblClick(index As Integer)
On Error GoTo LOKAL_ERROR
    
    If index = 17 Then
        gsARTNR = Text1(5).Text
        frmWKL67.Show 1
        gsARTNR = ""
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    iFocus = index
    
    If index >= 5 Then
        Label3(1).Caption = Label4(index - 5).Caption
        Label3(2).Caption = Trim$(Str$(index))
    Else
        Label3(1).Caption = Label1(index).Caption
        Label3(2).Caption = Trim$(Str$(index))
    End If
    
    Text1(index).SelStart = 0
    Text1(index).SelLength = Len(Text1(index).Text)
    Text1(index).BackColor = glSelBack1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    cZeichen = Chr$(KeyAscii)
    
    If gbTagAkt Then
        cZeichen = UCase(cZeichen)
    End If

    Select Case index
        Case 2, 6, 16, 23, 19, 7, 36, 37, 43, 44 'suche Artbez linr
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%!?" & Chr$(22) & Chr$(3) & Chr$(24)
        Case 1, 5, 28, 8, 13, 25, 18, 20, 21, 32 'suche Artnr/Ean
            cValid = "1234567890" & Chr$(8) & Chr$(22) & Chr$(3) & Chr$(24) 'Strg + V = 22 und Strg + C = 3 und Strg + X = 24
        Case 0, 3, 9, 41, 42, 35, 39 'suche Agn
            cValid = "1234567890" & Chr$(8) & Chr$(22) & Chr$(3) & Chr$(24)
        Case 26, 27, 10, 24, 31, 33, 38, 40 'J N
            cValid = "jnJN" & Chr$(8) & Chr$(22) & Chr$(3) & Chr$(24)
            cZeichen = UCase$(cZeichen)

        Case 11, 29, 12, 30, 14, 17, 45, 46, 48 'Preise
            cValid = "-1234567890," & Chr$(8) & Chr$(22) & Chr$(3) & Chr$(24)

        Case 15 'mwst
            cValid = "VEO" & Chr$(8) & Chr$(22) & Chr$(3) & Chr$(24)
            cZeichen = UCase$(cZeichen)
        Case 4 'suche libesnr
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%!?*" & Chr$(22) & Chr$(3) & Chr$(24)
        Case 22 'Inhalt auch Komma
            cValid = "1234567890," & Chr$(8) & Chr$(22) & Chr$(3) & Chr$(24)
    End Select

    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(cZeichen)
    End If


    'Hier die doppelten Kommas prüfen, nur bei den Preisen
    Select Case index
        Case 11, 29, 12, 30, 14, 45, 46 'Preise
            If cZeichen = "," Then
                If InStr(Text1(index).Text, ",") > 0 Then
                    KeyAscii = 0
                End If
            End If
        Case Else
    End Select
    'Ende*****Hier die doppelten Kommas prüfen, nur bei den Preisen

    Dim bSpringen As Boolean

    bSpringen = False

    If Len(Text1(index).Text) = Text1(index).MaxLength - 1 Then
        If KeyAscii <> 8 And KeyAscii <> 0 And KeyAscii <> 13 And KeyAscii <> 27 Then
            bSpringen = True
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim sAuswahlfeld As String
    Dim ctmp As String
    
    If KeyCode = vbKeyReturn Then
    
        If index = 48 Then
            'lek abschlagen
            If Text1(48).Text <> "" Then
                 If IsNumeric(Text1(48).Text) Then
                 
                    Text1(48).Text = SwapStr(Text1(48).Text, "-", "")
                 
                    Text1(11).Text = Format(CDbl(Text1(11).Text) - (CDbl(Text1(11).Text) * CDbl(Text1(48).Text) / 100), "####0.00")
                    Text1(11).SetFocus
                    
                 End If
            End If
            
            Exit Sub
            
        End If
    
        If index = 18 Or index = 20 Or index = 21 Then
        
        Else
            If index = 36 Or index = 7 Or index = 35 Or index = 41 Or index = 42 Or index < 5 Then
                Command1_Click 0
            Else
                Command5_Click 0
            End If
        End If
        
    End If
    
    If KeyCode = vbKeyEscape Then
        Command1_Click 2
    End If
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case index
            Case Is = 1     'ArtNr
                If Text1(7).Text = "" Then
                    MsgBox "Bitte einen Lieferanten angeben!", vbInformation, gsPname & " Hinweis:"
                    Exit Sub
                Else
                    If IsNumeric(Text1(7).Text) Then
                        gF2Prompt.cFeld = "ARTNR"
                        gF2Prompt.cWert = Text1(7).Text
                        frmWK00a.Show 1
                        If gF2Prompt.cWahl <> "" Then
                            Text1(index).Text = gF2Prompt.cWahl
                        End If
                    Else
                        MsgBox "Bitte einen Lieferanten angeben!", vbInformation, gsPname & " Hinweis:"
                        Exit Sub
                    End If
                End If
                
            Case Is = 0     'PGN
                gF2Prompt.cFeld = "PGN"
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(index).Text = gF2Prompt.cWahl
                    Label1(8).Caption = gF2Prompt.cWert
                End If
            Case Is = 2     'MARKE
                gF2Prompt.cFeld = "MARKE"
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(index).Text = gF2Prompt.cWahl
                End If

            Case Is = 11 'StaffelLEKPR
            
                gsLinr = cbo1.Text
                gsARTNR = Text1(5).Text
                frmWKL161.Show 1
                
                If LeseStaffelpreis(CLng(gsARTNR), CLng(gsLinr)) Then
                    Command1(27).BackColor = vbRed
                Else
                    Command1(27).BackColor = Command1(9).BackColor
                End If
                gsLinr = ""
                gsARTNR = ""
            Case Is = 30 'spezieller KVK
                gsARTNR = Text1(5).Text
                frmWKL144.Show 1
                
                If LeseSpezpreis(CLng(gsARTNR), 0) > 0 Then
                    Command1(26).BackColor = vbRed
                Else
                    Command1(26).BackColor = Command1(9).BackColor
                End If
                
                gsARTNR = ""
            Case Is = 14 'kalkulierte Spanne
                
                frmWKL165.Show 1
                
            Case Is = 37 'Protokoll der Artikelmerkmale
                Dim cPfad As String
                cPfad = gcDBPfad
                If Right(cPfad, 1) <> "\" Then
                    cPfad = cPfad & "\"
                End If
                
                zeigeHilfe "LPROTOK", "Artikelmerkmal.txt", cPfad
                
            Case Is = 35    'Linie
                ctmp = Text1(2).Text 'Marke
                ctmp = Trim$(ctmp)
                If ctmp = "" Then
                    ctmp = Text1(7).Text 'lieferant
                    ctmp = Trim$(ctmp)
                    If ctmp = "" Then
                        MsgBox "Bitte einen Lieferanten oder eine Marke angeben!", vbInformation, gsPname & " Hinweis:"
                        Text1(2).SetFocus
                        Exit Sub
                    Else
                        If IsNumeric(ctmp) = False Then
                            MsgBox "Bitte einen Lieferanten oder eine Marke angeben!", vbInformation, gsPname & " Hinweis:"
                            Text1(2).SetFocus
                            Exit Sub
                        Else
                            sAuswahlfeld = "LINR"
                        End If
                    End If
                Else
                    sAuswahlfeld = "MARKE"
                End If
                
                gF2Prompt.bMultiple = True
                gF2Prompt.cFeld = "LPZ"
                gF2Prompt.cWert = ctmp
                gF2Prompt.cEsFeld = sAuswahlfeld
            
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(index).Text = gF2Prompt.cWahl
                End If
                
                List3.Visible = False
                List3.Clear
                For lcount = 0 To 100
                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                        List3.Visible = True
                        Text1(index).Text = ""
                        
                        If gF2Prompt.cArray(lcount) <> "" Then
'                            List3.AddItem gF2Prompt.cArray(lcount) & Space(50) & Right(gF2Prompt.cArray(lcount), 6)
                            List3.AddItem gF2Prompt.cArray(lcount)
                        End If
                    Else
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List3.AddItem gF2Prompt.cArray(lcount)
                            Text1(index).Text = Left$(gF2Prompt.cArray(lcount), 3)
                        End If
                    End If
                Next lcount
                
            Case Is = 3
                gF2Prompt.bMultiple = True
                gF2Prompt.cFeld = "AGN"
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    If gF2Prompt.cWahl <> "" Then
                        Text1(index).Text = gF2Prompt.cWahl
                    End If
                End If
                
                List4.Visible = False
                List4.Clear
                For lcount = 0 To 100
                    If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                        List4.Visible = True
                        Text1(index).Text = ""
                        
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List4.AddItem gF2Prompt.cArray(lcount)
                        End If
                    Else
                        If gF2Prompt.cArray(lcount) <> "" Then
                            List4.AddItem gF2Prompt.cArray(lcount)
                            Text1(index).Text = Mid(gF2Prompt.cArray(lcount), 1, InStr(1, gF2Prompt.cArray(lcount), " ")) & " "
                        End If
                        
                    End If
                Next lcount
            Case Is = 7
                gF2Prompt.cFeld = "LINR"
                
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(index).Text = gF2Prompt.cWahl
                    Label1(10).Caption = gF2Prompt.cWert
                End If
            Case Is = 8
                If cbo1.Text = "" Then
                    MsgBox "Bitte einen Lieferanten angeben!", vbCritical, "STOP!"
                    Exit Sub
                End If
                gF2Prompt.cFeld = "LPZ"
                gF2Prompt.cWert = Trim$(cbo1.Text)
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(index).Text = gF2Prompt.cWahl
                End If
            Case Is = 9
                gF2Prompt.cFeld = "AGN"
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(index).Text = gF2Prompt.cWahl
                End If
                
            Case Is = 39
                gF2Prompt.cFeld = "PGN"
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(index).Text = gF2Prompt.cWahl
                End If
        End Select
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
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
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_Change(index As Integer)
On Error GoTo LOKAL_ERROR

    If index = 10 Then
        If Len(Text2(10).Text) >= 3 Then
            Label7(25).Caption = ErmittleGruppenbez(Text2(10).Text)
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_Change"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text2_KeyPress(index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    Select Case index
        
        Case 2, 3, 4, 5, 6, 0, 1, 10
            cValid = "1234567890" & Chr$(8)
        Case 7, 8, 9 'Modell, Material,Farbbez
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%!?"
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
    Fehler.gsFunktion = "Text2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text2_LostFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text2(index).BackColor = vbWhite
    
    Select Case index
        Case Is = 0
            PruefEanEingabe
        Case Is = 1

    End Select
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        Select Case index
            Case Is = 0
                cboGp.SetFocus
                
            Case Is = 1
                Command1_Click 5
            
        End Select
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PruefEanEingabe()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim sEAN    As String
    Dim rsArt   As Recordset
    Dim rsrs    As Recordset
    Dim i       As Integer
    
    If Len(Text2(0).Text) = 0 Then
        Label7(7).Caption = "Sie müssen erst die EAN des Einzelprodukts angeben!"
        Label7(7).Refresh
    Else
        sEAN = Trim(Text2(0).Text)
    
        cSQL = "select * from artikel where ean = '" & sEAN & "'"
        cSQL = cSQL & " or ean2 = '" & sEAN & "'"
        cSQL = cSQL & " or ean3 = '" & sEAN & "'"
        
        Set rsArt = gdBase.OpenRecordset(cSQL)
        If Not rsArt.RecordCount = 0 Then
            For i = 8 To 11
                Label7(i).Visible = True
            Next i
            
            Label7(10).Caption = IIf(IsNull(rsArt!artnr), "", rsArt!artnr)
            Label7(11).Caption = IIf(IsNull(rsArt!BEZEICH), "", rsArt!BEZEICH)
            
            
            cSQL = "Select * from Zuordean where ean = '" & sEAN & "' "
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                cboGp.Clear
                rsrs.MoveFirst
                i = 0
                Do While Not rsrs.EOF
                    i = i + 1
                    If Not IsNull(rsrs!GPEAN) Then
                        cboGp.AddItem rsrs!GPEAN
                    Else
                        cboGp.AddItem ""
                    End If
                    rsrs.MoveNext
                Loop
                Label7(7).Caption = "Es sind schon " & i & " Umverpackungs - EAN's angelegt."
                Label7(7).Refresh
            End If
            rsrs.Close: Set rsrs = Nothing
        Else
            For i = 8 To 11
                Label7(i).Visible = False
            Next i
            Label7(7).Caption = "Es wurde kein Artikel mit dieser EAN gefunden!"
            Label7(7).Refresh
            startframe6
        End If
        rsArt.Close: Set rsArt = Nothing
    
    End If
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PruefEanEingabe"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR

    Text2(index).SelStart = 0
    Text2(index).SelLength = Len(Text2(index).Text)
    Text2(index).BackColor = glSelBack1

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil  Artikel bearbeiten  ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
