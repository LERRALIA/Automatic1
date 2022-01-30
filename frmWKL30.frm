VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmWKL30 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Etiketten "
   ClientHeight    =   8625
   ClientLeft      =   2100
   ClientTop       =   2355
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
   Icon            =   "frmWKL30.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame10 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame8"
      Height          =   6135
      Left            =   240
      TabIndex        =   279
      Top             =   5040
      Visible         =   0   'False
      Width           =   4935
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   11
         Left            =   2880
         TabIndex        =   329
         Top             =   4080
         Visible         =   0   'False
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   10
         Left            =   7680
         TabIndex        =   328
         Top             =   2160
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   9
         Left            =   6720
         TabIndex        =   326
         Top             =   2160
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   8
         Left            =   5760
         TabIndex        =   324
         Top             =   2160
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   7
         Left            =   4800
         TabIndex        =   320
         Top             =   2160
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   6
         Left            =   4800
         TabIndex        =   313
         Top             =   2640
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   5
         Left            =   2880
         TabIndex        =   298
         Top             =   3120
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   4
         Left            =   3840
         TabIndex        =   297
         Top             =   3120
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   3
         Left            =   3840
         TabIndex        =   295
         Top             =   2640
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   294
         Top             =   2640
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   7
         Left            =   9720
         TabIndex        =   282
         Top             =   7320
         Width           =   2055
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   281
         Top             =   2160
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   280
         Top             =   2160
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   12
         Left            =   2880
         TabIndex        =   331
         Top             =   3600
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   13
         Left            =   2880
         TabIndex        =   341
         Top             =   4560
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   14
         Left            =   3840
         TabIndex        =   343
         Top             =   4560
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   15
         Left            =   8640
         TabIndex        =   344
         Top             =   2160
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "Drucken +Pfand"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   16
         Left            =   2880
         TabIndex        =   348
         Top             =   5040
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "Drucken +Pfand"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   17
         Left            =   5760
         TabIndex        =   349
         Top             =   2640
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   18
         Left            =   4800
         TabIndex        =   368
         Top             =   3120
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   19
         Left            =   2880
         TabIndex        =   372
         Top             =   5520
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   20
         Left            =   9720
         TabIndex        =   378
         Top             =   2160
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   21
         Left            =   2880
         TabIndex        =   382
         Top             =   6000
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "Drucken +Pfand"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "bestellen"
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
         Index           =   8
         Left            =   1800
         MouseIcon       =   "frmWKL30.frx":0442
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   384
         ToolTipText     =   "per Email bestellen"
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Kombi-Etikett 81 x 38"
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
         Index           =   34
         Left            =   120
         TabIndex        =   383
         Top             =   6000
         Width           =   2055
      End
      Begin VB.Label Label13 
         Caption         =   "Anzahl"
         Height          =   255
         Left            =   120
         TabIndex        =   380
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Var KW"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   33
         Left            =   9720
         TabIndex        =   379
         ToolTipText     =   "Variante 5"
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Etikett 49 x 36"
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
         Index           =   30
         Left            =   120
         TabIndex        =   373
         Top             =   5520
         Width           =   2055
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
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
         Index           =   2
         Left            =   10200
         MouseIcon       =   "frmWKL30.frx":074C
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   363
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
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
         Index           =   1
         Left            =   10200
         MouseIcon       =   "frmWKL30.frx":0A56
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   362
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
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
         Index           =   3
         Left            =   10200
         MouseIcon       =   "frmWKL30.frx":0D60
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   361
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "undefiniert"
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
         Index           =   2
         Left            =   6120
         MouseIcon       =   "frmWKL30.frx":106A
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   360
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "undefiniert"
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
         Index           =   1
         Left            =   6120
         MouseIcon       =   "frmWKL30.frx":1374
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   359
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "undefiniert"
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
         Index           =   3
         Left            =   6120
         MouseIcon       =   "frmWKL30.frx":167E
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   358
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label16 
         Height          =   255
         Index           =   2
         Left            =   7200
         TabIndex        =   357
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label16 
         Height          =   255
         Index           =   1
         Left            =   7200
         TabIndex        =   356
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label16 
         Height          =   255
         Index           =   3
         Left            =   7200
         TabIndex        =   355
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label17 
         Caption         =   $"frmWKL30.frx":1988
         Height          =   1095
         Index           =   0
         Left            =   2280
         TabIndex        =   354
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label15 
         Caption         =   "3."
         Height          =   255
         Index           =   2
         Left            =   5760
         TabIndex        =   353
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label15 
         Caption         =   "2."
         Height          =   255
         Index           =   1
         Left            =   5760
         TabIndex        =   352
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label15 
         Caption         =   "1."
         Height          =   255
         Index           =   0
         Left            =   5760
         TabIndex        =   351
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "Lieferantenangaben auf dem Dronova-Etikett:"
         Height          =   255
         Left            =   2280
         TabIndex        =   350
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Kombi-Etikett 69 x 38"
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
         Index           =   28
         Left            =   120
         TabIndex        =   347
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Dronova"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   27
         Left            =   8640
         TabIndex        =   345
         ToolTipText     =   "Variante 6"
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Etikett 50 x 40"
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
         Index           =   26
         Left            =   120
         TabIndex        =   342
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Etikett 50 x 37"
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
         Index           =   23
         Left            =   120
         TabIndex        =   332
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Edeka - Etikett"
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
         Index           =   24
         Left            =   120
         TabIndex        =   330
         Top             =   4080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Var 6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   22
         Left            =   7680
         TabIndex        =   327
         ToolTipText     =   "Variante 6"
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Var 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   21
         Left            =   6720
         TabIndex        =   325
         ToolTipText     =   "Variante 5"
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Var 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   20
         Left            =   5760
         TabIndex        =   323
         ToolTipText     =   "Variante 4"
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Var 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   17
         Left            =   4800
         TabIndex        =   314
         ToolTipText     =   "Variante 3"
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Etikett 40 x 18"
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
         Index           =   12
         Left            =   120
         TabIndex        =   299
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Etikett 70 x 35"
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
         Left            =   120
         TabIndex        =   296
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Regaletiketten, endlos"
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
         Height          =   615
         Index           =   9
         Left            =   120
         TabIndex        =   286
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Etikett 40 x 25"
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
         TabIndex        =   285
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Var 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   7
         Left            =   3000
         TabIndex        =   284
         ToolTipText     =   "Variante 1"
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Var 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   6
         Left            =   3840
         TabIndex        =   283
         ToolTipText     =   "Variante 2"
         Top             =   1800
         Width           =   735
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame8"
      Height          =   2535
      Left            =   8640
      TabIndex        =   268
      Top             =   4080
      Visible         =   0   'False
      Width           =   3495
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   20
         Left            =   6960
         TabIndex        =   321
         Top             =   1320
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   19
         Left            =   3720
         TabIndex        =   317
         Top             =   4200
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   18
         Left            =   4800
         TabIndex        =   316
         Top             =   4200
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   17
         Left            =   5880
         TabIndex        =   315
         Top             =   4200
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   16
         Left            =   5880
         TabIndex        =   312
         Top             =   2760
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   15
         Left            =   3720
         TabIndex        =   310
         Top             =   3720
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   14
         Left            =   3720
         TabIndex        =   308
         Top             =   3240
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   12
         Left            =   4800
         TabIndex        =   307
         Top             =   2760
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   13
         Left            =   3720
         TabIndex        =   305
         Top             =   2760
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   11
         Left            =   5880
         TabIndex        =   304
         Top             =   1320
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   10
         Left            =   4800
         TabIndex        =   302
         Top             =   2280
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   9
         Left            =   3720
         TabIndex        =   301
         Top             =   2280
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   8
         Left            =   5880
         TabIndex        =   300
         Top             =   2280
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   7
         Left            =   5880
         TabIndex        =   293
         Top             =   1800
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   6
         Left            =   5880
         TabIndex        =   291
         Top             =   840
         Width           =   915
         _ExtentX        =   1614
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
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   5
         Left            =   3720
         TabIndex        =   289
         Top             =   1800
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   4
         Left            =   4800
         TabIndex        =   288
         Top             =   1800
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   3
         Left            =   4800
         TabIndex        =   277
         Top             =   1320
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   276
         Top             =   1320
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   273
         Top             =   840
         Width           =   915
         _ExtentX        =   1614
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
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   272
         Top             =   840
         Width           =   915
         _ExtentX        =   1614
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
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   6
         Left            =   9720
         TabIndex        =   269
         Top             =   7320
         Width           =   2055
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   21
         Left            =   3720
         TabIndex        =   339
         Top             =   4680
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   22
         Left            =   6960
         TabIndex        =   369
         Top             =   1800
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   23
         Left            =   8040
         TabIndex        =   371
         Top             =   1320
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   24
         Left            =   9120
         TabIndex        =   375
         Top             =   1320
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   25
         Left            =   3720
         TabIndex        =   376
         Top             =   5160
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   26
         Left            =   4800
         TabIndex        =   381
         Top             =   3720
         Width           =   915
         _ExtentX        =   1614
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
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   27
         Left            =   6960
         TabIndex        =   387
         Top             =   2280
         Width           =   915
         _ExtentX        =   1614
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
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "bestellen"
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
         Index           =   4
         Left            =   2640
         MouseIcon       =   "frmWKL30.frx":1A2B
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   385
         ToolTipText     =   "per Email bestellen"
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   " Etikett 35 x 15"
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
         Index           =   32
         Left            =   120
         TabIndex        =   377
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Var 6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   31
         Left            =   9120
         TabIndex        =   374
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Var 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   29
         Left            =   8040
         TabIndex        =   370
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   " Etikett 48 x 18"
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
         Index           =   25
         Left            =   120
         TabIndex        =   340
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Var 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   19
         Left            =   6960
         TabIndex        =   322
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   " Etikett 30 x 15"
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
         Index           =   18
         Left            =   120
         TabIndex        =   318
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   " Etikett 44 x 21"
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
         Index           =   16
         Left            =   120
         TabIndex        =   311
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   " Etikett 49 x 19"
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
         Index           =   15
         Left            =   120
         TabIndex        =   309
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   " Etikett 51 x 19"
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
         Index           =   14
         Left            =   120
         TabIndex        =   306
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   " Etikett 38 x 23"
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
         Index           =   13
         Left            =   120
         TabIndex        =   303
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Var 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   10
         Left            =   5880
         TabIndex        =   292
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   " Etikett 45 x 23"
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
         TabIndex        =   290
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   " Etikett 40 x 18"
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
         Index           =   4
         Left            =   120
         TabIndex        =   278
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Var 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   275
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Var 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   274
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   " Schmucketikett 69 x 14"
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
         Index           =   1
         Left            =   120
         TabIndex        =   271
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   " Strichcodeetiketten, endlos"
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
         Index           =   0
         Left            =   120
         TabIndex        =   270
         Top             =   120
         Width           =   7815
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   3
      Left            =   4440
      TabIndex        =   4
      Top             =   1120
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   873
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
      Caption         =   "Preisliste"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   4
      Left            =   4440
      TabIndex        =   3
      ToolTipText     =   "38 x 21,2 mm"
      Top             =   600
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   873
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
      Caption         =   "Preisetiketten "
      Enabled         =   0   'False
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   495
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   873
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
      Caption         =   "Regal DINA4"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   873
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
      Caption         =   "Strichcode DINA4"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame1 
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
      Height          =   480
      Left            =   9600
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   600
      Begin VB.Frame Frame0 
         BackColor       =   &H00C0C0C0&
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
         Left            =   2160
         TabIndex        =   184
         Top             =   4680
         Width           =   9615
         Begin sevCommand3.Command Command3 
            Height          =   595
            Index           =   3
            Left            =   7320
            TabIndex        =   242
            Top             =   3360
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
            Caption         =   "a>A"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command3 
            Height          =   595
            Index           =   2
            Left            =   4440
            TabIndex        =   241
            Top             =   3360
            Width           =   2875
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
            Caption         =   "LÖSCHEN"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command3 
            Height          =   595
            Index           =   1
            Left            =   1560
            TabIndex        =   240
            Top             =   3360
            Width           =   2875
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
            Caption         =   "RÜCKGÄNGIG"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command3 
            Height          =   595
            Index           =   0
            Left            =   120
            TabIndex        =   239
            Top             =   3360
            Width           =   1435
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
            Caption         =   "A>a"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   595
            Index           =   53
            Left            =   7680
            TabIndex        =   238
            Top             =   2760
            Width           =   715
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
            Height          =   595
            Index           =   52
            Left            =   6960
            TabIndex        =   237
            Top             =   2760
            Width           =   715
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
            Height          =   595
            Index           =   51
            Left            =   6240
            TabIndex        =   236
            Top             =   2760
            Width           =   715
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
         Begin sevCommand3.Command Command0 
            Height          =   595
            Index           =   50
            Left            =   5520
            TabIndex        =   235
            Top             =   2760
            Width           =   715
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
            Height          =   595
            Index           =   49
            Left            =   4800
            TabIndex        =   234
            Top             =   2760
            Width           =   715
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
            Height          =   595
            Index           =   48
            Left            =   4080
            TabIndex        =   233
            Top             =   2760
            Width           =   715
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
            Height          =   595
            Index           =   47
            Left            =   3360
            TabIndex        =   232
            Top             =   2760
            Width           =   715
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
            Height          =   595
            Index           =   46
            Left            =   2640
            TabIndex        =   231
            Top             =   2760
            Width           =   715
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
            Height          =   595
            Index           =   45
            Left            =   1920
            TabIndex        =   230
            Top             =   2760
            Width           =   715
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
            Height          =   595
            Index           =   44
            Left            =   1200
            TabIndex        =   229
            Top             =   2760
            Width           =   715
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
            Height          =   595
            Index           =   43
            Left            =   8760
            TabIndex        =   228
            Top             =   2160
            Width           =   715
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
            Height          =   595
            Index           =   42
            Left            =   8040
            TabIndex        =   227
            Top             =   2160
            Width           =   715
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
            Height          =   595
            Index           =   41
            Left            =   7320
            TabIndex        =   226
            Top             =   2160
            Width           =   715
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
            Height          =   595
            Index           =   40
            Left            =   6600
            TabIndex        =   225
            Top             =   2160
            Width           =   715
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
            Height          =   595
            Index           =   39
            Left            =   5880
            TabIndex        =   224
            Top             =   2160
            Width           =   715
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
            Height          =   595
            Index           =   38
            Left            =   5160
            TabIndex        =   223
            Top             =   2160
            Width           =   715
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
            Height          =   595
            Index           =   37
            Left            =   4440
            TabIndex        =   222
            Top             =   2160
            Width           =   715
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
            Height          =   595
            Index           =   36
            Left            =   3720
            TabIndex        =   221
            Top             =   2160
            Width           =   715
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
            Height          =   595
            Index           =   35
            Left            =   3000
            TabIndex        =   220
            Top             =   2160
            Width           =   715
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
            Height          =   595
            Index           =   34
            Left            =   2280
            TabIndex        =   219
            Top             =   2160
            Width           =   715
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
            Height          =   595
            Index           =   33
            Left            =   1560
            TabIndex        =   218
            Top             =   2160
            Width           =   715
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
            Height          =   595
            Index           =   32
            Left            =   840
            TabIndex        =   217
            Top             =   2160
            Width           =   715
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
            Height          =   595
            Index           =   31
            Left            =   7680
            TabIndex        =   216
            Top             =   1560
            Width           =   715
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
            Height          =   595
            Index           =   30
            Left            =   6960
            TabIndex        =   215
            Top             =   1560
            Width           =   715
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
            Height          =   595
            Index           =   29
            Left            =   6240
            TabIndex        =   214
            Top             =   1560
            Width           =   715
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
            Height          =   595
            Index           =   28
            Left            =   5520
            TabIndex        =   213
            Top             =   1560
            Width           =   715
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
            Height          =   595
            Index           =   27
            Left            =   4800
            TabIndex        =   212
            Top             =   1560
            Width           =   715
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
            Height          =   595
            Index           =   26
            Left            =   4080
            TabIndex        =   211
            Top             =   1560
            Width           =   715
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
            Height          =   595
            Index           =   25
            Left            =   3360
            TabIndex        =   210
            Top             =   1560
            Width           =   715
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
            Height          =   595
            Index           =   24
            Left            =   2640
            TabIndex        =   209
            Top             =   1560
            Width           =   715
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
            Height          =   595
            Index           =   23
            Left            =   1920
            TabIndex        =   208
            Top             =   1560
            Width           =   715
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
            Height          =   595
            Index           =   22
            Left            =   1200
            TabIndex        =   207
            Top             =   1560
            Width           =   715
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
            Height          =   595
            Index           =   21
            Left            =   480
            TabIndex        =   206
            Top             =   1560
            Width           =   715
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
            Height          =   595
            Index           =   20
            Left            =   7320
            TabIndex        =   205
            Top             =   960
            Width           =   715
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
            Height          =   595
            Index           =   19
            Left            =   6600
            TabIndex        =   204
            Top             =   960
            Width           =   715
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
            Height          =   595
            Index           =   18
            Left            =   5880
            TabIndex        =   203
            Top             =   960
            Width           =   715
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
            Height          =   595
            Index           =   17
            Left            =   5160
            TabIndex        =   202
            Top             =   960
            Width           =   715
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
            Height          =   595
            Index           =   16
            Left            =   4440
            TabIndex        =   201
            Top             =   960
            Width           =   715
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
            Height          =   595
            Index           =   15
            Left            =   3720
            TabIndex        =   200
            Top             =   960
            Width           =   715
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
            Height          =   595
            Index           =   14
            Left            =   3000
            TabIndex        =   199
            Top             =   960
            Width           =   715
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
            Height          =   595
            Index           =   13
            Left            =   2280
            TabIndex        =   198
            Top             =   960
            Width           =   715
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
            Height          =   595
            Index           =   12
            Left            =   1560
            TabIndex        =   197
            Top             =   960
            Width           =   715
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
            Height          =   595
            Index           =   11
            Left            =   840
            TabIndex        =   196
            Top             =   960
            Width           =   715
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
            Height          =   595
            Index           =   10
            Left            =   120
            TabIndex        =   195
            Top             =   960
            Width           =   715
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
            Height          =   595
            Index           =   9
            Left            =   6600
            TabIndex        =   194
            Top             =   360
            Width           =   715
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
            Height          =   595
            Index           =   8
            Left            =   5880
            TabIndex        =   193
            Top             =   360
            Width           =   715
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
            Height          =   595
            Index           =   7
            Left            =   5160
            TabIndex        =   192
            Top             =   360
            Width           =   715
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
            Height          =   595
            Index           =   6
            Left            =   4440
            TabIndex        =   191
            Top             =   360
            Width           =   715
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
            Height          =   595
            Index           =   5
            Left            =   3720
            TabIndex        =   190
            Top             =   360
            Width           =   715
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
            Height          =   595
            Index           =   0
            Left            =   120
            TabIndex        =   189
            Top             =   360
            Width           =   715
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
         Begin sevCommand3.Command Command0 
            Height          =   595
            Index           =   1
            Left            =   840
            TabIndex        =   188
            Top             =   360
            Width           =   715
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
            Height          =   595
            Index           =   2
            Left            =   1560
            TabIndex        =   187
            Top             =   360
            Width           =   715
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
            Height          =   595
            Index           =   3
            Left            =   2280
            TabIndex        =   186
            Top             =   360
            Width           =   715
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
            Height          =   595
            Index           =   4
            Left            =   3000
            TabIndex        =   185
            Top             =   360
            Width           =   715
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
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "-1"
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
            Left            =   7560
            TabIndex        =   243
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin sevCommand3.Command Command2 
         Height          =   735
         Index           =   10
         Left            =   120
         TabIndex        =   46
         Top             =   7920
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
         Caption         =   "Standard"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Caption         =   "EURO-Preis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   10560
         TabIndex        =   161
         Top             =   4560
         Visible         =   0   'False
         Width           =   1335
         Begin VB.TextBox Text1 
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
            Index           =   39
            Left            =   2400
            TabIndex        =   170
            Text            =   "Text1"
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   38
            Left            =   2400
            TabIndex        =   169
            Text            =   "Text1"
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   37
            Left            =   2400
            TabIndex        =   168
            Text            =   "Text1"
            Top             =   1200
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "270°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   27
            Left            =   6240
            TabIndex        =   167
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "180°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   26
            Left            =   4920
            TabIndex        =   166
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "90°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   25
            Left            =   3720
            TabIndex        =   165
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "keine"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   24
            Left            =   2400
            TabIndex        =   164
            Top             =   840
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox Text1 
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
            Index           =   36
            Left            =   5760
            TabIndex        =   163
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   35
            Left            =   3240
            TabIndex        =   162
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 1, 2, 3, 4, 5, 6, 7, 8, 9)"
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
            Index           =   77
            Left            =   3360
            TabIndex        =   180
            Top             =   2280
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "vertik.Dehnung:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   76
            Left            =   120
            TabIndex        =   179
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 1, 2, 3, 4, 5, 6, 8)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   75
            Left            =   3360
            TabIndex        =   178
            Top             =   1800
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "horiz.Dehnung:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   74
            Left            =   120
            TabIndex        =   177
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 6, 7, 10, 12, 24)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   73
            Left            =   3360
            TabIndex        =   176
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Schriftgröße:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   72
            Left            =   120
            TabIndex        =   175
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Rotation:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   71
            Left            =   120
            TabIndex        =   174
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Oben"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   70
            Left            =   4920
            TabIndex        =   173
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Links"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   69
            Left            =   2400
            TabIndex        =   172
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Koordinaten:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   68
            Left            =   120
            TabIndex        =   171
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Caption         =   "DM-Preis"
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
         Index           =   7
         Left            =   3360
         TabIndex        =   141
         Top             =   1920
         Visible         =   0   'False
         Width           =   6975
         Begin VB.TextBox Text1 
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
            Index           =   34
            Left            =   2400
            TabIndex        =   150
            Text            =   "Text1"
            Top             =   2160
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   33
            Left            =   2400
            TabIndex        =   149
            Text            =   "Text1"
            Top             =   1680
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   32
            Left            =   2400
            TabIndex        =   148
            Text            =   "Text1"
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "270°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   23
            Left            =   6240
            TabIndex        =   147
            Top             =   840
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "180°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   22
            Left            =   4920
            TabIndex        =   146
            Top             =   840
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "90°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   21
            Left            =   3720
            TabIndex        =   145
            Top             =   840
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "keine"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   20
            Left            =   2400
            TabIndex        =   144
            Top             =   840
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox Text1 
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
            Index           =   31
            Left            =   5760
            TabIndex        =   143
            Text            =   "Text1"
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   30
            Left            =   3240
            TabIndex        =   142
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "LP/TLP 2642: 0,1,2"
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
            Index           =   67
            Left            =   3360
            TabIndex        =   160
            Top             =   2280
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "vertik.Dehnung:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   66
            Left            =   120
            TabIndex        =   159
            Top             =   2280
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "möglich sind bei Drucker:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   65
            Left            =   3360
            TabIndex        =   158
            Top             =   1800
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "horiz.Dehnung:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   64
            Left            =   120
            TabIndex        =   157
            Top             =   1800
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "TLP 2746: 2,3,4,5,6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   63
            Left            =   3360
            TabIndex        =   156
            Top             =   2640
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Schriftgröße:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   62
            Left            =   120
            TabIndex        =   155
            Top             =   1320
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Rotation:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   61
            Left            =   120
            TabIndex        =   154
            Top             =   840
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Oben"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   60
            Left            =   4920
            TabIndex        =   153
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Druckgeschwindigkeit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   59
            Left            =   360
            TabIndex        =   152
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Koordinaten:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   58
            Left            =   240
            TabIndex        =   151
            Top             =   2640
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Caption         =   "Druckdatum Etikett"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   121
         Top             =   3840
         Visible         =   0   'False
         Width           =   2055
         Begin VB.TextBox Text1 
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
            Index           =   29
            Left            =   2400
            TabIndex        =   130
            Text            =   "Text1"
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   28
            Left            =   2400
            TabIndex        =   129
            Text            =   "Text1"
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   27
            Left            =   2400
            TabIndex        =   128
            Text            =   "Text1"
            Top             =   1200
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "270°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   19
            Left            =   6240
            TabIndex        =   127
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "180°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   18
            Left            =   4920
            TabIndex        =   126
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "90°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   17
            Left            =   3720
            TabIndex        =   125
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "keine"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   16
            Left            =   2400
            TabIndex        =   124
            Top             =   840
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox Text1 
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
            Index           =   26
            Left            =   5760
            TabIndex        =   123
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   25
            Left            =   3240
            TabIndex        =   122
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 1, 2, 3, 4, 5, 6, 7, 8, 9)"
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
            Index           =   57
            Left            =   3360
            TabIndex        =   140
            Top             =   2280
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "vertik.Dehnung:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   56
            Left            =   120
            TabIndex        =   139
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 1, 2, 3, 4, 5, 6, 8)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   55
            Left            =   3360
            TabIndex        =   138
            Top             =   1800
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "horiz.Dehnung:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   54
            Left            =   120
            TabIndex        =   137
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 6, 7, 10, 12, 24)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   53
            Left            =   3360
            TabIndex        =   136
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Schriftgröße:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   52
            Left            =   120
            TabIndex        =   135
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Rotation:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   51
            Left            =   120
            TabIndex        =   134
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Oben"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   50
            Left            =   4920
            TabIndex        =   133
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Links"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   49
            Left            =   2400
            TabIndex        =   132
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Koordinaten:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   48
            Left            =   120
            TabIndex        =   131
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Caption         =   "Bar-Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   10680
         TabIndex        =   98
         Top             =   3480
         Visible         =   0   'False
         Width           =   1215
         Begin VB.TextBox Text1 
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
            Index           =   40
            Left            =   3240
            TabIndex        =   108
            Text            =   "Text1"
            Top             =   2640
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFF00&
            Caption         =   "lesbar"
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
            Left            =   3240
            TabIndex        =   120
            Top             =   3240
            Width           =   1815
         End
         Begin VB.TextBox Text1 
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
            Index           =   24
            Left            =   3240
            TabIndex        =   107
            Text            =   "Text1"
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   23
            Left            =   3240
            TabIndex        =   106
            Text            =   "Text1"
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   22
            Left            =   3240
            TabIndex        =   105
            Text            =   "Text1"
            Top             =   1200
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "270°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   15
            Left            =   7080
            TabIndex        =   104
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "180°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   5760
            TabIndex        =   103
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "90°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   4560
            TabIndex        =   102
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "keine"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   3240
            TabIndex        =   101
            Top             =   840
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox Text1 
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
            Index           =   21
            Left            =   5760
            TabIndex        =   100
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   20
            Left            =   3240
            TabIndex        =   99
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Höhe des BarCodes:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   78
            Left            =   120
            TabIndex        =   181
            Top             =   2760
            Width           =   2775
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "BarCode ist ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   47
            Left            =   120
            TabIndex        =   119
            Top             =   3240
            Width           =   2175
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind alle Werte von 2 - 30)"
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
            Index           =   46
            Left            =   4200
            TabIndex        =   118
            Top             =   2280
            Width           =   4455
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Breite breite Striche:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   45
            Left            =   120
            TabIndex        =   117
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 2, 3, 4)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   44
            Left            =   4200
            TabIndex        =   116
            Top             =   1800
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Breite schmale Striche:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   43
            Left            =   120
            TabIndex        =   115
            Top             =   1800
            Width           =   3015
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 8 für EAN8 oder 13 für EAN13)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   42
            Left            =   4200
            TabIndex        =   114
            Top             =   1320
            Width           =   5655
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Typ BarCode:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   41
            Left            =   120
            TabIndex        =   113
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Rotation:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   40
            Left            =   120
            TabIndex        =   112
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Oben"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   39
            Left            =   4920
            TabIndex        =   111
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Links"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   38
            Left            =   2400
            TabIndex        =   110
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Koordinaten:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   37
            Left            =   120
            TabIndex        =   109
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Caption         =   "Artikel - Bezeichnung"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   9840
         TabIndex        =   78
         Top             =   3120
         Visible         =   0   'False
         Width           =   2055
         Begin VB.TextBox Text1 
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
            Index           =   19
            Left            =   2400
            TabIndex        =   87
            Text            =   "Text1"
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   18
            Left            =   2400
            TabIndex        =   86
            Text            =   "Text1"
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   17
            Left            =   2400
            TabIndex        =   85
            Text            =   "Text1"
            Top             =   1200
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "270°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   6240
            TabIndex        =   84
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "180°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   4920
            TabIndex        =   83
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "90°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   3720
            TabIndex        =   82
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "keine"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   2400
            TabIndex        =   81
            Top             =   840
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox Text1 
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
            Index           =   16
            Left            =   5760
            TabIndex        =   80
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   15
            Left            =   3240
            TabIndex        =   79
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 1, 2, 3, 4, 5, 6, 7, 8, 9)"
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
            Index           =   36
            Left            =   3360
            TabIndex        =   97
            Top             =   2280
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "vertik.Dehnung:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   35
            Left            =   120
            TabIndex        =   96
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 1, 2, 3, 4, 5, 6, 8)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   34
            Left            =   3360
            TabIndex        =   95
            Top             =   1800
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "horiz.Dehnung:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   33
            Left            =   120
            TabIndex        =   94
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 6, 7, 10, 12, 24)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   32
            Left            =   3360
            TabIndex        =   93
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Schriftgröße:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   31
            Left            =   120
            TabIndex        =   92
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Rotation:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   30
            Left            =   120
            TabIndex        =   91
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Oben"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   28
            Left            =   4920
            TabIndex        =   90
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Links"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   26
            Left            =   2400
            TabIndex        =   89
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Koordinaten:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   25
            Left            =   120
            TabIndex        =   88
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Caption         =   "Trennlinie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   10560
         TabIndex        =   67
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
         Begin VB.TextBox Text1 
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
            Index           =   14
            Left            =   5760
            TabIndex        =   71
            Text            =   "Text1"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   13
            Left            =   3240
            TabIndex        =   70
            Text            =   "Text1"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   12
            Left            =   5760
            TabIndex        =   69
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   11
            Left            =   3240
            TabIndex        =   68
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "vertikal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   24
            Left            =   4800
            TabIndex        =   77
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "horizontal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   29
            Left            =   1920
            TabIndex        =   76
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Ausdehnung:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   27
            Left            =   120
            TabIndex        =   75
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Oben"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   23
            Left            =   4920
            TabIndex        =   74
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Links"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   22
            Left            =   2400
            TabIndex        =   73
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Koordinaten:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   21
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Caption         =   "Artikel - Bestellnummer "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   9720
         TabIndex        =   47
         Top             =   2400
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFF00&
            Caption         =   "drucke Bestell-Nr statt Firmenname"
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
            Index           =   1
            Left            =   2400
            TabIndex        =   183
            Top             =   2760
            Width           =   4335
         End
         Begin VB.TextBox Text1 
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
            Index           =   6
            Left            =   3240
            TabIndex        =   48
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   7
            Left            =   5760
            TabIndex        =   49
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "keine"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   2400
            TabIndex        =   50
            Top             =   840
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "90°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   3720
            TabIndex        =   51
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "180°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   4920
            TabIndex        =   52
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "270°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   6240
            TabIndex        =   53
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox Text1 
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
            Index           =   8
            Left            =   2400
            TabIndex        =   54
            Text            =   "Text1"
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   9
            Left            =   2400
            TabIndex        =   55
            Text            =   "Text1"
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   10
            Left            =   2400
            TabIndex        =   56
            Text            =   "Text1"
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Koordinaten:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   20
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Links"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   19
            Left            =   2400
            TabIndex        =   65
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Oben"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   18
            Left            =   4920
            TabIndex        =   64
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Rotation:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   17
            Left            =   120
            TabIndex        =   63
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Schriftgröße:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   16
            Left            =   120
            TabIndex        =   62
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 6, 7, 10, 12, 24)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   15
            Left            =   3360
            TabIndex        =   61
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "horiz.Dehnung:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   120
            TabIndex        =   60
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 1, 2, 3, 4, 5, 6, 8)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   3360
            TabIndex        =   59
            Top             =   1800
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "vertik.Dehnung:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   120
            TabIndex        =   58
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 1, 2, 3, 4, 5, 6, 7, 8, 9)"
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
            Index           =   11
            Left            =   3360
            TabIndex        =   57
            Top             =   2280
            Width           =   3615
         End
      End
      Begin sevCommand3.Command Command2 
         Height          =   735
         Index           =   9
         Left            =   120
         TabIndex        =   45
         Top             =   7080
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   44
         Top             =   6360
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
         Caption         =   "Euro-Preis"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   7
         Left            =   120
         TabIndex        =   43
         Top             =   5640
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
         Caption         =   "Druck"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   42
         Top             =   4920
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
         Caption         =   "Datum"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   41
         Top             =   4200
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
         Caption         =   "Bar-Code"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   40
         Top             =   3480
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
         Caption         =   "Artikel-Bez."
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   39
         Top             =   2760
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
         Caption         =   "Trennlinie"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   2040
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
         Caption         =   "Bestell-Nr"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   1320
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
         Caption         =   "Firmenname"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   600
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
         Caption         =   "Drucker"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Caption         =   "Firmen-Name und Druckposition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   2280
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFF00&
            Caption         =   "drucke Firmenname statt Bestell-Nr"
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
            Index           =   0
            Left            =   2400
            TabIndex        =   182
            Top             =   3360
            Value           =   1  'Aktiviert
            Width           =   4335
         End
         Begin VB.TextBox Text1 
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
            Index           =   5
            Left            =   2400
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   2640
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Index           =   4
            Left            =   2400
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Left            =   2400
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   1680
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "270°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   6240
            TabIndex        =   26
            Top             =   1320
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "180°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   4920
            TabIndex        =   25
            Top             =   1320
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "90°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3720
            TabIndex        =   24
            Top             =   1320
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFF00&
            Caption         =   "keine"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2400
            TabIndex        =   23
            Top             =   1320
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox Text1 
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
            Left            =   5760
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Left            =   3240
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Left            =   2400
            MaxLength       =   40
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   240
            Width           =   6735
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 1, 2, 3, 4, 5, 6, 7, 8, 9)"
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
            Index           =   10
            Left            =   3360
            TabIndex        =   35
            Top             =   2760
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "vertik.Dehnung:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   33
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 1, 2, 3, 4, 5, 6, 8)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   3360
            TabIndex        =   32
            Top             =   2280
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "horiz.Dehnung:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   30
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "(möglich sind 6, 7, 10, 12, 24)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   3360
            TabIndex        =   29
            Top             =   1800
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Schriftgröße:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   27
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Rotation:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   22
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Oben"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   4920
            TabIndex        =   20
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Links"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2400
            TabIndex        =   18
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Koordinaten:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            Caption         =   "Firmenname:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'Kein
         Caption         =   "Drucker-Anschluß"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   2280
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "LPT3"
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
            Index           =   2
            Left            =   240
            TabIndex        =   258
            Top             =   1320
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "LPT2"
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
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFF00&
            Caption         =   "LPT1"
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
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Label Label6 
         Caption         =   "Etiketten-Konfiguration"
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
         Left            =   120
         TabIndex        =   260
         Top             =   120
         Width           =   11655
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   6615
      Left            =   0
      TabIndex        =   249
      Top             =   1680
      Width           =   11775
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   4
         Left            =   9960
         TabIndex        =   319
         Top             =   1080
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
         Caption         =   "Etiketten weg?"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   2
         Left            =   9840
         TabIndex        =   267
         Top             =   6240
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
      Begin MSComctlLib.ProgressBar pbrZeit 
         Height          =   180
         Left            =   8040
         TabIndex        =   259
         Top             =   1560
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   1
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   3
         Left            =   8040
         TabIndex        =   9
         Top             =   120
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
         Caption         =   "Exportiere Etiketten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00808000&
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
         ForeColor       =   &H0000FFFF&
         Height          =   1335
         Left            =   8040
         TabIndex        =   254
         Top             =   2880
         Width           =   3735
         Begin VB.ComboBox cboStrichDINA4 
            Height          =   330
            Left            =   120
            TabIndex        =   346
            Text            =   "Combo1"
            Top             =   720
            Width           =   3495
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00808000&
            Caption         =   "Laser"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   256
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00808000&
            Caption         =   "Tintenstrahl"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   255
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Zentriert
            BackColor       =   &H00C0C000&
            BackStyle       =   0  'Transparent
            Caption         =   "bestellen"
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
            Index           =   5
            Left            =   120
            MouseIcon       =   "frmWKL30.frx":1D35
            MousePointer    =   99  'Benutzerdefiniert
            TabIndex        =   386
            ToolTipText     =   "per Email bestellen"
            Top             =   1080
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label12 
            BackColor       =   &H00808000&
            Caption         =   "Strichcode-Etiketten auf DIN A4"
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
            Left            =   0
            TabIndex        =   264
            Top             =   0
            Width           =   3615
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00808000&
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
         ForeColor       =   &H0000FFFF&
         Height          =   975
         Left            =   8040
         TabIndex        =   251
         Top             =   1800
         Width           =   3735
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
            Height          =   315
            Left            =   2640
            TabIndex        =   252
            Text            =   "Text2"
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label9 
            BackColor       =   &H00808000&
            Caption         =   " Anzahl Leer-Etiketten"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   0
            TabIndex        =   262
            Top             =   0
            Width           =   2535
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Leer-Etiketten dienen zum Überspringen bereits verbrauchter Etiketten auf dem Bogen Papier"
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   0
            TabIndex        =   253
            Top             =   360
            Width           =   3735
         End
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   1
         Left            =   8040
         TabIndex        =   8
         Top             =   600
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
         Caption         =   "Etikett-Layout"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   2
         Left            =   9960
         TabIndex        =   6
         Top             =   120
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
         Caption         =   "Markiere alle Sätze"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   375
         Index           =   0
         Left            =   9960
         TabIndex        =   7
         Top             =   600
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
         Caption         =   "Etiketten löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6150
         Left            =   120
         MultiSelect     =   1  '1 -Einfach
         TabIndex        =   5
         Top             =   360
         Width           =   7815
      End
      Begin VB.ListBox List1 
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
         TabIndex        =   250
         Top             =   120
         Width           =   7815
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00808000&
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
         ForeColor       =   &H0000FFFF&
         Height          =   1815
         Left            =   8040
         TabIndex        =   257
         Top             =   4320
         Width           =   3735
         Begin VB.TextBox txtTage 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   365
            Text            =   "60"
            Top             =   1200
            Width           =   615
         End
         Begin VB.ComboBox cboRegalDinA4 
            Height          =   330
            Left            =   120
            TabIndex        =   364
            Text            =   "Combo1"
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label18 
            Caption         =   "VK-Mengenangaben für spez. Etiketten:"
            Height          =   255
            Left            =   120
            TabIndex        =   367
            Top             =   840
            Width           =   3495
         End
         Begin VB.Label Label19 
            Caption         =   "Tage"
            Height          =   255
            Left            =   840
            TabIndex        =   366
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label11 
            BackColor       =   &H00808000&
            Caption         =   " Regaletiketten auf DIN A4 (alles in mm)"
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
            Left            =   0
            TabIndex        =   263
            Top             =   0
            Width           =   3735
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      Caption         =   "Etikettendatei:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   9120
      TabIndex        =   244
      Top             =   600
      Width           =   2655
      Begin VB.CheckBox Check13 
         Caption         =   "nur die Eigenen"
         Height          =   240
         Left            =   120
         TabIndex        =   265
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00808000&
         Caption         =   " Etikettendatei:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   261
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00808000&
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
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   248
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00808000&
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
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   247
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808000&
         Caption         =   "Anzahl Etiketten:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   246
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808000&
         Caption         =   "Anzahl Artikel:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   245
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   2
      RTSEnable       =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   8
      Left            =   2400
      TabIndex        =   287
      Top             =   1120
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   873
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
      Caption         =   "Regal endlos"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   266
      Top             =   1120
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   873
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
      Caption         =   "Strichcode endlos"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   9
      Left            =   8520
      TabIndex        =   338
      ToolTipText     =   "Anzahl Etiketten"
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   8
      Left            =   8520
      TabIndex        =   337
      ToolTipText     =   "Anzahl Etiketten"
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   7920
      TabIndex        =   336
      ToolTipText     =   "Anzal Artikel"
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
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
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   7920
      TabIndex        =   335
      ToolTipText     =   "Anzahl Artikel"
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "nicht markiert:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   6360
      TabIndex        =   334
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Rechts
      BackColor       =   &H00808000&
      Caption         =   "markiert:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6360
      TabIndex        =   333
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Etiketten drucken"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmWKL30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bOk             As Boolean
Dim iErsteZeile     As Integer
Dim iZweiteZeile    As Integer
Dim dePlus          As Double
Dim deNull          As Double
Dim deMinus         As Double
Private Sub PositionierenWKL30()
    On Error GoTo LOKAL_ERROR
    
    Frame0.Top = 4680
    Frame0.Left = 2160
    Frame0.Height = 4095
    Frame0.Width = 9615
    
    Frame1.Top = 0
    Frame1.Left = 0
    Frame1.Height = 9000
    Frame1.Width = 12000
    
    Frame2(0).Top = 600
    Frame2(0).Left = 2280
    Frame2(0).Height = 1815
    Frame2(0).Width = 2055
    
    Frame2(1).Top = 600
    Frame2(1).Left = 2280
    Frame2(1).Height = 3855
    Frame2(1).Width = 9495
    
    Frame2(2).Top = 600
    Frame2(2).Left = 2280
    Frame2(2).Height = 3255
    Frame2(2).Width = 9495
 
    Frame2(3).Top = 600
    Frame2(3).Left = 2280
    Frame2(3).Height = 1335
    Frame2(3).Width = 6975
    
    Frame2(4).Top = 600
    Frame2(4).Left = 2280
    Frame2(4).Height = 3255
    Frame2(4).Width = 9495
    
    Frame2(5).Top = 600
    Frame2(5).Left = 2280
    Frame2(5).Height = 3735
    Frame2(5).Width = 9495
    
    Frame2(6).Top = 600
    Frame2(6).Left = 2280
    Frame2(6).Height = 3255
    Frame2(6).Width = 9495
    
    Frame2(7).Top = 600
    Frame2(7).Left = 2280
    Frame2(7).Height = 3255
    Frame2(7).Width = 9495
    
    Frame2(8).Top = 600
    Frame2(8).Left = 2280
    Frame2(8).Height = 3255
    Frame2(8).Width = 9495
    
    Frame4.Top = 1680
    Frame4.Left = 0
    Frame4.Height = 7215
    Frame4.Width = 11895
    
    Frame8.Top = 480
    Frame8.Left = 0
    Frame8.Height = 8055
    Frame8.Width = 11895
    
    Frame10.Top = 480
    Frame10.Left = 0
    Frame10.Height = 8055
    Frame10.Width = 11895
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL30"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub BlendeAlleFramesAus()
    On Error GoTo LOKAL_ERROR
    Dim lcount As Long
    
    For lcount = 0 To 8
        Frame2(lcount).Visible = False
    Next lcount
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BlendeAlleFramesAus"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeEtikettenWKL30()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount      As Long
    Dim lCount2     As Long
    Dim lAnzahl     As Long
    Dim lAnz        As Long
    Dim cArtNr      As String
    Dim cPfad       As String
    Dim cSQL        As String
    Dim ctmp        As String
    Dim cDatum      As String
    Dim iFileNr     As Integer
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cDatum = Right(Trim$(Str$(Year(Now))), 2) & Format$(Month(Now), "00")
    loeschNEW "etidru2", gdBase
    
    cSQL = "Create Table ETIDRU2 "
    cSQL = cSQL & "( ARTNR Long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", BESTAND Long"
    cSQL = cSQL & ", ANZAHL Long"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", EAN Text(13)"
    cSQL = cSQL & ", LINR Long"
    cSQL = cSQL & ", LPZ Long"
    cSQL = cSQL & ", DATUM Text(6)"
    cSQL = cSQL & ", FILNR Long"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
        
    cSQL = "Delete from ETIDRU2"
    gdBase.Execute cSQL, dbFailOnError
        
    cSQL = "Delete from ETIDRU where ARTNR is NULL"
    gdBase.Execute cSQL, dbFailOnError
        
    lAnz = Val(Text2.Text)
    For lcount = 1 To lAnz
        cSQL = "Insert into ETIDRU2 "
        cSQL = cSQL & "(ARTNR, BEZEICH, VKPR, BESTAND, ANZAHL, LIBESNR, EAN, LINR, LPZ, DATUM, FILNR) "
        cSQL = cSQL & "values "
        cSQL = cSQL & "(NULL, '.' , NULL, NULL, " & Trim$(Str$(lAnz)) & ", NULL, NULL, NULL, NULL, NULL, NULL ) "
        gdBase.Execute cSQL, dbFailOnError
    Next lcount
    
    Dim cAnzEti As String
        
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            cArtNr = Left(List2.list(lcount), 6)
            cArtNr = Trim$(cArtNr)
            
            
            cAnzEti = Trim$(Mid(List2.list(lcount), Len(List2.list(lcount)) - 12, 10))
                        '//aenderung
                        ctmp = Left(cAnzEti, 3)
                        ctmp = Trim$(ctmp)
'                        cAnzEti = Val(ctmp)
                    
'            ctmp = Right(List2.list(lcount), 6)
'            ctmp = Left(ctmp, 2)
'            ctmp = Trim$(ctmp)

            lAnzahl = Val(ctmp)
            For lCount2 = 1 To lAnzahl
                cSQL = "Insert into ETIDRU2 Select "
                cSQL = cSQL & " artnr , BEZEICH, vkpr, BESTAND, ANZAHL, LIBESNR, EAN, LINR, LPZ, FILNR"
                cSQL = cSQL & " , '" & cDatum & "' as DATUM from ETIDRU where ARTNR = " & cArtNr
                gdBase.Execute cSQL, dbFailOnError
            Next lCount2
        End If
    Next lcount

    cSQL = "Delete from ETIDRU2 where ANZAHL <= 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    If Modul6.FindFile(gcDBPfad, "aWKL30t.rpt") Then
        reportbildschirm "spezial", "aWKL30t"
    Else
        reportbildschirm "WKL032", "aWKL30"
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeEtikettenWKL30"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeGrundPreisEtikettenWKL30(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount          As Long
    Dim cSQL            As String
    Dim cSQL1           As String
    Dim rsrs            As Recordset
    Dim rsZiel          As Recordset
    Dim dWert           As Double
    Dim cPfad           As String
    Dim cEAN            As String
    Dim cEANCode        As String
    Dim cDruckdatum     As String
    Dim cBezeich        As String
    Dim lAnz            As Long
    Dim dInhalt         As Double
    Dim cInhaltBez      As String
    Dim dVkPr           As Double
    Dim dGrundPreisDM   As Double
    Dim dGrundPreisEur  As Double
    Dim cGrundInhalt    As String
    Dim cLVK            As String
    Dim dLVK            As Double
    Dim iFileNr         As Integer
    Dim sEtikettKürzel  As String

    loeschNEW "DRU_GRUN", gdBase
    
    cSQL = "Create Table DRU_GRUN ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    cDruckdatum = Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")
    
    cSQL = "Select * from DRU_GRUN"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    lAnz = Val(Text2.Text)
    For lcount = 1 To lAnz
        rsZiel.AddNew
        rsZiel!FirmaName = ""
        rsZiel!artnr = Null
        rsZiel!BEZEICH = ""
        rsZiel!Barcode = ""
        rsZiel!LIBESNR = ""
        rsZiel!vkpr = Null
        rsZiel!vkpr_EUR = Null
        rsZiel!INHALT = Null
        rsZiel!INHALTBEZ = ""
        rsZiel!DRUCKDATUM = ""
        rsZiel!GRUNDPREIS = ""
        rsZiel!GRUND_INH = Null
        rsZiel!GRUND_DM = Null
        rsZiel!GRUND_EUR = Null
        rsZiel.Update
    Next lcount
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If rsrs.EOF Then

        Else
            rsZiel.AddNew
            rsZiel!FirmaName = gFirma.FirmaName
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich
            
            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            rsZiel!EAN = cEAN
            
            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode
            
            rsZiel!LIBESNR = rsrs!LIBESNR
            
            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If
                
            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = ermvorz(rsrs!artnr) & cDruckdatum
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If
                
                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If
            
            '//DRU_GRUN
            rsZiel.Update
            
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    '********************************************************
    '* Drucker-Formblatt-Konfiguration merken!
    '********************************************************
    dWert = 0
    If Option3(0).Value = True Then
        dWert = dWert + 0
    Else
        dWert = dWert + 1
    End If
    
    If cboRegalDinA4.Text = "selbstklebend 52,5 x 30" Then
        dWert = dWert + 2
    Else
        dWert = dWert + 4
    End If
    cBezeich = Trim$(Str$(dWert))
    
    iFileNr = FreeFile
    Open gcPfad & "ETIDRU.CFG" For Binary As #iFileNr
    Close iFileNr
    Kill gcPfad & "ETIDRU.CFG"
    iFileNr = FreeFile
    Open gcPfad & "ETIDRU.CFG" For Binary As #iFileNr
    Put #iFileNr, 1, cBezeich
    Close iFileNr
    
    '********************************************************
    '* Drucker-Formblatt-Konfiguration merken!
    '********************************************************
    
    If cboRegalDinA4.Text = "selbstklebend 52,5 x 30" Then
        If Modul6.FindFile(gcDBPfad, "aWKL30as.rpt") Then
            reportbildschirm "spezial", "aWKL30as"
        Else
            reportbildschirm "WKL017", "aWKL30a"
        End If
        
    ElseIf cboRegalDinA4.Text = "selbstklebend 70 x 36" Then
    
        If Modul6.FindFile(gcDBPfad, "aWKL30bs.rpt") Then
            reportbildschirm "spezial", "aWKL30bs"
        Else
            reportbildschirm "WKL017", "aWKL30b"
        End If
         
    ElseIf cboRegalDinA4.Text = "selbstklebend 35,6 x 16,9" Then
    
        reportbildschirm "WKL021e", "aWKL30eV3"

    ElseIf cboRegalDinA4.Text = "selbstklebend Sonder Etikett" Then
    
        If Modul6.FindFile(gcDBPfad, "aWKL30bd.rpt") Then
            reportbildschirm "WKL017", "aWKL30bd"
        End If
        
    ElseIf cboRegalDinA4.Text = "perforiert 38 x 70" Then
    
        If Modul6.FindFile(gcDBPfad, "aWKL30p7s.rpt") Then
            'Achtung bei diesem Spezialetikett Kommt dazu:
            'Netto-Preis
            'ListenVK
            SpalteAnfuegenNEW "DRU_GRUN", "LVK", "Text(20)", gdBase
            SpalteAnfuegenNEW "DRU_GRUN", "NETTOPR", "double", gdBase
            
            cSQL = "Select * from DRU_GRUN "
            Set rsZiel = gdBase.OpenRecordset(cSQL)
            
            Dim cNettoPr As String
            Dim cKVKPR As String
            Dim cArtNr As String

            If Not rsZiel.EOF Then
                rsZiel.MoveFirst
                Do While Not rsZiel.EOF
            
                    If Not IsNull(rsZiel!artnr) Then
                        cArtNr = rsZiel!artnr
                    End If
                    
                    If Val(cArtNr) > 0 Then
                
                        cSQL = "Select * from Artikel where artnr = " & cArtNr
                        Set rsrs = gdBase.OpenRecordset(cSQL)
                        If Not rsrs.EOF Then
                        
                            If Not IsNull(rsrs!KVKPR1) Then
                                dWert = rsrs!KVKPR1
                            End If
                        
                            cNettoPr = nettoR(dWert, rsrs!MWST)
                            cNettoPr = Format$(cNettoPr, "#####0.00")
                            cNettoPr = Trim(cNettoPr)
                            
                            dLVK = CDbl(cNettoPr) * 80 / 100
                            
                            cLVK = Format$(dLVK, "#####0.00")
                            
                            cLVK = SwapStr(cLVK, ",", "")
                            cLVK = "HWHP0000" & cLVK
        
                            rsZiel.Edit
                            rsZiel!NETTOPR = cNettoPr
                            rsZiel!LVK = cLVK
                            rsZiel.Update
                            
                        End If
                        rsrs.Close: Set rsrs = Nothing
                    
                    End If
                    rsZiel.MoveNext
                Loop
            End If
            rsZiel.Close: Set rsZiel = Nothing
            
            
            reportbildschirm "spezial", "aWKL30p7s"
        Else
            reportbildschirm "", "aWKL30p7"
        End If
        
    ElseIf cboRegalDinA4.Text = "perforiert 38 x 50" Then
    
        If Modul6.FindFile(gcDBPfad, "aWKL30p5_spez.rpt") Then
            reportbildschirm "", "aWKL30p5_spez"
        Else

            'Spalten anfügen
            If SpalteInTabellegefundenNEW("DRU_GRUN", "FARBTEXT", gdBase) = False Then
                
                SpalteAnfuegenNEW "DRU_GRUN", "FARBTEXT", "Text(35)", gdBase
                SpalteAnfuegenNEW "DRU_GRUN", "FARBwert", "integer", gdBase
                SpalteAnfuegenNEW "DRU_GRUN", "FARBwertS", "integer", gdBase
                SpalteAnfuegenNEW "DRU_GRUN", "FARBNR", "integer", gdBase
                SpalteAnfuegenNEW "DRU_GRUN", "AWM", "Text(2)", gdBase
            End If
            
            cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.Artnr = Artikel.artnr set DRU_GRUN.AWM = artikel.AWM "
            gdBase.Execute cSQL, dbFailOnError
        
            cSQL = "update DRU_GRUN set farbnr = Val(AWM) "
            gdBase.Execute cSQL, dbFailOnError
        
            BringFarbeInsSpiel "DRU_GRUN", gdBase
    
            reportbildschirm "", "aWKL30p5"
            
        End If
        
    ElseIf cboRegalDinA4.Text = "perforiert 38 x 50 Variante 2" Then
    
        cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.Artnr = Artikel.artnr set DRU_GRUN.libesnr = artikel.ean "
        gdBase.Execute cSQL, dbFailOnError
    
        reportbildschirm "", "aWKL30p16"
    
    ElseIf cboRegalDinA4.Text = "perforiert 26 x 35" Then
    
        reportbildschirm "", "aWKL30p10"
        
    ElseIf cboRegalDinA4.Text = "perforiert 39 x 35" Then
    
        reportbildschirm "", "aWKL30p11"
            
    ElseIf cboRegalDinA4.Text = "perforiert 39 x 33" Then
    
        reportbildschirm "", "aWKL30p13"
            
    ElseIf cboRegalDinA4.Text = "perforiert 50 x 40" Then
    
        If gsETILS <> "" Then
            
            For lcount = 0 To lAnzahl
                cSQL = "Update DRU_GRUN inner join LSTEETI on DRU_GRUN.artnr = LSTEETI.artnr "
                cSQL = cSQL & " set DRU_GRUN.VKPR = LSTEETI.vkpr where DRU_GRUN.artnr = " & acArtNr(lcount)
                gdBase.Execute cSQL, dbFailOnError
            Next lcount
        End If

        If SpalteInTabellegefundenNEW("DRU_GRUN", "EAN1", gdBase) = False Then
            SpalteAnfuegenNEW "DRU_GRUN", "EAN1", "Text(13)", gdBase
            SpalteAnfuegenNEW "DRU_GRUN", "Zusatz", "Text(50)", gdBase
        End If

        Dim cEAN13 As String
        cDruckdatum = Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")

        cSQL = "Select * from DRU_GRUN "
        Set rsZiel = gdBase.OpenRecordset(cSQL)
        
        If Not rsZiel.EOF Then
            rsZiel.MoveFirst
            Do While Not rsZiel.EOF
            
                If Not IsNull(rsZiel!artnr) Then
                    cArtNr = rsZiel!artnr
                End If
                
                cEAN13 = ""
            
                If Not IsNull(rsZiel!EAN) Then
                    cEAN13 = rsZiel!EAN 'cEAN13
                End If
                
                If cEAN13 = "" Then
                    If Not IsNull(rsZiel!EAN2) Then
                        cEAN13 = rsZiel!EAN2 'cEAN13
                    End If
                End If
                
                If cEAN13 = "" Then
                    If Not IsNull(rsZiel!EAN3) Then
                        cEAN13 = rsZiel!EAN3 'cEAN13
                    End If
                End If
                
                cEAN13 = Trim(cEAN13)
                
                If Val(cArtNr) > 0 Then
                    rsZiel.Edit
                    
                    rsZiel!EAN1 = cEAN13
                    
                    rsZiel!FirmaName = ""
                    If gbEtiExArtikel = True Then
                        rsZiel!FirmaName = "EX"
                    Else
                        If ermEX_INFO(rsZiel!artnr) = "J" Then
                            rsZiel!FirmaName = "EX"
                        End If
                    End If
                    
                    rsZiel!Zusatz = Left(ermLiefKürzelmitkleinstenLEKPR(cArtNr, gdBase), 3) & " " & cDruckdatum & " " & ermletztVKdurch2(cArtNr, CInt(txtTage.Text)) & ""
                    rsZiel!LIBESNR = ermLiefLIBESNRmitkleinstenLEKPR(cArtNr, gdBase)
                    
                    rsZiel.Update
                End If
        
                rsZiel.MoveNext
            Loop
        End If
        rsZiel.Close: Set rsZiel = Nothing

        reportbildschirm "", "aWKL30p5041"
            
    ElseIf cboRegalDinA4.Text = "perforiert 42 x 25" Then
    
            reportbildschirm "", "aWKL30p15"
            
    ElseIf cboRegalDinA4.Text = "perforiert 39 x 33 (EAN)" Then
    
        
    
        reportbildschirm "", "aWKL30p14"
            
    ElseIf cboRegalDinA4.Text = "35,6 x 16,9 Variante 2" Then
    
        If Modul6.FindFile(gcDBPfad, "aWKL30uV2.rpt") Then
            reportbildschirm "spezial", "aWKL30uV2"
        Else
            reportbildschirm "WKL021e", "aWKL30eV2"
        End If


    ElseIf cboRegalDinA4.Text = "perforiert 26 x 47" Then
    
        cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.Artnr = Artikel.artnr set DRU_GRUN.libesnr = artikel.ean "
        gdBase.Execute cSQL, dbFailOnError
        reportbildschirm "", "aWKL30p4"
    
    ElseIf cboRegalDinA4.Text = "perforiert 29 x 52" Then
    
        cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.Artnr = Artikel.artnr set DRU_GRUN.libesnr = artikel.ean "
        gdBase.Execute cSQL, dbFailOnError
        
        
        
        If Modul6.FindFile(gcDBPfad, "aWKL30p3_spez.rpt") Then
            reportbildschirm "", "aWKL30p3_spez"
        Else

            reportbildschirm "", "aWKL30p3"
        End If
        
    ElseIf cboRegalDinA4.Text = "perforiert 29 x 52 Variante 2" Then
    
        cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.Artnr = Artikel.artnr set DRU_GRUN.libesnr = artikel.ean "
        gdBase.Execute cSQL, dbFailOnError
        
        'Spalten anfügen
        If SpalteInTabellegefundenNEW("DRU_GRUN", "FARBTEXT", gdBase) = False Then
            SpalteAnfuegenNEW "DRU_GRUN", "FARBTEXT", "Text(35)", gdBase
            SpalteAnfuegenNEW "DRU_GRUN", "FARBwert", "integer", gdBase
            SpalteAnfuegenNEW "DRU_GRUN", "FARBwertS", "integer", gdBase
            SpalteAnfuegenNEW "DRU_GRUN", "FARBNR", "integer", gdBase
            SpalteAnfuegenNEW "DRU_GRUN", "AWM", "Text(2)", gdBase
        End If
        
        cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.Artnr = Artikel.artnr set DRU_GRUN.AWM = artikel.AWM "
        gdBase.Execute cSQL, dbFailOnError
    
        cSQL = "update DRU_GRUN set farbnr = Val(AWM) "
        gdBase.Execute cSQL, dbFailOnError
    
        BringFarbeInsSpiel "DRU_GRUN", gdBase
        
        reportbildschirm "", "aWKL30p3v2"
        
        
    ElseIf cboRegalDinA4.Text = "perforiert 26 x 47 (EAN 13)" Then
    
        cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.Artnr = Artikel.artnr set DRU_GRUN.libesnr = artikel.ean "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update DRU_GRUN d inner join Artikel a on d.Artnr = a.artnr set d.libesnr = a.ean2 "
        cSQL = cSQL & " where d.libesnr is null or d.libesnr = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update DRU_GRUN d inner join Artikel a on d.Artnr = a.artnr set d.libesnr = a.ean3 "
        cSQL = cSQL & " where d.libesnr is null or d.libesnr = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        SpalteAnfuegenNEW "DRU_GRUN", "BARCODE13", "Text(18)", gdBase
        
        cSQL = "Update DRU_GRUN set barcode13 =  '' "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Select * from DRU_GRUN "
        Set rsZiel = gdBase.OpenRecordset(cSQL)
        
        If Not rsZiel.EOF Then
            rsZiel.MoveFirst
            Do While Not rsZiel.EOF
            
                 If Len(Trim(rsZiel!LIBESNR)) = 13 Then
                    cEANCode = ean13(Trim(rsZiel!LIBESNR))
                    
                    rsZiel.Edit
                    rsZiel!BARCODE13 = cEANCode
                    rsZiel.Update
                End If
                
                If Len(Trim(rsZiel!LIBESNR)) = 8 Then
                
                    cEANCode = fnCodiereEANCode(Trim(rsZiel!LIBESNR))
                    
                    rsZiel.Edit
                    rsZiel!BARCODE13 = cEANCode
                    rsZiel.Update
                    
                End If
        
                rsZiel.MoveNext
            Loop
        End If
        rsZiel.Close: Set rsZiel = Nothing
        
        
        cSQL = "Update DRU_GRUN set barcode13 = barcode, libesnr = 'Art# ' & artnr where barcode13 = '' and artnr > 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        If Modul6.FindFile(gcDBPfad, "aWKL30p2_spez.rpt") Then
        
        Else
        
        
            SpalteAnfuegenNEW "DRU_GRUN", "Zusatz", "Text(30)", gdBase
        
        
            If lAnzahl > 5 Then
                loeschNEW "ARTLIEF", gdApp
                TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTLIEF"
            
                loeschNEW "KASSTAGE", gdBase
                
                cSQL = "Select sum(Menge) as SUMMENGE, Artnr into KASSTAGE from Kassjour where "
                cSQL = cSQL & " adate >= " & CLng(DateValue(Now) - CInt(txtTage.Text))
                cSQL = cSQL & " group by Artnr "
                gdBase.Execute cSQL, dbFailOnError
                
                cSQL = "Create Index ARTNR on KASSTAGE (ARTNR)"
                gdBase.Execute cSQL, dbFailOnError
            
                loeschNEW "KASSTAGE", gdApp
                TransferTab gdBase, App.Path & "\kissapp.mdb", "KASSTAGE"
            
            
            
                loeschNEW "LISRT", gdApp
                TransferTab gdBase, App.Path & "\kissapp.mdb", "LISRT"
                
                cSQL = "Create Index ARTNR on ARTLIEF (ARTNR)"
                gdApp.Execute cSQL, dbFailOnError
                
                cSQL = "Create Index LEKPR on ARTLIEF (LEKPR)"
                gdApp.Execute cSQL, dbFailOnError
                
                cSQL = "Create Index RKZ on ARTLIEF (RKZ)"
                gdApp.Execute cSQL, dbFailOnError
                
                cSQL = "Create Index linr on LISRT (linr)"
                gdApp.Execute cSQL, dbFailOnError
                
                cSQL = "Select * from DRU_GRUN "
                Set rsZiel = gdBase.OpenRecordset(cSQL)
                
                If Not rsZiel.EOF Then
                    rsZiel.MoveFirst
                    Do While Not rsZiel.EOF
                    
                        If Not IsNull(rsZiel!artnr) Then
                            cArtNr = rsZiel!artnr
                        End If
                        
                        
                        
                        
                        
                        
                        If Val(cArtNr) > 0 Then
                            rsZiel.Edit
                            rsZiel!Zusatz = ermZusatztext(cArtNr)
                            
                            rsZiel!DRUCKDATUM = ermvorz(cArtNr) & Trim(Left(ermLiefKürzelmitkleinstenLEKPR(cArtNr, gdApp), 4))
                            
                            
                            rsZiel!FirmaName = ""
                            If gbEtiExArtikel = True Then
                                rsZiel!FirmaName = "EX"
                            Else
                                If ermEX_INFO(rsZiel!artnr) = "J" Then
                                    rsZiel!FirmaName = "EX"
                                Else
                                    rsZiel!FirmaName = ermletztVKdurch2_APP_SPEZIAL(cArtNr, CInt(txtTage.Text), "KassTage", gdApp)
                                End If
                            End If
                            
                            
                            
                            
                            
                            rsZiel.Update
                        End If
                
                        rsZiel.MoveNext
                    Loop
                End If
                rsZiel.Close: Set rsZiel = Nothing
                
                
            Else
            
                cSQL = "Select * from DRU_GRUN "
                Set rsZiel = gdBase.OpenRecordset(cSQL)
                
                If Not rsZiel.EOF Then
                    rsZiel.MoveFirst
                    Do While Not rsZiel.EOF
                    
                        If Not IsNull(rsZiel!artnr) Then
                            cArtNr = rsZiel!artnr
                        End If
                        
                        If Val(cArtNr) > 0 Then
                            rsZiel.Edit
                            rsZiel!Zusatz = ermZusatztext(cArtNr)
                            
                            rsZiel!DRUCKDATUM = ermvorz(cArtNr) & Trim(Left(ermLiefKürzelmitkleinstenLEKPR(cArtNr, gdBase), 4))
                            
                            rsZiel!FirmaName = ""
                            If gbEtiExArtikel = True Then
                                rsZiel!FirmaName = "EX"
                            Else
                                If ermEX_INFO(rsZiel!artnr) = "J" Then
                                    rsZiel!FirmaName = "EX"
                                Else
                                    rsZiel!FirmaName = ermletztVKdurch2(cArtNr, CInt(txtTage.Text))
                                End If
                            End If
                            
                            
                            rsZiel.Update
                        End If
                
                        rsZiel.MoveNext
                    Loop
                End If
                rsZiel.Close: Set rsZiel = Nothing
                
                
            End If
        
        
        
        
        
        
            
            
            
        End If
        
        If Modul6.FindFile(gcDBPfad, "aWKL30p2_spez.rpt") Then
            reportbildschirm "", "aWKL30p2_spez"
        Else
            reportbildschirm "", "aWKL30p2"
        End If
        
        
    ElseIf cboRegalDinA4.Text = "perforiert 26 x 47 (EAN 13) W" Then
    
        cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.Artnr = Artikel.artnr set DRU_GRUN.libesnr = artikel.ean "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update DRU_GRUN d inner join Artikel a on d.Artnr = a.artnr set d.libesnr = a.ean2 "
        cSQL = cSQL & " where d.libesnr is null or d.libesnr = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update DRU_GRUN d inner join Artikel a on d.Artnr = a.artnr set d.libesnr = a.ean3 "
        cSQL = cSQL & " where d.libesnr is null or d.libesnr = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        SpalteAnfuegenNEW "DRU_GRUN", "BARCODE13", "Text(18)", gdBase
        
        cSQL = "Update DRU_GRUN set barcode13 =  '' "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Select * from DRU_GRUN "
        Set rsZiel = gdBase.OpenRecordset(cSQL)
        
        If Not rsZiel.EOF Then
            rsZiel.MoveFirst
            Do While Not rsZiel.EOF
            
                 If Len(Trim(rsZiel!LIBESNR)) = 13 Then
                    cEANCode = ean13(Trim(rsZiel!LIBESNR))
                    
                    rsZiel.Edit
                    rsZiel!BARCODE13 = cEANCode
                    rsZiel.Update
                End If
                
                If Len(Trim(rsZiel!LIBESNR)) = 8 Then
                
                    cEANCode = fnCodiereEANCode(Trim(rsZiel!LIBESNR))
                    
                    rsZiel.Edit
                    rsZiel!BARCODE13 = cEANCode
                    rsZiel.Update
                    
                End If
        
                rsZiel.MoveNext
            Loop
        End If
        rsZiel.Close: Set rsZiel = Nothing
        
        
        cSQL = "Update DRU_GRUN set barcode13 = barcode, libesnr = 'Art# ' & artnr where barcode13 = '' and artnr > 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        
        
        
        SpalteAnfuegenNEW "DRU_GRUN", "Zusatz", "Text(30)", gdBase
    
    
        If lAnzahl > 5 Then
            loeschNEW "ARTLIEF", gdApp
            TransferTab gdBase, App.Path & "\kissapp.mdb", "ARTLIEF"
        
            loeschNEW "KASSTAGE", gdBase
            
            cSQL = "Select sum(Menge) as SUMMENGE, Artnr into KASSTAGE from Kassjour where "
            cSQL = cSQL & " adate >= " & CLng(DateValue(Now) - CInt(txtTage.Text))
            cSQL = cSQL & " group by Artnr "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Create Index ARTNR on KASSTAGE (ARTNR)"
            gdBase.Execute cSQL, dbFailOnError
        
            loeschNEW "KASSTAGE", gdApp
            TransferTab gdBase, App.Path & "\kissapp.mdb", "KASSTAGE"
        
        
        
            loeschNEW "LISRT", gdApp
            TransferTab gdBase, App.Path & "\kissapp.mdb", "LISRT"
            
            cSQL = "Create Index ARTNR on ARTLIEF (ARTNR)"
            gdApp.Execute cSQL, dbFailOnError
            
            cSQL = "Create Index LEKPR on ARTLIEF (LEKPR)"
            gdApp.Execute cSQL, dbFailOnError
            
            cSQL = "Create Index RKZ on ARTLIEF (RKZ)"
            gdApp.Execute cSQL, dbFailOnError
            
            cSQL = "Create Index linr on LISRT (linr)"
            gdApp.Execute cSQL, dbFailOnError
            
            cSQL = "Select * from DRU_GRUN "
            Set rsZiel = gdBase.OpenRecordset(cSQL)
            
            If Not rsZiel.EOF Then
                rsZiel.MoveFirst
                Do While Not rsZiel.EOF
                
                    If Not IsNull(rsZiel!artnr) Then
                        cArtNr = rsZiel!artnr
                    End If
                    
                    If Val(cArtNr) > 0 Then
                        rsZiel.Edit
                        rsZiel!Zusatz = ermZusatztext(cArtNr)
                        
                        rsZiel!DRUCKDATUM = Trim(Left(ermLiefKürzelmitkleinstenLEKPR(cArtNr, gdApp), 5))
                        rsZiel!FirmaName = ermletztVKdurch2_APP_SPEZIAL(cArtNr, CInt(txtTage.Text), "KassTage", gdApp)
                        rsZiel.Update
                    End If
            
                    rsZiel.MoveNext
                Loop
            End If
            rsZiel.Close: Set rsZiel = Nothing
            
            
        Else
        
            cSQL = "Select * from DRU_GRUN "
            Set rsZiel = gdBase.OpenRecordset(cSQL)
            
            If Not rsZiel.EOF Then
                rsZiel.MoveFirst
                Do While Not rsZiel.EOF
                
                    If Not IsNull(rsZiel!artnr) Then
                        cArtNr = rsZiel!artnr
                    End If
                    
                    If Val(cArtNr) > 0 Then
                        rsZiel.Edit
                        rsZiel!Zusatz = ermZusatztext(cArtNr)
                        
                        rsZiel!DRUCKDATUM = Trim(Left(ermLiefKürzelmitkleinstenLEKPR(cArtNr, gdBase), 5))
                        rsZiel!FirmaName = ermletztVKdurch2(cArtNr, CInt(txtTage.Text))
                        rsZiel.Update
                    End If
            
                    rsZiel.MoveNext
                Loop
            End If
            rsZiel.Close: Set rsZiel = Nothing
                
        End If
        
        reportbildschirm "", "aWKL30p2W"
       
    ElseIf cboRegalDinA4.Text = "perforiert 38 x 42 (EAN 13)" Then
    
        cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.Artnr = Artikel.artnr set DRU_GRUN.libesnr = artikel.ean "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update DRU_GRUN d inner join Artikel a on d.Artnr = a.artnr set d.libesnr = a.ean2 "
        cSQL = cSQL & " where d.libesnr is null or d.libesnr = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update DRU_GRUN d inner join Artikel a on d.Artnr = a.artnr set d.libesnr = a.ean3 "
        cSQL = cSQL & " where d.libesnr is null or d.libesnr = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        SpalteAnfuegenNEW "DRU_GRUN", "EANKLAR", "Text(13)", gdBase
        
        cSQL = "Update DRU_GRUN set EANKLAR =  '' "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update DRU_GRUN d inner join Artikel a on d.Artnr = a.artnr set d.EANKLAR = a.ean "
        gdBase.Execute cSQL, dbFailOnError
        
        SpalteAnfuegenNEW "DRU_GRUN", "BARCODE13", "Text(18)", gdBase
        
        cSQL = "Update DRU_GRUN set barcode13 =  '' "
        gdBase.Execute cSQL, dbFailOnError
            
        cSQL = "Select * from DRU_GRUN "
        Set rsZiel = gdBase.OpenRecordset(cSQL)
        
        If Not rsZiel.EOF Then
            rsZiel.MoveFirst
            Do While Not rsZiel.EOF
            
                If Len(Trim(rsZiel!LIBESNR)) = 13 Then
                    cEANCode = ean13(Trim(rsZiel!LIBESNR))
                    
                    rsZiel.Edit
                    rsZiel!BARCODE13 = cEANCode
                    rsZiel.Update
                End If
                

                If Len(Trim(rsZiel!LIBESNR)) = 8 Then
                    cEANCode = fnCodiereEANCode(Trim(rsZiel!LIBESNR))
                    
                    rsZiel.Edit
                    rsZiel!BARCODE13 = cEANCode
                    rsZiel.Update
                End If
        
                rsZiel.MoveNext
            Loop
        End If
        rsZiel.Close: Set rsZiel = Nothing
        
        cSQL = "Update DRU_GRUN set barcode13 = barcode, libesnr = 'Art# ' & artnr where barcode13 = '' and artnr > 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        SpalteAnfuegenNEW "DRU_GRUN", "Zusatz", "Text(30)", gdBase
            
        cSQL = "Select * from DRU_GRUN "
        Set rsZiel = gdBase.OpenRecordset(cSQL)
        
        If Not rsZiel.EOF Then
            rsZiel.MoveFirst
            Do While Not rsZiel.EOF
            
                If Not IsNull(rsZiel!artnr) Then
                    cArtNr = rsZiel!artnr
                End If
                
                If Val(cArtNr) > 0 Then
                    rsZiel.Edit
                    rsZiel!Zusatz = ermZusatztext(cArtNr)
                    rsZiel.Update
                End If
        
                rsZiel.MoveNext
            Loop
        End If
        rsZiel.Close: Set rsZiel = Nothing
        
        cSQL = "Update DRU_GRUN set eanklar = '' where libesnr = eanklar "
        gdBase.Execute cSQL, dbFailOnError
        
        If Modul6.FindFile(gcDBPfad, "aWKL30p8_spez.rpt") Then
            reportbildschirm "spezial", "aWKL30p8_spez"
        Else
            reportbildschirm "", "aWKL30p8"
        End If
        
    ElseIf cboRegalDinA4.Text = "perforiert 22 x 22" Then
        cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.Artnr = Artikel.artnr set DRU_GRUN.libesnr = artikel.ean "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update DRU_GRUN d inner join Artikel a on d.Artnr = a.artnr set d.libesnr = a.ean2 "
        cSQL = cSQL & " where d.libesnr is null or d.libesnr = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update DRU_GRUN d inner join Artikel a on d.Artnr = a.artnr set d.libesnr = a.ean3 "
        cSQL = cSQL & " where d.libesnr is null or d.libesnr = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        SpalteAnfuegenNEW "DRU_GRUN", "BARCODE13", "Text(18)", gdBase
        
        cSQL = "Update DRU_GRUN set barcode13 =  '' "
        gdBase.Execute cSQL, dbFailOnError
            
            
        cSQL = "Select * from DRU_GRUN "
        Set rsZiel = gdBase.OpenRecordset(cSQL)
        
        If Not rsZiel.EOF Then
            rsZiel.MoveFirst
            Do While Not rsZiel.EOF
            
                 If Len(Trim(rsZiel!LIBESNR)) = 13 Then
                    cEANCode = ean13(Trim(rsZiel!LIBESNR))
                    
                    rsZiel.Edit
                    rsZiel!BARCODE13 = cEANCode
                    rsZiel.Update
                End If
                
                If Len(Trim(rsZiel!LIBESNR)) = 8 Then
                
                    cEANCode = fnCodiereEANCode(Trim(rsZiel!LIBESNR))
                    
                    rsZiel.Edit
                    rsZiel!BARCODE13 = cEANCode
                    rsZiel.Update
                    
                End If
        
                rsZiel.MoveNext
            Loop
        End If
        rsZiel.Close: Set rsZiel = Nothing
        
        cSQL = "Update DRU_GRUN set barcode13 = barcode, libesnr = 'Art# ' & artnr where barcode13 = '' and artnr > 0 "
        gdBase.Execute cSQL, dbFailOnError

        reportbildschirm "", "aWKL30p9"
        
    ElseIf cboRegalDinA4.Text = "perforiert 26 x 45 (EAN 13)" Then
        cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.Artnr = Artikel.artnr set DRU_GRUN.libesnr = artikel.ean "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update DRU_GRUN d inner join Artikel a on d.Artnr = a.artnr set d.libesnr = a.ean2 "
        cSQL = cSQL & " where d.libesnr is null or d.libesnr = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update DRU_GRUN d inner join Artikel a on d.Artnr = a.artnr set d.libesnr = a.ean3 "
        cSQL = cSQL & " where d.libesnr is null or d.libesnr = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        SpalteAnfuegenNEW "DRU_GRUN", "BARCODE13", "Text(18)", gdBase
        
        cSQL = "Update DRU_GRUN set barcode13 =  '' "
        gdBase.Execute cSQL, dbFailOnError
            
            
        cSQL = "Select * from DRU_GRUN "
        Set rsZiel = gdBase.OpenRecordset(cSQL)
        
        If Not rsZiel.EOF Then
            rsZiel.MoveFirst
            Do While Not rsZiel.EOF
            
                sEtikettKürzel = ermLiefKürzelmitkleinstenLEKPR(rsZiel!artnr, gdBase)
                
                rsZiel.Edit
                rsZiel!DRUCKDATUM = Left(sEtikettKürzel, 3)
                rsZiel.Update
            
                If Len(Trim(rsZiel!LIBESNR)) = 13 Then
                    cEANCode = ean13(Trim(rsZiel!LIBESNR))
                    
                    rsZiel.Edit
                    rsZiel!BARCODE13 = cEANCode
                    rsZiel.Update
                End If
                
                If Len(Trim(rsZiel!LIBESNR)) = 8 Then
                
                    cEANCode = fnCodiereEANCode(Trim(rsZiel!LIBESNR))
                    
                    rsZiel.Edit
                    rsZiel!BARCODE13 = cEANCode
                    rsZiel.Update
                End If
        
                rsZiel.MoveNext
            Loop
        End If
        rsZiel.Close: Set rsZiel = Nothing
        
        cSQL = "Update DRU_GRUN set barcode13 = barcode, libesnr = 'Art# ' & artnr where barcode13 = '' and artnr > 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        reportbildschirm "", "aWKL30p6"
        
    ElseIf cboRegalDinA4.Text = "perforiert 26 x 45 (EAN 13) V2" Then
        cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.Artnr = Artikel.artnr set DRU_GRUN.libesnr = artikel.ean "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update DRU_GRUN d inner join Artikel a on d.Artnr = a.artnr set d.libesnr = a.ean2 "
        cSQL = cSQL & " where d.libesnr is null or d.libesnr = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update DRU_GRUN d inner join Artikel a on d.Artnr = a.artnr set d.libesnr = a.ean3 "
        cSQL = cSQL & " where d.libesnr is null or d.libesnr = '' "
        gdBase.Execute cSQL, dbFailOnError
        
        SpalteAnfuegenNEW "DRU_GRUN", "BARCODE13", "Text(18)", gdBase
        
        cSQL = "Update DRU_GRUN set barcode13 =  '' "
        gdBase.Execute cSQL, dbFailOnError
            
            
        cSQL = "Select * from DRU_GRUN "
        Set rsZiel = gdBase.OpenRecordset(cSQL)
        
        If Not rsZiel.EOF Then
            rsZiel.MoveFirst
            Do While Not rsZiel.EOF
            
                rsZiel.Edit
                rsZiel!DRUCKDATUM = "" 'sEtikettKürzel
                rsZiel.Update
            
                If Len(Trim(rsZiel!LIBESNR)) = 13 Then
                    cEANCode = ean13(Trim(rsZiel!LIBESNR))
                    
                    rsZiel.Edit
                    rsZiel!BARCODE13 = cEANCode
                    rsZiel.Update
                End If
                
                If Len(Trim(rsZiel!LIBESNR)) = 8 Then
                
                    cEANCode = fnCodiereEANCode(Trim(rsZiel!LIBESNR))
                    
                    rsZiel.Edit
                    rsZiel!BARCODE13 = cEANCode
                    rsZiel.Update
                    
                End If
        
                rsZiel.MoveNext
            Loop
        End If
        rsZiel.Close: Set rsZiel = Nothing
        
        cSQL = "Update DRU_GRUN set barcode13 = barcode, libesnr = 'Art# ' & artnr where barcode13 = '' and artnr > 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        reportbildschirm "", "aWKL30p6"
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DruckeGrundPreisEtikettenWKL30"
        Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
        
        Fehlermeldung1

    End If
End Sub
Private Function ermvorz(cArtNr As String) As String
On Error GoTo LOKAL_ERROR

    Dim cKVKN As String
    Dim cek As String
    Dim cMWST As String
    Dim cnettosp As String
    Dim rsrs As Recordset
    Dim sSQL As String
    
    ermvorz = ""
    
    If cArtNr = "" Then
        Exit Function
    End If
    
    If IsNumeric(cArtNr) = False Then
        Exit Function
    End If
    
    sSQL = "Update artlief set lekpr =0 where lekpr is null and artnr = " & cArtNr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select artikel.kvkpr1,artlief.Lekpr,artikel.mwst from artikel "
    sSQL = sSQL & " inner join artlief on artikel.artnr = artlief.artnr "
    sSQL = sSQL & " and artikel.linr = artlief.linr"
    sSQL = sSQL & "  Where artikel.artnr = " & cArtNr
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!KVKPR1) Then
            cKVKN = rsrs!KVKPR1
        End If
        
        If Not IsNull(rsrs!lekpr) Then
            cek = rsrs!lekpr
        End If
        
        If Not IsNull(rsrs!MWST) Then
            cMWST = rsrs!MWST
        End If
        
        cnettosp = NettospanneInProzent(cKVKN, cek, cMWST)
        
        
        
        
        
       
        
'        MsgBox cnettosp
        If cnettosp <> "" Then
            If CDbl(cnettosp) > dePlus Then
                ermvorz = "+"
            ElseIf CDbl(cnettosp) > deNull And CDbl(cnettosp) < dePlus Then
                ermvorz = "~"
            ElseIf CDbl(cnettosp) < deMinus Then
                ermvorz = "-"
            End If
        End If
        
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermvorz"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
       Resume Next

End Function
Private Sub ermvorzNettospannen()
On Error GoTo LOKAL_ERROR

    
    Dim sSQL As String
    Dim rsVORZ As DAO.Recordset
        
    If NewTableSuchenDBKombi("ETINS", gdBase) = False Then
        CreateTableT2 "ETINS", gdBase
    End If
    
    sSQL = "Select * from ETINS "
    Set rsVORZ = gdBase.OpenRecordset(sSQL)
    
    If Not rsVORZ.EOF Then
    
        If Not IsNull(rsVORZ!ePLUS) Then
            dePlus = rsVORZ!ePLUS
        Else
            dePlus = 30
        End If
        
        If Not IsNull(rsVORZ!eNull) Then
            deNull = rsVORZ!eNull
        Else
            deNull = 10
        End If
        
        If Not IsNull(rsVORZ!eMinus) Then
            deMinus = rsVORZ!eMinus
        Else
            deMinus = 10
        End If
    End If
    rsVORZ.Close
        
        
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermvorzNettospannen"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub DruckeGrundPreisEtikettenLS(acArtNr() As String, lAnzahl As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cPfad As String
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    Dim lAnz As Long
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String
    
    Dim iFileNr As Integer

    loeschNEW "DRU_GRUN", gdBase
    
    cSQL = "Create Table DRU_GRUN ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", EAN1 Text(13)"
    cSQL = cSQL & ", NEPR Double"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "

    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")
    
    cSQL = "Select * from DRU_GRUN"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    lAnz = Val(Text2.Text)
    For lcount = 1 To lAnz
        rsZiel.AddNew
        rsZiel!FirmaName = ""
        rsZiel!artnr = Null
        rsZiel!BEZEICH = ""
        rsZiel!Barcode = ""
        rsZiel!LIBESNR = ""
        rsZiel!vkpr = Null
        rsZiel!vkpr_EUR = Null
        rsZiel!INHALT = Null
        rsZiel!INHALTBEZ = ""
        rsZiel!DRUCKDATUM = ""
        rsZiel!GRUNDPREIS = ""
        rsZiel!GRUND_INH = Null
        rsZiel!GRUND_DM = Null
        rsZiel!GRUND_EUR = Null
        rsZiel.Update
        
    Next lcount
    
    For lcount = 0 To lAnzahl
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If rsrs.EOF Then

        Else
            rsZiel.AddNew
            rsZiel!FirmaName = gFirma.FirmaName
            rsZiel!artnr = rsrs!artnr
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = fnEntferneLeerzeichen(cBezeich)
            rsZiel!BEZEICH = cBezeich

            cEAN = acArtNr(lcount)
            cEAN = fnMoveArtNr2EAN8(cEAN)
            
            rsZiel!EAN = cEAN
            rsZiel!EAN1 = rsrs!EAN 'cEAN

            cEANCode = fnCodiereEANCode(cEAN)
            rsZiel!Barcode = cEANCode

            rsZiel!LIBESNR = rsrs!LIBESNR

            rsZiel!vkpr = rsrs!KVKPR1
            If Not IsNull(rsrs!KVKPR1) Then
                dVkPr = rsrs!KVKPR1
            Else
                dVkPr = 0
            End If

            If dVkPr <> 0 Then
                dWert = dVkPr
            End If
            rsZiel!vkpr_EUR = dWert
            
            Dim cKVkPr1 As String
            Dim cNettoPr As String
'            If sart = "NETTO" Then
                cKVkPr1 = Format$(dWert, "#####0.00")
                cKVkPr1 = SwapStr(cKVkPr1, ",", ".")
                cKVkPr1 = Trim(cKVkPr1)
            
                cNettoPr = nettoR(dWert, rsrs!MWST)
                cNettoPr = Format$(cNettoPr, "#####0.00")
'                cNettoPr = SwapStr(cNettoPr, ",", ".")
                cNettoPr = Trim(cNettoPr)
                rsZiel!NEPR = cNettoPr
'            End If
            
            
            
            
            

            rsZiel!INHALT = rsrs!INHALT
            rsZiel!INHALTBEZ = rsrs!INHALTBEZ
            rsZiel!DRUCKDATUM = cDruckdatum
            rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
            If rsrs!GRUNDPREIS = "J" Then
                If Not IsNull(rsrs!INHALT) Then
                    dInhalt = rsrs!INHALT
                Else
                    dInhalt = 0
                End If
                If Not IsNull(rsrs!INHALTBEZ) Then
                    cInhaltBez = rsrs!INHALTBEZ
                Else
                    cInhaltBez = ""
                End If

                BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                If dGrundPreisDM > 0 Then
                    rsZiel!GRUND_INH = cGrundInhalt
                    rsZiel!GRUND_DM = dGrundPreisDM
                    rsZiel!GRUND_EUR = dGrundPreisEur
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
            Else
                rsZiel!GRUND_INH = Null
                rsZiel!GRUND_DM = Null
                rsZiel!GRUND_EUR = Null
            End If

            '//DRU_GRUN
            rsZiel.Update
            
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing


    
    '********************************************************
    '* Drucker-Formblatt-Konfiguration merken!
    '********************************************************
    dWert = 0
    If Option3(0).Value = True Then
        dWert = dWert + 0
    Else
        dWert = dWert + 1
    End If
    If cboRegalDinA4.Text = "selbstklebend 52,5 x 30" Then
        dWert = dWert + 2
    Else
        dWert = dWert + 4
    End If
    cBezeich = Trim$(Str$(dWert))
    
    iFileNr = FreeFile
    Open gcPfad & "ETIDRU.CFG" For Binary As #iFileNr
    Close iFileNr
    Kill gcPfad & "ETIDRU.CFG"
    iFileNr = FreeFile
    Open gcPfad & "ETIDRU.CFG" For Binary As #iFileNr
    Put #iFileNr, 1, cBezeich
    Close iFileNr
    
    '********************************************************
    '* Drucker-Formblatt-Konfiguration merken!
    '********************************************************
    If cboRegalDinA4.Text = "selbstklebend 52,5 x 30" Then

    Else


        If Modul6.FindFile(gcDBPfad, "aWKL30ls.rpt") Then

            reportbildschirm "spezial", "aWKL30ls"

        Else

        End If


    End If
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DruckeGrundPreisEtikettenLS"
        Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
End Sub

Private Sub DruckeGrundPreisEtiketten2WKL30(acArtNr() As String, lAnzahl As Long, acAnzEti() As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim lAnz As Long
    Dim cSQL As String
    Dim cSQL1 As String
    Dim rsrs As Recordset
    Dim rsZiel As Recordset
    Dim dWert As Double
    Dim cPfad As String
    Dim cEAN As String
    Dim cEANCode As String
    Dim cDruckdatum As String
    Dim cBezeich As String
    
    Dim dInhalt As Double
    Dim cInhaltBez As String
    Dim dVkPr As Double
    Dim dGrundPreisDM As Double
    Dim dGrundPreisEur As Double
    Dim cGrundInhalt As String
    
    Dim iFileNr As Integer
        
    loeschNEW "DRU_GRUN", gdBase
    
    cSQL = "Create Table DRU_GRUN ("
    cSQL = cSQL & "FIRMANAME Text(50)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", EAN Text(8)"
    cSQL = cSQL & ", BARCODE Text(11)"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", VKPR_EUR Double"
    cSQL = cSQL & ", INHALT Double"
    cSQL = cSQL & ", INHALTBEZ Text(3)"
    cSQL = cSQL & ", GRUNDPREIS Text(1)"
    cSQL = cSQL & ", GRUND_INH Text(7)"
    cSQL = cSQL & ", GRUND_DM Double"
    cSQL = cSQL & ", GRUND_EUR Double"
    cSQL = cSQL & ", DRUCKDATUM Text(5)"
    cSQL = cSQL & ") "

    gdBase.Execute cSQL, dbFailOnError
    
    cDruckdatum = "1" & Right(Format$(Year(Now), "00"), 2) & Format$(Month(Now), "00")
    
    cSQL = "Select * from DRU_GRUN"
    Set rsZiel = gdBase.OpenRecordset(cSQL)
    
    lAnz = Val(Text2.Text)
    For lcount = 1 To lAnz
        rsZiel.AddNew
        rsZiel!FirmaName = ""
        rsZiel!artnr = Null
        rsZiel!BEZEICH = "."
        rsZiel!Barcode = ""
        rsZiel!LIBESNR = ""
        rsZiel!vkpr = Null
        rsZiel!vkpr_EUR = Null
        rsZiel!INHALT = Null
        rsZiel!INHALTBEZ = ""
        rsZiel!DRUCKDATUM = ""
        rsZiel!GRUNDPREIS = ""
        rsZiel!GRUND_INH = Null
        rsZiel!GRUND_DM = Null
        rsZiel!GRUND_EUR = Null
        rsZiel.Update
        
    Next lcount
    

    For lcount = 0 To lAnzahl
    
        cSQL1 = "Select * from artikel where artnr = " & acArtNr(lcount)
        Set rsrs = gdBase.OpenRecordset(cSQL1)

        If rsrs.EOF Then
        Else
            For lAnz = 1 To Val(acAnzEti(lcount))
                rsZiel.AddNew
                rsZiel!FirmaName = gFirma.FirmaName
                rsZiel!artnr = rsrs!artnr
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                cBezeich = fnEntferneLeerzeichen(cBezeich)
                rsZiel!BEZEICH = cBezeich
                
                cEAN = acArtNr(lcount)
                cEAN = fnMoveArtNr2EAN8(cEAN)
'                rsZiel!EAN = cEAN
                
                'mal anders ab 02.01.14
                If Len(rsrs!EAN) > 8 Then
                    rsZiel!EAN = Right(rsrs!EAN, 8)
                Else
                    rsZiel!EAN = rsrs!EAN
                End If
                
                cEANCode = fnCodiereEANCode(cEAN)
                rsZiel!Barcode = cEANCode
                
                If gbEtiEan Then
                    rsZiel!LIBESNR = rsrs!EAN
                Else
                    rsZiel!LIBESNR = rsrs!LIBESNR
                End If
                
                rsZiel!vkpr = rsrs!KVKPR1
                If Not IsNull(rsrs!KVKPR1) Then
                    dVkPr = rsrs!KVKPR1
                Else
                    dVkPr = 0
                End If
                   
                rsZiel!vkpr_EUR = rsrs!vkpr
                
                rsZiel!INHALT = rsrs!INHALT
                rsZiel!INHALTBEZ = rsrs!INHALTBEZ
                rsZiel!DRUCKDATUM = cDruckdatum
                rsZiel!GRUNDPREIS = rsrs!GRUNDPREIS
                If rsrs!GRUNDPREIS = "J" Then
                    If Not IsNull(rsrs!INHALT) Then
                        dInhalt = rsrs!INHALT
                    Else
                        dInhalt = 0
                    End If
                    If Not IsNull(rsrs!INHALTBEZ) Then
                        cInhaltBez = rsrs!INHALTBEZ
                    Else
                        cInhaltBez = ""
                    End If
                    
                    BerechneGrundPreis dInhalt, cInhaltBez, dVkPr, cGrundInhalt, dGrundPreisDM, dGrundPreisEur
                    If dGrundPreisDM > 0 Then
                        rsZiel!GRUND_INH = cGrundInhalt
                        rsZiel!GRUND_DM = dGrundPreisDM
                        rsZiel!GRUND_EUR = dGrundPreisEur
                    Else
                        rsZiel!GRUND_INH = Null
                        rsZiel!GRUND_DM = Null
                        rsZiel!GRUND_EUR = Null
                    End If
                Else
                    rsZiel!GRUND_INH = Null
                    rsZiel!GRUND_DM = Null
                    rsZiel!GRUND_EUR = Null
                End If
                rsZiel.Update
                
            Next lAnz
        End If
        rsrs.Close: Set rsrs = Nothing
    Next lcount
    rsZiel.Close: Set rsZiel = Nothing
    
    '********************************************************
    '* Drucker-Formblatt-Konfiguration merken!
    '********************************************************
    dWert = 0
    If Option3(0).Value = True Then
        dWert = dWert + 0
    Else
        dWert = dWert + 1
    End If
    If cboRegalDinA4.Text = "selbstklebend 52,5 x 30" Then
        dWert = dWert + 2
    Else
        dWert = dWert + 4
    End If
    cBezeich = Trim$(Str$(dWert))
    
    iFileNr = FreeFile
    Open gcPfad & "ETIDRU.CFG" For Binary As #iFileNr
    Close iFileNr
    Kill gcPfad & "ETIDRU.CFG"
    iFileNr = FreeFile
    Open gcPfad & "ETIDRU.CFG" For Binary As #iFileNr
    Put #iFileNr, 1, cBezeich
    Close iFileNr
    
    '********************************************************
    '* Drucker-Formblatt-Konfiguration merken!
    '********************************************************
    If Option3(0).Value = True Then
        'Tintenstrahler
        
        If cboStrichDINA4.Text = "50 x 36" Then
            reportbildschirm "WKL021", "aWKL30c"
        ElseIf cboStrichDINA4.Text = "45,7 x 21,2" Then

            If Modul6.FindFile(gcDBPfad, "aWKL30s.rpt") Then
                reportbildschirm "spezial", "aWKL30s"
            Else
                reportbildschirm "WKL021f", "aWKL30f"
            End If
            
        ElseIf cboStrichDINA4.Text = "25,4 x 12,8" Then
'            If Modul6.FindFile(gcDBPfad, "aWKL30u.rpt") Then
'                reportbildschirm "spezial", "aWKL30u"
'            Else
                reportbildschirm "WKL021e", "aWKL30ej"
'            End If

        ElseIf cboStrichDINA4.Text = "35,6 x 16,9" Then
            If Modul6.FindFile(gcDBPfad, "aWKL30u.rpt") Then
                reportbildschirm "spezial", "aWKL30u"
            Else
                reportbildschirm "WKL021e", "aWKL30e"
            End If
            
        ElseIf cboStrichDINA4.Text = "35,6 x 16,9 Variante 2" Then
            If Modul6.FindFile(gcDBPfad, "aWKL30uV2.rpt") Then
                reportbildschirm "spezial", "aWKL30uV2"
            Else
                reportbildschirm "WKL021e", "aWKL30eV2"
            End If
            
        ElseIf cboStrichDINA4.Text = "35,6 x 16,9 Variante 4" Then
            If Modul6.FindFile(gcDBPfad, "aWKL30uV4.rpt") Then
                reportbildschirm "spezial", "aWKL30uV4"
            Else
                reportbildschirm "WKL021e", "aWKL30eV4"
            End If
            
        ElseIf cboStrichDINA4.Text = "35,6 x 16,9 Variante 5" Then
        
        
            cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.ARTNR = Artikel.ARTNR "
            cSQL = cSQL & " set DRU_GRUN.vkpr = Artikel.VKPR "
            cSQL = cSQL & "  where artikel.awm = '3' "
            gdBase.Execute cSQL, dbFailOnError
        
            If Modul6.FindFile(gcDBPfad, "aWKL30uV5.rpt") Then
                reportbildschirm "spezial", "aWKL30uV5"
            Else
                reportbildschirm "WKL021e", "aWKL30eV5"
            End If
        
        ElseIf cboStrichDINA4.Text = "48,5 x 25,4" Then '48,5 x 25,4
            If Modul6.FindFile(gcDBPfad, "aWKL30m.rpt") Then
                reportbildschirm "spezial", "aWKL30m"
            Else
                reportbildschirm "WKL021e", "aWKL30k"
            End If
            
        ElseIf cboStrichDINA4.Text = "45,7 x 21,2 Variante 2" Then

            reportbildschirm "", "aWKL30o"
            
        ElseIf cboStrichDINA4.Text = "52,5 x 30" Then

            reportbildschirm "", "aWKL30r"
        ElseIf cboStrichDINA4.Text = "52,5 x 21,2" Then

            If Modul6.FindFile(gcDBPfad, "aWKL30ps.rpt") Then
                reportbildschirm "spezial", "aWKL30ps"
            Else
                reportbildschirm "", "aWKL30p"
            End If
            
        ElseIf cboStrichDINA4.Text = "50 x 21,09" Then
        
        
            'direkt aus Artikel/Artlief
            
            If Not SpalteInTabellegefundenNEW("DRU_GRUN", "LINR", gdBase) Then
                SpalteAnfuegenNEW "DRU_GRUN", "LINR", "Long", gdBase
            End If
                
            If Not SpalteInTabellegefundenNEW("DRU_GRUN", "EANDRUCK", gdBase) Then
                SpalteAnfuegenNEW "DRU_GRUN", "EANDRUCK", "Text(13)", gdBase
            End If
            
            cSQL = "Update DRU_GRUN inner join Artlief on DRU_GRUN.ARTNR = Artlief.ARTNR "
            cSQL = cSQL & " set DRU_GRUN.LINR = Artlief.LINR "
            cSQL = cSQL & "  "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.ARTNR = Artikel.ARTNR "
            cSQL = cSQL & " set DRU_GRUN.EANDRUCK = Artikel.EAN "
            cSQL = cSQL & "  "
            gdBase.Execute cSQL, dbFailOnError
        
        
        
        
        
            reportbildschirm "", "aWKL30p1"
            
            
        ElseIf cboStrichDINA4.Text = "52,5 x 21,2 Variante 2" Then
        
        
            'direkt aus Artikel/Artlief
            
            If Not SpalteInTabellegefundenNEW("DRU_GRUN", "LINR", gdBase) Then
                SpalteAnfuegenNEW "DRU_GRUN", "LINR", "Long", gdBase
            End If
                
            If Not SpalteInTabellegefundenNEW("DRU_GRUN", "EANDRUCK", gdBase) Then
                SpalteAnfuegenNEW "DRU_GRUN", "EANDRUCK", "Text(13)", gdBase
            End If
            
            cSQL = "Update DRU_GRUN inner join Artlief on DRU_GRUN.ARTNR = Artlief.ARTNR "
            cSQL = cSQL & " set DRU_GRUN.LINR = Artlief.LINR "
            cSQL = cSQL & "  "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.ARTNR = Artikel.ARTNR "
            cSQL = cSQL & " set DRU_GRUN.EANDRUCK = Artikel.EAN "
            cSQL = cSQL & "  "
            gdBase.Execute cSQL, dbFailOnError
        
        
        
        
        
            reportbildschirm "", "aWKL30pm"
            
            
            
            
            
        End If
    Else
        'Laserdrucker
        If cboStrichDINA4.Text = "50 x 36" Then
        
            If Modul6.FindFile(gcDBPfad, "aWKL30x.rpt") Then
                reportbildschirm "spezial", "aWKL30x"
            Else
                reportbildschirm "WKL021f", "aWKL30d"
            End If
            
        ElseIf cboStrichDINA4.Text = "45,7 x 21,2" Then
            If Modul6.FindFile(gcDBPfad, "aWKL30s.rpt") Then
                reportbildschirm "spezial", "aWKL30s"
            Else
                reportbildschirm "WKL021f", "aWKL30h"
            End If

        ElseIf cboStrichDINA4.Text = "35,6 x 16,9" Then
            If Modul6.FindFile(gcDBPfad, "aWKL30v.rpt") Then
                reportbildschirm "spezial", "aWKL30v"
            Else
                reportbildschirm "WKL021e", "aWKL30g"
            End If
            
        ElseIf cboStrichDINA4.Text = "35,6 x 16,9 Variante 2" Then
            If Modul6.FindFile(gcDBPfad, "aWKL30vV2.rpt") Then
                reportbildschirm "spezial", "aWKL30vV2"
            Else
                reportbildschirm "WKL021e", "aWKL30gV2"
            End If
            
        ElseIf cboStrichDINA4.Text = "35,6 x 16,9 Variante 4" Then
            If Modul6.FindFile(gcDBPfad, "aWKL30vV4.rpt") Then
                reportbildschirm "spezial", "aWKL30vV4"
            Else
                reportbildschirm "WKL021e", "aWKL30gV4"
            End If
            
        ElseIf cboStrichDINA4.Text = "35,6 x 16,9 Variante 5" Then
        
        
        
            cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.ARTNR = Artikel.ARTNR "
            cSQL = cSQL & " set DRU_GRUN.vkpr = Artikel.VKPR "
            cSQL = cSQL & "  where artikel.awm = '3' "
            gdBase.Execute cSQL, dbFailOnError
        
            If Modul6.FindFile(gcDBPfad, "aWKL30vV5.rpt") Then
                reportbildschirm "spezial", "aWKL30vV5"
            Else
                reportbildschirm "WKL021e", "aWKL30gV5"
            End If
        
        ElseIf cboStrichDINA4.Text = "Sonder Etikett" Then 'Sonder Etikett


            'direkt aus Lieferschein
            
            If Not SpalteInTabellegefundenNEW("DRU_GRUN", "LINR", gdBase) Then
                SpalteAnfuegenNEW "DRU_GRUN", "LINR", "Long", gdBase
            End If
                
            If Not SpalteInTabellegefundenNEW("DRU_GRUN", "WEDate", gdBase) Then
                SpalteAnfuegenNEW "DRU_GRUN", "WEDate", "DATETIME", gdBase
            End If
            
            If Not SpalteInTabellegefundenNEW("DRU_GRUN", "LS", gdBase) Then
                SpalteAnfuegenNEW "DRU_GRUN", "LS", "Text(20)", gdBase
            End If
            
            cSQL = "Update DRU_GRUN inner join LSTEETI on DRU_GRUN.ARTNR = LSTEETI.ARTNR "
            cSQL = cSQL & " set DRU_GRUN.LINR = LSTEETI.LINR "
            cSQL = cSQL & " , DRU_GRUN.LIBESNR = LSTEETI.LIBESNR "
            cSQL = cSQL & " , DRU_GRUN.WEDate = LSTEETI.WEDate "
            cSQL = cSQL & " , DRU_GRUN.LS = LSTEETI.LS "
            gdBase.Execute cSQL, dbFailOnError
                 
            reportbildschirm "WKL021e", "aWKL30l"
            
        ElseIf cboStrichDINA4.Text = "48,5 x 25,4" Then '48,5 x 25,4
            If Modul6.FindFile(gcDBPfad, "aWKL30m.rpt") Then
                reportbildschirm "spezial", "aWKL30m"
            Else
                reportbildschirm "WKL021e", "aWKL30k"
            End If
            
        ElseIf cboStrichDINA4.Text = "45,7 x 21,2 Variante 2" Then

            reportbildschirm "", "aWKL30o"
            
        ElseIf cboStrichDINA4.Text = "52,5 x 30" Then

            reportbildschirm "", "aWKL30r"
        ElseIf cboStrichDINA4.Text = "52,5 x 21,2" Then
            
            If Modul6.FindFile(gcDBPfad, "aWKL30ps.rpt") Then
                reportbildschirm "spezial", "aWKL30ps"
            Else
                reportbildschirm "", "aWKL30p"
            End If
        
        ElseIf cboStrichDINA4.Text = "50 x 21,09" Then
        
            'direkt aus Artikel/Artlief
            
            If Not SpalteInTabellegefundenNEW("DRU_GRUN", "LINR", gdBase) Then
                SpalteAnfuegenNEW "DRU_GRUN", "LINR", "Long", gdBase
            End If
                
            If Not SpalteInTabellegefundenNEW("DRU_GRUN", "EANDRUCK", gdBase) Then
                SpalteAnfuegenNEW "DRU_GRUN", "EANDRUCK", "Text(13)", gdBase
            End If
            
            cSQL = "Update DRU_GRUN inner join Artlief on DRU_GRUN.ARTNR = Artlief.ARTNR "
            cSQL = cSQL & " set DRU_GRUN.LINR = Artlief.LINR "
            cSQL = cSQL & "  "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.ARTNR = Artikel.ARTNR "
            cSQL = cSQL & " set DRU_GRUN.EANDRUCK = Artikel.EAN "
            cSQL = cSQL & "  "
            gdBase.Execute cSQL, dbFailOnError
        
            reportbildschirm "", "aWKL30p1"
            
        ElseIf cboStrichDINA4.Text = "52,5 x 21,2 Variante 2" Then
        
        
            'direkt aus Artikel/Artlief
            
            If Not SpalteInTabellegefundenNEW("DRU_GRUN", "LINR", gdBase) Then
                SpalteAnfuegenNEW "DRU_GRUN", "LINR", "Long", gdBase
            End If
                
            If Not SpalteInTabellegefundenNEW("DRU_GRUN", "EANDRUCK", gdBase) Then
                SpalteAnfuegenNEW "DRU_GRUN", "EANDRUCK", "Text(13)", gdBase
            End If
            
            cSQL = "Update DRU_GRUN inner join Artlief on DRU_GRUN.ARTNR = Artlief.ARTNR "
            cSQL = cSQL & " set DRU_GRUN.LINR = Artlief.LINR "
            cSQL = cSQL & "  "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update DRU_GRUN inner join Artikel on DRU_GRUN.ARTNR = Artikel.ARTNR "
            cSQL = cSQL & " set DRU_GRUN.EANDRUCK = Artikel.EAN "
            cSQL = cSQL & "  "
            gdBase.Execute cSQL, dbFailOnError
        
        
        
        
        
            reportbildschirm "", "aWKL30pm"
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DruckeGrundPreisEtiketten2WKL30"
        Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub ExportiereMarkierteSaetzeWKL30()
    On Error GoTo LOKAL_ERROR

    Dim iFileNr As Integer
    Dim cLBSatz As String
    Dim ctmp As String
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim lartnr As Long
    Dim lFilNr As Long
    Dim lAnz As Long
    Dim lcount As Long
    Dim dWert As Double
    Dim rsrs As Recordset
    Dim rsExpo As Recordset
    Dim rsArtikel As Recordset
    Dim cSQL As String
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If

    loeschNEW "ETI_EXPO", gdBase

    cSQL = "Create Table ETI_EXPO "
    cSQL = cSQL & "("
    cSQL = cSQL & "ARTNR Long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", VKPR Text(10)"
    cSQL = cSQL & ", VKPR_EURO Text(10)"
    cSQL = cSQL & ", BESTAND Long"
    cSQL = cSQL & ", ANZAHL Long"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", EAN Text(13)"
    cSQL = cSQL & ", LINR Long"
    cSQL = cSQL & ", LPZ Long"
    cSQL = cSQL & ", FILNR Long"
    cSQL = cSQL & ")"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Select * from ETI_EXPO"
    Set rsExpo = gdBase.OpenRecordset(cSQL)
    
    lAnzSatz = List2.ListCount

    For lAktSatz = 0 To lAnzSatz - 1
        If List2.Selected(lAktSatz) = True Then
            cLBSatz = List2.list(lAktSatz)
            lartnr = Val(Left(cLBSatz, 6))
            lFilNr = Val(Right(cLBSatz, 2))
            
            If gsETILS <> "" Then
                cSQL = "Select * from LSTEETI where ARTNR = " & Trim$(Str$(lartnr)) & " and FILNR = " & Trim$(Str$(lFilNr)) & " "
            Else
                cSQL = "Select * from ETIDRU where ARTNR = " & Trim$(Str$(lartnr)) & " and FILNR = " & Trim$(Str$(lFilNr)) & " "
            End If
            
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!ANZAHL) Then
                    lAnz = rsrs!ANZAHL
                Else
                    lAnz = 0
                End If
                For lcount = 1 To lAnz
                    rsExpo.AddNew
                    rsExpo!artnr = rsrs!artnr
                    rsExpo!BEZEICH = rsrs!BEZEICH
                    
                    cSQL = "Select * from Artikel where artnr = " & rsrs!artnr
                    Set rsArtikel = gdBase.OpenRecordset(cSQL)
                    
                    If Not rsArtikel.EOF Then
                    
                        If Not IsNull(rsArtikel!vkpr) Then
                            rsExpo!vkpr = Format$(rsArtikel!vkpr, "######0.00") 'rsArtikel!VKPR
                        Else
                            rsExpo!vkpr = "0,00"
                        End If
                        
                        If Not IsNull(rsArtikel!KVKPR1) Then
                            rsExpo!VKPR_EURO = Format$(rsArtikel!KVKPR1, "######0.00") 'rsArtikel!VKPR
                        Else
                            rsExpo!VKPR_EURO = "0,00"
                        End If
                        
                        rsExpo!BESTAND = rsrs!BESTAND
                        rsExpo!ANZAHL = rsrs!ANZAHL
                        rsExpo!LIBESNR = rsrs!LIBESNR
                        rsExpo!EAN = rsrs!EAN
                        rsExpo!linr = rsrs!linr
                        rsExpo!LPZ = rsrs!LPZ
                        rsExpo!filnr = rsrs!filnr
                        rsExpo.Update
                    End If
                    rsArtikel.Close
                Next lcount
            End If
            rsrs.Close: Set rsrs = Nothing
        End If
    Next lAktSatz
    rsExpo.Close
    
    Kill cPfad & "ETI_EXPO.dbf"
    cSQL = "Select * into ETI_EXPO IN '" & cPfad & "' 'dbase IV;' from ETI_EXPO "
    gdBase.Execute cSQL, dbFailOnError

    MsgBox "Etikettendaten nach " & gcDBPfad & "\ETI_EXPO.DBF exportiert!", vbInformation, "Winkiss Hinweis:"

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    ElseIf err.Number = 70 Then
        MsgBox "Die Etikettendatei(ETI_EXPO.DBF) wird noch verwendet oder ist geöffnet. Bitte schließen Sie die Datei oder das Programm, dass diese Datei im Zugriff hat.", vbInformation, "Winkiss Hinweis:"

    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExportiereMarkierteSaetzeWKL30"
        Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Sub FuelleListeEtikettenWKL30()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cLBSatz As String
    Dim cFeld As String
    Dim dWert As Double
    
    List1.Clear
    List2.Clear
    List2.Visible = False
    
    List1.AddItem "ArtNr. Artikelbezeichnung                    KVK-Preis Anz.Etiketten Fil"
    
    If gsETILS <> "" Then
        cSQL = "Delete from LSTEETI where ANZAHL <= 0 and ls = '" & gsETILS & "'"
        gdBase.Execute cSQL, dbFailOnError
        cSQL = "Delete from LSTEETI where ANZAHL = null and ls = '" & gsETILS & "'"
        gdBase.Execute cSQL, dbFailOnError
        
        'Rabatt_OK
        
        If SpalteInTabellegefundenNEW("LSTEETI", "Rabatt_OK", gdBase) = False Then
            SpalteAnfuegenNEW "LSTEETI", "Rabatt_OK", "Text(1)", gdBase
            
            cSQL = "Update LSTEETI inner join Artikel on LSTEETI.artnr = Artikel.artnr set LSTEETI.rabatt_ok = artikel.rabatt_ok"
            gdBase.Execute cSQL, dbFailOnError
        End If
        
        If SpalteInTabellegefundenNEW("LSTEETI", "RKZ", gdBase) = False Then
            SpalteAnfuegenNEW "LSTEETI", "RKZ", "Text(1)", gdBase
            
            cSQL = "Update LSTEETI set RKZ = 'J'"
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update LSTEETI inner join Artlief on LSTEETI.artnr = Artlief.artnr "
            cSQL = cSQL & " set LSTEETI.rkz = 'N'"
            cSQL = cSQL & " where Artlief.rkz = 'N'"
            gdBase.Execute cSQL, dbFailOnError
        End If
            
        cSQL = "Select * "
        cSQL = cSQL & " from LSTEETI "
    Else

        cSQL = "Delete from ETIDRU where ANZAHL = null"
        gdBase.Execute cSQL, dbFailOnError
        
        'Rabatt_OK
        
        If SpalteInTabellegefundenNEW("ETIDRU", "Rabatt_OK", gdBase) = False Then
            SpalteAnfuegenNEW "ETIDRU", "Rabatt_OK", "Text(1)", gdBase
            
            cSQL = "Update ETIDRU inner join Artikel on Etidru.artnr = Artikel.artnr set etidru.rabatt_ok = artikel.rabatt_ok"
            gdBase.Execute cSQL, dbFailOnError
        End If
        
        If SpalteInTabellegefundenNEW("ETIDRU", "RKZ", gdBase) = False Then
            SpalteAnfuegenNEW "ETIDRU", "RKZ", "Text(1)", gdBase
            
            cSQL = "Update ETIDRU set RKZ = 'J'"
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update ETIDRU inner join Artlief on ETIDRU.artnr = Artlief.artnr "
            cSQL = cSQL & " set ETIDRU.rkz = 'N'"
            cSQL = cSQL & " where Artlief.rkz = 'N'"
            gdBase.Execute cSQL, dbFailOnError
        End If
        
        cSQL = "Select * from ETIDRU "
        If Check13.Value = vbChecked Then
            cSQL = cSQL & " where PCNAME = '" & srechnertab & "' "
        End If
    End If
    
    Select Case giSortierung
        Case Is = 0
            cSQL = cSQL & "order by FILNR, LINR, LPZ, BEZEICH"
        Case Is = 1
            cSQL = cSQL & "order by FILNR, LINR, LIBESNR, BEZEICH"
        Case Is = 2
            cSQL = cSQL & "order by FILNR, LINR, BEZEICH"
        Case Is = 3
            cSQL = cSQL & "order by FILNR, BEZEICH"
        Case Is = 4
            
            If gsETILS <> "" Then
            
            Else
                cSQL = cSQL & "order by lfnr" 'Originalreihenfolge
            End If
            
    End Select
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space(6 - Len(cFeld)) & cFeld
            cLBSatz = cFeld & " "
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!vkpr) Then
                dWert = rsrs!vkpr
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "###,##0.00")
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!ANZAHL) Then
                dWert = rsrs!ANZAHL
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "###,##0")
            cFeld = Space$(14 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!filnr) Then
                dWert = rsrs!filnr
            Else
                dWert = gcFilNr
            End If
            cFeld = Format$(dWert, "0")
            cFeld = Space$(2 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!RABATT_OK) Then
                cFeld = rsrs!RABATT_OK
            Else
                cFeld = "J"
            End If
            
            If cFeld = "N" Then
                cFeld = "*"
            Else
                cFeld = " "
            End If
        
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!RKZ) Then
                cFeld = rsrs!RKZ
            Else
                cFeld = "N"
            End If
            
            If cFeld = "J" Then
                cFeld = "EX"
            Else
                cFeld = "  "
            End If
        
            cLBSatz = cLBSatz & cFeld & " "
            
            List2.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    List2.Visible = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleListeEtikettenWKL30"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseKonfigurationEtikettWKL30()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cKommentar As String
    Dim cBefehl As String
    Dim llen As Long
    Dim lAnzKomma As Long
    Dim cZeichen As String
    Dim cZiel As String
    
    For lcount = 0 To 40
        Text1(lcount).Text = ""
    Next lcount
    
    cSQL = "Select * from SETELTRO"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!KOMMENTAR) Then
                cKommentar = rsrs!KOMMENTAR
            Else
                cKommentar = ""
            End If
            
            If Not IsNull(rsrs!BEFEHL) Then
                cBefehl = rsrs!BEFEHL
            Else
                cBefehl = ""
            End If
            
            Select Case cKommentar
                Case Is = "Druckerschnittstelle"
                    If cBefehl = "LPT1" Then
                        Option1(0).Value = True
                    ElseIf cBefehl = "LPT2" Then
                        Option1(1).Value = True
                    Else
                        Option1(2).Value = True
                    End If
                    
                Case Is = "Name"
                    Text1(0).Text = Trim$(cBefehl)
                    
                Case Is = "Ende der Textzeile"
                    'nix tun
                    
                Case Is = "Ende der Numerischen Zeile"
                    'nix tun
                    
                Case Is = "Position Firmenname"
                    llen = Len(cBefehl)
                    lAnzKomma = 0
                    For lcount = 2 To llen
                        cZeichen = Mid(cBefehl, lcount, 1)
                        If cZeichen <> "," Then
                            cZiel = cZiel & cZeichen
                        Else
                            lAnzKomma = lAnzKomma + 1
                            Select Case lAnzKomma
                                Case Is = 1
                                    Text1(1).Text = cZiel
                                Case Is = 2
                                    Text1(2).Text = cZiel
                                Case Is = 3
                                    Select Case cZiel
                                        Case Is = "0"
                                            Option2(0).Value = True
                                        Case Is = "1"
                                            Option2(1).Value = True
                                        Case Is = "2"
                                            Option2(2).Value = True
                                        Case Is = "3"
                                            Option2(3).Value = True
                                    End Select
                                Case Is = 4
                                    Select Case cZiel
                                        Case Is = "1"
                                            Text1(3).Text = "6"
                                        Case Is = "2"
                                            Text1(3).Text = "7"
                                        Case Is = "3"
                                            Text1(3).Text = "10"
                                        Case Is = "4"
                                            Text1(3).Text = "12"
                                        Case Is = "5"
                                            Text1(3).Text = "24"
                                    End Select
                                Case Is = 5
                                    Text1(4).Text = cZiel
                                Case Is = 6
                                    Text1(5).Text = cZiel
                                    
                            End Select
                            cZiel = ""
                        End If
                    Next lcount
                Case Is = "Position der Bestellnummer"
                    llen = Len(cBefehl)
                    lAnzKomma = 0
                    For lcount = 2 To llen
                        cZeichen = Mid(cBefehl, lcount, 1)
                        If cZeichen <> "," Then
                            cZiel = cZiel & cZeichen
                        Else
                            lAnzKomma = lAnzKomma + 1
                            Select Case lAnzKomma
                                Case Is = 1
                                    Text1(6).Text = cZiel
                                Case Is = 2
                                    Text1(7).Text = cZiel
                                Case Is = 3
                                    Select Case cZiel
                                        Case Is = "0"
                                            Option2(4).Value = True
                                        Case Is = "1"
                                            Option2(5).Value = True
                                        Case Is = "2"
                                            Option2(6).Value = True
                                        Case Is = "3"
                                            Option2(7).Value = True
                                    End Select
                                Case Is = 4
                                    Select Case cZiel
                                        Case Is = "1"
                                            Text1(8).Text = "6"
                                        Case Is = "2"
                                            Text1(8).Text = "7"
                                        Case Is = "3"
                                            Text1(8).Text = "10"
                                        Case Is = "4"
                                            Text1(8).Text = "12"
                                        Case Is = "5"
                                            Text1(8).Text = "24"
                                    End Select
                                Case Is = 5
                                    Text1(9).Text = cZiel
                                Case Is = 6
                                    Text1(10).Text = cZiel
                                    
                            End Select
                            cZiel = ""
                        End If
                    Next lcount
                    
                Case Is = "Position Linie"
                    llen = Len(cBefehl)
                    lAnzKomma = 0
                    For lcount = 3 To llen
                        cZeichen = Mid(cBefehl, lcount, 1)
                        If cZeichen <> "," Then
                            cZiel = cZiel & cZeichen
                        Else
                            lAnzKomma = lAnzKomma + 1
                            Select Case lAnzKomma
                                Case Is = 1
                                    Text1(11).Text = cZiel
                                Case Is = 2
                                    Text1(12).Text = cZiel
                                Case Is = 3
                                    Text1(13).Text = cZiel
                                Case Is = 4
                                    'es gibt hier kein 4.Komma
                            End Select
                            cZiel = ""
                        End If
                    Next lcount
                    Text1(14).Text = cZiel
                    cZiel = ""
                    
                Case Is = "Position Artikelbezeichnung"
                    llen = Len(cBefehl)
                    lAnzKomma = 0
                    For lcount = 2 To llen
                        cZeichen = Mid(cBefehl, lcount, 1)
                        If cZeichen <> "," Then
                            cZiel = cZiel & cZeichen
                        Else
                            lAnzKomma = lAnzKomma + 1
                            Select Case lAnzKomma
                                Case Is = 1
                                    Text1(15).Text = cZiel
                                Case Is = 2
                                    Text1(16).Text = cZiel
                                Case Is = 3
                                    Select Case cZiel
                                        Case Is = "0"
                                            Option2(8).Value = True
                                        Case Is = "1"
                                            Option2(9).Value = True
                                        Case Is = "2"
                                            Option2(10).Value = True
                                        Case Is = "3"
                                            Option2(11).Value = True
                                    End Select
                                Case Is = 4
                                    Select Case cZiel
                                        Case Is = "1"
                                            Text1(17).Text = "6"
                                        Case Is = "2"
                                            Text1(17).Text = "7"
                                        Case Is = "3"
                                            Text1(17).Text = "10"
                                        Case Is = "4"
                                            Text1(17).Text = "12"
                                        Case Is = "5"
                                            Text1(17).Text = "24"
                                    End Select
                                Case Is = 5
                                    Text1(18).Text = cZiel
                                Case Is = 6
                                    Text1(19).Text = cZiel
                                    
                            End Select
                            cZiel = ""
                        End If
                    Next lcount
                    
                Case Is = "Position Barcode"
                    llen = Len(cBefehl)
                    lAnzKomma = 0
                    For lcount = 2 To llen
                        cZeichen = Mid(cBefehl, lcount, 1)
                        If cZeichen <> "," Then
                            cZiel = cZiel & cZeichen
                        Else
                            lAnzKomma = lAnzKomma + 1
                            Select Case lAnzKomma
                                Case Is = 1
                                    Text1(20).Text = cZiel
                                Case Is = 2
                                    Text1(21).Text = cZiel
                                Case Is = 3
                                    Select Case cZiel
                                        Case Is = "0"
                                            Option2(12).Value = True
                                        Case Is = "1"
                                            Option2(13).Value = True
                                        Case Is = "2"
                                            Option2(14).Value = True
                                        Case Is = "3"
                                            Option2(15).Value = True
                                    End Select
                                Case Is = 4
                                    Select Case cZiel
                                        Case Is = "E80"
                                            Text1(22).Text = "8"
                                        Case Is = "E30"
                                            Text1(22).Text = "13"
                                    End Select
                                Case Is = 5
                                    Text1(23).Text = cZiel
                                Case Is = 6
                                    Text1(24).Text = cZiel
                                Case Is = 7
                                    Text1(40).Text = cZiel
                                Case Is = 8
                                    If cZiel = "B" Then
                                        Check1.Value = vbChecked
                                    Else
                                        Check1.Value = vbUnchecked
                                    End If
                            End Select
                            cZiel = ""
                        End If
                    Next lcount
                    
                Case Is = "Position Datum"
                    llen = Len(cBefehl)
                    lAnzKomma = 0
                    For lcount = 2 To llen
                        cZeichen = Mid(cBefehl, lcount, 1)
                        If cZeichen <> "," Then
                            cZiel = cZiel & cZeichen
                        Else
                            lAnzKomma = lAnzKomma + 1
                            Select Case lAnzKomma
                                Case Is = 1
                                    Text1(25).Text = cZiel
                                Case Is = 2
                                    Text1(26).Text = cZiel
                                Case Is = 3
                                    Select Case cZiel
                                        Case Is = "0"
                                            Option2(16).Value = True
                                        Case Is = "1"
                                            Option2(17).Value = True
                                        Case Is = "2"
                                            Option2(18).Value = True
                                        Case Is = "3"
                                            Option2(19).Value = True
                                    End Select
                                Case Is = 4
                                    Select Case cZiel
                                        Case Is = "1"
                                            Text1(27).Text = "6"
                                        Case Is = "2"
                                            Text1(27).Text = "7"
                                        Case Is = "3"
                                            Text1(27).Text = "10"
                                        Case Is = "4"
                                            Text1(27).Text = "12"
                                        Case Is = "5"
                                            Text1(27).Text = "24"
                                    End Select
                                Case Is = 5
                                    Text1(28).Text = cZiel
                                Case Is = 6
                                    Text1(29).Text = cZiel
                                    
                            End Select
                            cZiel = ""
                        End If
                    Next lcount
                    
                Case Is = "Position DM"
                    Text1(30).Text = Right(cBefehl, 1)
                    

                    
                Case Is = "Position Euro"
                    llen = Len(cBefehl)
                    lAnzKomma = 0
                    For lcount = 2 To llen
                        cZeichen = Mid(cBefehl, lcount, 1)
                        If cZeichen <> "," Then
                            cZiel = cZiel & cZeichen
                        Else
                            lAnzKomma = lAnzKomma + 1
                            Select Case lAnzKomma
                                Case Is = 1
                                    Text1(35).Text = cZiel
                                Case Is = 2
                                    Text1(36).Text = cZiel
                                Case Is = 3
                                    Select Case cZiel
                                        Case Is = "0"
                                            Option2(24).Value = True
                                        Case Is = "1"
                                            Option2(25).Value = True
                                        Case Is = "2"
                                            Option2(26).Value = True
                                        Case Is = "3"
                                            Option2(27).Value = True
                                    End Select
                                Case Is = 4
                                    Select Case cZiel
                                        Case Is = "1"
                                            Text1(37).Text = "6"
                                        Case Is = "2"
                                            Text1(37).Text = "7"
                                        Case Is = "3"
                                            Text1(37).Text = "10"
                                        Case Is = "4"
                                            Text1(37).Text = "12"
                                        Case Is = "5"
                                            Text1(37).Text = "24"
                                    End Select
                                Case Is = 5
                                    Text1(38).Text = cZiel
                                Case Is = 6
                                    Text1(39).Text = cZiel
                            End Select
                            cZiel = ""
                        End If
                    Next lcount
                    
                Case Is = "Name drucken (J/N)"
                    cBefehl = cBefehl
            End Select
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseKonfigurationEtikettWKL30"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeBarCodeEtikett3(lDelete As Long)
    On Error GoTo LOKAL_ERROR
    
    '*************************************************
    '* Diese Funktion dient zur Ansteuerung eines
    '* METO-Etiketten-Druckers, der am COM2-Port hängt
    '*************************************************
    
    Dim lStart As Long
    Dim lAktuell As Long
    
    Dim Dummy As Variant
    Dim cReturn As String
    Dim bOpen As Boolean
    Dim cSTX As String
    Dim cETX As String
    Dim cESC As String
    Dim cNEG As String
    Dim cPos As String
    Dim cXOFF As String
    Dim cGETSTAT As String
    Dim cSender As String
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    '************************
    '* Anfang des Datenteils
    '************************
    
    Dim cBezeich As String
    Dim cZeile1 As String
    Dim cZeile2 As String
    Dim cFilnr As String
    Dim lartnr As Long
    Dim cArtNr As String
    Dim cEAN As String
    Dim lMenge As Long
    Dim cMenge As String
    Dim dKVkPr1 As Double
    Dim dKVkPr1EUR As Double
    Dim cPreisDEM As String
    Dim cPreisEUR As String
    Dim cDatum As String
    
    '************************
    '* Ende des Datenteils
    '************************
    
    '***************************
    '* Standardwerte festlegen
    '***************************
    
    cSTX = Chr$(2)                      'Start Text
    cETX = Chr$(3)                      'Ende Text
    cESC = Chr$(27)                     'Escape
    cNEG = Chr$(254)                    'negative Antwort
    cPos = Chr$(255)                    'positive Antwort
    cXOFF = Chr$(19)                    'X-Off-Puffer voll
    
    MSComm1.CommPort = 2                'ComPort festlegen
    
    If MSComm1.PortOpen = False Then
        bOpen = False
        MSComm1.InputLen = 0                'wenn, dann gesamten Input-Buffer auslesen
        MSComm1.Settings = "1200,E,8,2"     'Übertragungsparameter Speed/Parity/DataBytes/CheckBytes
        MSComm1.Handshaking = comRTSXOnXOff 'Protokoll
        MSComm1.DTREnable = True
    Else
        MSComm1.PortOpen = False
        bOpen = False
        MSComm1.InputLen = 0                'wenn, dann gesamten Input-Buffer auslesen
        MSComm1.Settings = "1200,E,8,2"     'Übertragungsparameter Speed/Parity/DataBytes/CheckBytes
        MSComm1.Handshaking = comRTSXOnXOff 'Protokoll
        MSComm1.DTREnable = True
    End If
    
    If bOpen = False Then
        MSComm1.PortOpen = True             'ComPort öffnen
    End If
    bOpen = True
    
    cGETSTAT = cSTX                     'Startzeichen
    cGETSTAT = cGETSTAT & cESC          'Escape-Sequenz
    cGETSTAT = cGETSTAT & Space$(1)     'Status request
    cGETSTAT = cGETSTAT & cETX          'Endezeichen
    
    cSender = cSTX & cESC & "z"         'Setup-Modifikation
    cSender = cSender & "h" & "0"       'keine Barcode-Höhenmodifikation
    cSender = cSender & "L" & "3"       'Sprache = Deutsch
    cSender = cSender & "n" & "0"       'keine Messages vom Drucker
    cSender = cSender & vbCr & cETX     'Ende des Init-Strings
    
    
    MSComm1.Output = cGETSTAT           'explizites Abrufen des Druckerzustandes

    DoEvents
    
    lStart = Timer
    Do
        lAktuell = Timer
    Loop While lAktuell < lStart + 1
    
    cReturn = MSComm1.Input
    If Mid(cReturn, 3, 1) = cNEG Then
        'Der Drucker meldet einen Fehler
        MsgBox "Der Drucker meldet einen Fehler bzw. ist nicht bereit!", vbCritical, "STOP!"
        MSComm1.PortOpen = False
        Exit Sub
    End If
    
    MSComm1.Output = cSender
    DoEvents
    
    lStart = Timer
    Do
        lAktuell = Timer
    Loop While lAktuell < lStart + 1
    
    cSender = cSTX & cESC & "T" & "99" & cETX   'lösche alle Druckaufträge
    MSComm1.Output = cSender
    DoEvents
    
    lStart = Timer
    Do
        lAktuell = Timer
    Loop While lAktuell < lStart + 1
    
    '******************************************************
    '* große Etiketten mit DM und EURO
    '******************************************************
    
    cSender = cSTX & cESC & "D" & "99"      'Programmkopf setzen
    cSender = cSender & "05" & "2"          'nur noch 5 Felder wg. EURO / Programmtyp 2
    
    '* Artikelnummer
    cSender = cSender & vbCr                'Feldseparator
    cSender = cSender & "025"               'Position in mm X-Achse
    cSender = cSender & "011"               'Position in mm Y-Achse
    cSender = cSender & "016"               '16 Zeichen
    cSender = cSender & "000"               'Höhe Standard
    cSender = cSender & "1"                 'Schriftgröße
    cSender = cSender & "2"                 'Ausrichtung links->rechts
    cSender = cSender & "2"                 'Feldtyp Text
        
    '* Barcode
    cSender = cSender & vbCr                'Feldseparator
    cSender = cSender & "004"               'Position in mm X-Achse
    cSender = cSender & "010"               'Position in mm Y-Achse
    cSender = cSender & "000"               'Anzahl Zeichen
    cSender = cSender & "007"               'Feldhöhe in mm
    cSender = cSender & "0"                 'Vergrößerung 0=80%
    cSender = cSender & "2"                 'Ausrichtung links->rechts
    cSender = cSender & "1"                 'Feldtyp 1=BarCode
    cSender = cSender & "S"                 'BarCodeTyp 2=EAN8 / S=EAN8 nicht lesbar
    cSender = cSender & "1"                 'Prüfziffer berechnen 0=nein / 1=ja
    cSender = cSender & "00"                'Anzahl Vorschaltziffern
    cSender = cSender & ""                  'Vorschaltziffern
   
    '* Euro
    cSender = cSender & vbCr                'Feldseparator
    cSender = cSender & "003"               'Position in mm X-Achse
    cSender = cSender & "007"               'Position in mm Y-Achse
    cSender = cSender & "007"               'Anzahl Zeichen
    cSender = cSender & "000"               'Höhe Standard
    cSender = cSender & "1"                 'Schriftgröße
    cSender = cSender & "2"                 'Ausrichtung links->rechts
    cSender = cSender & "2"                 'Feldtyp 2=Text
    
'    '* DM
'    cSender = cSender & vbCr                'Feldseparator
'    cSender = cSender & "003"               'Position in mm X-Achse
'    cSender = cSender & "004"               'Position in mm Y-Achse
'    cSender = cSender & "007"               'Anzahl Zeichen
'    cSender = cSender & "000"               'Höhe Standard
'    cSender = cSender & "1"                 'Schriftgröße
'    cSender = cSender & "2"                 'Ausrichtung links->rechts
'    cSender = cSender & "2"                 'Feldtyp 2=Text
    
    '* 2.Textzeile
    cSender = cSender & vbCr                'Feldseparator
    cSender = cSender & "003"               'Position in mm X-Achse
    cSender = cSender & "017"               'Position in mm Y-Achse
    cSender = cSender & "018"               'Anzahl Zeichen
    cSender = cSender & "000"               'Höhe Standard
    cSender = cSender & "1"                 'Schriftgröße
    cSender = cSender & "2"                 'Ausrcihtung links-Rechts
    cSender = cSender & "2"                 'Feldtyp 2=Text
    
    '* 1.Textzeile
    cSender = cSender & vbCr                'Feldseparator
    cSender = cSender & "003"               'Position in mm X-Achse
    cSender = cSender & "020"               'Position in mm Y-Achse
    cSender = cSender & "018"               'Anzahl Zeichen
    cSender = cSender & "000"               'Höhe Standard
    cSender = cSender & "1"                 'Schriftgröße
    cSender = cSender & "2"                 'Ausrichtung links-Rechts
    cSender = cSender & "2"                 'Feldtyp 2=Text
    
    'Reserve
    cSender = cSender & vbCr                'Feldseparator
    cSender = cSender & "1"                 'Dummy-Felder
    cSender = cSender & vbCr                'Feldseparator
    cSender = cSender & "2"                 'Dummy-Felder
    cSender = cSender & vbCr                'Feldseparator
    cSender = cSender & "3"                 'Dummy-Felder
    cSender = cSender & vbCr                'Feldseparator
    cSender = cSender & "4"                 'Dummy-Felder
    cSender = cSender & vbCr                'Feldseparator
    cSender = cSender & "5"                 'Dummy-Felder
    cSender = cSender & vbCr                'Feldseparator
    cSender = cSender & cETX                'Textende
    
    MSComm1.Output = cSender
    DoEvents
    
    lStart = Timer
    Do
        lAktuell = Timer
    Loop While lAktuell < lStart + 1
    
    MSComm1.Output = cGETSTAT           'explizites Abrufen des Druckerzustandes
    'Do
    '    Dummy = DoEvents()
    'Loop Until MSComm1.InBufferCount > 0
    DoEvents
    lStart = Timer
    Do
        lAktuell = Timer
    Loop While lAktuell < lStart + 1
    cReturn = MSComm1.Input
    If Mid(cReturn, 3, 1) = cNEG Then
        'Der Drucker meldet einen Fehler
        MsgBox "Der Drucker meldet einen Fehler bzw. ist nicht bereit!", vbCritical, "STOP!"
        MSComm1.PortOpen = False
        Exit Sub
    End If
    
    cDatum = Right(Str$(Year(Now)), 2) & Trim$(Format$(Month(Now), "00")) & gcFilNr
    
    cSQL = "Select * from ETIDRU2"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!filnr) Then
                cFilnr = rsrs!filnr
            Else
                cFilnr = "0"
            End If
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            
            If Len(cBezeich) > 18 Then
                cZeile1 = Left(cBezeich, 18)
                cZeile2 = Mid(cBezeich, 19, Len(cBezeich) - 18)
            Else
                cZeile1 = cBezeich
                cZeile2 = ""
            End If
            
            If Not IsNull(rsrs!artnr) Then
                lartnr = rsrs!artnr
            Else
                lartnr = 0
            End If
            
            cArtNr = Trim$(Str$(lartnr))
            cArtNr = String$(6 - Len(cArtNr), "0") & cArtNr
            cEAN = fnMoveArtNr2EAN8(cArtNr)
            If Not IsNull(rsrs!ANZAHL) Then
                lMenge = rsrs!ANZAHL
            Else
                lMenge = 0
            End If
            
            If Not IsNull(rsrs!vkpr) Then
                dKVkPr1 = rsrs!vkpr
            Else
                dKVkPr1 = 0
            End If
            dKVkPr1EUR = dKVkPr1
            dKVkPr1EUR = (Fix((dKVkPr1EUR * 100) + 0.5)) / 100
            
            cPreisDEM = Format$(dKVkPr1, "######0.00")
            cPreisDEM = Space$(10 - Len(cPreisDEM)) & cPreisDEM
            cPreisDEM = gcWaehrung & cPreisDEM
            
            cPreisEUR = Format$(dKVkPr1EUR, "######0.00")
            cPreisEUR = Space$(10 - Len(cPreisEUR)) & cPreisEUR
            cPreisEUR = "EUR" & cPreisEUR
            
            cSender = cSTX & cESC & "C"
            cSender = cSender & "KISS"
            cSender = cSender & "99"
            cSender = cSender & vbCr & Trim$(Str$(lartnr))
            cSender = cSender & vbCr & cEAN
            cSender = cSender & vbCr & cPreisEUR
            cSender = cSender & vbCr & cZeile2
            cSender = cSender & vbCr & cZeile1
            cSender = cSender & vbCr & cETX
            MSComm1.Output = cSender
            
            cMenge = Trim$(Str$(lMenge))
            cMenge = String$(4 - Len(cMenge), "0") & cMenge
            cSender = cSTX & cESC & "f" & cMenge & cETX
            MSComm1.Output = cSender
            
            If lDelete = vbYes Then
            
            
                SicherInEtisic Trim$(Str$(lartnr)), cFilnr, Check13
                
                cSQL = "Delete from ETIDRU where ARTNR = " & Trim$(Str$(lartnr)) & " and FILNR = " & cFilnr & " "
                gdBase.Execute cSQL, dbFailOnError
                
                ZaehleEtikettenWKL30
    
                FuelleListeEtikettenWKL30
                
                
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    MSComm1.PortOpen = False
    
Exit Sub
LOKAL_ERROR:
    If bOpen Then
        MSComm1.PortOpen = False
    End If
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeBarCodeEtikett3"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeBarCodeEtikett5(lDelete As Long)
    On Error GoTo LOKAL_ERROR
    
    '*************************************************
    '* Diese Funktion dient zur Ansteuerung eines
    '* WAM-Etiketten-Druckers,
    '*************************************************
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim lStart As Long
    Dim lAktuell As Long
    
    Dim cSteuerwerte As String
    
    Dim lComPort As Long
    Dim cSetting As String
    
    ReDim acFelder(1 To 6, 1 To 2) As String
    ReDim cFormatNr(1 To 6) As String
    ReDim cFeldNr(1 To 6) As String
    ReDim cTextArt(1 To 6) As String
    ReDim cFeldLaenge(1 To 6) As String
    ReDim cZeichenGroesse(1 To 6) As String
    ReDim cTextInhalt(1 To 6) As String
    ReDim cFeld(1 To 6) As String
    
    Dim cSatz As String
    Dim cSenden As String
    Dim cPruefZiffer As String
    
    Dim lcount As Long
    Dim bOpen As Boolean
    Dim cDatum As String
    Dim iFileNr As Integer
    Dim cKommentar As String
    Dim cBefehl As String
    
    Dim cSpeed As String
    Dim cParms As String
    Dim cRueck As String
    
    Dim lartnr As Long
    Dim cArtNr As String
    Dim cZeile1 As String
    Dim cZeile2 As String
    Dim cEAN As String
    Dim cBezeich As String
    Dim cLiBesNr As String
    Dim lMenge As Long
    Dim cMenge As String
    Dim dKVkPr1 As Double
    Dim dKVkPr1DEM As Double
    Dim dKVkPr1EUR As Double
    Dim cPreisDEM As String
    Dim cPreisEUR As String
    
    '*********************************************
    '* Anfang des Datenteils (Steuerungszeichen)
    '*********************************************
    Dim cSTX As String
    Dim cETX As String
    Dim cEOT As String
    
    '*********************************************
    '* Ende des Datenteils (Steuerungszeichen)
    '*********************************************
    
    '***************************
    '* Standardwerte festlegen
    '***************************
    bOpen = False
    
    cSTX = Chr$(2)
    cETX = Chr$(3)
    cEOT = Chr$(4)
    
    lComPort = 1
    cSetting = "9600,E,7,1"
    
    acFelder(1, 1) = "006"      'Feld1 Horizontal
    acFelder(1, 2) = "015"      'Feld1 Vertikal
    acFelder(2, 1) = "011"      'Feld2 Horizontal
    acFelder(2, 2) = "015"      'Feld2 Vertikal
    acFelder(3, 1) = "023"      'Feld3 Horizontal
    acFelder(3, 2) = "030"      'Feld3 Vertikal
    acFelder(4, 1) = "028"      'Feld4 Horizontal
    acFelder(4, 2) = "030"      'Feld4 Vertikal
    acFelder(5, 1) = "038"      'Feld5 Horizontal
    acFelder(5, 2) = "030"      'Feld5 Vertikal
    acFelder(6, 1) = "047"      'Feld6 Horizontal
    acFelder(6, 2) = "030"      'Feld6 Vertikal
    
    cFormatNr(1) = "00"
    cFormatNr(2) = "00"
    cFormatNr(3) = "00"
    cFormatNr(4) = "00"
    cFormatNr(5) = "00"
    cFormatNr(6) = "00"
    
    cFeldNr(1) = "01"
    cFeldNr(2) = "02"
    cFeldNr(3) = "03"
    cFeldNr(4) = "04"
    cFeldNr(5) = "05"
    cFeldNr(6) = "06"
    
    cTextArt(1) = "A2"
    cTextArt(2) = "A2"
    cTextArt(3) = "F4"
    cTextArt(4) = "A2"
    cTextArt(5) = "A2"
    cTextArt(6) = "A2"
    
    cFeldLaenge(1) = "120"
    cFeldLaenge(2) = "120"
    cFeldLaenge(3) = "060"
    cFeldLaenge(4) = "120"
    cFeldLaenge(5) = "120"
    cFeldLaenge(6) = "120"
    
    cZeichenGroesse(1) = "001"
    cZeichenGroesse(2) = "001"
    cZeichenGroesse(3) = "010"
    cZeichenGroesse(4) = "001"
    cZeichenGroesse(5) = "002"
    cZeichenGroesse(6) = "002"
    
    cTextInhalt(1) = "B00"
    cTextInhalt(2) = "B00"
    cTextInhalt(3) = "M20"
    cTextInhalt(4) = "B00"
    cTextInhalt(5) = "B00"
    cTextInhalt(6) = "B00"
    
    iFileNr = FreeFile
    Open gcDBPfad & "\SETWAM.DBF" For Binary As #iFileNr
    If LOF(iFileNr) = 0 Then
        Close iFileNr
        loesch "SETWAM"
    Else
        Close iFileNr
        
        cSQL = "Select * from SETWAM"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!KOMMENTAR) Then
                    cKommentar = rsrs!KOMMENTAR
                Else
                    cKommentar = ""
                End If
                cKommentar = UCase$(cKommentar)
                
                If Not IsNull(rsrs!BEFEHL) Then
                    cBefehl = rsrs!BEFEHL
                Else
                    cBefehl = ""
                End If
                cBefehl = UCase$(cBefehl)
                
                Select Case cKommentar
                    Case Is = "COM-SCHNITTSTELLE"
                        lComPort = Val(cBefehl)
                        If lComPort = 0 Then
                            lComPort = 1
                        End If
                        
                    Case Is = "SCHNITTSTELLENGESCHWINDIGKEIT"
                        cSpeed = cBefehl
                        If cSpeed = "" Then
                            cSpeed = "9600"
                        End If
                        
                    Case Is = "SCHNITTSTELLENFORMAT"
                        cParms = cBefehl
                        If cParms = "" Then
                            cParms = "E,7,1"
                        End If
                        
                    Case Is = "HORIZONTALE POSITION FELD 1"
                        acFelder(1, 1) = cBefehl
                        
                    Case Is = "VERTIKALE POSITION FELD 1"
                        acFelder(1, 2) = cBefehl
                        
                    Case Is = "HORIZONTALE POSITION FELD 2"
                        acFelder(2, 1) = cBefehl
                        
                    Case Is = "VERTIKALE POSITION FELD 2"
                        acFelder(2, 2) = cBefehl
                    
                    Case Is = "HORIZONTALE POSITION FELD 3"
                        acFelder(3, 1) = cBefehl
                    
                    Case Is = "VERTIKALE POSITION FELD 3"
                        acFelder(3, 2) = cBefehl
                
                    Case Is = "HORIZONTALE POSITION FELD 4"
                        acFelder(4, 1) = cBefehl
                    
                    Case Is = "VERTIKALE POSITION FELD 4"
                        acFelder(4, 2) = cBefehl
                    
                    Case Is = "HORIZONTALE POSITION FELD 5"
                        acFelder(5, 1) = cBefehl
                    
                    Case Is = "VERTIKALE POSITION FELD 5"
                        acFelder(5, 2) = cBefehl
                    
                    Case Is = "HORIZONTALE POSITION FELD 6"
                        acFelder(6, 1) = cBefehl
                    
                    Case Is = "VERTIKALE POSITION FELD 6"
                        acFelder(6, 2) = cBefehl
                        
                End Select
                
                rsrs.MoveNext
            Loop
            cSetting = cSpeed & "," & cParms
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    '***********************************
    '* Init-Routine für Drucker
    '***********************************
    
    For lcount = 1 To 6
        cFeld(lcount) = cFormatNr(lcount) & cFeldNr(lcount) & cTextArt(lcount) _
                        & acFelder(lcount, 1) & acFelder(lcount, 2) & cFeldLaenge(lcount) _
                        & cZeichenGroesse(lcount) & cTextInhalt(lcount)
    Next lcount
                
    cSatz = Chr$(27) & Chr$(66) & Chr$(27) & Chr$(69) & Chr$(48)
    For lcount = 1 To 6
        cSatz = cSatz & cFeld(lcount) & Chr$(13)
    Next lcount
                
    cPruefZiffer = Mid(Trim$(Str$((100000 + (Len(cSatz) + 2)))), 2, 5)
    
    cSenden = cSTX & cSatz & cETX & cPruefZiffer & cEOT
    
    '***********************************
    '* Verbindung zum Drucker aufbauen
    '***********************************
    
    MSComm1.CommPort = lComPort         'ComPort festlegen
    
    If MSComm1.PortOpen = False Then
        bOpen = False
        MSComm1.InputLen = 0                'wenn, dann gesamten Input-Buffer auslesen
        MSComm1.Settings = cSetting         'Übertragungsparameter Speed/Parity/DataBytes/CheckBytes
        MSComm1.Handshaking = comRTSXOnXOff 'Protokoll
        MSComm1.DTREnable = True
    Else
        MSComm1.PortOpen = False
        bOpen = False
        MSComm1.InputLen = 0                'wenn, dann gesamten Input-Buffer auslesen
        MSComm1.Settings = cSetting         'Übertragungsparameter Speed/Parity/DataBytes/CheckBytes
        MSComm1.Handshaking = comRTSXOnXOff 'Protokoll
        MSComm1.DTREnable = True
    End If
    
    If bOpen = False Then
        MSComm1.PortOpen = True             'ComPort öffnen
    End If
    bOpen = True
    
    '**************************
    '* Senden
    '**************************
    MSComm1.Output = cSenden
    DoEvents
    
    '**************************
    '* Druckdaten lesen
    '**************************
    
    cDatum = Format$(Fix(Now), "DD.MM.YYYY")
    cDatum = gcFilNr & Right(cDatum, 1) & Mid(cDatum, 4, 2)
    
    cSQL = "Select * from ETIDRU2"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                lartnr = rsrs!artnr
            Else
                lartnr = 0
            End If
            cArtNr = Trim$(Str$(lartnr))
            cEAN = fnMoveArtNr2EAN8(cArtNr)
            cEAN = Mid(cEAN, 2, 6)
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = cBezeich & Space$(35 - Len(cBezeich))
            cZeile1 = Mid(cBezeich, 1, 18)
            cZeile2 = Mid(cBezeich, 19, 17)
            
            If Not IsNull(rsrs!LIBESNR) Then
                cLiBesNr = rsrs!LIBESNR
            Else
                cLiBesNr = ""
            End If
            
            If Not IsNull(rsrs!ANZAHL) Then
                lMenge = rsrs!ANZAHL
            Else
                lMenge = 0
            End If
            cMenge = Trim$(Str$(lMenge))
            cMenge = String$(3 - Len(cMenge), "0") & cMenge
            
            If Not IsNull(rsrs!vkpr) Then
                dKVkPr1 = rsrs!vkpr
            Else
                dKVkPr1 = 0
            End If

            
            If gcWaehrung = "EUR" Then
                dKVkPr1DEM = dKVkPr1
                dKVkPr1DEM = (Fix((dKVkPr1DEM * 100) + 0.5)) / 100
            Else
                dKVkPr1DEM = dKVkPr1
            End If
            
            cPreisDEM = Format$(dKVkPr1DEM, "######0.00")
            cPreisDEM = gcWaehrung & "  " & cPreisDEM

            
            cSatz = Chr$(27) & Chr$(65) & Chr$(27) & Chr$(70) & Chr$(48)
            cSatz = cSatz & cZeile1 & Chr$(13)
            cSatz = cSatz & cZeile2 & Chr$(13)
            cSatz = cSatz & cEAN & Chr$(13)
            cSatz = cSatz & cDatum & "/" & gcFilNr & "/" & Trim$(Str$(lartnr)) & Chr$(13)
            cSatz = cSatz & cPreisDEM & Chr$(13)

            cSatz = cSatz & Chr$(27) & Chr$(75) & cMenge
            
            cPruefZiffer = Mid(Trim$(Str$((100000 + (Len(cSatz) + 2)))), 2, 5)
    
            cSenden = cSTX & cSatz & cETX & cPruefZiffer & cEOT
            
            '**************************
            '* Senden
            '**************************
            MSComm1.Output = cSenden
            DoEvents
            
            '*******************************************************
            '* Der WAM-Drucker braucht elend lange,
            '* bis er das nächste Datenpaket verarbeiten kann,
            '* deshalb 3 Sekunden Pause (bei 2 Sek. unterschlägt er
            '* das nächste Datenpaket)
            '*******************************************************
            
            
            cRueck = MSComm1.Input
            Select Case Mid(cRueck, 3, 1)
                Case Is = "L"
                    'alles okay
                Case Is = "M"
                    MsgBox "Papierfehler am Drucker!", vbCritical, "STOP!"
                    Exit Sub
                Case Is = "N"
                    MsgBox "Druckkopf nicht positioniert!", vbCritical, "STOP!"
                    Exit Sub
                Case Is = "O"
                    MsgBox "Übertragungsfehler (Checksumme)!", vbCritical, "STOP!"
                    Exit Sub
                Case Is = "P"
                    MsgBox "Druck unterbrochen!", vbCritical, "STOP!"
                    Exit Sub
                Case Is = "Q"
                    MsgBox "Druck fortgesetzt!", vbCritical, "STOP!"
                Case Is = "R"
                    MsgBox "Etikett oder Format nicht definiert!", vbCritical, "STOP!"
                    Exit Sub
                Case Is = "S"
                    lStart = Timer
                    Do
                        lAktuell = Timer
                    Loop While lAktuell < lStart + (Fix(lMenge / 3)) + 2
            End Select
                
            lStart = Timer
            Do
                lAktuell = Timer
            Loop While lAktuell < lStart + (Fix(lMenge / 3)) + 2
            
            If lDelete = vbYes Then
                rsrs.delete
            End If
            
            rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    '*******************************************
    '* 2 Etiketten nachschieben, damit sauber
    '* abgerissen werden kann
    '*******************************************
    lStart = Timer
    Do
        lAktuell = Timer
    Loop While lAktuell < lStart + 2
    
    
    '***********************************
    '* Init-Routine für Drucker
    '***********************************
    For lcount = 1 To 6
        cFeld(lcount) = cFormatNr(lcount) & cFeldNr(lcount) & cTextArt(lcount) _
                        & acFelder(lcount, 1) & acFelder(lcount, 2) & cFeldLaenge(lcount) _
                        & cZeichenGroesse(lcount) & cTextInhalt(lcount)
    Next lcount
                
    cSatz = Chr$(27) & Chr$(66) & Chr$(27) & Chr$(69) & Chr$(48)
    For lcount = 1 To 6
        cSatz = cSatz & cFeld(lcount) & Chr$(13)
    Next lcount
                
    cPruefZiffer = Mid(Trim$(Str$((100000 + (Len(cSatz) + 2)))), 2, 5)
    
    cSenden = cSTX & cSatz & cETX & cPruefZiffer & cEOT
    
    
    '**************************
    '* Senden
    '**************************
    MSComm1.Output = cSenden
    DoEvents
    
    '**************************
    '* Pause
    '**************************
    lStart = Timer
    Do
        lAktuell = Timer
    Loop While lAktuell < lStart + 2
    
    '***********************************
    '* 2 Leeretiketten drucken
    '***********************************
    
    cSatz = Chr$(27) & Chr$(65) & Chr$(27) & Chr$(70) & Chr$(48)
    cSatz = cSatz & "" & Chr$(13)
    cSatz = cSatz & "" & Chr$(13)
    cSatz = cSatz & "" & Chr$(13)
    cSatz = cSatz & "" & Chr$(13)
    cSatz = cSatz & "" & Chr$(13)
    cSatz = cSatz & "" & Chr$(13)
    cSatz = cSatz & Chr$(27) & Chr$(75) & "2"
    
    cPruefZiffer = Mid(Trim$(Str$((100000 + (Len(cSatz) + 2)))), 2, 5)

    cSenden = cSTX & cSatz & cETX & cPruefZiffer & cEOT
    
    '**************************
    '* Senden
    '**************************
    MSComm1.Output = cSenden
    DoEvents
    
    '**************************
    '* Pause
    '**************************
    lStart = Timer
    Do
        lAktuell = Timer
    Loop While lAktuell < lStart + 2
    
    '***********************************
    '* Init-Routine für Drucker
    '***********************************
    For lcount = 1 To 6
        cFeld(lcount) = cFormatNr(lcount) & cFeldNr(lcount) & cTextArt(lcount) _
                        & acFelder(lcount, 1) & acFelder(lcount, 2) & cFeldLaenge(lcount) _
                        & cZeichenGroesse(lcount) & cTextInhalt(lcount)
    Next lcount
                
    cSatz = Chr$(27) & Chr$(66) & Chr$(27) & Chr$(69) & Chr$(48)
    For lcount = 1 To 6
        cSatz = cSatz & cFeld(lcount) & Chr$(13)
    Next lcount
                
    cPruefZiffer = Mid(Trim$(Str$((100000 + (Len(cSatz) + 2)))), 2, 5)
    
    cSenden = cSTX & cSatz & cETX & cPruefZiffer & cEOT
    
    
    '**************************
    '* Senden
    '**************************
    MSComm1.Output = cSenden
    DoEvents
    
    '**************************
    '* Pause
    '**************************
    lStart = Timer
    Do
        lAktuell = Timer
    Loop While lAktuell < lStart + 3
    
    '**************************
    '* Com-Port schließen
    '**************************
    MSComm1.PortOpen = False
    
Exit Sub
LOKAL_ERROR:
    If bOpen Then
        MSComm1.PortOpen = False
    End If
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeBarCodeEtikett5"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckeBarCodeEtikett4(lDelete As Long)
    On Error GoTo LOKAL_ERROR
    
    '*************************************************
    '* Diese Funktion dient zur Ansteuerung eines
    '* ZEBRA-Etiketten-Druckers über LPT-Port
    '*************************************************
    Dim cFilnr As String
    Dim lartnr As Long
    Dim cArtNr As String
    Dim cBezeich As String
    Dim cLiBesNr As String
    Dim cEAN As String
    Dim cFirma As String
    Dim lMenge As Long
    Dim dKVkPr1 As Double
    Dim dKVkPr1EUR As Double
    Dim dKVkPr1DEM As Double
    Dim cPreisEUR As String
    Dim cPreisDEM As String

    Dim lStart As Long
    Dim lAktuell As Long

    Dim lcount As Long
    Dim lAnzZeile As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim aDeviceName As String
    
    ReDim cDruckZeile(1 To 1) As String
    
    Dim cDruck As String
    Dim lGruppen As Long
    Dim lGruppenCount As Long
    Dim lMaxIndex As Long
    
    '************************
    '* Anfang des Datenteils
    '************************
        
    Dim cAnfang As String
    Dim cEnde As String
    Dim cSTX As String
    Dim cFS As String
    Dim cData As String
    Dim cFontD As String
    Dim cFontB As String
    Dim cFontE As String
    Dim cBarCom As String
    Dim cDatum As String
    
    '************************
    '* Ende des Datenteils
    '************************
    
    '********************************
    '* Drucker-Parameter festlegen
    '********************************
    
    cAnfang = "XA"
    cEnde = "XZ"
    cSTX = Chr$(94)
    cFS = "FS"
    cData = "FD"
    cFontD = "CFD"
    cFontB = "CFB"
    cFontE = "CFE"
    cBarCom = "B8N,30,Y,N,N"
    
    cDatum = Format$(Fix(Now), "DD.MM.YYYY")
    cDatum = gcFilNr & Right(cDatum, 1) & Mid(cDatum, 4, 2)
    
    '********************************
    '* zu druckende Daten lesen
    '********************************
    
    cSQL = "Select * from ETIDRU2"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!filnr) Then
                cFilnr = rsrs!filnr
            Else
                cFilnr = "0"
            End If
            
            If Not IsNull(rsrs!artnr) Then
                lartnr = rsrs!artnr
            Else
                lartnr = 0
            End If
            cArtNr = Trim$(Str$(lartnr))
            cEAN = fnMoveArtNr2EAN8(cArtNr)
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            KonvertAnsiAscii cBezeich
            
            If Not IsNull(rsrs!LIBESNR) Then
                cLiBesNr = rsrs!LIBESNR
            Else
                cLiBesNr = ""
            End If
            
            If Not IsNull(rsrs!ANZAHL) Then
                lMenge = rsrs!ANZAHL
            Else
                lMenge = 0
            End If
            
            If Not IsNull(rsrs!vkpr) Then
                dKVkPr1 = rsrs!vkpr
            Else
                dKVkPr1 = 0
            End If
            If gcWaehrung <> "EUR" Then
                dKVkPr1EUR = dKVkPr1
                dKVkPr1EUR = (Fix((dKVkPr1EUR * 100) + 0.5)) / 100
            Else
                dKVkPr1EUR = dKVkPr1
            End If
            
            If gcWaehrung = "EUR" Then
                dKVkPr1DEM = dKVkPr1
                dKVkPr1DEM = (Fix((dKVkPr1DEM * 100) + 0.5)) / 100
            Else
                dKVkPr1DEM = dKVkPr1
            End If
            
            cPreisDEM = Format$(dKVkPr1DEM, "######0.00")
            
            cPreisEUR = Format$(dKVkPr1EUR, "######0.00")
            
            cFirma = gFirma.FirmaName
            KonvertAnsiAscii cFirma

            cDruck = cSTX & cAnfang & cSTX & cFontD
            cDruck = cDruck & cSTX & "FO076,020" & cSTX & cData & cFirma & cSTX & cFS
            cDruck = cDruck & cSTX & "FO003,030" & cSTX & cData & String$(28, "-") & cSTX & cFS
            cDruck = cDruck & cSTX & cFontB
            cDruck = cDruck & cSTX & "FO003,045" & cSTX & cData & cBezeich & cSTX & cFS
            cDruck = cDruck & cSTX & cFontD
            cDruck = cDruck & cSTX & "FO085,060" & cSTX & cData & cLiBesNr & cSTX & cFS
            cDruck = cDruck & cSTX & cFontD
            cDruck = cDruck & cSTX & "FO085,70" & cSTX & cBarCom
            cDruck = cDruck & cSTX & cData & cEAN & cSTX & cFS
            cDruck = cDruck & cSTX & cFontE
            cDruck = cDruck & cSTX & "FO003,125" & cSTX & cData & gcWaehrung & " " & cSTX & cFS
            cDruck = cDruck & cSTX & "FO080,125" & cSTX & cData & cPreisDEM & cSTX & cFS
            cDruck = cDruck & cSTX & cFontB
            cDruck = cDruck & cSTX & "FO279,145" & cSTX & cData & cDatum & cSTX & cFS
            cDruck = cDruck & cSTX & cFontE
            cDruck = cDruck & cSTX & "FO003,152" & cSTX & cData & "EUR " & cSTX & cFS
            cDruck = cDruck & cSTX & "FO080,152" & cSTX & cData & cPreisEUR & cSTX & cFS
            cDruck = cDruck & cSTX & cEnde
            
            aDeviceName = gcEtikettenDrucker
            
            ReDim cDruckZeile(1 To lMenge) As String
            For lcount = 1 To lMenge
                cDruckZeile(lcount) = cDruck
            Next lcount
            lAnzZeile = lMenge
            OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
            
            If lDelete = vbYes Then
            
                SicherInEtisic Trim$(Str$(lartnr)), cFilnr, Check13
            
                cSQL = "Delete from ETIDRU where ARTNR = " & Trim$(Str$(lartnr)) & " and FILNR = " & cFilnr & " "
                gdBase.Execute cSQL, dbFailOnError
                
                ZaehleEtikettenWKL30
    
                FuelleListeEtikettenWKL30
                
            End If
            
            '*******************************************************************
            '* Lt. Kunde bricht der ZEBRA-Drucker nach ca. 16 Etiketten
            '*******************************************************************
            
            lStart = Timer
            Do
                lAktuell = Timer
            Loop While lAktuell < lStart + 2
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
Exit Sub
LOKAL_ERROR:

    
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeBarCodeEtikett4"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub SchreibeKonfigurationEtikettWKL30()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cKommentar As String
    Dim cBefehl As String
    Dim ctmp As String
    Dim llen As Long
    Dim lAnzKomma As Long
    Dim cZeichen As String
    Dim cZiel As String
    
    'Schnittstelle
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Druckerschnittstelle'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Druckerschnittstelle"
    If Option1(0).Value = True Then
        rsrs!BEFEHL = "LPT1"
    ElseIf Option1(1).Value = True Then
        rsrs!BEFEHL = "LPT2"
    ElseIf Option1(2).Value = True Then
        rsrs!BEFEHL = "LPT3"
    End If
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Name der Firma
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Name'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Name"
    rsrs!BEFEHL = Trim$(Text1(0).Text)
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Druckposition Firmenname
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Firmenname'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Firmenname"
    cBefehl = "A"
    ctmp = Text1(1).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(2).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    For lcount = 0 To 3
        If Option2(lcount).Value = True Then
            ctmp = Trim$(Str$(lcount))
            Exit For
        End If
    Next lcount
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(3).Text
    ctmp = Trim$(ctmp)
    Select Case ctmp
        Case Is = 6
            ctmp = "1"
        Case Is = 7
            ctmp = "2"
        Case Is = 10
            ctmp = "3"
        Case Is = 12
            ctmp = "4"
        Case Is = 24
            ctmp = "5"
    End Select
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(4).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(5).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    cBefehl = cBefehl & "N,"
    
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Druckposition Bestellnummer
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position der Bestellnummer'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position der Bestellnummer"
    cBefehl = "A"
    ctmp = Text1(6).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(7).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    For lcount = 4 To 7
        If Option2(lcount).Value = True Then
            ctmp = Trim$(Str$(lcount - 4))
            Exit For
        End If
    Next lcount
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(8).Text
    ctmp = Trim$(ctmp)
    Select Case ctmp
        Case Is = 6
            ctmp = "1"
        Case Is = 7
            ctmp = "2"
        Case Is = 10
            ctmp = "3"
        Case Is = 12
            ctmp = "4"
        Case Is = 24
            ctmp = "5"
    End Select
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(9).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(10).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    cBefehl = cBefehl & "N,"
    
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Druckposition Linie
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Linie'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Linie"
    cBefehl = "LO"
    ctmp = Text1(11).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(12).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(13).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(14).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp
    
    'Kein Komma am Ende, da kein dynamischer Input bei Linie!
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    'Druckposition Artikelbezeichnung
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Artikelbezeichnung'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Artikelbezeichnung"
    cBefehl = "A"
    ctmp = Text1(15).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(16).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    For lcount = 8 To 11
        If Option2(lcount).Value = True Then
            ctmp = Trim$(Str$(lcount - 8))
            Exit For
        End If
    Next lcount
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(17).Text
    ctmp = Trim$(ctmp)
    Select Case ctmp
        Case Is = 6
            ctmp = "1"
        Case Is = 7
            ctmp = "2"
        Case Is = 10
            ctmp = "3"
        Case Is = 12
            ctmp = "4"
        Case Is = 24
            ctmp = "5"
    End Select
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(18).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(19).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    cBefehl = cBefehl & "N,"
    
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    'Position Barcode
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Barcode'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Barcode"
    cBefehl = "B"
    ctmp = Text1(20).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(21).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    For lcount = 12 To 15
        If Option2(lcount).Value = True Then
            ctmp = Trim$(Str$(lcount - 12))
            Exit For
        End If
    Next lcount
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(22).Text
    ctmp = Trim$(ctmp)
    Select Case ctmp
        Case Is = "8"
            ctmp = "E80"
        Case Is = "13"
            ctmp = "E30"
    End Select
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(23).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(24).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(40).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    If Check1.Value = vbChecked Then
        ctmp = "B"
    Else
        ctmp = "N"
    End If
    cBefehl = cBefehl & ctmp & ","
    
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    'Druckposition Datum
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Datum'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Datum"
    cBefehl = "A"
    ctmp = Text1(25).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(26).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    For lcount = 16 To 19
        If Option2(lcount).Value = True Then
            ctmp = Trim$(Str$(lcount - 16))
            Exit For
        End If
    Next lcount
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(27).Text
    ctmp = Trim$(ctmp)
    Select Case ctmp
        Case Is = 6
            ctmp = "1"
        Case Is = 7
            ctmp = "2"
        Case Is = 10
            ctmp = "3"
        Case Is = 12
            ctmp = "4"
        Case Is = 24
            ctmp = "5"
    End Select
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(28).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(29).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    cBefehl = cBefehl & "N,"
    
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    'Druckposition DM
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position DM'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position DM"
    cBefehl = "S"
    ctmp = Text1(30).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp '& ","
    

    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    'Druckposition Euro
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Euro'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Euro"
    cBefehl = "A"
    ctmp = Text1(35).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(36).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    For lcount = 24 To 27
        If Option2(lcount).Value = True Then
            ctmp = Trim$(Str$(lcount - 24))
            Exit For
        End If
    Next lcount
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(37).Text
    ctmp = Trim$(ctmp)
    Select Case ctmp
        Case Is = 6
            ctmp = "1"
        Case Is = 7
            ctmp = "2"
        Case Is = 10
            ctmp = "3"
        Case Is = 12
            ctmp = "4"
        Case Is = 24
            ctmp = "5"
    End Select
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(38).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    ctmp = Text1(39).Text
    ctmp = Trim$(ctmp)
    cBefehl = cBefehl & ctmp & ","
    
    cBefehl = cBefehl & "N,"
    
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Firmenname oder Bestellnummer drucken
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Name drucken (J/N)'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Name drucken (J/N)"
    If Check2(0).Value = vbChecked Then
        cBefehl = "J"
    Else
        cBefehl = "N"
    End If
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeKonfigurationEtikettWKL30"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SchreibeStandardEtikettWKL30()
    On Error GoTo LOKAL_ERROR
    
    If Len(gcEtikettenDrucker) >= 11 Then
        If UCase$(Left(gcEtikettenDrucker, 11)) = "ELTRON 2746" Then
            SchreibeStandardEtikettEltron2746WKL30
        Else
            SchreibeStandardEtikettEltronWKL30
        End If
    Else
        SchreibeStandardEtikettEltronWKL30
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeStandardEtikettWKL30"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SchreibeStandardEtikettEltronWKL30()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cKommentar As String
    Dim cBefehl As String
    Dim ctmp As String
    Dim llen As Long
    Dim lAnzKomma As Long
    Dim cZeichen As String
    Dim cZiel As String
    
    'Schnittstelle
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Druckerschnittstelle'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Druckerschnittstelle"
    rsrs!BEFEHL = "LPT1"
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    'Name der Firma
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Name'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Name"
    rsrs!BEFEHL = gFirma.FirmaName
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Druckposition Firmenname
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Firmenname'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Firmenname"
    cBefehl = "A270,0,0,1,1,1,N,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    'Druckposition Bestellnummer
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position der Bestellnummer'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position der Bestellnummer"
    cBefehl = "A270,0,0,1,1,1,N,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Druckposition Linie
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Linie'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Linie"
    cBefehl = "LO270,17,290,2"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Druckposition Artikelbezeichnung
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Artikelbezeichnung'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Artikelbezeichnung"
    cBefehl = "A270,25,0,1,1,1,N,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Position Barcode
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Barcode'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Barcode"
    cBefehl = "B315,42,0,E80,2,2,32,B,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Druckposition Datum
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Datum'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Datum"
    cBefehl = "A550,77,3,1,1,1,N,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Druckposition DM
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position DM'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position DM"
    cBefehl = "A270,112,0,2,2,1,N,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Druckposition Euro
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Euro'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Euro"
    cBefehl = "A270,87,0,2,2,1,N,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Firmenname oder Bestellnummer drucken
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Name drucken (J/N)'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Name drucken (J/N)"
    cBefehl = "J"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeStandardEtikettEltronWKL30"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub SchreibeStandardEtikettEltron2746WKL30()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cKommentar As String
    Dim cBefehl As String
    Dim ctmp As String
    Dim llen As Long
    Dim lAnzKomma As Long
    Dim cZeichen As String
    Dim cZiel As String
    
    'Schnittstelle
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Druckerschnittstelle'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Druckerschnittstelle"
    rsrs!BEFEHL = "LPT1"
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Name der Firma
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Name'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Name"
    rsrs!BEFEHL = gFirma.FirmaName
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Druckposition Firmenname
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Firmenname'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Firmenname"
    cBefehl = "A530,20,0,1,1,1,N,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    'Druckposition Bestellnummer
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position der Bestellnummer'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position der Bestellnummer"
    cBefehl = "A530,20,0,1,1,1,N,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Druckposition Linie
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Linie'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Linie"
    cBefehl = "LO530,37,290,2"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    'Druckposition Artikelbezeichnung
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Artikelbezeichnung'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Artikelbezeichnung"
    cBefehl = "A530,45,0,1,1,1,N,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    'Position Barcode
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Barcode'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Barcode"
    cBefehl = "B580,62,0,E80,2,2,32,B,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    'Druckposition Datum
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Datum'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Datum"
    cBefehl = "A800,97,3,1,1,1,N,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    'Druckposition DM
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position DM'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position DM"
    cBefehl = "A530,132,0,2,2,1,N,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
    'Druckposition Euro
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Position Euro'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Position Euro"
    cBefehl = "A530,107,0,2,2,1,N,"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    'Firmenname oder Bestellnummer drucken
    cSQL = "Select * from SETELTRO where KOMMENTAR = 'Name drucken (J/N)'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!KOMMENTAR = "Name drucken (J/N)"
    cBefehl = "J"
    rsrs!BEFEHL = cBefehl
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeStandardEtikettEltron2746WKL30"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZaehleEtikettenWKL30()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzArt As Long
    Dim dAnzEti As Double
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "delete from ETIDRU where vkpr > 9999999999"
    gdBase.Execute cSQL, dbFailOnError
    
    If gsETILS <> "" Then
        cSQL = "Select count(ARTNR) as ANZART from LSTEETI "
    Else
        cSQL = "Select count(ARTNR) as ANZART from ETIDRU"
        If Check13.Value = vbChecked Then
            cSQL = cSQL & " where PCNAME = '" & srechnertab & "' "
        End If
    End If
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!anzart) Then
            lAnzArt = rsrs!anzart
        Else
            lAnzArt = 0
        End If
    Else
        lAnzArt = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If gsETILS <> "" Then
        cSQL = "Select SUM(ANZAHL) as ANZETI from LSTEETI "
    Else
        cSQL = "Select SUM(ANZAHL) as ANZETI from ETIDRU"
        
        If Check13.Value = vbChecked Then
            cSQL = cSQL & " where PCNAME = '" & srechnertab & "' "
        End If
    End If
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!ANZETI) Then
            dAnzEti = rsrs!ANZETI
        Else
            dAnzEti = 0
        End If
    Else
        dAnzEti = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Label4(2).Caption = Format$(lAnzArt, "###,###,###,##0")
    Label4(3).Caption = Format$(dAnzEti, "###,###,###,##0")
    
    
    Label4(6).Caption = "0"
    Label4(9).Caption = "0"
    
    Label4(7).Caption = Format$(lAnzArt, "###,###,###,##0")
    Label4(8).Caption = Format$(dAnzEti, "###,###,###,##0")
    
    Label4(7).Caption = SwapStr(Label4(7).Caption, ".", "")
    Label4(8).Caption = SwapStr(Label4(8).Caption, ".", "")
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3021 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ZaehleEtikettenWKL30"
        Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub

Private Sub cboStrichDINA4_Click()
On Error GoTo LOKAL_ERROR

    Label1(5).Visible = False

    Select Case cboStrichDINA4.Text
            
        Case "35,6 x 16,9"
        
            Label1(5).Visible = True
        
        Case "35,6 x 16,9 Variante 2"
        
            Label1(5).Visible = True
    
    End Select

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboStrichDINA4_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check13_Click()
On Error GoTo LOKAL_ERROR

ZaehleEtikettenWKL30
FuelleListeEtikettenWKL30
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check13_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function checkAuswahlEtiketten() As Boolean
On Error GoTo LOKAL_ERROR

    Dim bFound As Boolean
    Dim lcount As Long

    checkAuswahlEtiketten = False
    bFound = False
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            bFound = True
        End If
    Next lcount
    If Not bFound Then
        MsgBox "Zum Drucken bitte mindestens einen Listeneintrag auswählen!", vbInformation, "Winkiss Hinweis:"
        List2.SetFocus
    Else
        checkAuswahlEtiketten = True
    End If
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkAuswahlEtiketten"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Check2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    If Check2(0).Visible Then
        If Check2(0).Value = vbChecked Then
            Check2(1).Value = vbUnchecked
        Else
            Check2(1).Value = vbChecked
        End If
    End If
    If Check2(1).Visible Then
        If Check2(1).Value = vbChecked Then
            Check2(0).Value = vbUnchecked
        Else
            Check2(0).Value = vbChecked
        End If
    End If
Exit Sub
LOKAL_ERROR:
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Check2_Click"
        Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Label3.Caption).Text = Text1(Label3.Caption).Text & Command0(Index).Caption
    Text1(Label3.Caption).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As Long
    Dim cDrucker As String
    Dim bReturn As Boolean
    Dim lAnz As Long
    Dim lcount As Long
    Dim iRet As Integer
    Dim aDeviceName As String
    Dim cEscapeSequenz As String
    ReDim acArtNr(0 To 0) As String
    ReDim acAnzEti(0 To 0) As String
    Dim bFound As Boolean
    Dim lAnzahl As Long
    Dim cLBSatz As String
    Dim cArtNr As String
    Dim cAnzEti As String
    Dim lLoeschen As Long
    Dim cSQL As String
    Dim lRecordCount As Long
    Dim lAkt As Long
    Dim cFilnr As String
    
    Screen.MousePointer = 11
    
    lLoeschen = vbNo
    
    Select Case Index
    
        Case Is = 5 'Strichcode Spezial Zebra TLP 2844
            '1. Checken ob Etiketten ausgewählt wurden
            
            If checkAuswahlEtiketten = False Then Exit Sub
            
            Frame4.Visible = False
            Frame8.Visible = True
            
            ZeigeFokus_Endlos "STRICH"
            
        Case Is = 6 'Zurück Strichcode Spezial Zebra TLP 2844
            Frame8.Visible = False
            Frame4.Visible = True
        Case Is = 7 'Zurück TLP2844 Regaletiketten
            voreinstellungspeichern30
            Frame10.Visible = False
            Frame4.Visible = True
            
        Case Is = 8 'TLP2844 Regaletiketten
            '1. Checken ob Etiketten ausgewählt wurden
            
            voreinstellungspeichern30
    
            If checkAuswahlEtiketten = False Then Exit Sub
            
            Frame4.Visible = False
            Frame10.Visible = True
            
            Label13.Caption = ""
            
            ZeigeFokus_Endlos "REGAL"
            
        Case Is = 0  '//StrichcodeEtiketten
            If IsAktionZulaessig("Etiketten drucken") = False Then
                Exit Sub
            End If
            
            lLoeschen = MsgBox("Druckdaten nach dem Drucken löschen?", vbQuestion + vbYesNo, "DRUCKDATEN LÖSCHEN")
            If gcEtikettenDrucker = gcListenDrucker Then
                bFound = False
                lAnzahl = -1
                For lcount = 0 To List2.ListCount - 1
                    If List2.Selected(lcount) = True Then
                        bFound = True
                        lAnzahl = lAnzahl + 1
                        ReDim Preserve acArtNr(0 To lAnzahl) As String
                        ReDim Preserve acAnzEti(0 To lAnzahl) As String
                        cLBSatz = List2.list(lcount)
                        cArtNr = Left(cLBSatz, 6)
                        cArtNr = Trim$(cArtNr)
                        acArtNr(lAnzahl) = cArtNr
                        
                        cAnzEti = Trim$(Mid(cLBSatz, Len(cLBSatz) - 18, 10))
                        '//aenderung
                        ctmp = Left(cAnzEti, 3)
                        ctmp = Trim$(ctmp)
                        cAnzEti = Val(ctmp)
                        '//End Aenderung
                        acAnzEti(lAnzahl) = cAnzEti
                    End If
                Next lcount
                If Not bFound Then
                    MsgBox "Zum Drucken bitte mindestens einen Listeneintrag auswählen!", vbCritical, "STOP!"
                    List2.SetFocus
                Else
                    'Etikettendruck über CrystalReport-Formular
                    DruckeGrundPreisEtiketten2WKL30 acArtNr(), lAnzahl, acAnzEti()
                    Erase acArtNr
                End If

            Else
                iRet = MsgBox("Wollen Sie die Etiketten auf Ihrem Listendrucker erstellen?", vbQuestion + vbYesNoCancel, "DRUCKER")
                If iRet = vbCancel Then
                    Exit Sub
                ElseIf iRet = vbYes Then
                    'Hier den Listendrucker aktivieren
                    bFound = False
                    lAnzahl = -1
                    For lcount = 0 To List2.ListCount - 1
                        If List2.Selected(lcount) = True Then
                            bFound = True
                            lAnzahl = lAnzahl + 1
                            ReDim Preserve acArtNr(0 To lAnzahl) As String
                            ReDim Preserve acAnzEti(0 To lAnzahl) As String
                            cLBSatz = List2.list(lcount)
                            cArtNr = Left(cLBSatz, 6)
                            cArtNr = Trim$(cArtNr)
                            acArtNr(lAnzahl) = cArtNr
                            cAnzEti = Trim$(Mid(cLBSatz, Len(cLBSatz) - 18, 10))
                            '//aenderung
                            ctmp = Left(cAnzEti, 3)
                            ctmp = Trim$(ctmp)
                            cAnzEti = Val(ctmp)
                            '//End Aenderung
                            
                            acAnzEti(lAnzahl) = cAnzEti
                        End If
                    Next lcount
                    If Not bFound Then
                        MsgBox "Zum Drucken bitte mindestens einen Listeneintrag auswählen!", vbCritical, "STOP!"
                        List2.SetFocus
                    Else
                        'Etikettendruck über CrystalReport-Formular
                        DruckeGrundPreisEtiketten2WKL30 acArtNr(), lAnzahl, acAnzEti()
                        Erase acArtNr
                    End If
                    
                    
                    
                    
                ElseIf iRet = vbNo Then  '//vbNo, Etiketten nicht im Listendrucker erstellt
                
                
                    iRet = MsgBox("DIN A4 Blatt - Etiketten, ist das richtig?", vbQuestion + vbYesNoCancel, "DRUCKER")
                    If iRet = vbCancel Then
                        Exit Sub
                    ElseIf iRet = vbYes Then
                    
                        Dim cDruckertemp As String
                        cDruckertemp = gcListenDrucker
                        gcListenDrucker = gcEtikettenDrucker
                    
                        'Hier den Listendrucker aktivieren
                        bFound = False
                        lAnzahl = -1
                        For lcount = 0 To List2.ListCount - 1
                            If List2.Selected(lcount) = True Then
                                bFound = True
                                lAnzahl = lAnzahl + 1
                                ReDim Preserve acArtNr(0 To lAnzahl) As String
                                ReDim Preserve acAnzEti(0 To lAnzahl) As String
                                cLBSatz = List2.list(lcount)
                                cArtNr = Left(cLBSatz, 6)
                                cArtNr = Trim$(cArtNr)
                                acArtNr(lAnzahl) = cArtNr
                                
                                cAnzEti = Trim$(Mid(cLBSatz, Len(cLBSatz) - 18, 10))
                                '//aenderung
                                ctmp = Left(cAnzEti, 3)
                                ctmp = Trim$(ctmp)
                                cAnzEti = Val(ctmp)
                                '//End Aenderung
                                
                            
                                
                                acAnzEti(lAnzahl) = cAnzEti
                            End If
                        Next lcount
                        If Not bFound Then
                            MsgBox "Zum Drucken bitte mindestens einen Listeneintrag auswählen!", vbCritical, "STOP!"
                            List2.SetFocus
                        Else
                            'Etikettendruck über CrystalReport-Formular
                            DruckeGrundPreisEtiketten2WKL30 acArtNr(), lAnzahl, acAnzEti()
                            Erase acArtNr
                        End If
                        
                        gcListenDrucker = cDruckertemp
                
                    ElseIf iRet = vbNo Then
                    
                        'hier das Alte
                        If Modul6.FindFile(App.Path, "aWOKINE.rpt") Then
                            bFound = False
                            lAnzahl = -1
                            For lcount = 0 To List2.ListCount - 1
                                If List2.Selected(lcount) = True Then
                                    bFound = True
                                    lAnzahl = lAnzahl + 1
                                    ReDim Preserve acArtNr(0 To lAnzahl) As String
                                    ReDim Preserve acAnzEti(0 To lAnzahl) As String
                                    cLBSatz = List2.list(lcount)
                                    cArtNr = Left(cLBSatz, 6)
                                    cArtNr = Trim$(cArtNr)
                                    acArtNr(lAnzahl) = cArtNr
                                    cAnzEti = Trim$(Right(cLBSatz, 15))
                                    ctmp = Left(cAnzEti, 2)
                                    ctmp = Trim$(ctmp)
                                    cAnzEti = Val(ctmp)
                                    
                                    acAnzEti(lAnzahl) = cAnzEti
                                End If
                            Next lcount
                            If Not bFound Then
                                MsgBox "Zum Drucken bitte mindestens einen Listeneintrag auswählen!", vbCritical, "STOP!"
                                List2.SetFocus
                            Else
                                DruckeNettoStrichcode acArtNr(), lAnzahl, acAnzEti()
                                Erase acArtNr
                                
                            End If
                        
                            
                        Else
                            If Modul6.FindFile(gcDBPfad, "aWKL30ys.rpt") Then 'neu Strichcode über Crystal auf etidrucker
                                bFound = False
                                lAnzahl = -1
                                For lcount = 0 To List2.ListCount - 1
                                    If List2.Selected(lcount) = True Then
                                        bFound = True
                                        lAnzahl = lAnzahl + 1
                                        ReDim Preserve acArtNr(0 To lAnzahl) As String
                                        ReDim Preserve acAnzEti(0 To lAnzahl) As String
                                        cLBSatz = List2.list(lcount)
                                        cArtNr = Left(cLBSatz, 6)
                                        cArtNr = Trim$(cArtNr)
                                        acArtNr(lAnzahl) = cArtNr
                                        cAnzEti = Trim$(Right(cLBSatz, 15))
                                        ctmp = Left(cAnzEti, 2)
                                        ctmp = Trim$(ctmp)
                                        cAnzEti = Val(ctmp)
                                        
                                        acAnzEti(lAnzahl) = cAnzEti
                                    End If
                                Next lcount
                                If Not bFound Then
                                    MsgBox "Zum Drucken bitte mindestens einen Listeneintrag auswählen!", vbCritical, "STOP!"
                                    List2.SetFocus
                                Else
                                    DruckeStrichcodeY acArtNr(), lAnzahl, acAnzEti()
                                    reportbildschirmToPrinterETI "aWKL30ys", gcEtikettenDrucker, True
                                    Erase acArtNr
                                    
                                End If
                            Else
                                aDeviceName = gcEtikettenDrucker
                                cEscapeSequenz = ""
                                DruckeBarCodeEtikett2 aDeviceName, cEscapeSequenz, lcount
                            End If
                        End If
                    
                    
                    
                    
                    
                    End If
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                    
                    
                    
                End If
            End If
            
            
            
            
            
            
            If lLoeschen = vbYes Then
                For lcount = 0 To List2.ListCount - 1
                    If List2.Selected(lcount) = True Then
                    
                    End If
                Next lcount
            End If
            
            AktionAustragen "Etiketten drucken"
        Case Is = 1
            'lLoeschen = MsgBox("Druckdaten nach dem Drucken löschen?", vbQuestion + vbYesNo, "DRUCKDATEN LÖSCHEN")
            LeseKonfigurationEtikettWKL30
            Frame3.Visible = False
            Frame4.Visible = False
            
            Frame1.Visible = True
            Command1(0).Visible = False
            Command1(1).Visible = False
            Command1(2).Visible = False
            Command1(3).Visible = False
            Command1(4).Visible = False
            Command4(1).Visible = False
            Command2(0).SetFocus
            
        Case Is = 2
        
            Dim sSQL As String
            Dim i As Integer
            
            If Not tableSuchenDBKombi("VOREINAP", 2) Then
                sSQL = "Create table VOREINAP ( Schluessel Text(30),Wert Text(30) )"
                gdApp.Execute sSQL, dbFailOnError
            End If
            
            sSQL = "Delete from Voreinap where schluessel = 'Tintenstrahl'"
            gdApp.Execute sSQL, dbFailOnError
                
            sSQL = "Delete from Voreinap where schluessel = 'Laser'"
            gdApp.Execute sSQL, dbFailOnError
            
            For lcount = 0 To cboRegalDinA4.ListCount - 1
                sSQL = "Delete from Voreinap where schluessel = '" & cboRegalDinA4.list(lcount) & "'"
                gdApp.Execute sSQL, dbFailOnError
            Next lcount
            
            For lcount = 0 To cboStrichDINA4.ListCount - 1
                sSQL = "Delete from Voreinap where schluessel = '" & cboStrichDINA4.list(lcount) & "'"
                gdApp.Execute sSQL, dbFailOnError
            Next lcount
            
            If cboRegalDinA4.Text <> "bitte auswählen" Then
                sSQL = "Insert into Voreinap (schluessel,wert) values ('" & cboRegalDinA4.Text & "','EIN')"
                gdApp.Execute sSQL, dbFailOnError
            End If
            
            
            If cboStrichDINA4.Text <> "bitte auswählen" Then
                sSQL = "Insert into Voreinap (schluessel,wert) values ('" & cboStrichDINA4.Text & "','EIN')"
                gdApp.Execute sSQL, dbFailOnError
            End If
            
            If Option3(0).Value = True Then
                SchreibeVoreinap Option3(0).Caption, "EIN"
            ElseIf Option3(1).Value = True Then
                SchreibeVoreinap Option3(1).Caption, "EIN"
            End If
            
            
            
            
            voreinstellungspeichern30
            
            Unload frmWKL30
    
        Case Is = 3
            If List2.ListCount > 0 Then
                setzedrucker gcListenDrucker
                    
                Printer.FontName = "Courier New"
                Printer.FontSize = 10
                
                Printer.Print
                Printer.Print
                Printer.Print Space$(10) & "PREISLISTE"
                Printer.Print Space$(10) & "----------"
                Printer.Print
                Printer.Print
                
                lAkt = 6
                For lcount = 0 To List2.ListCount - 1
                    cLBSatz = List2.list(lcount)
                    Printer.Print Space$(10) & cLBSatz
                    lAkt = lAkt + 1
                    If lAkt / 5 = Fix(lAkt / 5) Then
                        Printer.Print
                        lAkt = lAkt + 1
                    End If
                    If lAkt > 60 Then
                        Printer.NewPage
                        Printer.Print
                        Printer.Print
                        Printer.Print Space$(10) & "PREISLISTE"
                        Printer.Print Space$(10) & "----------"
                        Printer.Print
                        Printer.Print
                        lAkt = 6
                    End If
                Next lcount
                
                Printer.EndDoc
            End If
        Case Is = 4  '//Preisetiketten drücken

            If IsAktionZulaessig("Etiketten drucken") = False Then
                Exit Sub
            End If
            
            lLoeschen = MsgBox("Druckdaten nach dem Drucken löschen?", vbQuestion + vbYesNo, "DRUCKDATEN LÖSCHEN")
            
            setzedrucker gcListenDrucker
            
            bFound = False
            lAnzahl = -1
            For lcount = 0 To List2.ListCount - 1
                If List2.Selected(lcount) = True Then
                    bFound = True
                    lAnzahl = lAnzahl + 1
                    ReDim Preserve acArtNr(0 To lAnzahl) As String
                    ReDim Preserve acAnzEti(0 To lAnzahl) As String
                    cLBSatz = List2.list(lcount)
                    cArtNr = Left(cLBSatz, 6)
                    cArtNr = Trim$(cArtNr)
                    acArtNr(lAnzahl) = cArtNr
                    cAnzEti = Trim$(Right(cLBSatz, 10))
                    acAnzEti(lAnzahl) = cAnzEti
                End If
            Next lcount
            If Not bFound Then
                MsgBox "Zum Drucken bitte mindestens einen Listeneintrag auswählen!", vbCritical, "STOP!"
                List2.SetFocus
            Else
                'Etikettendruck über CrystalReport-Formular
                DruckeEtikettenWKL30
            End If
            
            AktionAustragen "Etiketten drucken"
    End Select
    
    If lLoeschen = vbYes Then
        lRecordCount = List2.ListCount - 1
        For lcount = 0 To lRecordCount
            If lcount > lRecordCount Then
                Exit For
            End If
            If List2.Selected(lcount) = True Then
                cLBSatz = List2.list(lcount)
                cArtNr = Left(cLBSatz, 6)
                cArtNr = Trim$(cArtNr)
                cFilnr = Mid(cLBSatz, 70, 2)
                cFilnr = Trim$(cFilnr)

                cAnzEti = Mid(cLBSatz, Len(cLBSatz) - 18, 10)
                cAnzEti = Trim$(cAnzEti)
                
                
                
                
                
                
                
'                cFilnr = Trim$(Right(cLBSatz, 2))
'
'                cAnzEti = Trim$(Right(cLBSatz, 15))
'
'
'                ctmp = Left(cAnzEti, 2)
'                ctmp = Trim$(ctmp)
'                cAnzEti = Val(ctmp)
    
                If gsETILS <> "" Then
                    cSQL = "Delete from ETIDRULS where ARTNR = " & cArtNr & " And FILNR = " & cFilnr & " And Anzahl = " & cAnzEti & ""
                Else
                    
                    SicherInEtisic cArtNr, cFilnr, Check13
                    cSQL = "Delete from ETIDRU where ARTNR = " & cArtNr & " and FILNR = " & cFilnr & " "
                End If
                
                gdBase.Execute cSQL, dbFailOnError
                
                List2.RemoveItem lcount
                lcount = lcount - 1
                lRecordCount = lRecordCount - 1
            End If
        Next lcount
        
        ZaehleEtikettenWKL30
        FuelleListeEtikettenWKL30
    End If
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
'    Resume Next
End Sub
Private Sub SchreibeVoreinap(schluessel As String, Wert As String)
    On Error GoTo LOKAL_ERROR
        Dim sSQL As String
        
        sSQL = "Insert into VOREINAP (Schluessel,Wert) values ('" & schluessel & "', '" & Wert & "')"
        gdApp.Execute sSQL, dbFailOnError

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeVoreinap"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
End Sub

Private Sub DruckeBarCodeEtikett2(aDeviceName, cEscapeSequenz, lDelete As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim lhPrinter As Long
    Dim lReturn As Long
    Dim lpcWritten As Long
    Dim lDoc As Long
    Dim lRet As Long
    Dim MyDocInfo As DOCINFO
    Dim lSize As Long
    Dim aDat As String
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cKommentar As String
    Dim cBefehl As String
    ReDim cDruckBefehl(1 To 13) As String
    
    Dim cFilnr As String
    Dim cArtNr As String
    Dim cArtNrMerker As String
    Dim cBezeich As String
    Dim cEAN As String
    Dim cVKPR As String
    Dim cEuro As String
    Dim dVkPr As Double
    Dim dEuro As Double
    Dim dAnzahl As Double
    Dim cLiBesNr As String
    Dim cDatum As String
    Dim lPos As Long
    Dim lWert As Long
    Dim lPruef As Long
    Dim cZeichen As String
    Dim cLetzter As String
    Dim cZiel As String
    Dim cDBFeld As String
    Dim lcount As Long
    Dim iFileNr As Integer
    
    '***************************************************
    '* zu druckende Daten in ETIDRU2 schreiben
    '***************************************************
    
    loeschNEW "ETIDRU2", gdBase
    
    cSQL = "Create Table ETIDRU2 "
    cSQL = cSQL & "( ARTNR Long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", BESTAND Long"
    cSQL = cSQL & ", ANZAHL Long"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", EAN Text(13)"
    cSQL = cSQL & ", LINR Long"
    cSQL = cSQL & ", LPZ Long"
    cSQL = cSQL & ", FILNR Long"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
        
        
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            cArtNr = Left(List2.list(lcount), 6)
            cArtNr = Trim$(cArtNr)
            cFilnr = Trim$(Right(List2.list(lcount), 2))
            
            If gsETILS <> "" Then
                cSQL = "Insert into ETIDRU2 Select "
                cSQL = cSQL & " ARTNR"
                cSQL = cSQL & " , Bezeich "
                cSQL = cSQL & " , VKPR "
                cSQL = cSQL & " , Bestand "
                cSQL = cSQL & " , Anzahl "
                cSQL = cSQL & " , Libesnr"
                cSQL = cSQL & " , Ean "
                cSQL = cSQL & " , Linr "
                cSQL = cSQL & " , Lpz "
                cSQL = cSQL & " , Filnr "
                cSQL = cSQL & " from LSTEETI where ARTNR = " & cArtNr
                cSQL = cSQL & " and FILNR = " & cFilnr
                gdBase.Execute cSQL, dbFailOnError
                

            Else
            
                cSQL = "Insert into ETIDRU2 Select "
                cSQL = cSQL & " ARTNR"
                cSQL = cSQL & " , Bezeich "
                cSQL = cSQL & " , VKPR "
                cSQL = cSQL & " , Bestand "
                cSQL = cSQL & " , Anzahl "
                cSQL = cSQL & " , Libesnr"
                cSQL = cSQL & " , Ean "
                cSQL = cSQL & " , Linr "
                cSQL = cSQL & " , Lpz "
                cSQL = cSQL & " , Filnr "
                cSQL = cSQL & " from ETIDRU where ARTNR = " & cArtNr
                cSQL = cSQL & " and FILNR = " & cFilnr
                gdBase.Execute cSQL, dbFailOnError
            End If
        End If
    Next lcount
    
    cSQL = "Delete from ETIDRU2 where ANZAHL <= 0"
    gdBase.Execute cSQL, dbFailOnError
    
    If InStr(UCase$(aDeviceName), "METO") > 0 Then
        DruckeBarCodeEtikett3 lDelete
        Exit Sub
    End If
    
    If InStr(UCase$(aDeviceName), "ZEBRA") > 0 Then
        DruckeBarCodeEtikett4 lDelete
        Exit Sub
    End If
    
    If InStr(UCase$(aDeviceName), "WAM") > 0 Then
        DruckeBarCodeEtikett5 lDelete
        Exit Sub
    End If
    
    cArtNrMerker = ""
    
'    cDatum = Right(Str$(Year(Now)), 2) & Trim$(Format$(Month(Now), "00")) & gcFilNr
    
    
    lReturn = OpenPrinter(aDeviceName, lhPrinter, 0)
    If lReturn = 0 Then
        MsgBox "Drucker " & aDeviceName & " nicht gefunden!", vbCritical, "STOP!"
        Exit Sub
    End If
    
    
    If Not NewTableSuchenDBKombi("SETELTRO", gdBase) Then
        MsgBox "Die Tabelle SETELTRO wurde nicht gefunden!", vbCritical, "STOP!"
        Exit Sub
    Else
    
        cSQL = "Select * from SETELTRO"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                cDBFeld = "Kommentar"
                If Not IsNull(rsrs!KOMMENTAR) Then
                    cKommentar = rsrs!KOMMENTAR
                Else
                    cKommentar = ""
                End If
                cDBFeld = "Befehl"""
                If Not IsNull(rsrs!BEFEHL) Then
                    cBefehl = rsrs!BEFEHL
                Else
                    cBefehl = ""
                End If
                
                
                Select Case cKommentar
                    Case Is = "Druckerschnittstelle"
                        cDruckBefehl(1) = cBefehl
                        
                    Case Is = "Name"
                        KonvertAnsiAscii cBefehl
                        cDruckBefehl(2) = cBefehl
                        
                    Case Is = "Ende der Textzeile"
                        cDruckBefehl(3) = cBefehl
                        
                    Case Is = "Ende der Numerischen Zeile"
                        cDruckBefehl(4) = cBefehl
                        
                    Case Is = "Position Firmenname"
                        cDruckBefehl(5) = cBefehl
                        
                    Case Is = "Position der Bestellnummer"
                        cDruckBefehl(6) = cBefehl
                        
                    Case Is = "Position Linie"
                        cDruckBefehl(7) = cBefehl
                        
                    Case Is = "Position Artikelbezeichnung"
                        cDruckBefehl(8) = cBefehl
                        
                    Case Is = "Position Barcode"
                        cDruckBefehl(9) = cBefehl
                        
                    Case Is = "Position Datum"
                        cDruckBefehl(10) = cBefehl
                        
                    Case Is = "Position DM"
                        cDruckBefehl(11) = cBefehl
                        
                    Case Is = "Position Euro"
                        cDruckBefehl(12) = cBefehl
                        
                    Case Is = "Name drucken (J/N)"
                        cDruckBefehl(13) = cBefehl
                        
                End Select
                        
                rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    MyDocInfo.pDocName = "Print BarCode"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    If lDoc = 0 Then
        MsgBox "StartDocPrinter ( " & Trim$(Str$(lhPrinter)) & " ) gescheitert!", vbCritical, "STOP!"
    End If
    
    lRet = StartPagePrinter(lhPrinter)
    If lRet = 0 Then
        MsgBox "StartPagePrinter ( " & Trim$(Str$(lhPrinter)) & " ) gescheitert!", vbCritical, "STOP!"
    End If
          
    cSQL = "Select * from ETIDRU2 where ANZAHL > 0"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!filnr) Then
                cFilnr = rsrs!filnr
            Else
                cFilnr = "0"
            End If
                        
            cDatum = Right(Str$(Year(Now)), 2) & Trim$(Format$(Month(Now), "00")) & cFilnr
                        
            cDBFeld = "ARTNR"
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            Else
                cArtNr = "-1"
            End If
            cArtNr = Trim$(cArtNr)
            cArtNr = String$(6 - Len(cArtNr), "0") & cArtNr
        
            cDBFeld = "BEZEICH"
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = rsrs!BEZEICH
            Else
                cBezeich = ""
            End If
            cBezeich = Trim$(cBezeich)
        
            cBezeich = fnEntferneLeerzeichen(cBezeich)
        
            cEAN = "2" & cArtNr
            For lPos = 1 To Len(cEAN)
                lWert = Val(Mid(cEAN, lPos, 1))
                If lPos / 2 = Fix(lPos / 2) Then
                    lPruef = lPruef + (lWert * 1)
                Else
                    lPruef = lPruef + (lWert * 3)
                End If
            Next lPos
            lWert = 10 - (lPruef Mod 10)
            If lWert = 10 Then
                lWert = 0
            End If
            cEAN = cEAN & Trim$(Str$(lWert))
        
            cDBFeld = "VKPR"
            If Not IsNull(rsrs!vkpr) Then
                dVkPr = rsrs!vkpr
            Else
                dVkPr = 0
            End If
            cVKPR = Format$(dVkPr, "#####0.00")
            cVKPR = Space$(9 - Len(cVKPR)) & cVKPR
            cVKPR = gcWaehrung & cVKPR
        
            dEuro = dVkPr
            cEuro = Format$(dEuro, "#####0.00")
            cEuro = Space$(9 - Len(cEuro)) & cEuro
            cEuro = "EUR" & cEuro
        
            cDBFeld = "ANZAHL"
            If Not IsNull(rsrs!ANZAHL) Then
                dAnzahl = rsrs!ANZAHL
            Else
                dAnzahl = 0
            End If
                    
            cDBFeld = "LIBESNR"
            If Not IsNull(rsrs!LIBESNR) Then
                cLiBesNr = rsrs!LIBESNR
            Else
                cLiBesNr = ""
            End If
            cLiBesNr = Trim$(cLiBesNr)
                
            'Init Barcode-Drucker
            aDat = "N" & Chr$(10)
            
            aDat = aDat & cDruckBefehl(11) & Chr$(10)
            
            
    
            'Firmenname oder Artikel-Bestell-Nummer drucken
            If cDruckBefehl(13) = "J" Then
                aDat = aDat & cDruckBefehl(5) & Chr$(34) & Trim$(cDruckBefehl(2)) & Chr$(34) & Chr$(10)
            Else
                aDat = aDat & cDruckBefehl(6) & Chr$(34) & cLiBesNr & Chr$(34) & Chr$(10)
            End If
                
            'Trennlinie drucken
            aDat = aDat & cDruckBefehl(7) + Chr$(10) + Chr$(10)
        
            'Produktname drucken
            aDat = aDat & cDruckBefehl(8) & Chr$(34) & cBezeich & Chr$(34) & Chr$(10)
        
            'Barcode drucken
            aDat = aDat & cDruckBefehl(9) & Chr$(34) & cEAN & Chr$(34) & Chr$(10)
               
            'DM drucken
            aDat = aDat & cDruckBefehl(11) & Chr$(34) & cVKPR & Chr$(34) & Chr$(10)
                
            'Datum drucken
            aDat = aDat & cDruckBefehl(10) & Chr$(34) & cDatum & Chr$(34) & Chr$(10)
                
            'Euro drucken
            aDat = aDat & cDruckBefehl(12) & Chr$(34) & cEuro & Chr$(34) & Chr$(10)
                
            'Anzahl Etiketten senden
            aDat = aDat & "P" & Format$(dAnzahl, "####0") & Chr$(10)
            lSize = Len(aDat)
            lReturn = 0
'            MsgBox aDat
            lReturn = WritePrinter(lhPrinter, ByVal aDat, lSize, lpcWritten)
            If lReturn = 0 Then
                MsgBox "WritePrinter (" & Trim$(Str$(lhPrinter)) & " / " & aDat & ") gescheitert!", vbCritical, "STOP!"
            End If
                
            If lDelete = vbYes Then
            
                SicherInEtisic cArtNr, cFilnr, Check13
            
                cSQL = "Delete from ETIDRU where ARTNR = " & cArtNr & " and FILNR = " & cFilnr & " "
                gdBase.Execute cSQL, dbFailOnError
                
                ZaehleEtikettenWKL30
    
                FuelleListeEtikettenWKL30
    
            End If
                
            rsrs.MoveNext
            
            Dim lStart As Long
            Dim lAktuell As Long
            
            lStart = Timer
            Do
                lAktuell = Timer
            Loop While lAktuell < lStart + (Fix(dAnzahl / 3)) + 2
        
            
        Loop
    
        
        'NACHLAUF
        'Init Barcode-Drucker
        aDat = "N" & Chr$(10)
        
        'Endetext drucken
        aDat = aDat & cDruckBefehl(8) & Chr$(34) & "ENDE ETIKETTENDRUCK" & Chr$(34) & Chr$(10)
        
        aDat = aDat & "P1" + Chr$(10)
    
        'Endetext drucken
        aDat = aDat & "N" & Chr$(10)
    
        aDat = aDat & "A270,24,0,2,2,1,N," + Chr$(34) + "############" + Chr$(34) + Chr$(10)
    
        aDat = aDat & "A270,64,0,2,2,1,N," + Chr$(34) + "############" + Chr$(34) + Chr$(10)
    
        aDat = aDat & "A270,104,0,2,2,1,N," + Chr$(34) + "############" + Chr$(34) + Chr$(10)

        aDat = aDat & "P1" + Chr$(10)
        lSize = Len(aDat)
        lReturn = 0
        lReturn = WritePrinter(lhPrinter, ByVal aDat, lSize, lpcWritten)
        If lReturn = 0 Then
            MsgBox "WritePrinter (" & Trim$(Str$(lhPrinter)) & " / " & aDat & ") gescheitert!", vbCritical, "STOP!"
        End If
    Else
        MsgBox "Keine Daten für Etikettendruck gefunden!", vbInformation, "INFO"
    
    End If
    
    lReturn = 0
    lReturn = EndPagePrinter(lhPrinter)
    If lReturn = 0 Then
        MsgBox "EndPagePrinter (" & Trim$(Str$(lhPrinter)) & ") gescheitert!", vbCritical, "STOP!"
    End If
    
    lReturn = 0
    lReturn = EndDocPrinter(lhPrinter)
    If lReturn = 0 Then
        MsgBox "EndDocPrinter (" & Trim$(Str$(lhPrinter)) & ") gescheitert!", vbCritical, "STOP!"
    End If
    
    lReturn = 0
    lReturn = ClosePrinter(lhPrinter)
    If lReturn = 0 Then
        MsgBox "ClosePrinter (" & Trim$(Str$(lhPrinter)) & ") gescheitert!", vbCritical, "STOP!"
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3010 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DruckeBarCodeEtikett2"
        Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub

Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    
    BlendeAlleFramesAus
    
    Select Case Index
        Case Is = 0
            
            Frame2(Index).Visible = True
'            Option1(0).SetFocus
            
        Case Is = 1
            
            Frame2(Index).Visible = True
            Text1(0).SetFocus
            
        Case Is = 2
            
            Frame2(Index).Visible = True
            Text1(6).SetFocus
            
        Case Is = 3
            
            Frame2(Index).Visible = True
            Text1(11).SetFocus
            
        Case Is = 4
            
            Frame2(Index).Visible = True
            Text1(15).SetFocus
            
        Case Is = 5
            
            Frame2(Index).Visible = True
            Text1(20).SetFocus
            
        Case Is = 6
            
            Frame2(Index).Visible = True
            Text1(25).SetFocus
            
        Case Is = 7
            
            Frame2(Index).Visible = True
            Text1(30).SetFocus
            
        Case Is = 8
            
            Frame2(Index).Visible = True
            Text1(35).SetFocus
            
        Case Is = 9
            SchreibeKonfigurationEtikettWKL30
            Command1(0).Visible = True
            Command1(1).Visible = True
            Command1(2).Visible = True
            Command1(3).Visible = True
            Command1(4).Visible = True
            Command4(1).Visible = True
            Frame3.Visible = True
            Frame4.Visible = True
            Frame1.Visible = False
            
        Case Is = 10
            iRet = MsgBox("Wollen Sie das Etikett auf die Standardwerte zurücksetzen?", vbYesNo + vbQuestion, "STANDARD-WERTE")
            If iRet = vbYes Then
                SchreibeStandardEtikettWKL30
                LeseKonfigurationEtikettWKL30
            End If
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    Select Case Index
    
        Case Is = 0
            For lcount = 0 To 53
                Command0(lcount).Caption = LCase$(Command0(lcount).Caption)
            Next lcount
        Case Is = 1
            If Trim(Label3.Caption) = "-1" Then
            
            Else
                If Trim(Label3.Caption) <> "" Then
                    If Text1(Label3.Caption).Text <> "" Then
                        Text1(Label3.Caption).Text = Left(Text1(Label3.Caption).Text, Len(Text1(Label3.Caption).Text) - 1)
                        Text1(Label3.Caption).SetFocus
                    End If
                End If
            End If
        Case Is = 2
            If Trim(Label3.Caption) = "-1" Then
            
            Else
                If Trim(Label3.Caption) <> "" Then
                    Text1(Label3.Caption).Text = ""
                    Text1(Label3.Caption).SetFocus
                End If
            End If
        Case Is = 3
            For lcount = 0 To 53
                Command0(lcount).Caption = UCase$(Command0(lcount).Caption)
            Next lcount
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten." & Index
    
    Fehlermeldung1
End Sub
Private Sub Command4_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Dim lCount2           As Long
    Dim lAnzahlEtikett    As Long
    Dim lAnzahl           As Long
    Dim lcount            As Long
    Dim counter           As Long
    
    Dim ctmp              As String
    Dim cLBSatz           As String
    Dim cArtNr            As String
    Dim cFilnr            As String
    Dim cAnzahl           As String
    Dim cAnzEti           As String
    Dim cSQL              As String
    ReDim acArtNr(0 To 0) As String
    Dim i                 As Integer
    Dim iRet              As Integer
    Dim bFound            As Boolean
    Dim bVorhanden        As Boolean
    Dim rsrs              As Recordset
    
    Select Case Index
    
        '***********
        Case Is = 0
'            AktionAustragen "Etiketten drucken"
            If IsAktionZulaessig("Etiketten drucken") = False Then
                Exit Sub
            End If
        
            bFound = False
            
            cSQL = "update ETIDRU set FILNR = 0 where FILNR = NULL"
            gdBase.Execute cSQL, dbFailOnError
            
            For lcount = 0 To List2.ListCount - 1
                If List2.Selected(lcount) = True Then
                        bFound = True
                End If
            Next lcount
            
            
            If bFound Then
            
                iRet = MsgBox("Wollen Sie nur die markierten Etiketten löschen ?", vbYesNoCancel + vbQuestion, "STOP!")
                    If iRet = vbCancel Then
                        Screen.MousePointer = 0
                        Exit Sub
                    ElseIf iRet = vbYes Then
                        
                        pbrZeit.Max = 50
                        pbrZeit.Visible = True
                        
                        For lcount = 0 To List2.ListCount - 1
                        
                            If counter = 50 Then
                            counter = 0
                            End If
                            
                            counter = counter + 1
                            
                            pbrZeit.Value = counter
                            
                            If List2.Selected(lcount) = True Then
                                cLBSatz = List2.list(lcount)
                                cArtNr = Left(cLBSatz, 6)
                                cArtNr = Trim$(cArtNr)
                                cFilnr = Mid(cLBSatz, 70, 2)
                                cFilnr = Trim$(cFilnr)
'                                cFilnr = Trim$(Right(cLBSatz, 3))
                                cAnzEti = Mid(cLBSatz, Len(cLBSatz) - 18, 10)
                                cAnzEti = Trim$(cAnzEti)
                                
                                SicherInEtisic cArtNr, cFilnr, Check13
                                
                                If gsETILS <> "" Then
                                    cSQL = "Delete from ETIDRULS where ARTNR = " & cArtNr & " And FILNR = " & cFilnr & " "
                                    cSQL = cSQL & " and Anzahl = " & cAnzEti
                                    gdBase.Execute cSQL, dbFailOnError
                                    
                                    cSQL = "Delete from LSTEETI where ARTNR = " & cArtNr & " And FILNR = " & cFilnr & " "
                                    cSQL = cSQL & " and Anzahl = " & cAnzEti
                                    gdBase.Execute cSQL, dbFailOnError
                                Else
                                    cSQL = "Delete from ETIDRU where ARTNR = " & cArtNr
                                    cSQL = cSQL & " and FILNR = " & cFilnr
                                    gdBase.Execute cSQL, dbFailOnError
                                End If
                                
                                
                            End If
                        Next lcount
'                        etidrukomp
                        pbrZeit.Visible = False
                        
                    Else
                        iRet = MsgBox("Wollen Sie wirklich alle Etiketten löschen ?", vbYesNoCancel + vbQuestion, "Winkiss Frage:")
                        If iRet = vbCancel Then
                            Screen.MousePointer = 0
                            Exit Sub
                        ElseIf iRet = vbYes Then
                        
                            SicherInEtisicALL Check13
                        
                            If gsETILS <> "" Then
                                cSQL = "Delete from lsteeti"
                                gdBase.Execute cSQL, dbFailOnError
                            Else
                                cSQL = "Delete from ETIDRU"
                                gdBase.Execute cSQL, dbFailOnError
                            End If

                        End If
                        
                    End If
            Else
            
                iRet = MsgBox("Wollen Sie wirklich alle Etiketten löschen ?", vbYesNoCancel + vbQuestion, "Winkiss Frage:")
                If iRet = vbCancel Then
                    Screen.MousePointer = 0
                    Exit Sub
                ElseIf iRet = vbYes Then
                
                    SicherInEtisicALL Check13
                    
                    If gsETILS <> "" Then
                        cSQL = "Delete from lsteeti"
                        gdBase.Execute cSQL, dbFailOnError
                    Else
                        cSQL = "Delete from ETIDRU"
                        gdBase.Execute cSQL, dbFailOnError
                    End If
                    
                End If
                    
            End If
                      
            ZaehleEtikettenWKL30
            FuelleListeEtikettenWKL30
            
            AktionAustragen "Etiketten drucken"
    
        
        Case Is = 1     '** Drucke Grundpreis-Etikett **
        
            If IsAktionZulaessig("Etiketten drucken") = False Then
                Exit Sub
            End If
            
            voreinstellungspeichern30
            
            bFound = False
            lAnzahl = -1
            For lcount = 0 To List2.ListCount - 1
                If List2.Selected(lcount) = True Then
                    bFound = True
                    lAnzahl = lAnzahl + 1
                    ReDim Preserve acArtNr(0 To lAnzahl) As String
                    cLBSatz = List2.list(lcount)
                    cArtNr = Left(cLBSatz, 6)
                    cArtNr = Trim$(cArtNr)
                    acArtNr(lAnzahl) = cArtNr
                End If
            Next lcount
            
            If Not bFound Then
                MsgBox "Zum Drucken bitte mindestens einen Listeneintrag auswählen!", vbCritical, "STOP!"
                List2.SetFocus
            Else
            
                If Modul6.FindFile(App.Path, "aWOKIBR.rpt") And Modul6.FindFile(App.Path, "aWOKINE.rpt") Then
                    iRet = MsgBox("Möchten Sie Netto - Etiketten drucken?", vbQuestion + vbYesNo, "Winkiss Frage")
                    If iRet = vbYes Then
                        DruckeGrundPreisEtikettenWKL30Jebe acArtNr(), lAnzahl, "NETTO"
                    Else
                        DruckeGrundPreisEtikettenWKL30Jebe acArtNr(), lAnzahl, "BRUTTO"
                    End If
                Else
                    If Modul6.FindFile(App.Path, "aWKL30xs.rpt") Then
                        DruckeGrundPreisEtikettenWKL30kleinspezial acArtNr(), lAnzahl, "aWKL30xs"
                    Else
                        If Modul6.FindFile(gcDBPfad, "aWKL30ls.rpt") Then
           
                            DruckeGrundPreisEtikettenLS acArtNr(), lAnzahl
                        Else
                            DruckeGrundPreisEtikettenWKL30 acArtNr(), lAnzahl '*hier Standard*hier Standard*hier Standard
                        End If
                        
                    End If
                End If
                

                '** Regalletikett : gedrückte Anzahl immer 1 **
                Erase acArtNr
            End If
            
            AktionAustragen "Etiketten drucken"
                       
        Case Is = 2     '** Markiere alle **
            List2.Visible = False
            For lcount = List2.ListCount - 1 To 0 Step -1
                List2.Selected(lcount) = True
            Next lcount
            List2.Visible = True
            List2.SetFocus
            
            Label4(6).Caption = Label4(2).Caption
            Label4(9).Caption = Label4(3).Caption
            
            Label4(6).Caption = SwapStr(Label4(6).Caption, ".", "")
            Label4(9).Caption = SwapStr(Label4(9).Caption, ".", "")
            
            Label4(7).Caption = "0"
            Label4(8).Caption = "0"
    
        Case Is = 3     '** Exportiere markierte Sätze **
            ExportiereMarkierteSaetzeWKL30
        Case Is = 4
            
            Screen.MousePointer = 0
            frmWKL182.Show 1
            
            ZaehleEtikettenWKL30
            FuelleListeEtikettenWKL30
        
    End Select
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim ctmp As Long
    Dim lcount As Long
    ReDim acArtNr(0 To 0) As String
    ReDim acAnzEti(0 To 0) As String
    Dim lAnzahl As Long
    Dim cLBSatz As String
    Dim cArtNr As String
    Dim cAnzEti As String

    lAnzahl = -1
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            lAnzahl = lAnzahl + 1
            ReDim Preserve acArtNr(0 To lAnzahl) As String
            ReDim Preserve acAnzEti(0 To lAnzahl) As String
            cLBSatz = List2.list(lcount)
            cArtNr = Left(cLBSatz, 6)
            cArtNr = Trim$(cArtNr)
            acArtNr(lAnzahl) = cArtNr
            cAnzEti = Trim$(Right(cLBSatz, 15))
            ctmp = Left(cAnzEti, 2)
            ctmp = Trim$(ctmp)
            cAnzEti = Val(ctmp)
            acAnzEti(lAnzahl) = cAnzEti
        End If
    Next lcount
    
    merke_endlos Index, "STRICH"
    ZeigeFokus_Endlos "STRICH"

    Select Case Index
        Case 0  'Schmucketikett 69x14 Variante 1
            DruckeSchmucketikett69x14Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL311a", gcEtikettenDrucker, True
            Erase acArtNr
                                
        Case 1  'Schmucketikett 69x14 Variante 2
            DruckeSchmucketikett69x14Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL311b", gcEtikettenDrucker, True
            Erase acArtNr
            
        Case 2  'Etikett 40x18 Variante 1
            DruckeEtikett40x18Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL312a", gcEtikettenDrucker, True
            Erase acArtNr
        
        Case 3  'Etikett 40x18 Variante 2
            DruckeEtikett40x18Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL312b", gcEtikettenDrucker, True
            Erase acArtNr
        Case 11  'Etikett 40x18 Variante 3
            DruckeEtikett40x18Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL312c", gcEtikettenDrucker, True
            Erase acArtNr
        Case 20  'Etikett 40x18 Variante 4
            DruckeEtikett40x18Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL312d", gcEtikettenDrucker, True
            Erase acArtNr
        Case 5  'Etikett 45x23 Variante 1
            DruckeEtikett45x23Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL313a", gcEtikettenDrucker, True
            Erase acArtNr
        Case 4  'Etikett 45x23 Variante 2
            DruckeEtikett45x23Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL313b", gcEtikettenDrucker, True
            Erase acArtNr
        Case 6  'Schmucketikett 69x14 Variante 3
            DruckeSchmucketikett69x14Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL311c", gcEtikettenDrucker, True
            Erase acArtNr
        Case 7  'Etikett 45x23 Variante 3
            DruckeEtikett45x23Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL313c", gcEtikettenDrucker, True
            Erase acArtNr
        Case 9  'Etikett 38x23 Variante 1
            DruckeEtikett45x23Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL314a", gcEtikettenDrucker, True
            Erase acArtNr
        Case 10  'Etikett 38x23 Variante 2
            DruckeEtikett45x23Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL314b", gcEtikettenDrucker, True
            Erase acArtNr
        Case 8  'Etikett 38x23 Variante 3
            DruckeEtikett45x23Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL314c", gcEtikettenDrucker, True
            Erase acArtNr
        Case 27  'Etikett 38x23 Variante 4
            DruckeEtikett45x23Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL314d", gcEtikettenDrucker, True
            Erase acArtNr
        Case 13  'Etikett 51x19 Variante 1
            DruckeEtikett51x19Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL315a", gcEtikettenDrucker, True
            Erase acArtNr
        Case 12  'Etikett 51x19 Variante 2
            DruckeEtikett51x19Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL315b", gcEtikettenDrucker, True
            Erase acArtNr
        Case 14  'Etikett 49x19 Variante 1
            DruckeEtikett49x19Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL316a", gcEtikettenDrucker, True
            Erase acArtNr
        Case 15  'Etikett 44x21 Variante 1
            DruckeEtikett44x21Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL317a", gcEtikettenDrucker, True
            Erase acArtNr
        Case 26  'Etikett 44x21 Variante 2
            DruckeEtikett44x21Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL317b", gcEtikettenDrucker, True
            Erase acArtNr
        Case 16  'Etikett 51x19 Variante 3
            DruckeEtikett51x19Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL315c", gcEtikettenDrucker, True
            Erase acArtNr
            
        Case 19  'Etikett 30x15 Variante 1
            DruckeEtikett45x23Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL3015a", gcEtikettenDrucker, True
            Erase acArtNr
        Case 18  'Etikett 30x15 Variante 2
            DruckeEtikett45x23Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL3015b", gcEtikettenDrucker, True
            Erase acArtNr
        Case 17  'Etikett 30x15 Variante 3
            DruckeEtikett45x23Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL3015c", gcEtikettenDrucker, True
            Erase acArtNr
            
        Case 21  'Etikett 48x18 Variante 1
            DruckeEtikett48x18Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL319a", gcEtikettenDrucker, True
            Erase acArtNr
            
        Case 22  'Etikett 45x23 Variante 4
            DruckeEtikett45x23Variante1 acArtNr(), lAnzahl, acAnzEti()
            reportbildschirmToPrinterETI "aWKL313d", gcEtikettenDrucker, True
            Erase acArtNr
            
        Case 23  'Etikett 40x18 Variante 5
            DruckeEtikett40x18Variante5 acArtNr(), lAnzahl, acAnzEti()
'            reportbildschirmToPrinterETI "aWKL312a", gcEtikettenDrucker, True

            reportbildschirmToPrinterETI "aWKL312e", gcEtikettenDrucker, True
            Erase acArtNr
        Case 24  'Etikett 40x18 Variante 6
            DruckeEtikett40x18Variante1 acArtNr(), lAnzahl, acAnzEti()
'            reportbildschirmToPrinterETI "aWKL312a", gcEtikettenDrucker, True
            reportbildschirmToPrinterETI "aWKL312f", gcEtikettenDrucker, True
            Erase acArtNr
        Case 25 'Etikett 35x15 Variante 1
            DruckeEtikett35x15Variante1 acArtNr(), lAnzahl, acAnzEti()
'            reportbildschirmToPrinterETI "aWKL316a", gcEtikettenDrucker, True
            reportbildschirmToPrinterETI "aWKL322a", gcEtikettenDrucker, True
            Erase acArtNr
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ZeigeFokus_Endlos(sArt As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim iIndex As Integer
    Dim i As Integer
    
    iIndex = -1
    sSQL = "select EtiIndex from FOKUSENDLOS where Art = '" & sArt & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!EtiIndex) Then
            iIndex = rsrs!EtiIndex
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

    If sArt = "REGAL" And iIndex > -1 Then
        For i = 0 To 20
            Command6(i).ForeColor = glButtonForecolor
        Next i
        
        Command6(iIndex).ForeColor = glWarn
    End If
    
    If sArt = "STRICH" And iIndex > -1 Then
        For i = 0 To 26
            Command5(i).ForeColor = glButtonForecolor
        Next i
        
        Command5(iIndex).ForeColor = glWarn
    End If
         
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeFokus_Endlos"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub merke_endlos(Index As Integer, sArt As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    sSQL = "Delete from FOKUSENDLOS where Art = '" & sArt & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into FOKUSENDLOS (EtiIndex,Art) values (" & Index & ",'" & sArt & "')"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "merke_endlos"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command6_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim ctmp As Long
    Dim lcount As Long
    ReDim acArtNr(0 To 0) As String
    ReDim acAnzEti(0 To 0) As String
    Dim lAnzahl As Long
    Dim cLBSatz As String
    Dim cArtNr As String
    Dim cAnzEti As String

    lAnzahl = -1
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            lAnzahl = lAnzahl + 1
            ReDim Preserve acArtNr(0 To lAnzahl) As String
            cLBSatz = List2.list(lcount)
            cArtNr = Left(cLBSatz, 6)
            cArtNr = Trim$(cArtNr)
            acArtNr(lAnzahl) = cArtNr
        End If
    Next lcount
    
    merke_endlos Index, "REGAL"
    ZeigeFokus_Endlos "REGAL"
            
    Select Case Index
        Case 0  'TLP2844 etikett 40x25 Variante 1
            DruckeTLPRegaletikett40x25Variante1 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
                
            reportbildschirmToPrinterETI "aWKL311i", gcEtikettenDrucker, True
            Erase acArtNr
                                
        Case 1  'TLP2844 etikett 40x25 Variante 2
            DruckeTLPRegaletikett40x25Variante1 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311j", gcEtikettenDrucker, True
            Erase acArtNr
            
        Case 2  'TLP2844 etikett 70x35 Variante 1
            DruckeTLPRegaletikett40x25Variante1 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311k", gcEtikettenDrucker, True
            Erase acArtNr
                                
        Case 3  'TLP2844 etikett 70x35 Variante 2
            DruckeTLPRegaletikett40x25Variante1 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311l", gcEtikettenDrucker, True
            Erase acArtNr
            
        Case 5  'TLP2844 etikett 40x18 Variante 1
            DruckeTLPRegaletikett40x25Variante1 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311m", gcEtikettenDrucker, True
            Erase acArtNr
                                
        Case 4  'TLP2844 etikett 40x18 Variante 2
            DruckeTLPRegaletikett40x25Variante1 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311n", gcEtikettenDrucker, True
            Erase acArtNr
        Case 6  'TLP2844 etikett 70x35 Variante 3 - Achtung für Haarmarkt Wagner
            spezial_DruckeTLPRegaletikett70x35Variante1 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311w", gcEtikettenDrucker, True
            Erase acArtNr
            
        Case 7  'TLP2844 etikett 40x25 Variante 3
            DruckeTLPRegaletikett40x25Variante1 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311o", gcEtikettenDrucker, True
            Erase acArtNr
        Case 8  'TLP2844 etikett 40x25 Variante 4
            DruckeTLPRegaletikett40x25Variante4 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311p", gcEtikettenDrucker, True
            Erase acArtNr
            
        Case 9  'TLP2844 etikett 40x25 Variante 5
            DruckeTLPRegaletikett40x25Variante5 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311q", gcEtikettenDrucker, True
            Erase acArtNr
            
        Case 10  'TLP2844 etikett 40x25 Variante 6
            DruckeTLPRegaletikett40x25Variante1 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311r", gcEtikettenDrucker, True
            Erase acArtNr
        Case 11  'EDEKA TLP2844 etikett 40x25 Variante 7
            DruckeTLPRegaletikett40x25VarianteEdeka acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311s", gcEtikettenDrucker, True
            Erase acArtNr
        Case 12  'TLP2844 etikett 50x37 Variante 1
            DruckeTLPRegaletikett50x37Variante1 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL318a", gcEtikettenDrucker, True
            Erase acArtNr
        Case 13  'TLP2844 etikett 50x40 Variante 1
            DruckeTLPRegaletikett50x40Variante1 acArtNr(), lAnzahl, CInt(txtTage.Text)
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL318b", gcEtikettenDrucker, True
            Erase acArtNr
        Case 14  'TLP2844 etikett 50x40 Variante 2
            DruckeTLPRegaletikett50x40Variante2 acArtNr(), lAnzahl, CInt(txtTage.Text), Label13
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL318c", gcEtikettenDrucker, True
        
            Erase acArtNr
               
        Case 15  'TLP2844 etikett 40x25 Variante 7
            DruckeTLPRegaletikett40x25Var_Dronova acArtNr(), lAnzahl
            
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311t", gcEtikettenDrucker, True
            Erase acArtNr
            
        Case 16  'TLP2844 etikett 69x38 Variante 1
            DruckeTLPRegaletikett69x38Var_Kombi acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311u", gcEtikettenDrucker, True
            Erase acArtNr
            
            
        Case 17  'TLP2844 etikett 70x35 Variante 4
            DruckeTLPRegaletikett69x38Var4 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311v", gcEtikettenDrucker, True 'awkl311l
            Erase acArtNr
            
        Case 18  'TLP2844 etikett 40x18 Variante 3
            DruckeTLPRegaletikett40x25Variante1 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311d", gcEtikettenDrucker, True
            Erase acArtNr
        Case 19  'TLP2844 etikett 49x36 Variante 1
            DruckeTLPRegaletikett49x36Variante1 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL318d", gcEtikettenDrucker, True

            Erase acArtNr
            
        Case 20  'TLP2844 etikett 40x25 Variante 7
            DruckeTLPRegaletikett40x25Variante7 acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL311q", gcEtikettenDrucker, True
            Erase acArtNr
            
        Case 21  'TLP2844 etikett 81x38 Variante 1
            DruckeTLPRegaletikett81x38Var_Kombi acArtNr(), lAnzahl
            If gsETILS <> "" Then
                Update_Preis_Terminpreis acArtNr(), lAnzahl
            End If
            
            reportbildschirmToPrinterETI "aWKL312u", gcEtikettenDrucker, True
            Erase acArtNr
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim cDrucker As String
    Dim bReturn As Boolean
    Dim lAnz As Long
    Dim lcount As Long
    Dim iFileNr As Integer
    Dim ctmp As String
    Dim sSQL As String
    
    PositionierenWKL30
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Label1(0)
    
    If NewTableSuchenDBKombi("FOKUSENDLOS", gdBase) = False Then
        CreateTableT2 "FOKUSENDLOS", gdBase
    End If
    
    ermvorzNettospannen
    
    cDrucker = gcEtikettenDrucker
    bReturn = SetDefaultPrinter(cDrucker)
    If bReturn = False Then
        MsgBox "Etiketten-Drucker konnte nicht initialisiert werden!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    lAnz = Printers.Count
    For lcount = 0 To lAnz - 1
        If Printers(lcount).DeviceName = gcEtikettenDrucker Then
            Set Printer = Printers(lcount)
            Exit For
        End If
    Next lcount
    
    If InStr(UCase$(cDrucker), "ELTRON") = 0 Then
        Command1(1).Visible = False
    Else
        Command1(1).Visible = True
    End If
    
    If gsETILS <> "" Then
    
    End If
    
    ZaehleEtikettenWKL30

    iFileNr = FreeFile
    Open gcPfad & "ETIDRU.CFG" For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        ctmp = Space$(LOF(iFileNr))
        Get #iFileNr, 1, ctmp
        Close iFileNr
    Else
        Close iFileNr
        Kill gcPfad & "ETIDRU.CFG"
        ctmp = "2"
    End If
    
    Select Case ctmp
        Case Is = 0
            'nicht möglich!
        Case Is = 1
            'nicht möglich
        Case Is = 2
            Option3(0).Value = True
            cboRegalDinA4.Text = "selbstklebend 52,5 x 30"
        
        Case Is = 3
            Option3(1).Value = True
            cboRegalDinA4.Text = "selbstklebend 52,5 x 30"
        Case Is = 4
            Option3(0).Value = True
            cboRegalDinA4.Text = "selbstklebend 70 x 36"
        Case Is = 5
            Option3(1).Value = True
            cboRegalDinA4.Text = "selbstklebend 70 x 36"
    End Select
    
    cboStrichDINA4.AddItem "50 x 36"
    cboStrichDINA4.AddItem "45,7 x 21,2"
    cboStrichDINA4.AddItem "35,6 x 16,9"
    cboStrichDINA4.AddItem "35,6 x 16,9 Variante 2"
    cboStrichDINA4.AddItem "35,6 x 16,9 Variante 4"
    cboStrichDINA4.AddItem "35,6 x 16,9 Variante 5"
    cboStrichDINA4.AddItem "48,5 x 25,4"
    cboStrichDINA4.AddItem "Sonder Etikett"
    cboStrichDINA4.AddItem "45,7 x 21,2 Variante 2"
    cboStrichDINA4.AddItem "52,5 x 30"
    cboStrichDINA4.AddItem "52,5 x 21,2"
    cboStrichDINA4.AddItem "52,5 x 21,2 Variante 2"
    cboStrichDINA4.AddItem "50 x 21,09"
    cboStrichDINA4.AddItem "25,4 x 12,8"
    cboStrichDINA4.Text = "bitte auswählen"
    
    If Not tableSuchenDBKombi("VOREINAP", 2) Then
        sSQL = "Create table VOREINAP ( Schluessel Text(30),Wert Text(30) )"
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    Label1(5).Visible = False
    
    If LeseVoreinap("50 x 36") = True Then
        cboStrichDINA4.Text = "50 x 36"
    ElseIf LeseVoreinap("45,7 x 21,2") = True Then
        cboStrichDINA4.Text = "45,7 x 21,2"
    ElseIf LeseVoreinap("35,6 x 16,9") = True Then
        Label1(5).Visible = True
        cboStrichDINA4.Text = "35,6 x 16,9"
    ElseIf LeseVoreinap("35,6 x 16,9 Variante 2") = True Then
        Label1(5).Visible = True
        cboStrichDINA4.Text = "35,6 x 16,9 Variante 2"
    ElseIf LeseVoreinap("35,6 x 16,9 Variante 4") = True Then
        Label1(5).Visible = True
        cboStrichDINA4.Text = "35,6 x 16,9 Variante 4"
    ElseIf LeseVoreinap("35,6 x 16,9 Variante 5") = True Then
        Label1(5).Visible = True
        cboStrichDINA4.Text = "35,6 x 16,9 Variante 5"
    ElseIf LeseVoreinap("48,5 x 25,4") = True Then
        cboStrichDINA4.Text = "48,5 x 25,4"
    ElseIf LeseVoreinap("Sonder Etikett") = True Then
        cboStrichDINA4.Text = "Sonder Etikett"
    ElseIf LeseVoreinap("45,7 x 21,2 Variante 2") = True Then
        cboStrichDINA4.Text = "45,7 x 21,2 Variante 2"
    ElseIf LeseVoreinap("52,5 x 30") = True Then
        cboStrichDINA4.Text = "52,5 x 30"
    ElseIf LeseVoreinap("52,5 x 21,2") = True Then
        cboStrichDINA4.Text = "52,5 x 21,2"
    ElseIf LeseVoreinap("52,5 x 21,2 Variante 2") = True Then
        cboStrichDINA4.Text = "52,5 x 21,2 Variante 2"
    ElseIf LeseVoreinap("50 x 21,09") = True Then
        cboStrichDINA4.Text = "50 x 21,09"
    ElseIf LeseVoreinap("25,4 x 12,8") = True Then
        cboStrichDINA4.Text = "25,4 x 12,8"
    End If
    
    
    
    
    'neu für Regal DinA4
    cboRegalDinA4.AddItem "selbstklebend 70 x 36"
    cboRegalDinA4.AddItem "selbstklebend 52,5 x 30"
    cboRegalDinA4.AddItem "selbstklebend 35,6 x 16,9"
    cboRegalDinA4.AddItem "selbstklebend Sonder Etikett"
    cboRegalDinA4.AddItem "perforiert 38 x 70"
    cboRegalDinA4.AddItem "perforiert 26 x 47"
    cboRegalDinA4.AddItem "perforiert 26 x 47 (EAN 13)"
    cboRegalDinA4.AddItem "perforiert 38 x 42 (EAN 13)"
    cboRegalDinA4.AddItem "perforiert 22 x 22"
    cboRegalDinA4.AddItem "perforiert 26 x 45 (EAN 13)"
    cboRegalDinA4.AddItem "perforiert 26 x 45 (EAN 13) V2"
    cboRegalDinA4.AddItem "perforiert 38 x 50"
    cboRegalDinA4.AddItem "perforiert 38 x 50 Variante 2"
    cboRegalDinA4.AddItem "perforiert 29 x 52"
    cboRegalDinA4.AddItem "perforiert 29 x 52 Variante 2"
    cboRegalDinA4.AddItem "perforiert 26 x 35"
    cboRegalDinA4.AddItem "perforiert 39 x 35"
    cboRegalDinA4.AddItem "perforiert 39 x 33"
    cboRegalDinA4.AddItem "perforiert 42 x 25"
    cboRegalDinA4.AddItem "perforiert 39 x 33 (EAN)"
    cboRegalDinA4.AddItem "perforiert 50 x 40"
    cboRegalDinA4.AddItem "perforiert 26 x 47 (EAN 13) W"
    cboRegalDinA4.AddItem "35,6 x 16,9 Variante 2"
    
    cboRegalDinA4.Text = "bitte auswählen"
    
    If LeseVoreinap("selbstklebend 70 x 36") = True Then
        cboRegalDinA4.Text = "selbstklebend 70 x 36"
    ElseIf LeseVoreinap("selbstklebend 35,6 x 16,9") = True Then
        cboRegalDinA4.Text = "selbstklebend 35,6 x 16,9"
    ElseIf LeseVoreinap("selbstklebend 52,5 x 30") = True Then
        cboRegalDinA4.Text = "selbstklebend 52,5 x 30"
    ElseIf LeseVoreinap("selbstklebend Sonder Etikett") = True Then
        cboRegalDinA4.Text = "selbstklebend Sonder Etikett"
    ElseIf LeseVoreinap("perforiert 38 x 70") = True Then
        cboRegalDinA4.Text = "perforiert 38 x 70"
    ElseIf LeseVoreinap("perforiert 26 x 47") = True Then
        cboRegalDinA4.Text = "perforiert 26 x 47"
    ElseIf LeseVoreinap("perforiert 26 x 47 (EAN 13)") = True Then
        cboRegalDinA4.Text = "perforiert 26 x 47 (EAN 13)"
    ElseIf LeseVoreinap("perforiert 38 x 42 (EAN 13)") = True Then
        cboRegalDinA4.Text = "perforiert 38 x 42 (EAN 13)"
    ElseIf LeseVoreinap("perforiert 22 x 22") = True Then
        cboRegalDinA4.Text = "perforiert 22 x 22"
    ElseIf LeseVoreinap("perforiert 26 x 45 (EAN 13)") = True Then
        cboRegalDinA4.Text = "perforiert 26 x 45 (EAN 13)"
    ElseIf LeseVoreinap("perforiert 26 x 45 (EAN 13) V2") = True Then
        cboRegalDinA4.Text = "perforiert 26 x 45 (EAN 13) V2"
    ElseIf LeseVoreinap("perforiert 38 x 50") = True Then
        cboRegalDinA4.Text = "perforiert 38 x 50"
    ElseIf LeseVoreinap("perforiert 38 x 50 Variante 2") = True Then
        cboRegalDinA4.Text = "perforiert 38 x 50 Variante 2"
    ElseIf LeseVoreinap("perforiert 29 x 52") = True Then
        cboRegalDinA4.Text = "perforiert 29 x 52"
    ElseIf LeseVoreinap("perforiert 29 x 52 Variante 2") = True Then
        cboRegalDinA4.Text = "perforiert 29 x 52 Variante 2"
    ElseIf LeseVoreinap("perforiert 26 x 35") = True Then
        cboRegalDinA4.Text = "perforiert 26 x 35"
    ElseIf LeseVoreinap("perforiert 39 x 35") = True Then
        cboRegalDinA4.Text = "perforiert 39 x 35"
    ElseIf LeseVoreinap("perforiert 39 x 33") = True Then
        cboRegalDinA4.Text = "perforiert 39 x 33"
    ElseIf LeseVoreinap("perforiert 42 x 25") = True Then
        cboRegalDinA4.Text = "perforiert 42 x 25"
    ElseIf LeseVoreinap("perforiert 39 x 33 (EAN)") = True Then
        cboRegalDinA4.Text = "perforiert 39 x 33 (EAN)"
    ElseIf LeseVoreinap("35,6 x 16,9 Variante 2") = True Then
        cboRegalDinA4.Text = "35,6 x 16,9 Variante 2"
    ElseIf LeseVoreinap("perforiert 50 x 40") = True Then
        cboRegalDinA4.Text = "perforiert 50 x 40"
    ElseIf LeseVoreinap("perforiert 26 x 47 (EAN 13) W") = True Then
        cboRegalDinA4.Text = "perforiert 26 x 47 (EAN 13) W"
    End If
    
    'Ende
    
    If tableSuchenDBKombi("VOREINAP", 2) Then
    
        Option3(0).Value = LeseVoreinap("Tintenstrahl")
        Option3(1).Value = LeseVoreinap("Laser")
        
    End If


    Text2.Text = "0"

    List1.Clear
    List2.Clear
    List1.AddItem "ArtNr. Artikelbezeichnung                     VK-Preis Anz.Etiketten Fil"
    
    FuelleListeEtikettenWKL30
    
    Command1(0).Enabled = True
    Command1(2).Enabled = True
    Command1(3).Enabled = True
    Command1(4).Enabled = True
    Command4(1).Enabled = True
    
    Frame4.Enabled = True
    
    If Val(gsEdeka) > 0 Then
        Label7(24).Visible = True
        Command6(11).Visible = True
    End If
    
    If NewTableSuchenDBKombi("E30", gdBase) Then
        If SpalteInTabellegefundenNEW("E30", "VKTAGE", gdBase) = False Then
            SpalteAnfuegenNEW "E30", "VKTAGE", "integer", gdBase
        End If
        voreinstellungladenWKL30
    End If
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub voreinstellungspeichern30()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    Dim lLinr1 As Long
    Dim lLinr2 As Long
    Dim lLinr3 As Long
    Dim lVKTAGE As Long

    loeschNEW "E30", gdBase
    CreateTableT2 "E30", gdBase

    lLinr1 = Val(Label1(3).Caption)
    lLinr2 = Val(Label1(1).Caption)
    lLinr3 = Val(Label1(2).Caption)
    lVKTAGE = Val(txtTage.Text)
    
    sSQL = "Insert into E30 ( LIEF_1,LIEF_2,LIEF_3,VKTAGE) "
    sSQL = sSQL & " values (" & lLinr1 & "," & lLinr2 & "," & lLinr3 & "," & lVKTAGE & ")"
    gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern30"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungladenWKL30()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    
    Set rs = gdBase.OpenRecordset("E30")
    If Not rs.EOF Then
    
        If Not IsNull(rs!LIEF_1) Then
            Label1(3).Caption = rs!LIEF_1
            Label16(3).Caption = ermLiefBez(Label1(3).Caption)
        Else
            Label1(3).Caption = "undefiniert"
            Label16(3).Caption = ""
        End If
        
        If Not IsNull(rs!LIEF_2) Then
            Label1(1).Caption = rs!LIEF_2
            Label16(1).Caption = ermLiefBez(Label1(1).Caption)
        Else
            Label1(1).Caption = "undefiniert"
            Label16(1).Caption = ""
        End If
        
        If Not IsNull(rs!LIEF_3) Then
            Label1(2).Caption = rs!LIEF_3
            Label16(2).Caption = ermLiefBez(Label1(2).Caption)
        Else
            Label1(2).Caption = "undefiniert"
            Label16(2).Caption = ""
        End If
        
        txtTage.Text = "60"
        If Not IsNull(rs!VKTAGE) Then
            txtTage.Text = rs!VKTAGE
        End If
    
    End If
    rs.Close: Set rs = Nothing
    
    If Label1(3).Caption = "0" Then Label1(3).Caption = "undefiniert"
    If Label1(1).Caption = "0" Then Label1(1).Caption = "undefiniert"
    If Label1(2).Caption = "0" Then Label1(2).Caption = "undefiniert"

     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladenWKL30"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    LogtoEnd Me
    
    Dim sSQL As String
    
    If gsETILS <> "" Then
        'Rabatt_OK
        If SpalteInTabellegefundenNEW("LSTEETI", "Rabatt_OK", gdBase) = True Then
            sSQL = "Alter table LSTEETI drop column Rabatt_OK "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        If SpalteInTabellegefundenNEW("LSTEETI", "RKZ", gdBase) = True Then
            sSQL = "Alter table LSTEETI drop column RKZ "
            gdBase.Execute sSQL, dbFailOnError
        End If
    Else
        If SpalteInTabellegefundenNEW("ETIDRU", "Rabatt_OK", gdBase) = True Then
            sSQL = "Alter table ETIDRU drop column Rabatt_OK "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        If SpalteInTabellegefundenNEW("ETIDRU", "RKZ", gdBase) = True Then
            sSQL = "Alter table ETIDRU drop column RKZ "
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    
'    AktionAustragen "Etiketten drucken"
    
    gsETILS = "" 'Etiketten aus Lieferscheinen auf Null setzen
    setzedrucker gcListenDrucker
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Frame10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(4).ForeColor = glS1
    Label1(8).ForeColor = glS1
    Label1(1).ForeColor = glS1
    Label1(2).ForeColor = glS1
    Label1(3).ForeColor = glS1
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame10_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(5).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame6_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(4).ForeColor = glS1
    
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame8_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim sMessText As String

    Select Case Index
    
        Case 1, 2, 3
    
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        gF2Prompt.cFeld = "LINR"
            
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
            If gF2Prompt.cWahl <> "" Then
                Label1(Index).Caption = gF2Prompt.cWahl
                Label16(Index).Caption = ermLiefBez(Label1(Index).Caption)
            End If
        End If
        
        voreinstellungspeichern30
        
        Case 4
            sMessText = "Klebeetiketten, 40 mm breit x 18 mm hoch, auf Rolle, 1 Verpackungseinheit = 1 Rolle = 3500 Etiketten = 10,15 €/Brutto" & vbCrLf & vbCrLf
            sMessText = sMessText & "Denken Sie bitte auch an Carbonfarbband! 1 Rolle pro 3500 Etiketten (ArtNr: 500945, Preis 1 Rolle = 7,14 €/Brutto )" & vbCrLf & vbCrLf
            sMessText = sMessText & "verwendeter Etikettendrucker: " & gcEtikettenDrucker
            
            Etikettenbestellung_Per_Mail 501000, sMessText
            
        Case 5
        
            Select Case cboStrichDINA4.Text
            
                Case "35,6 x 16,9"
                
                    sMessText = "Klebeetiketten DIN A4, 35,6 mm breit x 16,9 mm hoch, 1 Verpackungseinheit = 1 Blatt á 80 Etiketten = 80 Etiketten = 1,31 €/Brutto" & vbCrLf & vbCrLf
                    sMessText = sMessText & "verwendeter Etikettendrucker: " & gcEtikettenDrucker
                
                    Etikettenbestellung_Per_Mail 501367, sMessText
                
                Case "35,6 x 16,9 Variante 2"
                
                    sMessText = "Klebeetiketten DIN A4, 35,6 mm breit x 16,9 mm hoch, 1 Verpackungseinheit = 1 Blatt á 80 Etiketten = 80 Etiketten = 1,31 €/Brutto" & vbCrLf & vbCrLf
                    sMessText = sMessText & "verwendeter Etikettendrucker: " & gcEtikettenDrucker
                
                    Etikettenbestellung_Per_Mail 501367, sMessText
            
            End Select
        
        Case 8
            sMessText = "Perforierte Kartonetiketten, 40 mm breit x 25 mm hoch, auf Rolle, 1 Verpackungseinheit = 1 Rolle = 1000 Etiketten = 34,51 €/Brutto" & vbCrLf & vbCrLf
            sMessText = sMessText & "Denken Sie bitte auch an Carbonfarbband! 1 Rolle pro 2000 Etiketten (ArtNr: 500945, Preis 1 Rolle = 7,14 €/Brutto)" & vbCrLf & vbCrLf
            sMessText = sMessText & "verwendeter Etikettendrucker: " & gcEtikettenDrucker
            
            Etikettenbestellung_Per_Mail 500915, sMessText
    End Select

     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Select Case Index
    
        Case 1, 2, 3
    
            Label1(Index).ForeColor = glLink
        Case 4
            Label1(Index).ForeColor = glLink
        Case 5
            Label1(Index).ForeColor = glLink
        
        Case 8
            Label1(Index).ForeColor = glLink
    End Select

     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label17_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Index > 0 Then
     
        Label1(Index).Caption = "undefiniert"
        Label16(Index).Caption = ""
        
        voreinstellungspeichern30
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label17_Click"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
     Dim lOk        As Long
    
    If KeyCode = 16 Then
        If bOk = False Then
            If List2.Selected(List2.ListIndex) = True Then
                iErsteZeile = List2.ListIndex
                bOk = True
                Exit Sub
            End If
        End If
        If List2.Selected(List2.ListIndex) = True Then iZweiteZeile = List2.ListIndex
        For lOk = iErsteZeile To iZweiteZeile
            List2.Selected(lOk) = True
        Next
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    
    If List2.ListCount > 0 Then
        If List2.Selected(List2.ListIndex) = True Then
            'Zähler hoch
            Zähler "hoch", Mid(List2.list(List2.ListIndex), 63, 6)
        Else
            'Zähler runter
            Zähler "runter", Mid(List2.list(List2.ListIndex), 63, 6)
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Zähler(sArt As String, sAnzEti As String)
On Error GoTo LOKAL_ERROR
    
    sAnzEti = SwapStr(sAnzEti, ".", "")
    Select Case sArt
        Case "hoch"
            Label4(6).Caption = CLng(Label4(6).Caption) + 1
            Label4(7).Caption = CLng(Label4(7).Caption) - 1
            
            Label4(9).Caption = CLng(Label4(9).Caption) + Val(sAnzEti)
            Label4(8).Caption = CLng(Label4(8).Caption) - Val(sAnzEti)
        Case "runter"
            Label4(6).Caption = CLng(Label4(6).Caption) - 1
            Label4(7).Caption = CLng(Label4(7).Caption) + 1
            
            Label4(9).Caption = CLng(Label4(9).Caption) - Val(sAnzEti)
            Label4(8).Caption = CLng(Label4(8).Caption) + Val(sAnzEti)
    End Select

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zähler"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = Len(Text1(Index).Text)
    Label3.Caption = Index
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text2.BackColor = glSelBack1
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr$(8)
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text2.BackColor = vbWhite

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

