VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmWK25f 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Rabatt-Verkäufe"
   ClientHeight    =   8625
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   11910
   Icon            =   "frmWK25f.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame3 
      Height          =   3855
      Left            =   120
      TabIndex        =   46
      Top             =   3600
      Visible         =   0   'False
      Width           =   11655
      Begin sevCommand3.Command Command1 
         Height          =   310
         Index           =   6
         Left            =   11160
         TabIndex        =   60
         ToolTipText     =   "Starten Sie hier die Anzeige"
         Top             =   240
         Width           =   330
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MaskColor       =   16777215
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
         UseMaskColor    =   -1  'True
         Version3        =   -1  'True
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   4560
         ScaleHeight     =   2055
         ScaleWidth      =   2055
         TabIndex        =   52
         Top             =   960
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   120
         ScaleHeight     =   36.248
         ScaleMode       =   6  'Millimeter
         ScaleWidth      =   36.248
         TabIndex        =   47
         Top             =   960
         Width           =   2055
      End
      Begin sevCommand3.Command cmdMinus 
         Height          =   310
         Left            =   120
         TabIndex        =   58
         ToolTipText     =   "Starten Sie hier die Anzeige"
         Top             =   240
         Width           =   450
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MaskColor       =   16777215
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
         UseMaskColor    =   -1  'True
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdPlus 
         Height          =   310
         Left            =   3360
         TabIndex        =   57
         ToolTipText     =   "Starten Sie hier die Anzeige"
         Top             =   240
         Width           =   450
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         MaskColor       =   16777215
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
         UseMaskColor    =   -1  'True
         Version3        =   -1  'True
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
         Index           =   9
         Left            =   2160
         TabIndex        =   56
         Top             =   240
         Width           =   1095
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
         Index           =   8
         Left            =   840
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label8 
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
         Height          =   135
         Index           =   0
         Left            =   6960
         TabIndex        =   54
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label7 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   0
         Left            =   6720
         TabIndex        =   53
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "nur die Rabatte betrachtet"
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
         Left            =   4560
         TabIndex        =   51
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Rabattverteilung"
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
         TabIndex        =   50
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label6 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label5 
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
         Index           =   0
         Left            =   480
         TabIndex        =   48
         Top             =   3120
         Width           =   4695
      End
   End
   Begin VB.Frame Frame0 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Height          =   975
      Left            =   0
      TabIndex        =   15
      Top             =   7560
      Width           =   12015
      Begin sevCommand3.Command Command0 
         Height          =   855
         Index           =   0
         Left            =   480
         TabIndex        =   17
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
      Begin sevCommand3.Command Command0 
         Height          =   855
         Index           =   1
         Left            =   1320
         TabIndex        =   18
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
         Index           =   2
         Left            =   2160
         TabIndex        =   19
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
         Index           =   3
         Left            =   3000
         TabIndex        =   20
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
         Index           =   4
         Left            =   3840
         TabIndex        =   21
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
         Index           =   5
         Left            =   4680
         TabIndex        =   22
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
         Index           =   6
         Left            =   5520
         TabIndex        =   23
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
         Index           =   7
         Left            =   6360
         TabIndex        =   24
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
         Index           =   8
         Left            =   7200
         TabIndex        =   25
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
         Index           =   9
         Left            =   8040
         TabIndex        =   26
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
         Index           =   10
         Left            =   8880
         TabIndex        =   27
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
         Index           =   11
         Left            =   9720
         TabIndex        =   28
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
         Index           =   12
         Left            =   10560
         TabIndex        =   29
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
      Begin VB.Label Label0 
         BackStyle       =   0  'Transparent
         Caption         =   "-1"
         Height          =   255
         Left            =   120
         TabIndex        =   16
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
      Height          =   3855
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   11895
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   5
         Left            =   9600
         TabIndex        =   59
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
         Caption         =   "Verteilung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
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
         Height          =   2055
         Left            =   0
         TabIndex        =   40
         Top             =   840
         Width           =   2415
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
            Index           =   8
            Left            =   120
            TabIndex        =   61
            Top             =   1800
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            Index           =   4
            Left            =   120
            TabIndex        =   41
            Top             =   1440
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
            TabIndex        =   45
            Top             =   120
            Width           =   2175
         End
      End
      Begin sevCommand3.Command Command1 
         Height          =   405
         Index           =   1
         Left            =   5520
         TabIndex        =   39
         Top             =   0
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
         Left            =   5520
         TabIndex        =   38
         Top             =   480
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
         Left            =   5520
         TabIndex        =   37
         Top             =   960
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
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bediener"
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
         Index           =   3
         Left            =   7440
         TabIndex        =   35
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kundennummer"
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
         Left            =   7440
         TabIndex        =   34
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Rabatt in Euro"
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
         Index           =   1
         Left            =   7440
         TabIndex        =   33
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Datum"
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
         Index           =   0
         Left            =   7440
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   0
         Left            =   9600
         TabIndex        =   6
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
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   2
         Left            =   9600
         TabIndex        =   7
         Top             =   1320
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   3
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   4
         Left            =   4200
         TabIndex        =   4
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#######"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   5
         Left            =   4200
         TabIndex        =   5
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###"
         PromptChar      =   "_"
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   20
         Left            =   2520
         TabIndex        =   62
         ToolTipText     =   "Kalender"
         Top             =   0
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
         Left            =   2520
         TabIndex        =   63
         ToolTipText     =   "Kalender"
         Top             =   480
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
         Left            =   7440
         TabIndex        =   36
         Top             =   0
         Width           =   1455
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
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   1095
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
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikel-Nr.:"
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
         Left            =   3120
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Left            =   3120
         TabIndex        =   11
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Left            =   3120
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bed.Nr."
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
         Left            =   3120
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Label lblAnzeige 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   31
      Top             =   7080
      Width           =   10815
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Rabatt-Protokoll"
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
      TabIndex        =   30
      Top             =   0
      Width           =   10935
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
End
Attribute VB_Name = "frmWK25f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim glfarbeR(18) As Long
Dim bErste As Boolean
Private Sub cmdMinus_Click()
On Error GoTo LOKAL_ERROR

    Dim lDatum As Long
    
    lDatum = DateValue(Label1(8).Caption)
    lDatum = lDatum - 1
    Label1(8).Caption = Format(lDatum, "DD.MM.YYYY")
    Label1(8).Refresh
    
    Label1(9).Caption = Format(lDatum, "DD.MM.YYYY")
    Label1(9).Refresh
    
    MaskEdBox1(0).Text = Label1(8).Caption
    MaskEdBox1(1).Text = Label1(9).Caption
    
    zeigeRabattkreis

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdMinus_Click"
    Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub cmdPlus_Click()
On Error GoTo LOKAL_ERROR

    Dim lDatum As Long
    
    lDatum = DateValue(Label1(9).Caption)
    lDatum = lDatum + 1
    Label1(8).Caption = Format(lDatum, "DD.MM.YYYY")
    Label1(8).Refresh
    
    Label1(9).Caption = Format(lDatum, "DD.MM.YYYY")
    Label1(9).Refresh
    
    MaskEdBox1(0).Text = Label1(8).Caption
    MaskEdBox1(1).Text = Label1(9).Caption
    
    zeigeRabattkreis

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdPlus_Click"
    Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Activate()
    On Error GoTo LOKAL_ERROR
    
    Me.Refresh
    
    If bErste = True Then
        zeigeRabattkreis
        bErste = False
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdPlus_Click"
    Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "vkpro1", gdBase
    loeschNEW "vkpro", gdBase
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
Private Sub LeseDatenWK25f()
    On Error GoTo LOKAL_ERROR
    
    Dim dPreis As Double
    Dim iMenge As Integer
    Dim dVkPr  As Double
    
    Dim iFileNr As Integer
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim bAnd As Boolean
    
    Dim cFeld As String
    Dim dWert As Double
    Dim cLBSatz As String
    
    Dim lDatVon As Long
    Dim lDatBis As Long
    Dim lartnr As Long
    Dim lLinr As Long
    Dim lKUNDNR As Long
    Dim lBedNr As Long
    Dim ctmp As String
    
    Dim dSumUmsatz As Double
    Dim dSumSoll As Double
    Dim lSumMenge As Long
    Dim lMenge As Long
    Dim lAnz As Long
    Dim dRabProz As Double
    Dim dIst As Double
    Dim dSoll As Double
    
    loeschNEW "vkpro1", gdBase
    loeschNEW "vkpro", gdBase
    
    lblanzeige.Caption = "Daten werden ermittelt..."
    lblanzeige.Refresh
    
    If MaskEdBox1(0).Text <> "__.__.____" Then
        ctmp = MaskEdBox1(0).Text
        lDatVon = DateValue(ctmp)
    Else
        lDatVon = -1
    End If
    
    If MaskEdBox1(1).Text <> "__.__.____" Then
        ctmp = MaskEdBox1(1).Text
        lDatBis = DateValue(ctmp)
    Else
        lDatBis = -1
    End If
    
    If MaskEdBox1(2).Text <> "______" Then
        ctmp = MaskEdBox1(2).Text
        lartnr = Val(ctmp)
    Else
        lartnr = -1
    End If
    
    If MaskEdBox1(3).Text <> "______" Then
        ctmp = MaskEdBox1(3).Text
        lLinr = Val(ctmp)
    Else
        lLinr = -1
    End If
    
    If MaskEdBox1(4).Text <> "_______" Then
        ctmp = MaskEdBox1(4).Text
        lKUNDNR = Val(ctmp)
    Else
        lKUNDNR = -1
    End If
    
    If MaskEdBox1(5).Text <> "___" Then
        ctmp = MaskEdBox1(5).Text
        lBedNr = Val(ctmp)
    Else
        lBedNr = -1
    End If
    
    cSQL = "Select * into vkpro1 from Kassjour where "
    
    bAnd = False
    
    If lDatVon > -1 Then
        cSQL = cSQL & "ADATE >= " & Trim$(Str$(lDatVon))
        bAnd = True
    End If
    
    If lDatBis > -1 Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        cSQL = cSQL & "ADATE <= " & Trim$(Str$(lDatBis))
        bAnd = True
    End If
    
    If lartnr > -1 Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        cSQL = cSQL & "ARTNR = " & Trim$(Str$(lartnr))
        bAnd = True
    End If
    
    If lLinr > -1 Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        cSQL = cSQL & "LINR = " & Trim$(Str$(lLinr))
        bAnd = True
    End If
    
    If lKUNDNR > -1 Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        cSQL = cSQL & "KUNDNR = " & Trim$(Str$(lKUNDNR))
        bAnd = True
    End If
    
    If lBedNr > -1 Then
        If bAnd Then
            cSQL = cSQL & " and "
        End If
        cSQL = cSQL & "BEDIENER = " & Trim$(Str$(lBedNr))
        bAnd = True
    End If
    
    cSQL = cSQL & " and Abs(PREIS) <> Abs(MENGE * VKPR) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update vkpro1 set MOPREIS = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from vkpro1 where (MENGE * VKPR) < 0.01 and (MENGE * VKPR) > -0.01"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "select * from vkpro1"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Preis) Then
                dPreis = rsrs!Preis
            Else
                dPreis = 0
            End If
            If Not IsNull(rsrs!Menge) Then
                iMenge = rsrs!Menge
            Else
                iMenge = 0
            End If
            If Not IsNull(rsrs!vkpr) Then
                dVkPr = rsrs!vkpr
            Else
                dVkPr = 0
            End If
            rsrs.Edit
            dWert = dPreis - (iMenge * dVkPr)
            
            dWert = Format$(dWert, "#####0.00")
            rsrs!MOPREIS = dWert
            rsrs.Update
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
        
    cSQL = "Delete from vkpro1 where MOPREIS = 0 "
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Delete from vkpro1 where MOPREIS is null "
    gdBase.Execute cSQL, dbFailOnError
            
    loeschNEW "vkpro", gdBase
    CreateTable "VKPRO", gdBase
    
    cSQL = "Insert into vkpro Select "
    cSQL = cSQL & " artnr "
    cSQL = cSQL & ", bezeich "
    cSQL = cSQL & ", MENGE  "
    cSQL = cSQL & ", AGN  "
    cSQL = cSQL & ", Preis  "
    cSQL = cSQL & ", ADATE  "
    cSQL = cSQL & ", AZEIT  "
    cSQL = cSQL & ", BEDIENER  "
    cSQL = cSQL & ", KUNDNR  "
    cSQL = cSQL & ", LINR  "
    cSQL = cSQL & ", EAN  "
    cSQL = cSQL & ", FILIALE  "
    cSQL = cSQL & ", KASNUM  "
    cSQL = cSQL & ", MWST "
    cSQL = cSQL & ", EKPR   "
    cSQL = cSQL & ", VKPR  "
    cSQL = cSQL & ", MOPREIS  "
    cSQL = cSQL & ", BELEGNR  "
    cSQL = cSQL & ", UMS_OK "
    cSQL = cSQL & ", KK_ART "
    cSQL = cSQL & ", Best1  "
    cSQL = cSQL & ", Rabkenn "
    cSQL = cSQL & ", LPZ "

    cSQL = cSQL & " from vkpro1 "
    cSQL = cSQL & " where MOPREIS <> 0 and MOPREIS is not null "
    cSQL = cSQL & " order by ADATE, AZEIT, ARTNR "
    gdBase.Execute cSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("vkpro", dbOpenTable)
    If rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        lblanzeige.Caption = "Es wurden keine Daten ermittelt."
        lblanzeige.Refresh
    Else
        rsrs.Close: Set rsrs = Nothing
        If Option1(0).Value = True Then                         'Datum
            reportbildschirm "dWKL36", "aWKL36"
        ElseIf Option1(1).Value = True Then                     'Rabat
            Sortierung 1
            reportbildschirm "dWKL36a", "aWKL36a"
        ElseIf Option1(2).Value = True Then                     'kundnr
            Sortierung 2
            reportbildschirm "dWKL36a", "aWKL36a"
        ElseIf Option1(3).Value = True Then                     'Bednr
            Sortierung 3
            reportbildschirm "dWKL36a", "aWKL36a"
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
        Fehler.gsFunktion = "LeseDatenWK25f"
        Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten. "
    
        Fehlermeldung1
    End If
End Sub
      
Private Sub Rabatthöhen(lDatVon As Long, lDatBis As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim dPreis          As Double
    Dim dVkPr           As Double
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim dWert           As Double
    Dim lMenge          As Long
    Dim dSollumsatz     As Double
    Dim dRabattEuro     As Double
    Dim dRabattProz     As Double
    Dim dStart          As Double
    Dim dEnd            As Double
    Dim i               As Integer
    Dim k               As Integer
    Dim dErg            As Double
    Dim dAnteil         As Double
    Dim dSumSollUmsatz  As Double
    
    loeschNEW "VKRAB", gdBase
    CreateTableT2 "VKRAB", gdBase
    
    cSQL = "Insert into VKRAB Select Preis,Menge,vkpr,0 as mopreis from Kassjour where ADATE >= " & Trim$(Str$(lDatVon))
    cSQL = cSQL & " and ADATE <= " & Trim$(Str$(lDatBis))
    cSQL = cSQL & " and Abs(PREIS) < Abs(MENGE * VKPR) "
    cSQL = cSQL & " and menge > 0"
    cSQL = cSQL & " and preis > 0"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from VKRAB where (MENGE * VKPR) < 0.01 and (MENGE * VKPR) > -0.01"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "select * from VKRAB"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Preis) Then
                dPreis = rsrs!Preis
            Else
                dPreis = 0
            End If
            If Not IsNull(rsrs!Menge) Then
                lMenge = rsrs!Menge
            Else
                lMenge = 0
            End If
            If Not IsNull(rsrs!vkpr) Then
                dVkPr = rsrs!vkpr
            Else
                dVkPr = 0
            End If
            rsrs.Edit
            dWert = dPreis - (lMenge * dVkPr)
            
            dWert = Format$(dWert, "#####0.00")
            rsrs!MOPREIS = dWert
            rsrs.Update
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
        
    cSQL = "Delete from VKRAB where MOPREIS = 0 "
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Delete from VKRAB where MOPREIS is null "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "select * from VKRAB"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!vkpr) Then
                dVkPr = rsrs!vkpr
            Else
                dVkPr = 0
            End If
            
            If Not IsNull(rsrs!Menge) Then
                lMenge = rsrs!Menge
            Else
                lMenge = 0
            End If
            
            If Not IsNull(rsrs!Preis) Then
                dPreis = rsrs!Preis
            Else
                dPreis = 0
            End If
            
            dSollumsatz = dVkPr * lMenge
            dRabattEuro = Round(dSollumsatz - dPreis, 2)
            
            If dSollumsatz <> 0 Then
                dRabattProz = Round(100 * dRabattEuro / dSollumsatz, 0)
            Else
                dRabattProz = 0
            End If
            
            rsrs.Edit
            rsrs!Sollumsatz = dSollumsatz
            rsrs!Rabatteuro = dRabattEuro
            rsrs!RabattProz = dRabattProz
            rsrs.Update
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    loeschNEW "VKRABANT", gdBase
    CreateTableT2 "VKRABANT", gdBase
    
    cSQL = "Insert into VKRABANT Select "
    cSQL = cSQL & " sum(Sollumsatz) as SollumsatzSum "
    cSQL = cSQL & ", sum(RabattEuro) as RabattEuroSum "
    cSQL = cSQL & ", RabattProz  "
    cSQL = cSQL & " from VKRAB group by RabattProz "
    gdBase.Execute cSQL, dbFailOnError
    
    dSumSollUmsatz = 0
    
    cSQL = "select sum(SollumsatzSum) as Maxi from VKRABANT"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dSumSollUmsatz = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    If dSumSollUmsatz <> 0 Then
        cSQL = "Update VKRABANT set ANTEIL = 100 * Sollumsatzsum / '" & dSumSollUmsatz & "'"
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    i = -1
    
    dStart = 0.001
    dEnd = 0
    
    Picture2.ScaleMode = vbPixels
    
    cSQL = "select * from VKRABANT order by anteil desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            If Not IsNull(rsrs!Anteil) Then
                dAnteil = rsrs!Anteil
            Else
                dAnteil = 0
            End If
            
            If dAnteil > 0 Then

                dEnd = dEnd + dAnteil
                i = i + 1
                
                
                Dim j As Double
                Dim l As Double
                
                Dim siRadius As Single
                
                siRadius = CDbl(Picture2.Height) / 34
                
                For j = dStart To dEnd
'                    l = dStart + j - 0.002

                    j = j + 1
                    If j > 100 Then j = 100
                    Call DrawPiePiece(glfarbeR(i), dStart, j, Picture2, siRadius)
                
                    PauseSi (0.01)
                Next j
                
                Call DrawPiePiece(glfarbeR(i), dStart, dEnd, Picture2, siRadius)
'                PauseSi (0.2)
                
                
                dStart = dStart + dAnteil
                
                
                dErg = CDbl(rsrs!Sollumsatzsum - rsrs!RabattEurosum)
                
                If i = 0 Then
                    k = i
                    Label7(k).BackColor = glfarbeR(i)
                    Label7(k).Visible = True
                    Label7(k).Refresh
                    Label8(k).Caption = Format(dAnteil, "#####0.00") & " % der Rabatte werden mit einem Rabatt von " & rsrs!RabattProz & " % erzielt"
                    Label8(k).Visible = True
                    Label8(k).Refresh
                Else
                    k = i
                    
                    Load Label7(k)
                    Label7(k).Top = Label7(0).Top
                    Label7(k).Top = Label7(0).Top + Label7(0).Height * (k) + 20 * (k)
                    Label7(k).BackColor = glfarbeR(i)
                    Label7(k).Visible = True
                    
                    Load Label8(k)
                    Label8(k).Top = Label8(0).Top
                    Label8(k).Top = Label8(0).Top + Label7(0).Height * (k) + 20 * (k)
                    Label8(k).Caption = Format(dAnteil, "#####0.00") & " % der Rabatte werden mit einem Rabatt von " & rsrs!RabattProz & " % erzielt"
                    Label8(k).Visible = True
                    Label8(k).Refresh
                End If
            End If
            
            If dStart > 100 Then
                Exit Do
            End If
            
            If i = 18 Then
                Exit Do
            End If
        
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 360 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Rabatthöhen"
        Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten. "
    
        Fehlermeldung1
        
'        Resume Next
    End If
End Sub
Private Sub Rabattgrob(lDatVon As Long, lDatBis As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim dPreis          As Double
    Dim dVkPr           As Double
    Dim cSQL            As String
    Dim rsrs            As Recordset
    Dim lMenge          As Long
    Dim dSollumsatz     As Double
    Dim dRabattEuro     As Double
    Dim dRabattProz     As Double
    Dim dSumSollUmsatz  As Double
    Dim dAnteil         As Double
    Dim dStart          As Double
    Dim dEnd            As Double
    Dim i               As Integer
    Dim k               As Integer
    Dim dErg            As Double
    
    loeschNEW "VKRAB", gdBase
    CreateTableT2 "VKRAB", gdBase
    
    cSQL = "Insert into vkRAB select  "
    cSQL = cSQL & " MENGE  "
    cSQL = cSQL & ", Preis  "
    cSQL = cSQL & ", VKPR  "
    cSQL = cSQL & ", vkpr * menge as Sollumsatz  "
    cSQL = cSQL & ", round((vkpr * menge) - preis, 2) As Rabatteuro "
    cSQL = cSQL & " from Kassjour where ADATE >= " & Trim$(Str$(lDatVon))
    cSQL = cSQL & " and ADATE <= " & Trim$(Str$(lDatBis))
    cSQL = cSQL & " and preis > 0 "
    gdBase.Execute cSQL, dbFailOnError
    
'    cSQL = "select * from VKRAB"
'    Set rsrs = gdBase.OpenRecordset(cSQL)
'    If Not rsrs.EOF Then
'        rsrs.MoveFirst
'        Do While Not rsrs.EOF
'
'            If Not IsNull(rsrs!vkpr) Then
'                dVkPr = rsrs!vkpr
'            Else
'                dVkPr = 0
'            End If
'
'            If Not IsNull(rsrs!Menge) Then
'                lMenge = rsrs!Menge
'            Else
'                lMenge = 0
'            End If
'
'            If Not IsNull(rsrs!Preis) Then
'                dPreis = rsrs!Preis
'            Else
'                dPreis = 0
'            End If
'
'            dSollumsatz = dVkPr * lMenge
'            dRabattEuro = Round(dSollumsatz - dPreis, 2)
'
'            If dSollumsatz <> 0 Then
'                dRabattProz = Round(100 * dRabattEuro / dSollumsatz, 0)
'            Else
'                dRabattProz = 0
'            End If
'
'            rsrs.Edit
'            rsrs!Sollumsatz = dSollumsatz
'            rsrs!RabattEuro = dRabattEuro
'            rsrs!RabattProz = dRabattProz
'            rsrs.Update
'            rsrs.MoveNext
'        Loop
'    End If
'    rsrs.Close: Set rsrs = Nothing
    
'    cSQL = "Update VKRAB Set Sollumsatz = vkpr * menge   "
'    cSQL = cSQL & " , RabattEuro = round((vkpr*menge)-preis,2) "
'    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update VKRAB Set  RabattProz = Round(100 * RabattEuro / Sollumsatz, 0) "
    cSQL = cSQL & " where Sollumsatz <> 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "VKRABANT", gdBase
    CreateTableT2 "VKRABANT", gdBase
    
    cSQL = "Insert into VKRABANT Select "
    cSQL = cSQL & " sum(Sollumsatz) as SollumsatzSum "
    cSQL = cSQL & ", sum(RabattEuro) as RabattEuroSum "
    cSQL = cSQL & " from VKRAB where rabattproz = 0 " 'group by RabattProz"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update VKRABANT Set RabattProz = 0   "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into VKRABANT Select "
    cSQL = cSQL & " sum(Sollumsatz) as SollumsatzSum "
    cSQL = cSQL & ", sum(RabattEuro) as RabattEuroSum "
    cSQL = cSQL & " from VKRAB where rabattproz > 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    dSumSollUmsatz = 0
    
    cSQL = "select sum(SollumsatzSum) as Maxi from VKRABANT"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dSumSollUmsatz = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    If dSumSollUmsatz <> 0 Then
        cSQL = "Update VKRABANT set ANTEIL = 100 * Sollumsatzsum / '" & dSumSollUmsatz & "'"
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    i = -1
    
    dStart = 0.001
    dEnd = 0
    
    Picture1.ScaleMode = vbPixels
    
    cSQL = "select * from VKRABANT"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            If Not IsNull(rsrs!Anteil) Then
                dAnteil = rsrs!Anteil
            Else
                dAnteil = 0
            End If
            
            If dAnteil > 0 Then

                dEnd = dEnd + dAnteil
                i = i + 1
                
                Dim j As Double
                Dim l As Double
                Dim siRadius As Single
                
                siRadius = CDbl(Picture1.Height) / 34
                For j = dStart To dEnd
                    j = j + 1
                    If j > 100 Then j = 100
                    Call DrawPiePiece(glfarbeR(i + 17), dStart, j, Picture1, siRadius)
                
                    PauseSi (0.02)
                Next j
                
                
                Call DrawPiePiece(glfarbeR(i + 17), dStart, dEnd, Picture1, siRadius)
                
                dStart = dStart + dAnteil
                
                dErg = CDbl(rsrs!Sollumsatzsum - rsrs!RabattEurosum)
                
                If i = 0 Then
                    k = i
                    Label6(k).BackColor = glfarbeR(i + 17)
                    Label6(k).Visible = True
                    Label6(k).Refresh
                    
                    If rsrs!RabattProz = 0 Then
                        Label5(k).Caption = Format(dAnteil, "#####0") & " % des Umsatzes ohne Rabatt erzielt (" & Format(dErg, "#####0.00") & ")"
                    Else
                        Label5(k).Caption = Format(dAnteil, "#####0") & " % des Umsatzes mit Rabatt erzielt (" & Format(dErg, "#####0.00") & ")"
                    End If
                    Label5(k).Visible = True
                    Label5(k).Refresh
                Else
                    k = i
                    
                    Load Label6(k)
                    Label6(k).Top = Label6(0).Top
                    Label6(k).Top = Label6(0).Top + Label6(0).Height * (k) + 60 * (k)
                    Label6(k).BackColor = glfarbeR(i + 17)
                    Label6(k).Visible = True
                    
                    Load Label5(k)
                    Label5(k).Top = Label5(0).Top
                    Label5(k).Top = Label5(0).Top + Label5(0).Height * (k) + 60 * (k)
                    If rsrs!RabattProz = 0 Then
                        Label5(k).Caption = Format(dAnteil, "#####0") & " % des Umsatzes ohne Rabatt erzielt (" & Format(dErg, "#####0.00") & ")"
                    Else
                        Label5(k).Caption = Format(dAnteil, "#####0") & " % des Umsatzes mit Rabatt erzielt (" & Format(dErg, "#####0.00") & ")"
                    End If
                    Label5(k).Visible = True
                    Label5(k).Refresh
                End If
            End If
        
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 360 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Rabattgrob"
        Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten. "
    
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iFeld As Integer
    Dim iCount As Integer
    Dim ctmp As String
    Dim cZiel As String
    Dim cZeichen As String
    
    iFeld = Val(Label0.Caption)
    
    ctmp = MaskEdBox1(iFeld).Text
    
    Select Case Index
        Case 0 To 9
            cZeichen = Command0(Index).Caption
            For iCount = 1 To Len(ctmp)
                If Mid(ctmp, iCount, 1) = "_" Then
                    Mid(ctmp, iCount, 1) = cZeichen
                    Exit For
                End If
            Next iCount
            MaskEdBox1(iFeld).Text = ctmp
        Case Is = 10
            If iFeld > 0 Then
                iFeld = iFeld - 1
            End If
            
        Case Is = 11
            If iFeld < 5 Then
                iFeld = iFeld + 1
            End If
        
        Case Is = 12
            If iFeld < 2 Then
                MaskEdBox1(iFeld).Text = "__.__.____"
            End If
        Case Is = 20        ' Kalender
            MaskEdBox1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
            
            Label1(8).Caption = MaskEdBox1(0).Text
            Label1(8).Refresh
            
            
            If Frame3.Visible = True Then
                zeigeRabattkreis
            End If
        Case Is = 21        ' Kalender
            MaskEdBox1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
            
            Label1(9).Caption = MaskEdBox1(1).Text
            Label1(9).Refresh
            
            If Frame3.Visible = True Then
                zeigeRabattkreis
            End If
    End Select
    
    MaskEdBox1(iFeld).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim lDatVon As Long
    Dim lDatBis As Long
    
    Select Case Index
        Case Is = 0     'Zeige
            iRet = fnPruefeEingabeWK25f()
            Select Case iRet
                Case Is = 0     'alles okay!
                    Screen.MousePointer = 11
                    LeseDatenWK25f
                    Screen.MousePointer = 0
                Case Is = 1     'Von Datum falsch
                    MsgBox "Das eingegebene VON-Datum ist falsch!", vbCritical, "STOP!"
                    MaskEdBox1(0).SetFocus
                Case Is = 2     'Bis Datum falsch
                    MsgBox "Das eingegebene BIS-Datum ist falsch!", vbCritical, "STOP!"
                    MaskEdBox1(1).SetFocus
                Case Is = 3     'Von Datum > Bis Datum
                    MsgBox "Das eingegebene VON-Datum ist größer als das BIS-Datum!", vbCritical, "STOP!"
                    MaskEdBox1(0).SetFocus
                Case Is = 4
                Case Is = 5
            End Select
        Case Is = 2     'schließe
            Unload frmWK25f
        Case 1
            MaskEdBox1_KeyUp 3, vbKeyF2, 0
        Case 3
            MaskEdBox1_KeyUp 4, vbKeyF2, 0
        Case 4
            MaskEdBox1_KeyUp 5, vbKeyF2, 0
        Case 5
            zeigeRabattkreis
        Case 6
            Frame3.Visible = False
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub zeigeRabattkreis()
    On Error GoTo LOKAL_ERROR
    
    Dim lDatVon As Long
    Dim lDatBis As Long
    
    Dim i As Integer
    For i = 1 To 20
    
        Unload Label5(i)
        Unload Label6(i)
        Unload Label7(i)
        Unload Label8(i)
        
    
    Next i
    
    Label5(0).Visible = False
    Label6(0).Visible = False
    Label7(0).Visible = False
    Label8(0).Visible = False
    
    Picture1.FillStyle = 1
    Picture1.Refresh
    Picture2.FillStyle = 1
    Picture2.Refresh
       
    
    lDatVon = -1
    lDatBis = -1
    
    lDatVon = DateValue(Label1(8).Caption)
    lDatBis = DateValue(Label1(9).Caption)

    If lDatVon > -1 And lDatBis > -1 Then
    
        Frame3.Visible = True
        Me.Refresh
    
        Rabattgrob lDatVon, lDatBis
        Rabatthöhen lDatVon, lDatBis
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 340 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "zeigeRabattkreis"
        Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten. "
    
        Fehlermeldung1
'        Resume Next
    End If
'
End Sub
Private Function fnPruefeEingabeWK25f()
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld As String
    Dim bgefunden As Boolean
    Dim lVon As Long
    Dim lBis As Long
    
    fnPruefeEingabeWK25f = 0
    
    bgefunden = False
    
    cFeld = MaskEdBox1(0).Text
    cFeld = Trim$(cFeld)
    If cFeld = "__.__.____" Then
        lVon = 0
    ElseIf Not IsDate(cFeld) Then
        fnPruefeEingabeWK25f = 1
        Exit Function
    Else
        lVon = DateValue(cFeld)
    End If
    
    
    
    cFeld = MaskEdBox1(1).Text
    cFeld = Trim$(cFeld)
    If cFeld = "__.__.____" Then
        lBis = 0
    ElseIf Not IsDate(cFeld) Then
        fnPruefeEingabeWK25f = 2
        Exit Function
    Else
        lBis = DateValue(cFeld)
    End If
    
    If lVon = 0 And lBis = 0 Then
        lVon = Fix(Now)
        MaskEdBox1(0).Text = Format$(lVon, "DD.MM.YYYY")
        lBis = Fix(Now)
        MaskEdBox1(1).Text = Format$(lBis, "DD.MM.YYYY")
    End If
    
    If lVon = 0 And lBis <> 0 Then
        cFeld = Format$(lBis, "DD.MM.YYYY")
        MaskEdBox1(0).Text = cFeld
        lVon = lBis
    End If
    If lVon <> 0 And lBis = 0 Then
        cFeld = Format$(lVon, "DD.MM.YYYY")
        MaskEdBox1(1).Text = cFeld
        lBis = lVon
    End If
        
    
    If lVon > lBis Then
        fnPruefeEingabeWK25f = 3
        Exit Function
    End If
    
    
    cFeld = MaskEdBox1(2).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "______" Then
        bgefunden = True
    End If
    
    cFeld = MaskEdBox1(3).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        bgefunden = True
    End If
    
    cFeld = MaskEdBox1(4).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        bgefunden = True
    End If
    
    cFeld = MaskEdBox1(5).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        bgefunden = True
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEinhabeWK25d"
    Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Function
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    bErste = True
    
    Option1(0).Value = True
    Option1(6).Value = True
    
    glfarbeR(0) = &H8000000F
    glfarbeR(1) = &HFFFF&
    glfarbeR(2) = &HC000&
    glfarbeR(3) = &HFF&
    glfarbeR(4) = &HC0FFFF
    glfarbeR(5) = &H80FF&
    glfarbeR(6) = &HFF00FF
    glfarbeR(7) = &HFFFF00
    glfarbeR(8) = &HC0C0FF
    glfarbeR(9) = &HFFC0C0
    glfarbeR(10) = &H8080FF
    glfarbeR(11) = &HC0FFC0
    glfarbeR(12) = &HFF8080
    glfarbeR(13) = &H40C0&
    glfarbeR(14) = &H800080
    glfarbeR(15) = &H80&
    glfarbeR(16) = &H808000
    glfarbeR(17) = &HC0C0&
    glfarbeR(18) = &HFF80FF
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten. " 'Fehlerstufe = " & Trim$(Str$(iFehler))

    Fehlermeldung1
End Sub
Private Sub MaskEdBox1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    MaskEdBox1(Index).BackColor = glSelBack1
    MaskEdBox1(Index).SelStart = 0
    MaskEdBox1(Index).SelLength = Len(MaskEdBox1(Index).Text)
    Label0.Caption = Trim$(Str$(Index))
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub
Private Sub MaskEdBox1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            Case Is = 2
                ctmp = MaskEdBox1(3).Text
                ctmp = Trim$(Str$(Val(ctmp)))
                If MaskEdBox1(3).Text = "______" Then
                    MsgBox "Bitte einen Lieferanten angeben!", vbCritical, "STOP!"
                    MaskEdBox1(3).SetFocus
                    Exit Sub
                End If
                gF2Prompt.cFeld = "ARTNR"
                gF2Prompt.cWert = ctmp
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    If gF2Prompt.cWahl <> "" Then
                        ctmp = gF2Prompt.cWahl
                        Select Case Index
                            Case Is = 5
                                ctmp = ctmp & String$(3 - Len(ctmp), "_")
                            Case Else
                                ctmp = ctmp & String$(6 - Len(ctmp), "_")
                        End Select
                
                        MaskEdBox1(Index).Text = ctmp
                    End If
                    MaskEdBox1(Index).SetFocus
                End If
            Case Is = 3
                gF2Prompt.cFeld = "LINR"
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    If gF2Prompt.cWahl <> "" Then
                        ctmp = gF2Prompt.cWahl
                        Select Case Index
                            Case Is = 5
                                ctmp = ctmp & String$(3 - Len(ctmp), "_")
                            Case Else
                                ctmp = ctmp & String$(6 - Len(ctmp), "_")
                        End Select
                
                        MaskEdBox1(Index).Text = ctmp
                    End If
                    MaskEdBox1(Index).SetFocus
                End If
            Case Is = 4
                gF2Prompt.cFeld = "KUN"
                
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    If gF2Prompt.cWahl <> "" Then
                        ctmp = gF2Prompt.cWahl
                        
                        ctmp = ctmp & String$(7 - Len(ctmp), "_")
                       
                
                        MaskEdBox1(Index).Text = ctmp
                    End If
                    MaskEdBox1(Index).SetFocus
                End If
            Case Is = 5
                gF2Prompt.cFeld = "BED"
                If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    ctmp = gF2Prompt.cWahl
                    Select Case Index
                        Case Is = 5
                            ctmp = ctmp & String$(3 - Len(ctmp), "_")
                        Case Else
                            ctmp = ctmp & String$(6 - Len(ctmp), "_")
                    End Select
            
                    MaskEdBox1(Index).Text = ctmp
                End If
                MaskEdBox1(Index).SetFocus
        End If
            
        End Select
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub
Private Sub MaskEdBox1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    MaskEdBox1(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Rabattprotokoll ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub

Private Sub Option1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    
    Select Case Index
    
        Case Is = 4   'vormonat
        
            If Month(DateValue(Now)) = 1 Then
                MaskEdBox1(0).Text = Format("01.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
                MaskEdBox1(1).Text = Format("31.12." & Year(DateValue(Now)) - 1, "DD.MM.YYYY")
            Else
                MaskEdBox1(0).Text = Format("01." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                Select Case Month(DateValue(Now)) - 1
                    Case 1, 3, 5, 7, 8, 10, 12
                        MaskEdBox1(1).Text = Format("31." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                    
                    Case 2
                        If Year(DateValue(Now)) = 2016 Then
                            MaskEdBox1(1).Text = Format("29." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        Else
                            MaskEdBox1(1).Text = Format("28." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                        End If
                    
                    Case Else
                        MaskEdBox1(1).Text = Format("30." & Month(DateValue(Now)) - 1 & "." & Year(DateValue(Now)), "DD.MM.YYYY")
                End Select
            End If
            
            Label1(8).Caption = MaskEdBox1(0).Text
            Label1(8).Refresh
            
            Label1(9).Caption = MaskEdBox1(1).Text
            Label1(9).Refresh
            
            If Frame3.Visible = True Then
                zeigeRabattkreis
            End If
                
        Case Is = 5     'ak monat
            MaskEdBox1(0).Text = Format("01." & Month(DateValue(Now)) & "." & Year(DateValue(Now)), "DD.MM.YYYY")
            MaskEdBox1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
            
            Label1(8).Caption = MaskEdBox1(0).Text
            Label1(8).Refresh
            
            Label1(9).Caption = MaskEdBox1(1).Text
            Label1(9).Refresh
            If Frame3.Visible = True Then
                zeigeRabattkreis
            End If
        
        Case Is = 6     'gestern
            MaskEdBox1(0).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
            MaskEdBox1(1).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
            
            Label1(8).Caption = MaskEdBox1(0).Text
            Label1(8).Refresh
            
            Label1(9).Caption = MaskEdBox1(1).Text
            Label1(9).Refresh
            If Frame3.Visible = True Then

                zeigeRabattkreis
            End If
        Case Is = 7     'heute
            MaskEdBox1(0).Text = Format(DateValue(Now), "DD.MM.YYYY")
            MaskEdBox1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
            
            Label1(8).Caption = MaskEdBox1(0).Text
            Label1(8).Refresh
            
            Label1(9).Caption = MaskEdBox1(1).Text
            Label1(9).Refresh
            
            If Frame3.Visible = True Then
                zeigeRabattkreis
            End If
            
        Case Is = 8     'akt Jahr
            MaskEdBox1(0).Text = Format("01.01." & Year(DateValue(Now)), "DD.MM.YYYY")
            MaskEdBox1(1).Text = Format(DateValue(Now), "DD.MM.YYYY")
            
            Label1(8).Caption = MaskEdBox1(0).Text
            Label1(8).Refresh
            
            Label1(9).Caption = MaskEdBox1(1).Text
            Label1(9).Refresh
            If Frame3.Visible = True Then
                zeigeRabattkreis
            End If
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
