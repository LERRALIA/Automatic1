VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWKL112 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Artikel retournieren"
   ClientHeight    =   8625
   ClientLeft      =   2145
   ClientTop       =   2655
   ClientWidth     =   11910
   Icon            =   "frmWKL112.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'Kein
      Height          =   3255
      Left            =   120
      TabIndex        =   84
      Top             =   7440
      Visible         =   0   'False
      Width           =   2175
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
         Index           =   5
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   95
         Top             =   5160
         Width           =   1095
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
         Height          =   2580
         Left            =   120
         TabIndex        =   86
         Top             =   1680
         Width           =   11415
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   4
         Left            =   10440
         TabIndex        =   85
         Top             =   4440
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
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   6
         Left            =   9360
         TabIndex        =   88
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
         Index           =   7
         Left            =   9360
         TabIndex        =   89
         Top             =   5520
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
         Caption         =   "Abschicken"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List8 
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
         TabIndex        =   87
         Top             =   1440
         Width           =   11415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Belegnr:"
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
         Left            =   9360
         TabIndex        =   103
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Linr:"
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
         Left            =   8160
         TabIndex        =   102
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   3840
         TabIndex        =   96
         Top             =   5520
         Width           =   5415
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
         Index           =   12
         Left            =   120
         TabIndex        =   94
         Top             =   4440
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
         Index           =   11
         Left            =   1800
         TabIndex        =   93
         Top             =   3120
         Width           =   4335
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   "nur diese Artikel werden als Retoure versendet"
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
         Index           =   0
         Left            =   120
         TabIndex        =   92
         Top             =   240
         Width           =   8175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel, die nicht enthalten sind, gehören nicht zu dem Lieferanten oder haben keine Bestellnummer hinterlegt."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   120
         TabIndex        =   91
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Belegnr:"
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
         Left            =   8160
         TabIndex        =   90
         Top             =   5160
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      Height          =   1215
      Left            =   10560
      TabIndex        =   37
      Top             =   3240
      Visible         =   0   'False
      Width           =   1320
      Begin sevCommand3.Command Command2 
         Height          =   760
         Index           =   16
         Left            =   9480
         TabIndex        =   26
         Top             =   940
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   20
         Left            =   8640
         TabIndex        =   25
         Top             =   940
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   19
         Left            =   7800
         TabIndex        =   24
         Top             =   940
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   18
         Left            =   6960
         TabIndex        =   23
         Top             =   940
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   14
         Left            =   4440
         TabIndex        =   22
         Top             =   940
         Width           =   2480
         _ExtentX        =   4366
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   13
         Left            =   1920
         TabIndex        =   21
         Top             =   940
         Width           =   2480
         _ExtentX        =   4366
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   12
         Left            =   1080
         TabIndex        =   20
         Top             =   940
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   11
         Left            =   240
         TabIndex        =   19
         Top             =   940
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   17
         Left            =   9480
         TabIndex        =   18
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   10
         Left            =   8640
         TabIndex        =   17
         Top             =   120
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   9
         Left            =   7800
         TabIndex        =   16
         Top             =   120
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   8
         Left            =   6960
         TabIndex        =   15
         Top             =   120
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   7
         Left            =   6120
         TabIndex        =   14
         Top             =   120
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   6
         Left            =   5280
         TabIndex        =   13
         Top             =   120
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   5
         Left            =   4440
         TabIndex        =   12
         Top             =   120
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   4
         Left            =   3600
         TabIndex        =   11
         Top             =   120
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   3
         Left            =   2760
         TabIndex        =   10
         Top             =   120
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   2
         Left            =   1920
         TabIndex        =   9
         Top             =   120
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   1
         Left            =   1080
         TabIndex        =   8
         Top             =   120
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Height          =   760
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   800
         _ExtentX        =   1402
         _ExtentY        =   1349
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
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'Kein
      Height          =   2775
      Left            =   -3840
      TabIndex        =   62
      Top             =   5880
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CheckBox Check21 
         Caption         =   "im Anschluss an Budni"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   97
         Top             =   5760
         Width           =   2175
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   77
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
         TabIndex        =   63
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
         TabIndex        =   64
         Top             =   1080
         Width           =   5535
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   8
         Left            =   10440
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
         Top             =   840
         Width           =   5535
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   11
         Left            =   9360
         TabIndex        =   66
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
         TabIndex        =   65
         Top             =   5160
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
         Caption         =   "Retournieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSComctlLib.ProgressBar pbr1 
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   6360
         Visible         =   0   'False
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Die Artikel, die auf der linken Seite angezeigt sind, können jetzt retourniert werden."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   4560
         TabIndex        =   78
         Top             =   5160
         Width           =   4695
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         Left            =   1800
         TabIndex        =   73
         Top             =   3120
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
         TabIndex        =   72
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
         TabIndex        =   71
         Top             =   600
         Width           =   5295
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'Kein
      Height          =   1095
      Left            =   1080
      TabIndex        =   55
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
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
         Index           =   2
         Left            =   9120
         MaxLength       =   6
         TabIndex        =   79
         Top             =   5640
         Width           =   2415
      End
      Begin sevCommand3.Command Command7 
         Height          =   495
         Index           =   3
         Left            =   9120
         TabIndex        =   59
         Top             =   6360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
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
         Left            =   9120
         TabIndex        =   58
         Top             =   6960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
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
         TabIndex        =   57
         Top             =   6360
         Width           =   8895
      End
      Begin VB.TextBox txtStatus 
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Top             =   6000
         Visible         =   0   'False
         Width           =   975
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   2
         Left            =   9120
         TabIndex        =   80
         Top             =   4680
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
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
         Height          =   495
         Index           =   10
         Left            =   9120
         TabIndex        =   81
         Top             =   5160
         Width           =   2415
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
         TabIndex        =   76
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
         TabIndex        =   61
         Top             =   240
         Width           =   6135
      End
      Begin VB.Image Image2 
         Height          =   915
         Left            =   10800
         MouseIcon       =   "frmWKL112.frx":0442
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL112.frx":074C
         ToolTipText     =   "Klicken Sie hier, wenn Sie Daten aus dem MDE - Gerät einlesen möchten"
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00C0C000&
         Caption         =   $"frmWKL112.frx":0D2F
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
         TabIndex        =   60
         Top             =   2160
         Width           =   7935
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      Height          =   975
      Left            =   120
      TabIndex        =   50
      Top             =   7680
      Width           =   2535
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         Caption         =   "Wie möchten Sie bei der Artikelretoure vorgehen?"
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
         TabIndex        =   54
         Top             =   360
         Width           =   10935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'Kein
      Height          =   5535
      Left            =   120
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CheckBox Check8 
         Caption         =   "Menge halten"
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
         Left            =   960
         TabIndex        =   104
         Top             =   2760
         Width           =   1455
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   0
         Left            =   5640
         TabIndex        =   49
         Top             =   2880
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
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
         Caption         =   "Bestandshistorie"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command8 
         Height          =   375
         Left            =   5880
         TabIndex        =   48
         Top             =   2280
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
         Top             =   3480
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
         Top             =   3840
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
         Top             =   840
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
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2760
         Width           =   1215
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Left            =   10080
         TabIndex        =   6
         Top             =   4725
         Width           =   1575
         _ExtentX        =   2778
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   15
         Left            =   10080
         TabIndex        =   5
         Top             =   4200
         Width           =   1575
         _ExtentX        =   2778
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
         Top             =   240
         Width           =   5055
      End
      Begin sevCommand3.Command Command1 
         Height          =   615
         Left            =   9240
         TabIndex        =   2
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
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
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   0
         Left            =   8400
         TabIndex        =   82
         Top             =   4200
         Width           =   1650
         _ExtentX        =   2910
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   1
         Left            =   8400
         TabIndex        =   83
         Top             =   4725
         Width           =   1650
         _ExtentX        =   2910
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
         Caption         =   "Druck leeren"
         PictureAlign    =   2
         Version3        =   -1  'True
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
         Height          =   1530
         ItemData        =   "frmWKL112.frx":0E36
         Left            =   240
         List            =   "frmWKL112.frx":0E3D
         TabIndex        =   98
         Top             =   3600
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
         ItemData        =   "frmWKL112.frx":0E48
         Left            =   240
         List            =   "frmWKL112.frx":0E4A
         TabIndex        =   99
         Top             =   3360
         Width           =   7695
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   3
         Left            =   8040
         TabIndex        =   100
         Top             =   3360
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
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
         Caption         =   "an Budni"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
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
         TabIndex        =   101
         Top             =   3000
         Width           =   3735
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
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   5
         Left            =   9600
         TabIndex        =   42
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "K-Vk:"
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
         Index           =   9
         Left            =   8160
         TabIndex        =   41
         Top             =   2760
         Width           =   1335
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
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Abgang:"
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
         Left            =   2520
         TabIndex        =   36
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "L-Vk:"
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
         Index           =   5
         Left            =   8160
         TabIndex        =   35
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 Euro"
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
         Left            =   1080
         TabIndex        =   33
         Top             =   2160
         Width           =   3135
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
         Left            =   4320
         TabIndex        =   31
         Top             =   2160
         Width           =   1335
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
         Top             =   360
         Width           =   2175
      End
   End
   Begin sevCommand3.Command Command2 
      Height          =   360
      Index           =   21
      Left            =   11280
      TabIndex        =   46
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
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   11640
      Y1              =   720
      Y2              =   720
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
      TabIndex        =   47
      Top             =   6480
      Width           =   11415
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Artikel retournieren"
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
      Left            =   360
      TabIndex        =   43
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmWKL112"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim new19Artikel As ArtikelTyp
Dim iBudniRetourenzaehler As Integer
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim iRet As Integer

Select Case Index
    Case 0
'        Frame1.Visible = False
        
        iRet = MsgBox("Möchten Sie die unten angezeigten Daten ausdrucken?", vbYesNo + vbQuestion, "Winkiss Frage:")
        If iRet = vbYes Then
            sSQL = "Update RETPRINT inner join Artlief on RETPRINT.linr = Artlief.linr and RETPRINT.artnr = Artlief.artnr set "
            sSQL = sSQL & " RETPRINT.libesnr = Artlief.libesnr"
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update RETPRINT inner join Artikel on  RETPRINT.artnr = Artikel.artnr set "
            sSQL = sSQL & " RETPRINT.BESTANDneu = Artikel.Bestand"
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update RETPRINT inner join lisrt on RETPRINT.linr = lisrt.linr set "
            sSQL = sSQL & " RETPRINT.liefbez = lisrt.liefbez"
            sSQL = sSQL & " ,RETPRINT.KUNDNR = lisrt.KUNDNR"
            sSQL = sSQL & " ,RETPRINT.PLZ = lisrt.PLZ"
            sSQL = sSQL & " ,RETPRINT.STADT = lisrt.STADT"
            sSQL = sSQL & " ,RETPRINT.STRASSE = lisrt.STRASSE"
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update RETPRINT set Firma_Name = '" & gFirma.FirmaName & "'"
            sSQL = sSQL & " ,Firma_Strasse = '" & gFirma.strasse & "'"
            sSQL = sSQL & " ,Firma_Plz = '" & gFirma.Plz & "'"
            sSQL = sSQL & " ,Firma_Ort = '" & gFirma.Ort & "'"
            gdBase.Execute sSQL, dbFailOnError
        
            reportbildschirm "", "aZEN124a"
        End If
    
        
    Case 1
    
       
'        Frame1.Visible = False
        
        iRet = MsgBox("Möchten Sie die unten angezeigten Daten löschen?", vbYesNo + vbQuestion, "Winkiss Frage:")
        If iRet = vbYes Then
            delRETPRINT
            fuellelist List5
        End If
        
    
        
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command6_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    Dim bWarning As Boolean
    Dim sMessText As String
    
    Select Case Index
        Case 0
            gsARTNR = Label2(2).Caption
            frmWKL78.Show 1
        Case 1
            einlesenausmdeVorschlag
        Case 2
            Text1_KeyUp 2, vbKeyF2, 0
        Case 3
            VerarbeiteKumuliereBudni
        Case 4 'drucken endgültig für Budni
        
        
            UpdateKundnrfromLinr Label7(3).Caption
        
            reportbildschirm "umv1", "aWKL112d"
        Case 6
            Frame2.Visible = False
        Case 7  'versenden an Budni
            Dim iRet As Integer
        
            If iBudniRetourenzaehler > 0 Then
                iRet = MsgBox("Möchten Sie diese Retoure zum  " & iBudniRetourenzaehler + 1 & ". Mal abschicken?", vbYesNo + vbDefaultButton2 + vbCritical, "Winkiss Frage:")
                If iRet = vbNo Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            
            End If
            
            sendenAnBudni Text1(5).Text
            iBudniRetourenzaehler = iBudniRetourenzaehler + 1
            
            
        Case 8
            Drucke_Nicht
        Case 12
        
            'sicherheitscheck auf Bestellnr wenn budni
            If Check21.Value = vbChecked Then
            
                sSQL = "Select * from RETOURKB where Status = 'vorhanden' and libesnr = '' "
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    'Achtung
                    bWarning = True
                End If
                rsrs.Close
            
            End If
        
            If bWarning = False Then
                einlesenausMDE Text1(2).Text, Check21.Value
            Else
                sMessText = "Es sind Artikel enthalten, denen keine Bestellnummer zugeordnet ist." & vbCrLf & vbCrLf
                sMessText = sMessText & "Bitte in der Artikelbearbeitung die Bestellnummer nachtragen und den Vorgang wiederholen!"
                MsgBox sMessText, vbCritical, "Winkiss Hinweis:"
            End If
            
        Case 11  'zurück Dateien Zentrale 1
            Frame8.Visible = False
            Frame5.Visible = True
            Text1(2).SetFocus
            anzeigeNew "normal", "", Label5
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command7_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0
            Unload frmWKL112
            
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
            
            
            
           
            Command6(3).Visible = False
            
        Case Is = 3
        
            anzeige "", "", Label5
            
            If Text1(2).Text = "" Then
                anzeige "rot", "Geben Sie bitte einen Lieferanten an!", Label5
                Text1(2).SetFocus
                Exit Sub
            End If
            
            If IsNumeric(Text1(2).Text) = False Then
                anzeige "rot", "Geben Sie bitte einen Lieferanten an!", Label5
                Text1(2).SetFocus
                Exit Sub
            End If
        
            If MDEeinlesenOhneLinr(Label5, txtStatus, picprogress, frmWKL19) = False Then
                anzeige "rot", "Es konnten keine Daten aus dem MDE - Gerät ausgelesen werden.", Label5
            Else
                Frame5.Visible = False
                Frame8.Visible = True
                anzeigeNew "normal", "", Label5
                MdeVerarbeitung1 Text1(2).Text
                
                
                anzeige "normal", "Die Artikel, die auf der linken Seite angezeigt sind, können jetzt retourniert werden.", Label7(0)
                
                
                 'ist es Budni?
                Dim rsLi            As DAO.Recordset
                Dim sSQL            As String

                Check21.Visible = False
                Check21.Value = vbUnchecked
    
                sSQL = "select KUNDNR,linr from LISRT where FORMAT = 'EDIBUDNI' and linr = " & Text1(2).Text
                Set rsLi = gdBase.OpenRecordset(sSQL)
                If Not rsLi.EOF Then
                    Check21.Visible = True
                    Check21.Value = vbChecked
                End If
                rsLi.Close: Set rsLi = Nothing
                
                
                
            End If
            
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub einlesenausMDE(cLinr As String, bimAnschlussAnBudni As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rs          As Recordset
    Dim lCounter    As Long
    Dim dBestand    As Double
    Screen.MousePointer = 11
    
    loeschNEW "RETOURY", gdBase
    CreateTableT2 "RETOURY", gdBase
    
    sSQL = "Insert into RETOURY Select sum(r.bestvor) as lmaxanz"
    sSQL = sSQL & ", r.ARTNR "
    sSQL = sSQL & ", r.BEZEICH "
    sSQL = sSQL & ", a.LINR "
    sSQL = sSQL & ", r.LPZ "
    sSQL = sSQL & ", r.LIBESNR "
    sSQL = sSQL & ", r.KVKPR1 "
    sSQL = sSQL & ", r.BESTAND "
    sSQL = sSQL & ", r.FILIALE "
    sSQL = sSQL & ", r.LEKPR "
    sSQL = sSQL & ", r.SEK "
    sSQL = sSQL & ", 0 as BESTANDN "
    sSQL = sSQL & ", 0 as FARBNR "
    sSQL = sSQL & ", 0 as FARBwert "
    sSQL = sSQL & ", 0 as FARBwertS "
    sSQL = sSQL & ", '' as FARBTEXT "
    sSQL = sSQL & ", '' as KUNDNR "
    sSQL = sSQL & " from RETOURKB r"
    
    If cLinr <> "" Then
        sSQL = sSQL & " inner join Artlief a on r.artnr = a.artnr "
    End If
    sSQL = sSQL & " where r.Status = 'vorhanden' "
    
    If cLinr <> "" Then
        sSQL = sSQL & " and a.linr = " & cLinr
    End If
    
    sSQL = sSQL & " group by "
    sSQL = sSQL & " r.ARTNR "
    sSQL = sSQL & ", r.BEZEICH "
    sSQL = sSQL & ", a.LINR "
    sSQL = sSQL & ", r.LPZ "
    sSQL = sSQL & ", r.LIBESNR "
    sSQL = sSQL & ", r.KVKPR1 "
    sSQL = sSQL & ", r.BESTAND "
    sSQL = sSQL & ", r.FILIALE "
    sSQL = sSQL & ", r.LEKPR "
    sSQL = sSQL & ", r.SEK "
    sSQL = sSQL & " order by a.linr,r.lpz "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Update RETOURY inner join Artikel on RETOURY.ARTNR = Artikel.Artnr "
    sSQL = sSQL & " set  RETOURY.FARBNR = val(ARTIKEL.awm) "
    gdBase.Execute sSQL, dbFailOnError

    BringFarbeInsSpiel "RETOURY", gdBase
    
    Set rs = gdBase.OpenRecordset("RETOURY")
    
    If rs.EOF Then
    
        If cLinr <> "" Then
            MsgBox "Keine Artikeldaten zur Verarbeitung vorhanden (eventuell gehören diese Artikel nicht zum Lieferant: " & cLinr & ")", vbCritical, "Winkiss Hinweis:"
        
        Else
            MsgBox "Keine Artikeldaten zur Verarbeitung vorhanden", vbCritical, "Winkiss Hinweis:"
        End If
    
        Screen.MousePointer = 0
        rs.Close: Set rs = Nothing
        Exit Sub
    End If
    
    pbr1.Max = 50
    pbr1.Visible = True
    
    lCounter = 0
    rs.MoveFirst
    If Not rs.EOF Then
        anzeigeNew "normal", "Die Artikelretoure wird jetzt eingelesen...", Label5
        Do While Not rs.EOF
            If lCounter = 50 Then
                lCounter = 0
            End If
            lCounter = lCounter + 1
            pbr1.Value = lCounter
            
            If Not IsNull(rs!artnr) Then
                If Not IsNull(rs!lmaxanz) Then
                    InsertRetoure_MDE rs!artnr, CLng(rs!lmaxanz)
                End If
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    
    
    pbr1.Visible = False

    anzeigeNew "normal", "Die Retournierung wurde erfolgreich durchgeführt.", Label7(0)

    Command6(12).Visible = False
    Check21.Visible = False
    Check21.Value = vbUnchecked

    Screen.MousePointer = 0
    
    If bimAnschlussAnBudni = True Then
        AnzeigenRetoury cLinr
    Else
    
        UpdateKundnrfromLinr cLinr
        reportbildschirm "umv1", "aWKL112d"
    End If
    
    

    Exit Sub
LOKAL_ERROR:

    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "einlesenausmde"
        Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub UpdateKundnrfromLinr(cLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsLi        As Recordset
    Dim cKundnr     As String
    
    sSQL = "select KUNDNR from LISRT where linr = " & cLinr
    Set rsLi = gdBase.OpenRecordset(sSQL)
    If Not rsLi.EOF Then
        cKundnr = Trim(rsLi!Kundnr)
    End If
    rsLi.Close: Set rsLi = Nothing
    
    If Val(cKundnr) > 0 Then
        sSQL = " Update RETOURY set kundnr = '" & cKundnr & "'"
        gdBase.Execute sSQL, dbFailOnError
    End If
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UpdateKundnrfromLinr"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub sendenAnBudni(sBelegnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim sBudniKundnr    As String
    Dim sBudniLinr      As String
    Dim rsLi            As DAO.Recordset
    Dim cQuelle         As String
    Dim cZiel           As String
    Dim cDatname        As String
    Dim lRet            As Long
    Dim lfail           As Long
    Dim ctmp            As String
    
    Screen.MousePointer = 11
    
    sBudniLinr = ""

    Set rsLi = gdBase.OpenRecordset("select linr from LISRT where FORMAT = 'EDIBUDNI' ")
    If Not rsLi.EOF Then
        sBudniLinr = Trim(rsLi!linr)
    End If
    rsLi.Close: Set rsLi = Nothing
    
    sBudniKundnr = ""
    
    sSQL = "select KUNDNR,linr from LISRT where FORMAT = 'EDIBUDNI' and linr = " & sBudniLinr
    Set rsLi = gdBase.OpenRecordset(sSQL)
    If Not rsLi.EOF Then
        sBudniKundnr = Trim(rsLi!Kundnr)
    End If
    rsLi.Close: Set rsLi = Nothing
    
    If Val(sBudniKundnr) > 0 Then
    
        erstelle_RetourenDatei sBudniLinr, sBudniKundnr, sBelegnr
        
        cQuelle = App.Path
        If Right$(cQuelle, 1) <> "\" Then
            cQuelle = cQuelle & "\"
        End If
        cQuelle = cQuelle & "EDI\EDI.TXT"
        cQuelle = ShortPath(cQuelle)
        
        cZiel = gcDBPfad      'dabapfad + RETOURE
        If Right$(cZiel, 1) <> "\" Then
            cZiel = cZiel & "\"
        End If
        cZiel = ShortPath(cZiel)
    
        sBudniKundnr = String$(6 - Len(sBudniKundnr), "0") & sBudniKundnr
        
        cDatname = "BUDNIRET+1+"
        cDatname = cDatname & sBudniKundnr & "+"
        cDatname = cDatname & Format(DateValue(Now), "YYYYMMDD") & "+"
        cDatname = cDatname & Format(TimeValue(Now), "HHMMSS") & ".001"
            
        cZiel = cZiel & "\RETOURE\" & cDatname
            
        lRet = CopyFile(cQuelle, cZiel, lfail)
        
        cZiel = App.Path
        cZiel = ShortPath(cZiel)
        
        sBudniKundnr = String$(6 - Len(sBudniKundnr), "0") & sBudniKundnr
        
        cZiel = cZiel & "\EDI\" & cDatname
            
        lRet = CopyFile(cQuelle, cZiel, lfail)
    
        If lRet <> 0 Then
            Dim i As Integer
            Dim bmerke As Boolean
            bmerke = gbFTPautomatic
            gbFTPautomatic = True
            
            cZiel = App.Path
            cZiel = ShortPath(cZiel)
            
            Kill cZiel & "\EDI\EDI.TXT"
        
            giKissFtpMode = 35
            frmWKL38.Show 1

            gbFTPautomatic = bmerke
            
            Screen.MousePointer = 0
            
            ctmp = "Ihre Retoure wurde übertragen" & vbCrLf & vbCrLf
            ctmp = ctmp & "Drucken Sie sich bitte im Anschluss Ihre Retourendaten aus!"
            
            MsgBox ctmp, vbInformation + vbOKOnly, "Winkiss Hinweis:"
            
            anzeige "normal", "Die Retournierung wurde erfolgreich durchgeführt.", Label7(1)
    
            Command6(12).Visible = False

            UpdateKundnrfromLinr sBudniLinr
    
            reportbildschirm "umv1", "aWKL112d"
            
        Else
            Screen.MousePointer = 0
            MsgBox "Die Retourendatei konnte nicht kopiert werden.", vbCritical, "Winkiss Hinweis:"
            anzeige "rot", "Fehler", Label7(1)
        End If
    Else
        Screen.MousePointer = 0
        anzeige "rot", "Kundennummer bei Lieferant " & sBudniLinr & " fehlt.", Label7(1)
    End If
    
    Exit Sub
LOKAL_ERROR:

    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "sendenAnBudni"
        Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub VerarbeiteKumuliereBudni()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As DAO.Recordset
    Dim sBudniLinr      As String
    Dim rsLi            As DAO.Recordset

    sBudniLinr = ""

    sSQL = "select linr from LISRT where FORMAT = 'EDIBUDNI' "
    Set rsLi = gdBase.OpenRecordset(sSQL)
    If Not rsLi.EOF Then
        sBudniLinr = Trim(rsLi!linr)
    End If
    rsLi.Close: Set rsLi = Nothing
    
    If sBudniLinr = "" Then
        anzeige "rot", "kein gültiger Budni - Lieferant vorhanden", Label7(9)
        Exit Sub
    End If
    
    anzeige "normal", sBudniLinr, Label7(3)
    
    
    Screen.MousePointer = 11
    
    loeschNEW "RETOURY", gdBase
    CreateTableT2 "RETOURY", gdBase
    
    sSQL = "Insert into RETOURY Select sum(r.Menge) as lmaxanz"
    sSQL = sSQL & ", r.ARTNR "
    sSQL = sSQL & ", r.BEZEICH "
    sSQL = sSQL & ", a.LINR "
    sSQL = sSQL & ", 0 as LPZ "
    sSQL = sSQL & ", a.LIBESNR "
    sSQL = sSQL & ", 0 as KVKPR1 "
    sSQL = sSQL & ", 0 as BESTAND "
    sSQL = sSQL & ", 0 as FILIALE "
    sSQL = sSQL & ", r.LEKPR "
    sSQL = sSQL & ", 0 as SEK "
    sSQL = sSQL & ", 0 as BESTANDN "
    sSQL = sSQL & ", 0 as FARBNR "
    sSQL = sSQL & ", 0 as FARBwert "
    sSQL = sSQL & ", 0 as FARBwertS "
    sSQL = sSQL & ", '' as FARBTEXT "
    sSQL = sSQL & ", '' as KUNDNR "
    sSQL = sSQL & " from RETPRINT r"
    sSQL = sSQL & " inner join Artlief a on r.artnr = a.artnr "
    sSQL = sSQL & " where a.linr = " & sBudniLinr
    sSQL = sSQL & " group by "
    sSQL = sSQL & " r.ARTNR "
    sSQL = sSQL & ", r.BEZEICH "
    sSQL = sSQL & ", a.LINR "
    sSQL = sSQL & ", a.LIBESNR "
    sSQL = sSQL & ", r.LEKPR "
    sSQL = sSQL & " order by a.linr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Update RETOURY inner join Artikel on RETOURY.ARTNR = Artikel.Artnr "
    sSQL = sSQL & " set  RETOURY.FARBNR = val(ARTIKEL.awm) "
    sSQL = sSQL & " ,  RETOURY.KVKPR1 = ARTIKEL.KVKPR1 "
    sSQL = sSQL & " ,  RETOURY.SEK = ARTIKEL.EKPR "
    sSQL = sSQL & " , RETOURY.lpz = Artikel.lpz "
    gdBase.Execute sSQL, dbFailOnError
    
    

    BringFarbeInsSpiel "RETOURY", gdBase
    
    Screen.MousePointer = 0
    
    AnzeigenRetoury sBudniLinr
    
    
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "VerarbeiteKumuliereBudni"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub AnzeigenRetoury(sBudniLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As DAO.Recordset
    
    Screen.MousePointer = 11
    
    sSQL = "Delete from RETOURY where libesnr ='' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from RETOURY where libesnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    Frame2.Visible = True
    iBudniRetourenzaehler = 0
    
    Label7(3).Caption = sBudniLinr
    
    anzeige "normal", "Artikel, die nicht enthalten sind, gehören nicht zu dem Lieferanten: " & sBudniLinr & " oder haben keine Bestellnummer hinterlegt.", Label7(9)

    Set rsrs = gdBase.OpenRecordset("RETOURY")
    List8.Clear
    List8.AddItem " Artnr  Artikelbezeichnung                    Menge       L-EK        BestellNr      Lieferant"
    List2.Clear
    List2.Visible = False
    
    Dim cFeld   As String
    Dim cLBSatz As String
    Dim lGesbestand As Long
    lGesbestand = 0
    
    If Not rsrs.EOF Then
    
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            End If
    
            cLBSatz = cFeld & Space(8 - Len(cFeld))
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(38 - Len(cFeld))
            
            If Not IsNull(rsrs!lmaxanz) Then
                cFeld = rsrs!lmaxanz
                lGesbestand = lGesbestand + Val(cFeld)
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(12 - Len(cFeld))
            
            If Not IsNull(rsrs!lekpr) Then
                cFeld = Format(rsrs!lekpr, "###0.00")
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(12 - Len(cFeld))
            
            If Not IsNull(rsrs!LIBESNR) Then
                cFeld = rsrs!LIBESNR
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(15 - Len(cFeld))
            
            If Not IsNull(rsrs!linr) Then
                cFeld = rsrs!linr
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(6 - Len(cFeld))
            
            List2.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    Else
        anzeige "rot", "Keine Artikeldaten zur Verarbeitung vorhanden (eventuell gehören diese Artikel nicht zum Lieferant: " & sBudniLinr & " oder es ist keine Bestellnummer eingetragen)", Label7(9)
    End If
    
    List2.Visible = True
    rsrs.Close: Set rsrs = Nothing
    
    If List2.ListCount = 0 Then
        anzeige "rot", "keine Artikel", Label7(12)
    Else
        anzeige "normal", List2.ListCount & " verschiedene Artikel mit einer Gesamtmenge: " & lGesbestand, Label7(12)
    End If
    
    Screen.MousePointer = 0
    



    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AnzeigenRetoury"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub erstelle_RetourenDatei(cLinr As String, sBudniKundnr As String, sBelegnr As String)
    On Error GoTo LOKAL_ERROR

    Dim lrow            As Long
    Dim lPos            As Long
    Dim lWert           As Long
    Dim cSatz           As String
    Dim cSQL            As String
    Dim cBestMenge      As String
    Dim cEAN            As String
    Dim cLiBesNr        As String
    Dim cPfad           As String
    Dim iFileNr         As Integer
    Dim rsrs            As Recordset
    Dim rsArt           As Recordset
    Dim sTime           As String
    Dim ctmp            As String
    Dim cArtNr          As String
    Dim cMinMen         As String
    Dim sLiefArtnr      As String
    Dim lLFNR           As Long
    
    lLFNR = 0

    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    sTime = Left(sTime, 5)

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM")

    ctmp = ctmp & sTime
    ctmp = SwapStr(ctmp, ".", "")
    ctmp = SwapStr(ctmp, ":", "")

    cPfad = gcPfad    'AppPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "EDI\"
    
    Kill cPfad & "*.*"
    
    Kill cPfad & "EDI.txt"

    iFileNr = FreeFile
    Open cPfad & "EDI.txt" For Binary As #iFileNr
    
    sBudniKundnr = String$(6 - Len(sBudniKundnr), "0") & sBudniKundnr
    sBelegnr = String$(10 - Len(sBelegnr), "0") & sBelegnr
    
    
    cSQL = "Select * from RETOURY  "
    Set rsArt = gdBase.OpenRecordset(cSQL)
    If Not rsArt.EOF Then
        rsArt.MoveFirst
        
        Do While Not rsArt.EOF
                    
            cBestMenge = ""
            If Not IsNull(rsArt!lmaxanz) Then
                cBestMenge = Trim(rsArt!lmaxanz)
            End If
            
            cArtNr = ""
            If Not IsNull(rsArt!artnr) Then
                cArtNr = Trim(rsArt!artnr)
            End If
            
            If Val(cBestMenge) > 0 Then
        
                cSQL = "Select * from ARTLIEF where ARTNR = " & cArtNr & " and LINR = " & cLinr & " "
                Set rsrs = gdBase.OpenRecordset(cSQL)
                If Not rsrs.EOF Then
                    rsrs.MoveFirst
    
                    cLiBesNr = ""
                    
                    If Not IsNull(rsrs!LIBESNR) Then
                        cLiBesNr = Trim(rsrs!LIBESNR)
                    End If
                    
                End If
                rsrs.Close
                    
                cSQL = "Select EAN,ean2,ean3 from ARTIKEL where ARTNR = " & cArtNr
                Set rsrs = gdBase.OpenRecordset(cSQL)
                If Not rsrs.EOF Then
                    rsrs.MoveFirst
                    cEAN = ""
                    If Not IsNull(rsrs!EAN) Then
                        cEAN = Trim(rsrs!EAN)
                    End If
                    
                    If cEAN = "" Then
                        If Not IsNull(rsrs!EAN2) Then
                            cEAN = Trim(rsrs!EAN2)
                        End If
                    End If
                    
                    If cEAN = "" Then
                        If Not IsNull(rsrs!EAN3) Then
                            cEAN = Trim(rsrs!EAN3)
                        End If
                    End If
                    
                End If
                rsrs.Close
                
                lLFNR = lLFNR + 1
               
                cSatz = lLFNR & vbTab
                cSatz = cSatz & cLiBesNr & vbTab
                cSatz = cSatz & cBestMenge & vbTab
                cSatz = cSatz & cEAN & vbTab
                cSatz = cSatz & cArtNr & vbTab
                
                
                cSatz = cSatz & sBelegnr & vbTab
                cSatz = cSatz & Format(DateValue(Now), "YYYYMMDD") & vbTab
                cSatz = cSatz & sBudniKundnr & vbTab
                cSatz = cSatz & vbCrLf
                
                lPos = LOF(iFileNr)
                lPos = lPos + 1
                Put #iFileNr, lPos, cSatz
                
            End If
        rsArt.MoveNext
        Loop
    End If
    rsArt.Close
                
    
    cSatz = "1" & vbTab & vbTab & "00" & vbTab & "1111111111111" & vbTab & "111111" & vbTab & "111111" & vbTab & "11111111" & vbTab & "111111" & vbCrLf
            
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    cSatz = "1" & vbTab & vbTab & "00" & vbTab & "1111111111111" & vbTab & "111111" & vbTab & "111111" & vbTab & "11111111" & vbTab & "111111" & vbCrLf
    
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    cSatz = "1" & vbTab & vbTab & "00" & vbTab & "1111111111111" & vbTab & "111111" & vbTab & "111111" & vbTab & "11111111" & vbTab & "111111" & vbCrLf
    
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz

    Close iFileNr

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "erstelle_RetourenDatei"
        Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten."
    
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Sub einlesenausmdeVorschlag()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rs          As Recordset
    Dim lCounter    As Long
    Dim dBestand    As Double
    Screen.MousePointer = 11
    
    loeschNEW "RETOURVOR", gdBase
    
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
    
    sSQL = sSQL & " into RETOURVOR from RETOURKB where Status = 'vorhanden' "
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
    
    Set rs = gdBase.OpenRecordset("RETOURVOR")
    
    If rs.EOF Then
        anzeigeNew "rot", "Keine Artikeldaten zur Druckansicht vorhanden", Label5
        Screen.MousePointer = 0
        rs.Close: Set rs = Nothing
        Exit Sub
    End If
    

    Screen.MousePointer = 0
    anzeigeNew "normal", "", Label5
    
    reportbildschirm "umv1", "aWKL112g"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "einlesenausmdeVorschlag"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Drucke_Nicht()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rs          As Recordset
    Dim lCounter    As Long
    Dim dBestand    As Double
    Screen.MousePointer = 11
    
    loeschNEW "RETOURVOR", gdBase
    
    sSQL = "Select * into RETOURVOR from RETOURKB where left(Status,5) = 'nicht' "
    gdBase.Execute sSQL, dbFailOnError
    
    Set rs = gdBase.OpenRecordset("RETOURVOR")
    
    If rs.EOF Then
        anzeigeNew "rot", "Keine Artikeldaten zur Druckansicht vorhanden", Label5
        Screen.MousePointer = 0
        rs.Close: Set rs = Nothing
        Exit Sub
    End If
    

    Screen.MousePointer = 0
    anzeigeNew "normal", "", Label5
    
    reportbildschirm "umv1", "aWKL112c"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Drucke_Nicht"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
    
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
        
        
        If Budnivorhanden = True Then
            Command6(3).Visible = True
        End If

    ElseIf Option2(2).Value = True Then     'Mde
        Frame5.Visible = True
        
        List5.Visible = False
        List6.Visible = False
        Label11(3).Visible = False
        Frame1.Visible = False
        Command2(21).Visible = False
        
        
        
        
        Command6(3).Visible = False
        
        Text1(2).SetFocus
        
        
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
    
    
    
    
    End If
    
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "RETOURVOR", gdBase
    loeschNEW "RETOURKB", gdBase
    loeschNEW "RETOURY", gdBase
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
    
    Frame2.Top = 840
    Frame2.Left = 120
    Frame2.Height = 7695
    Frame2.Width = 11655
   
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
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub LeereDialogWKL15()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = ""
    
    If Check8.Value = vbUnchecked Then
        Text1(1).Text = ""
    End If
    
    If Option1(1).Value Then
        Text1(4).Text = ""
        Label2(4).Caption = ""
    End If
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
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
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
    
    bDebug = False
    bgefunden = True
    bEAN = True
    
    cSuch = Text1(0).Text
    cSuch = Trim$(cSuch)
    
    If cSuch = "" Then
        MsgBox "Bitte Wert eingeben!", vbCritical, "STOP!"
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
        
        
        cSQL = "Select B.ARTNR, A.BEZEICH, B.LINR, A.BESTAND, A.VKPR, A.KVKPR1, A.MINBEST, B.LIBESNR, A.EAN "
        cSQL = cSQL & "from ARTIKEL A, ARTLIEF B where A.ARTNR = B.ARTNR "
    End If
    
    If Len(cSuch) <= 6 Then
        cSQL = "Select B.ARTNR, A.BEZEICH, B.LINR, A.BESTAND, A.VKPR, A.KVKPR1, A.MINBEST, B.LIBESNR, A.EAN "
        cSQL = cSQL & "from ARTIKEL A, ARTLIEF B where B.ARTNR = " & cSuch & " and A.ARTNR = B.ARTNR "
    Else
        If Len(cSuch) <= 8 And Left(cSuch, 1) = "2" Then ' Or Left(cSuch, 1) = "0") Then
            cSuch = Mid(cSuch, 2, 6)
            cSQL = cSQL & "and B.ARTNR = " & cSuch & " "
        Else
            If bEAN Then
            
                'Ean ist richtig!
                'aber Zeitung?
            
                If Len(cSuch) >= 13 And Left(cSuch, 3) = "419" Then
                    cSQL = cSQL & "and B.ARTNR = " & Zeitungs_EAN_ZU_Artnr(cSuch, "V") & " "
                    
                ElseIf Len(cSuch) >= 13 And Left(cSuch, 3) = "419" Then
                    cSQL = cSQL & "and B.ARTNR = " & Zeitungs_EAN_ZU_Artnr(cSuch, "V") & " "
                    
                ElseIf Len(cSuch) >= 13 And Left(cSuch, 3) = "434" Then
                    cSQL = cSQL & "and B.ARTNR = " & Zeitungs_EAN_ZU_Artnr(cSuch, "E") & " "
                ElseIf Len(cSuch) >= 13 And Left(cSuch, 3) = "434" Then
                    cSQL = cSQL & "and B.ARTNR = " & Zeitungs_EAN_ZU_Artnr(cSuch, "E") & " "
                Else
                    'hier wie immer
                    cSQL = cSQL & "and (A.EAN = '" & cSuch & "' "
                    cSQL = cSQL & "or A.EAN2 = '" & cSuch & "' "
                    cSQL = cSQL & "or A.EAN3 = '" & cSuch & "' )"
                    
                End If
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
            Else

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
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
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
        SucheArtikelWKL15
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
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    Dim lcount As Long
    Dim ctmp As String
    Dim lBestneu As Long
    Dim lBestalt As Long
    Dim lAbgang As Long
    Dim iRet As Long
    Dim bneg As Boolean
    bneg = True
    
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
                    Screen.MousePointer = 0
                    Exit Sub
                Else
                    Text1(0).Text = Label2(2).Caption
                End If
            End If
            
            
            lBestalt = Label2(1).Caption
            lAbgang = Text1(1).Text
            lBestneu = lBestalt - lAbgang
            
            
            
            If lAbgang > 999 Then
                iRet = MsgBox("Diese ungewöhnlich hohe Abgangsmenge von " & lAbgang & " zulassen?", vbYesNo + vbQuestion, "Winkiss Frage:")
                If iRet = vbNo Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
            
            
            


            bneg = True
            If lBestneu < 0 Then
                
                iRet = MsgBox("negativen Bestandseintrag zulassen? neuer Bestand: " & lBestneu, vbYesNo + vbQuestion, "Winkiss Frage:")
                If iRet = vbNo Then
                    bneg = False
                End If
            End If
            
            InsertRetoureWKL20 Label2(2).Caption, bneg
            
            LeereDialogWKL15
            Text1(0).SetFocus

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
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub fuellelist(lst As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim cFeld   As String
    Dim cLBSatz As String
    
    cSQL = "Select * from RETPRINT order by lfnr desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    lst.Clear
    lst.Visible = False
    
    If Not rsrs.EOF Then
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            End If
    
            cLBSatz = cFeld & Space(8 - Len(cFeld))
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(38 - Len(cFeld))
            
            If Not IsNull(rsrs!Menge) Then
                cFeld = rsrs!Menge
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(12 - Len(cFeld))
            
            If Not IsNull(rsrs!lekpr) Then
                cFeld = Format(rsrs!lekpr, "###0.00")
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(12 - Len(cFeld))
            
            If Not IsNull(rsrs!BESTANDneu) Then
                cFeld = rsrs!BESTANDneu
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(6 - Len(cFeld))
            
            lst.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    lst.Visible = True
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellelist"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub InsertRetoureWKL20(cArtNr As String, bNegativzulassen As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim lLinr           As Long
    Dim dEkpr           As Double
    Dim dWert           As Double
    Dim ctmp            As String
    Dim cSQL            As String
    Dim rsKJ            As Recordset
    Dim rsArt           As Recordset
    Dim rsRP            As Recordset
    Dim lAbgang         As Long
    Dim lBestandneu     As Long
    
    '*** KASSJOUR-Felder ****
    Dim cKJBezeich      As String
    Dim cKJMenge        As String
    Dim cKJAZeit        As String
    Dim cKJLPZ          As String
    Dim cKJAGN          As String
    Dim cKJEAN          As String
    Dim cKJMwst         As String
    Dim lBestandAlt     As Long
    Dim dKJPreis        As Double
    Dim dKJVkpr         As Double
    Dim lKJADate        As Long
    Dim cKJMopreis      As String
    
    cSQL = "Select * from Retoure"
    Set rsKJ = gdBase.OpenRecordset(cSQL)
    
    cSQL = "Select * from RETPRINT"
    Set rsRP = gdBase.OpenRecordset(cSQL)
    
    cKJBezeich = ""
    cKJMopreis = ""
    cKJMenge = ""
    dKJPreis = 0
    lKJADate = 0
    cKJAZeit = ""
    cKJLPZ = ""
    cKJAGN = ""
    cKJEAN = ""
    cKJMwst = ""
    dEkpr = 0
    dKJVkpr = 0
    lLinr = 0
        
    cSQL = "Select * from Artikel where Artnr = " & cArtNr
    Set rsArt = gdBase.OpenRecordset(cSQL)
    
    If Not rsArt.EOF Then
    
        If Not IsNull(rsArt!BEZEICH) Then
            cKJBezeich = rsArt!BEZEICH
        Else
            cKJBezeich = ""
        End If
        
        If Not IsNull(rsArt!BESTAND) Then
            lBestandAlt = rsArt!BESTAND
        Else
            lBestandAlt = 0
        End If
        
        If Not IsNull(rsArt!LPZ) Then
            cKJLPZ = rsArt!LPZ
        Else
            cKJLPZ = ""
        End If
        
        If Not IsNull(rsArt!AGN) Then
            cKJAGN = rsArt!AGN
        Else
            cKJAGN = ""
        End If
        
        If Not IsNull(rsArt!EAN) Then
            cKJEAN = rsArt!EAN
        Else
            cKJEAN = ""
        End If
        
        If Not IsNull(rsArt!vkpr) Then
            dKJVkpr = rsArt!vkpr
        Else
            dKJVkpr = 0
        End If
        
        If Not IsNull(rsArt!MWST) Then
            cKJMwst = rsArt!MWST
        Else
            cKJMwst = "V"
        End If
        
        If Not IsNull(rsArt!KVKPR1) Then
            dKJPreis = rsArt!KVKPR1
        Else
            dKJPreis = 0
        End If
        
        If Not IsNull(rsArt!linr) Then
            lLinr = rsArt!linr
        Else
            lLinr = 0
        End If
    Else
        dEkpr = 0
    End If
    rsArt.Close: Set rsArt = Nothing
    
    lAbgang = CLng(Text1(1).Text)
    lBestandneu = lBestandAlt - lAbgang
    
    'Achtung neu: Retouren ins negative nicht möglich
    If bNegativzulassen = True Then
    
    Else
        If lBestandneu < 0 Then lBestandneu = 0
    End If
    
    Bestandsveraenderung cArtNr, lBestandneu, "BEA Retoure"
    ABinBESTAKT cArtNr, lAbgang, "BEA Retoure"
    
    dEkpr = ermLEKPR(cArtNr, lLinr)
    
    lKJADate = Fix(Now)
    cKJAZeit = Format$(Now, "HH:MM:SS")
                  
    rsKJ.AddNew
    rsKJ!artnr = Val(cArtNr)
    rsKJ!BEZEICH = cKJBezeich
    rsKJ!Menge = lAbgang
    rsKJ!Preis = dKJPreis
    rsKJ!ADATE = lKJADate
    rsKJ!AZEIT = cKJAZeit
    rsKJ!BEDIENER = gcBedienerNr
    rsKJ!Kundnr = 0
    rsKJ!FILIALE = CInt(gcFilNr)
    rsKJ!kasnum = gcKasNum
    rsKJ!linr = lLinr
    rsKJ!LPZ = Val(cKJLPZ)
    rsKJ!AGN = Val(cKJAGN)
    rsKJ!EAN = cKJEAN
    rsKJ!MWST = cKJMwst
    rsKJ!ekpr = dEkpr
    rsKJ!vkpr = dKJVkpr
    rsKJ!BELEGNR = 9999
    rsKJ!best1 = lBestandAlt
    rsKJ!MOPPREIS = 0
    rsKJ!SENDOK = False
    rsKJ.Update
    
    rsRP.AddNew
    rsRP!artnr = Val(cArtNr)
    rsRP!BEZEICH = cKJBezeich
    rsRP!Menge = lAbgang
    rsRP!linr = lLinr
    rsRP!lekpr = dEkpr
    rsRP!BESTAND = lBestandAlt
    rsRP!BESTANDneu = lBestandneu
    rsRP.Update
    
    rsRP.Close
    rsKJ.Close
    
    fuellelist List5
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InsertRetoureWKL20"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub InsertRetoure_MDE(cArtNr As String, lAbgang As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim lLinr           As Long
    Dim dEkpr           As Double
    Dim dWert           As Double
    Dim ctmp            As String
    Dim cSQL            As String
    Dim rsKJ            As Recordset
    Dim rsArt           As Recordset
    Dim lBestandneu     As Long
    
    '*** KASSJOUR-Felder ****
    Dim cKJBezeich      As String
    Dim cKJMenge        As String
    Dim cKJAZeit        As String
    Dim cKJLPZ          As String
    Dim cKJAGN          As String
    Dim cKJEAN          As String
    Dim cKJMwst         As String
    Dim lBestandAlt     As Long
    Dim dKJPreis        As Double
    Dim dKJVkpr         As Double
    Dim lKJADate        As Long
    Dim cKJMopreis      As String
    
    cSQL = "Select * from Retoure"
    Set rsKJ = gdBase.OpenRecordset(cSQL)
    
    cKJBezeich = ""
    cKJMopreis = ""
    cKJMenge = ""
    dKJPreis = 0
    lKJADate = 0
    cKJAZeit = ""
    cKJLPZ = ""
    cKJAGN = ""
    cKJEAN = ""
    cKJMwst = ""
    dEkpr = 0
    dKJVkpr = 0
    lLinr = 0
        
    cSQL = "Select * from Artikel where Artnr = " & cArtNr
    Set rsArt = gdBase.OpenRecordset(cSQL)
    
    If Not rsArt.EOF Then
    
        If Not IsNull(rsArt!BEZEICH) Then
            cKJBezeich = rsArt!BEZEICH
        Else
            cKJBezeich = ""
        End If
        
        If Not IsNull(rsArt!BESTAND) Then
            lBestandAlt = rsArt!BESTAND
        Else
            lBestandAlt = 0
        End If
        
        If Not IsNull(rsArt!LPZ) Then
            cKJLPZ = rsArt!LPZ
        Else
            cKJLPZ = ""
        End If
        
        If Not IsNull(rsArt!AGN) Then
            cKJAGN = rsArt!AGN
        Else
            cKJAGN = ""
        End If
        
        If Not IsNull(rsArt!EAN) Then
            cKJEAN = rsArt!EAN
        Else
            cKJEAN = ""
        End If
        
        If Not IsNull(rsArt!vkpr) Then
            dKJVkpr = rsArt!vkpr
        Else
            dKJVkpr = 0
        End If
        
        If Not IsNull(rsArt!MWST) Then
            cKJMwst = rsArt!MWST
        Else
            cKJMwst = "V"
        End If
        
        If Not IsNull(rsArt!KVKPR1) Then
            dKJPreis = rsArt!KVKPR1
        Else
            dKJPreis = 0
        End If
        
        If Not IsNull(rsArt!linr) Then
            lLinr = rsArt!linr
        Else
            lLinr = 0
        End If
    Else
        dEkpr = 0
    End If
    rsArt.Close: Set rsArt = Nothing
    
    lBestandneu = lBestandAlt - lAbgang
    
    'Achtung neu: Retouren ins negative nicht möglich
    If lBestandneu < 0 Then lBestandneu = 0
    
    Bestandsveraenderung cArtNr, lBestandneu, "BEA Retoure MDE"
    ABinBESTAKT cArtNr, lAbgang, "BEA Retoure MDE"
    
    dEkpr = ermLEKPR(cArtNr, lLinr)
    
    lKJADate = Fix(Now)
    cKJAZeit = Format$(Now, "HH:MM:SS")
                  
    rsKJ.AddNew
    rsKJ!artnr = Val(cArtNr)
    rsKJ!BEZEICH = cKJBezeich
    rsKJ!Menge = lAbgang
    rsKJ!Preis = dKJPreis
    rsKJ!ADATE = lKJADate
    rsKJ!AZEIT = cKJAZeit
    rsKJ!BEDIENER = gcBedienerNr
    rsKJ!Kundnr = 0
    rsKJ!FILIALE = CInt(gcFilNr)
    rsKJ!kasnum = gcKasNum
    rsKJ!linr = lLinr
    rsKJ!LPZ = Val(cKJLPZ)
    rsKJ!AGN = Val(cKJAGN)
    rsKJ!EAN = cKJEAN
    rsKJ!MWST = cKJMwst
    rsKJ!ekpr = dEkpr
    rsKJ!vkpr = dKJVkpr
    rsKJ!BELEGNR = 9999
    rsKJ!best1 = lBestandAlt
    rsKJ!MOPPREIS = 0
    rsKJ!SENDOK = False
    rsKJ.Update
    
    rsKJ.Close
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InsertRetoure_MDE"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command3_Click()
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKL112
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
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
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
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
    List6.AddItem " Artnr  Artikelbezeichnung                    Menge       EKPR        neuer Bestand"
    Screen.MousePointer = 0
    
    Frame1.Visible = False

    List5.Visible = False
    List6.Visible = False
    Label11(3).Visible = False
    Command2(21).Visible = False
    
    
    Command6(3).Visible = False
    
    
    
    fuellelist List5
    
    Option2(Leselast19Einstellung).Value = True
    Option2(2).Caption = Option2(2).Caption & " (" & gsMDEGERAET & ")"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Function Budnivorhanden() As Boolean
On Error GoTo LOKAL_ERROR

    Budnivorhanden = False

    Dim rsLi            As DAO.Recordset

    Set rsLi = gdBase.OpenRecordset("select linr from LISRT where FORMAT = 'EDIBUDNI' ")
    If Not rsLi.EOF Then
        Budnivorhanden = True
    End If
    rsLi.Close: Set rsLi = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Budnivorhanden"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Function
Private Sub anzeigeMDE()
    
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim sSQL        As String
    Dim cLBSatz     As String
    Dim cArtNr      As String
    Dim cBez        As String
    Dim ckPr        As String
    Dim cMenge      As String
    Dim cLinr       As String
    Dim cLiBesNr    As String
    Dim cLfnr       As String
    Dim iZaehler    As Integer
    
    List12.Clear
    List11.Clear
    List12.AddItem "Artnr  Bezeichnung      VK-Preis Menge  Lief.  BestellNr"
    
    sSQL = "Select * from RETOURKB where Status = 'vorhanden' "
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
            cLiBesNr = IIf(IsNull(rsrs!LIBESNR), "", rsrs!LIBESNR)
            
            cLBSatz = cArtNr & Space$(7 - Len(cArtNr))
            If Len(cBez) > 15 Then
                cBez = Left(cBez, 15) & "..."
            End If
            cLBSatz = cLBSatz & cBez & Space$(19 - Len(cBez))
            
            cLBSatz = cLBSatz & ckPr & Space$(7 - Len(ckPr))
            cLBSatz = cLBSatz & cMenge & Space$(7 - Len(cMenge))
            cLBSatz = cLBSatz & cLinr & Space$(7 - Len(cLinr))
            
            If cLiBesNr = "" Then cLiBesNr = "keine Angabe"
            cLBSatz = cLBSatz & cLiBesNr & Space$(14 - Len(cLiBesNr))
            
            List11.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
        Label7(5).Caption = "insgesamt: " & iZaehler & " verschiedene Artikel"
        Label7(5).Refresh
    End If
    rsrs.Close: Set rsrs = Nothing
    
    List4.Clear
    List3.Clear
    List4.AddItem "EANCODE/BEZ        Menge Reihenf Lieferant"
    
    sSQL = "Select * from RETOURKB where left(Status,5) = 'nicht' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    iZaehler = 0
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iZaehler = iZaehler + 1
            
            cBez = IIf(IsNull(rsrs!BEZEICH), "", rsrs!BEZEICH)
            cMenge = IIf(IsNull(rsrs!BESTVOR), "0", rsrs!BESTVOR)
            cLfnr = IIf(IsNull(rsrs!lfnr), "0", rsrs!lfnr)
            cLinr = IIf(IsNull(rsrs!linr), "0", rsrs!linr)
            
            If Len(cBez) > 15 Then
                cBez = Left(cBez, 15) & "..."
            End If
            cLBSatz = cLBSatz & cBez & Space$(19 - Len(cBez))
        
            cLBSatz = cBez & Space$(19 - Len(cBez))
            cLBSatz = cLBSatz & cMenge & Space$(6 - Len(cMenge)) & cLfnr
            cLBSatz = cLBSatz & cMenge & Space$(7 - Len(cMenge)) & cLinr
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
    
    Me.Refresh
    Frame5.Visible = False
    Me.Refresh
    Frame8.Visible = True
    Me.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "anzeigeMDE"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub delRETPRINT()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String

    cSQL = "Delete from RETPRINT "
    gdBase.Execute cSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "delRETPRINT"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MdeVerarbeitung1(sLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsMDE       As Recordset
    Dim rsFilB      As Recordset
    Dim rsArt       As Recordset
    Dim rsArtlief   As Recordset
    Dim seekEAN     As String
    Dim lBestand    As Long
    Dim lartnr      As Long
    
    Screen.MousePointer = 11
    
    Command6(12).Visible = False
    
    loeschNEW "RETOURKB", gdBase
    CreateTableT2 "RETOURKB", gdBase
    
    Set rsFilB = gdBase.OpenRecordset("RETOURKB")
    
    Set rsMDE = gdBase.OpenRecordset("mdeinh")
    If Not rsMDE.EOF Then
        rsMDE.MoveFirst
        Do While Not rsMDE.EOF
            If Not IsNull(rsMDE!eancode) Then
            
                seekEAN = Trim(rsMDE!eancode)
                seekEAN = checkean(seekEAN)
                
                
                If Len(seekEAN) = 11 Then
                    seekEAN = "0" & seekEAN
            
'                    sSQL = "select * from artikel where ean = '" & seekEAN & "'"
'                    sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
'                    sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                    
                    sSQL = "select * from artikel where ((ean = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean3 = '" & seekEAN & "')"
                    sSQL = sSQL & " or artikel.artnr in (Select artnr from artean_K where EAN = '" & seekEAN & "')) "
                    
                    
                    
                    
                ElseIf Len(seekEAN) = 8 Then
                    If Left(seekEAN, 1) = "2" Then
                        seekEAN = Mid$(seekEAN, 2, 6)
                        sSQL = "select * from artikel where artnr = " & seekEAN
                    Else
                        sSQL = "select * from artikel where ((ean = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean3 = '" & seekEAN & "')"
                        sSQL = sSQL & " or artikel.artnr in (Select artnr from artean_K where EAN = '" & seekEAN & "')) "
                        
                        
                        
                        
                        
                        
                        
                        
                    End If
                
                Else
                
                
                    'Ean ist richtig!
                    'aber Zeitung?
                
                    If Len(seekEAN) >= 13 And Left(seekEAN, 3) = "419" Then
                        sSQL = "select * from artikel where artnr = " & Zeitungs_EAN_ZU_Artnr(seekEAN, "V") & " "
                    ElseIf Len(seekEAN) >= 13 And Left(seekEAN, 3) = "419" Then
                        sSQL = "select * from artikel where artnr = " & Zeitungs_EAN_ZU_Artnr(seekEAN, "V") & " "
                    ElseIf Len(seekEAN) >= 13 And Left(seekEAN, 3) = "434" Then
                        sSQL = "select * from artikel where artnr = " & Zeitungs_EAN_ZU_Artnr(seekEAN, "E") & " "
                    ElseIf Len(seekEAN) >= 13 And Left(seekEAN, 3) = "434" Then
                        sSQL = "select * from artikel where artnr = " & Zeitungs_EAN_ZU_Artnr(seekEAN, "E") & " "
                    Else
'                        sSQL = "select * from artikel where ean = '" & seekEAN & "'"
'                        sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
'                        sSQL = sSQL & " or ean3 = '" & seekEAN & "'"


                        sSQL = "select * from artikel where ((ean = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean3 = '" & seekEAN & "')"
                        sSQL = sSQL & " or artikel.artnr in (Select artnr from artean_K where EAN = '" & seekEAN & "')) "
                    End If
                    
                End If

                lartnr = 0
                Set rsArt = gdBase.OpenRecordset(sSQL)
                If Not rsArt.EOF Then
                    
                    If Not IsNull(rsArt!artnr) Then
                        lartnr = rsArt!artnr
                    End If
                End If
                rsArt.Close: Set rsArt = Nothing
                
                If lartnr > 0 Then
                
                    sSQL = "select "
                    sSQL = sSQL & " a.ARTNR "
                    sSQL = sSQL & ", a.BEZEICH "
                    sSQL = sSQL & ", a.Bestand "
                    sSQL = sSQL & ", a.EKPR "
                    sSQL = sSQL & ", a.LPZ "
                    sSQL = sSQL & ", a.KVKPR1 "
                    sSQL = sSQL & ", l.Linr "
                    sSQL = sSQL & ", l.LIBESNR "
                    sSQL = sSQL & ", l.LEKPR "
                    sSQL = sSQL & " from Artlief L , Artikel A where L.artnr = A.artnr and L.artnr  = " & lartnr & " and L.Linr = " & sLinr
                    Set rsArtlief = gdBase.OpenRecordset(sSQL)
                            
                    lBestand = 0
                    
                    If Not rsArtlief.EOF Then 'hier die bekannten
                        rsFilB.AddNew
                        
                        If Not IsNull(rsArtlief!artnr) Then
                            rsFilB!artnr = rsArtlief!artnr
                        Else
                            rsFilB!artnr = 0
                        End If
                        
                        If Not IsNull(rsArtlief!BEZEICH) Then
                            rsFilB!BEZEICH = rsArtlief!BEZEICH
                        Else
                            rsFilB!BEZEICH = ""
                        End If
                        
                        If Not IsNull(rsArtlief!linr) Then
                            rsFilB!linr = rsArtlief!linr
                        Else
                            rsFilB!linr = 0
                        End If
                        
                        If Not IsNull(rsArtlief!LIBESNR) Then
                            rsFilB!LIBESNR = rsArtlief!LIBESNR
                        Else
                            rsFilB!LIBESNR = ""
                        End If
                        
                        If Not IsNull(rsArtlief!LPZ) Then
                            rsFilB!LPZ = rsArtlief!LPZ
                        Else
                            rsFilB!LPZ = 0
                        End If
                        
                        If Not IsNull(rsArtlief!ekpr) Then
                            rsFilB!sEK = rsArtlief!ekpr
                        Else
                            rsFilB!sEK = 0
                        End If
                        
                        If Not IsNull(rsArtlief!lekpr) Then
                            rsFilB!lekpr = rsArtlief!lekpr
                        Else
                            rsFilB!lekpr = 0
                        End If
                        
                        If Not IsNull(rsArtlief!KVKPR1) Then
                            rsFilB!KVKPR1 = rsArtlief!KVKPR1
                        Else
                            rsFilB!KVKPR1 = 0
                        End If
                        
                        If Not IsNull(rsArtlief!BESTAND) Then
                            rsFilB!BESTAND = rsArtlief!BESTAND
                            lBestand = rsArtlief!BESTAND
                        Else
                            rsFilB!BESTAND = 0
                            lBestand = 0
                        End If
                        
                        rsFilB!BESTVOR = rsMDE!Menge
                        rsFilB!BESTANDN = lBestand + CLng(rsMDE!Menge)
                        rsFilB!FILIALE = CByte(gcFilNr)
                        rsFilB!Status = "vorhanden"
                        rsFilB.Update
                    Else 'hier die die nicht zu linr passen
                    
                    
                    
                        sSQL = "select Top 1 "
                        sSQL = sSQL & " a.ARTNR "
                        sSQL = sSQL & ", a.BEZEICH "
                        sSQL = sSQL & ", a.Bestand "
                        sSQL = sSQL & ", a.EKPR "
                        sSQL = sSQL & ", a.LPZ "
                        sSQL = sSQL & ", a.KVKPR1 "
                        sSQL = sSQL & ", l.Linr "
                        sSQL = sSQL & ", l.LIBESNR "
                        sSQL = sSQL & ", l.LEKPR "
                        sSQL = sSQL & " from Artlief L , Artikel A where L.artnr = A.artnr and L.artnr  = " & lartnr & " and L.Linr <> " & sLinr
                        Set rsArt = gdBase.OpenRecordset(sSQL)
                                
                        lBestand = 0
                        
                        If Not rsArt.EOF Then 'hier die bekannten
                            rsFilB.AddNew
                            
                            If Not IsNull(rsArt!artnr) Then
                                rsFilB!artnr = rsArt!artnr
                            Else
                                rsFilB!artnr = 0
                            End If
                            
                            If Not IsNull(rsArt!BEZEICH) Then
                                rsFilB!BEZEICH = rsArt!BEZEICH
                            Else
                                rsFilB!BEZEICH = ""
                            End If
                            
                            If Not IsNull(rsArt!linr) Then
                                rsFilB!linr = rsArt!linr
                            Else
                                rsFilB!linr = 0
                            End If
                            
                            If Not IsNull(rsArt!LIBESNR) Then
                                rsFilB!LIBESNR = rsArt!LIBESNR
                            Else
                                rsFilB!LIBESNR = ""
                            End If
                            
                            If Not IsNull(rsArt!LPZ) Then
                                rsFilB!LPZ = rsArt!LPZ
                            Else
                                rsFilB!LPZ = 0
                            End If
                            
                            If Not IsNull(rsArt!ekpr) Then
                                rsFilB!sEK = rsArt!ekpr
                            Else
                                rsFilB!sEK = 0
                            End If
                            
                            If Not IsNull(rsArt!lekpr) Then
                                rsFilB!lekpr = rsArt!lekpr
                            Else
                                rsFilB!lekpr = 0
                            End If
                            
                            If Not IsNull(rsArt!KVKPR1) Then
                                rsFilB!KVKPR1 = rsArt!KVKPR1
                            Else
                                rsFilB!KVKPR1 = 0
                            End If
                            
                            If Not IsNull(rsArt!BESTAND) Then
                                rsFilB!BESTAND = rsArt!BESTAND
                                lBestand = rsArt!BESTAND
                            Else
                                rsFilB!BESTAND = 0
                                lBestand = 0
                            End If
                            
                            rsFilB!BESTVOR = rsMDE!Menge
                            rsFilB!BESTANDN = lBestand + CLng(rsMDE!Menge)
                            rsFilB!FILIALE = CByte(gcFilNr)
                            rsFilB!Status = "nicht bei " & sLinr
                            rsFilB.Update
                        
                            
                        End If
                        rsArt.Close: Set rsArt = Nothing
                    
                    
                    
                        
                    End If
                    rsArtlief.Close: Set rsArtlief = Nothing
                Else 'hier die völlig unbekannten
                    
                    rsFilB.AddNew
                    rsFilB!BEZEICH = seekEAN
                    rsFilB!BESTVOR = rsMDE!Menge
                    rsFilB!Status = "nicht vorhanden"
                    rsFilB!FILIALE = CByte(gcFilNr)
                    rsFilB.Update
                
                End If
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
        Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
        
        Fehlermeldung1
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
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherlast19Einstellung(i As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschapp "E112X"
    CreateTable "E112X", gdApp
    
    sSQL = "Insert into E112X (Ind) values (" & i & ")"
    gdApp.Execute sSQL, dbFailOnError
    
     
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherlast19Einstellung"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Leselast19Einstellung() As Byte
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    Leselast19Einstellung = 0
    
    If Not NewTableSuchenDBKombi("E112X", gdApp) Then
        CreateTable "E112X", gdApp
        
        sSQL = "Insert into E112X (Ind) values (0)"
        gdApp.Execute sSQL, dbFailOnError
    End If
    
    Set rsrs = gdApp.OpenRecordset("E112X")
    If Not rsrs.EOF Then
        Leselast19Einstellung = rsrs!ind
    End If
    rsrs.Close: Set rsrs = Nothing

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Leselast19Einstellung"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Text1_Change(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If Text1(1).Text <> "" And Label2(0).Caption <> "unbekannt" Then
        Command2(15).Caption = "Speichern"
    Else
        Command2(15).Caption = "Leeren"
    End If
    
    If Index = 2 Then
        LiefKuerzelAufloesung Label1(10), Text1(2)
    End If
    
    If Index = 5 Then
        If Len(Text1(5)) > 3 Then
            Command6(7).Enabled = True
        Else
            Command6(7).Enabled = False
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
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
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
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
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%!?" & Chr$(22) & Chr$(3) & Chr$(24)
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
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            Command1_Click
        End If
        If Index = 1 Then
            Command2_Click 15
        End If
        
        If Index = 2 Then
            Command7_Click 3
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
            
            Case Is = 2    'Lieferant
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
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
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
    Fehler.gsFehlertext = "Im Programmteil Artikelretoure ist ein Fehler aufgetreten. "
    
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

