VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL14 
   Caption         =   "Arbeitszeitauswertung"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   8760
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   8040
      Top             =   120
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   6840
      TabIndex        =   272
      Top             =   8160
      Visible         =   0   'False
      Width           =   1575
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   280
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Zurück"
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
         Index           =   6
         Left            =   9480
         TabIndex        =   277
         Top             =   4440
         Width           =   2055
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   240
         TabIndex        =   276
         Top             =   2520
         Width           =   4455
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         TabIndex        =   275
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Einfügen"
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
         Left            =   9480
         TabIndex        =   274
         Top             =   3480
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Löschen"
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
         Index           =   4
         Left            =   9480
         TabIndex        =   273
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label19 
         Caption         =   "h Pause"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   283
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "h"
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
         Left            =   2040
         TabIndex        =   282
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Rechts
         Caption         =   "über"
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
         Left            =   240
         TabIndex        =   281
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Pausenzeiten bearbeiten"
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
         Left            =   240
         TabIndex        =   279
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label10 
         Caption         =   "HH:MM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   278
         Top             =   1920
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   10440
      TabIndex        =   228
      Top             =   7800
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Korrigieren"
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
         Index           =   11
         Left            =   240
         TabIndex        =   336
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5280
         TabIndex        =   303
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Einfügen"
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
         Index           =   10
         Left            =   5280
         TabIndex        =   302
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Löschen"
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
         Index           =   9
         Left            =   6960
         TabIndex        =   301
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         TabIndex        =   300
         Text            =   "Combo1"
         Top             =   960
         Width           =   5415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Löschen"
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
         Left            =   3600
         TabIndex        =   237
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Einfügen"
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
         Left            =   1920
         TabIndex        =   236
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   3000
         TabIndex        =   233
         Top             =   1440
         Width           =   1695
         Begin VB.OptionButton Option1 
            Caption         =   "geht"
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
            Left            =   120
            TabIndex        =   235
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "kommt"
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
            Left            =   120
            TabIndex        =   234
            Top             =   120
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         TabIndex        =   232
         Text            =   "Text1"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   240
         TabIndex        =   231
         Top             =   2520
         Width           =   4935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Zurück"
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
         Index           =   0
         Left            =   9480
         TabIndex        =   230
         Top             =   4440
         Width           =   2055
      End
      Begin sevCommand3.Command Command8 
         Height          =   225
         Left            =   1560
         TabIndex        =   338
         Top             =   1860
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   397
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
         Height          =   225
         Left            =   1560
         TabIndex        =   339
         Top             =   1560
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   397
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
      Begin VB.Label Label23 
         Caption         =   "HH:MM"
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
         Left            =   6600
         TabIndex        =   304
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "Bemerkung"
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
         Left            =   5280
         TabIndex        =   299
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label14 
         Caption         =   "HH:MM"
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
         Left            =   2040
         TabIndex        =   239
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Datum"
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
         Left            =   240
         TabIndex        =   238
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Arbeitszeit bearbeiten"
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
         Left            =   240
         TabIndex        =   229
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   11655
      Begin VB.CommandButton Command1 
         Caption         =   "Export"
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
         Index           =   12
         Left            =   5160
         TabIndex        =   337
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Drucken"
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
         Index           =   8
         Left            =   7320
         TabIndex        =   297
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Pausenzeiten"
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
         Index           =   7
         Left            =   9480
         TabIndex        =   284
         Top             =   4440
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   30
         Left            =   9960
         TabIndex        =   271
         Text            =   "Text2"
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   29
         Left            =   9960
         TabIndex        =   270
         Text            =   "Text2"
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   28
         Left            =   9960
         TabIndex        =   269
         Text            =   "Text2"
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   27
         Left            =   9960
         TabIndex        =   268
         Text            =   "Text2"
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   26
         Left            =   9960
         TabIndex        =   267
         Text            =   "Text2"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   25
         Left            =   9960
         TabIndex        =   266
         Text            =   "Text2"
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   24
         Left            =   9960
         TabIndex        =   265
         Text            =   "Text2"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   23
         Left            =   9960
         TabIndex        =   264
         Text            =   "Text2"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   22
         Left            =   9960
         TabIndex        =   263
         Text            =   "Text2"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   21
         Left            =   9960
         TabIndex        =   262
         Text            =   "Text2"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   20
         Left            =   9960
         TabIndex        =   261
         Text            =   "Text2"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   19
         Left            =   9960
         TabIndex        =   260
         Text            =   "Text2"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   18
         Left            =   9960
         TabIndex        =   259
         Text            =   "Text2"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   17
         Left            =   9960
         TabIndex        =   258
         Text            =   "Text2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   16
         Left            =   9960
         TabIndex        =   257
         Text            =   "Text2"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   15
         Left            =   3840
         TabIndex        =   256
         Text            =   "Text2"
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   14
         Left            =   3840
         TabIndex        =   255
         Text            =   "Text2"
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   13
         Left            =   3840
         TabIndex        =   254
         Text            =   "Text2"
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   12
         Left            =   3840
         TabIndex        =   253
         Text            =   "Text2"
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   11
         Left            =   3840
         TabIndex        =   252
         Text            =   "Text2"
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   10
         Left            =   3840
         TabIndex        =   251
         Text            =   "Text2"
         Top             =   3000
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   9
         Left            =   3840
         TabIndex        =   250
         Text            =   "Text2"
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   8
         Left            =   3840
         TabIndex        =   249
         Text            =   "Text2"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   3840
         TabIndex        =   248
         Text            =   "Text2"
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   3840
         TabIndex        =   247
         Text            =   "Text2"
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   3840
         TabIndex        =   246
         Text            =   "Text2"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   3840
         TabIndex        =   245
         Text            =   "Text2"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   3840
         TabIndex        =   244
         Text            =   "Text2"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   3840
         TabIndex        =   243
         Text            =   "Text2"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   3840
         TabIndex        =   242
         Text            =   "Text2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   3840
         TabIndex        =   241
         Text            =   "Text2"
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   41
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   40
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   2
         Left            =   3480
         TabIndex        =   39
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   38
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   4
         Left            =   3480
         TabIndex        =   37
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   5
         Left            =   3480
         TabIndex        =   36
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   6
         Left            =   3480
         TabIndex        =   35
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   7
         Left            =   3480
         TabIndex        =   34
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   8
         Left            =   3480
         TabIndex        =   33
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   9
         Left            =   3480
         TabIndex        =   32
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   10
         Left            =   3480
         TabIndex        =   31
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   11
         Left            =   3480
         TabIndex        =   30
         Top             =   3240
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   12
         Left            =   3480
         TabIndex        =   29
         Top             =   3480
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   13
         Left            =   3480
         TabIndex        =   28
         Top             =   3720
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   14
         Left            =   3480
         TabIndex        =   27
         Top             =   3960
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   15
         Left            =   3480
         TabIndex        =   26
         Top             =   4200
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   16
         Left            =   9600
         TabIndex        =   25
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   17
         Left            =   9600
         TabIndex        =   24
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   18
         Left            =   9600
         TabIndex        =   23
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   19
         Left            =   9600
         TabIndex        =   22
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   20
         Left            =   9600
         TabIndex        =   21
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   21
         Left            =   9600
         TabIndex        =   20
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   22
         Left            =   9600
         TabIndex        =   19
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   23
         Left            =   9600
         TabIndex        =   18
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   24
         Left            =   9600
         TabIndex        =   17
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   25
         Left            =   9600
         TabIndex        =   16
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   26
         Left            =   9600
         TabIndex        =   15
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   27
         Left            =   9600
         TabIndex        =   14
         Top             =   3240
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   28
         Left            =   9600
         TabIndex        =   13
         Top             =   3480
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   29
         Left            =   9600
         TabIndex        =   12
         Top             =   3720
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Height          =   195
         Index           =   30
         Left            =   9600
         TabIndex        =   11
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   30
         Left            =   9360
         TabIndex        =   335
         Top             =   3960
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   29
         Left            =   9360
         TabIndex        =   334
         Top             =   3720
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   28
         Left            =   9360
         TabIndex        =   333
         Top             =   3480
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   27
         Left            =   9360
         TabIndex        =   332
         Top             =   3240
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   26
         Left            =   9360
         TabIndex        =   331
         Top             =   3000
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   25
         Left            =   9360
         TabIndex        =   330
         Top             =   2760
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   24
         Left            =   9360
         TabIndex        =   329
         Top             =   2520
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   9360
         TabIndex        =   328
         Top             =   2280
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   22
         Left            =   9360
         TabIndex        =   327
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   9360
         TabIndex        =   326
         Top             =   1800
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   20
         Left            =   9360
         TabIndex        =   325
         Top             =   1560
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   9360
         TabIndex        =   324
         Top             =   1320
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   9360
         TabIndex        =   323
         Top             =   1080
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   9360
         TabIndex        =   322
         Top             =   840
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   9360
         TabIndex        =   321
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   3240
         TabIndex        =   320
         Top             =   4200
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   3240
         TabIndex        =   319
         Top             =   3960
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   3240
         TabIndex        =   318
         Top             =   3720
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   3240
         TabIndex        =   317
         Top             =   3480
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   3240
         TabIndex        =   316
         Top             =   3240
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   3240
         TabIndex        =   315
         Top             =   3000
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   3240
         TabIndex        =   314
         Top             =   2760
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   3240
         TabIndex        =   313
         Top             =   2520
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   3240
         TabIndex        =   312
         Top             =   2280
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   3240
         TabIndex        =   311
         Top             =   2040
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   3240
         TabIndex        =   310
         Top             =   1800
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   309
         Top             =   1560
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3240
         TabIndex        =   308
         Top             =   1320
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   307
         Top             =   1080
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   306
         Top             =   840
         Width           =   195
      End
      Begin VB.Label Label24 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   3240
         TabIndex        =   305
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label20 
         Caption         =   "Tag"
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
         Index           =   11
         Left            =   6000
         TabIndex        =   296
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "erzielte Arbeitszeit"
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
         Index           =   10
         Left            =   10560
         TabIndex        =   295
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Pause"
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
         Left            =   9960
         TabIndex        =   294
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Unter brechung"
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
         Index           =   8
         Left            =   8760
         TabIndex        =   293
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "bis"
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
         Left            =   8040
         TabIndex        =   292
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "von"
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
         Left            =   6960
         TabIndex        =   291
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Tag"
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
         Left            =   240
         TabIndex        =   290
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "erzielte Arbeitszeit"
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
         Index           =   5
         Left            =   4440
         TabIndex        =   289
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Pause"
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
         Left            =   3840
         TabIndex        =   288
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Unter brechung"
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
         Index           =   2
         Left            =   2640
         TabIndex        =   287
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "bis"
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
         Left            =   1920
         TabIndex        =   286
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "von"
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
         Left            =   1080
         TabIndex        =   285
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "1."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   227
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "2."
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   226
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "3."
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   225
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "4."
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   224
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "5."
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   223
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "6."
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   222
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "7."
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   221
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "8."
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   220
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   219
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   218
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   217
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   216
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   215
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   214
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   213
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   212
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   211
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "9."
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   210
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   209
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "10."
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   208
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   207
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "11."
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   206
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   205
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "12."
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   204
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   203
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "13."
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   202
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   201
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "14."
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   200
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   199
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "15."
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   198
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   600
         TabIndex        =   197
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "16."
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   196
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   6360
         TabIndex        =   195
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "17."
         Height          =   255
         Index           =   16
         Left            =   5880
         TabIndex        =   194
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   6360
         TabIndex        =   193
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "18."
         Height          =   255
         Index           =   17
         Left            =   5880
         TabIndex        =   192
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Left            =   6360
         TabIndex        =   191
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "19."
         Height          =   255
         Index           =   18
         Left            =   5880
         TabIndex        =   190
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Index           =   19
         Left            =   6360
         TabIndex        =   189
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "20."
         Height          =   255
         Index           =   19
         Left            =   5880
         TabIndex        =   188
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Index           =   20
         Left            =   6360
         TabIndex        =   187
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "21."
         Height          =   255
         Index           =   20
         Left            =   5880
         TabIndex        =   186
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Index           =   21
         Left            =   6360
         TabIndex        =   185
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "22."
         Height          =   255
         Index           =   21
         Left            =   5880
         TabIndex        =   184
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Index           =   22
         Left            =   6360
         TabIndex        =   183
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "23."
         Height          =   255
         Index           =   22
         Left            =   5880
         TabIndex        =   182
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Index           =   23
         Left            =   6360
         TabIndex        =   181
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "24."
         Height          =   255
         Index           =   23
         Left            =   5880
         TabIndex        =   180
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Index           =   24
         Left            =   6360
         TabIndex        =   179
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "25."
         Height          =   255
         Index           =   24
         Left            =   5880
         TabIndex        =   178
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Index           =   25
         Left            =   6360
         TabIndex        =   177
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "26."
         Height          =   255
         Index           =   25
         Left            =   5880
         TabIndex        =   176
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Index           =   26
         Left            =   6360
         TabIndex        =   175
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "27."
         Height          =   255
         Index           =   26
         Left            =   5880
         TabIndex        =   174
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Index           =   27
         Left            =   6360
         TabIndex        =   173
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "28."
         Height          =   255
         Index           =   27
         Left            =   5880
         TabIndex        =   172
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Index           =   28
         Left            =   6360
         TabIndex        =   171
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "29."
         Height          =   255
         Index           =   28
         Left            =   5880
         TabIndex        =   170
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Index           =   29
         Left            =   6360
         TabIndex        =   169
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "30."
         Height          =   255
         Index           =   29
         Left            =   5880
         TabIndex        =   168
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Day"
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
         Index           =   30
         Left            =   6360
         TabIndex        =   167
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         Caption         =   "31."
         Height          =   255
         Index           =   30
         Left            =   5880
         TabIndex        =   166
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   165
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   164
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   163
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   162
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   161
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   160
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   159
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   158
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   157
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   156
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   155
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   154
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   153
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   152
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   151
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   1080
         TabIndex        =   150
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   6960
         TabIndex        =   149
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   6960
         TabIndex        =   148
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   6960
         TabIndex        =   147
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Index           =   19
         Left            =   6960
         TabIndex        =   146
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Index           =   20
         Left            =   6960
         TabIndex        =   145
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Index           =   21
         Left            =   6960
         TabIndex        =   144
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Index           =   22
         Left            =   6960
         TabIndex        =   143
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Index           =   23
         Left            =   6960
         TabIndex        =   142
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Index           =   24
         Left            =   6960
         TabIndex        =   141
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Index           =   25
         Left            =   6960
         TabIndex        =   140
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Index           =   26
         Left            =   6960
         TabIndex        =   139
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Index           =   27
         Left            =   6960
         TabIndex        =   138
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Index           =   28
         Left            =   6960
         TabIndex        =   137
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Index           =   29
         Left            =   6960
         TabIndex        =   136
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Index           =   30
         Left            =   6960
         TabIndex        =   135
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   134
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   133
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   132
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   131
         Top             =   600
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         Index           =   1
         X1              =   5640
         X2              =   5640
         Y1              =   4440
         Y2              =   600
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   130
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   129
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   128
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   127
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   126
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   7
         Left            =   1920
         TabIndex        =   125
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   8
         Left            =   1920
         TabIndex        =   124
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   9
         Left            =   1920
         TabIndex        =   123
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   10
         Left            =   1920
         TabIndex        =   122
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   11
         Left            =   1920
         TabIndex        =   121
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   12
         Left            =   1920
         TabIndex        =   120
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   13
         Left            =   1920
         TabIndex        =   119
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   14
         Left            =   1920
         TabIndex        =   118
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   15
         Left            =   1920
         TabIndex        =   117
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   116
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   115
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   114
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   113
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   112
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   111
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   110
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   4
         Left            =   4560
         TabIndex        =   109
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   108
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   107
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   106
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   6
         Left            =   4560
         TabIndex        =   105
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   7
         Left            =   2640
         TabIndex        =   104
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   7
         Left            =   4560
         TabIndex        =   103
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   8
         Left            =   2640
         TabIndex        =   102
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   101
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   9
         Left            =   2640
         TabIndex        =   100
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   9
         Left            =   4560
         TabIndex        =   99
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   10
         Left            =   2640
         TabIndex        =   98
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   10
         Left            =   4560
         TabIndex        =   97
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   11
         Left            =   2640
         TabIndex        =   96
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   11
         Left            =   4560
         TabIndex        =   95
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   12
         Left            =   2640
         TabIndex        =   94
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   12
         Left            =   4560
         TabIndex        =   93
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   13
         Left            =   2640
         TabIndex        =   92
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   13
         Left            =   4560
         TabIndex        =   91
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   14
         Left            =   2640
         TabIndex        =   90
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   14
         Left            =   4560
         TabIndex        =   89
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   15
         Left            =   2640
         TabIndex        =   88
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   15
         Left            =   4560
         TabIndex        =   87
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   16
         Left            =   8040
         TabIndex        =   86
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   16
         Left            =   8760
         TabIndex        =   85
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   16
         Left            =   10680
         TabIndex        =   84
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   17
         Left            =   8040
         TabIndex        =   83
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   17
         Left            =   8760
         TabIndex        =   82
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   17
         Left            =   10680
         TabIndex        =   81
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   18
         Left            =   8040
         TabIndex        =   80
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   18
         Left            =   8760
         TabIndex        =   79
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   18
         Left            =   10680
         TabIndex        =   78
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   19
         Left            =   8040
         TabIndex        =   77
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   19
         Left            =   8760
         TabIndex        =   76
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   19
         Left            =   10680
         TabIndex        =   75
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   20
         Left            =   8040
         TabIndex        =   74
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   20
         Left            =   8760
         TabIndex        =   73
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   20
         Left            =   10680
         TabIndex        =   72
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   21
         Left            =   8040
         TabIndex        =   71
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   21
         Left            =   8760
         TabIndex        =   70
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   21
         Left            =   10680
         TabIndex        =   69
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   22
         Left            =   8040
         TabIndex        =   68
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   22
         Left            =   8760
         TabIndex        =   67
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   22
         Left            =   10680
         TabIndex        =   66
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   23
         Left            =   8040
         TabIndex        =   65
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   23
         Left            =   8760
         TabIndex        =   64
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   23
         Left            =   10680
         TabIndex        =   63
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   24
         Left            =   8040
         TabIndex        =   62
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   24
         Left            =   8760
         TabIndex        =   61
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   24
         Left            =   10680
         TabIndex        =   60
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   25
         Left            =   8040
         TabIndex        =   59
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   25
         Left            =   8760
         TabIndex        =   58
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   25
         Left            =   10680
         TabIndex        =   57
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   26
         Left            =   8040
         TabIndex        =   56
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   26
         Left            =   8760
         TabIndex        =   55
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   26
         Left            =   10680
         TabIndex        =   54
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   27
         Left            =   8040
         TabIndex        =   53
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   27
         Left            =   8760
         TabIndex        =   52
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   27
         Left            =   10680
         TabIndex        =   51
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   28
         Left            =   8040
         TabIndex        =   50
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   28
         Left            =   8760
         TabIndex        =   49
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   28
         Left            =   10680
         TabIndex        =   48
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   29
         Left            =   8040
         TabIndex        =   47
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   29
         Left            =   8760
         TabIndex        =   46
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   29
         Left            =   10680
         TabIndex        =   45
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Index           =   30
         Left            =   8040
         TabIndex        =   44
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Index           =   30
         Left            =   8760
         TabIndex        =   43
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   30
         Left            =   10680
         TabIndex        =   42
         Top             =   3960
         Width           =   735
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
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
      Left            =   6240
      TabIndex        =   8
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
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
      Left            =   480
      TabIndex        =   7
      Top             =   1560
      Width           =   615
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   7200
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7200
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Schließen"
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
      Index           =   2
      Left            =   9600
      TabIndex        =   0
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Zentriert
      Caption         =   "Monat"
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
      Left            =   2880
      TabIndex        =   298
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label15 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   240
      Top             =   8040
      Width           =   9375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      Caption         =   "Jahr"
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
      Left            =   4680
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      Caption         =   "Monat"
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
      Left            =   1560
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Bediener"
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
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Bednr"
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
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   975
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
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Arbeitszeitauswertung"
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
      TabIndex        =   1
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmWKL14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim byteMonat       As Byte
Dim lJahr           As Long
Dim Ermpausen()      As PauseA
Dim iZaehler        As Integer
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "stemptemp", gdBase
    loeschNEW "DRU_ZEI", gdBase
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
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim dDat As Date
Dim czeit As String
Dim cKrit As String
Dim dKrit As Double
Dim cART As String
Dim llf As Long
Dim cLBSatz As String
    
    Screen.MousePointer = 11
    Select Case Index
        
        Case 0
            Frame2.Visible = False
            Me.Refresh
            ZeigeZeiten Label1.Caption, byteMonat, lJahr
        Case 1
            'einfügen
            
            If Option1(0).Value = True Then
                cART = "kommt"
            ElseIf Option1(1).Value = True Then
                cART = "geht"
            End If
            
            If fnPruefeUhrzeit(Text1.Text) = 0 Then
                czeit = Text1.Text
            Else
                Text1.SetFocus
                Screen.MousePointer = 0
                anzeige "rot", "richtiges Zeitformat wird erwartet", Label15
                Exit Sub
            End If
            
            dDat = DateValue(Label13.Caption)
            Speicherstempel Label1.Caption, Label2.Caption, dDat, cART, czeit
            LeseDatenBeaZeiten List3, Label1.Caption, byteMonat, lJahr, CInt(Left(Label13.Caption, 2))
        
        Case 2     'Beenden
            Unload frmWKL14
        Case 3 'Löschen
            If List3.ListIndex < 0 Then
                anzeige "rot", "Bitte einen Eintrag in der Liste markieren!", Label15
                List3.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            llf = CLng(Trim(Right(List3.list(List3.ListIndex), 5)))
            Delstempel llf
            
            LeseDatenBeaZeiten List3, Label1.Caption, byteMonat, lJahr, CInt(Left(Label13.Caption, 2))
        Case 4
            If List4.ListIndex < 0 Then
                anzeige "rot", "Bitte einen Eintrag in der Liste markieren!", Label15
                List4.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            llf = CLng(Trim(Right(List4.list(List4.ListIndex), 5)))
            DelPause llf
            
            ZeigePausenZ
        Case 5
            If fnPruefeUhrzeit(Text4.Text) = 0 Then
                czeit = Text4.Text
            Else
                Text4.SetFocus
                Screen.MousePointer = 0
                anzeige "rot", "richtiges Zeitformat wird erwartet", Label15
                Exit Sub
            End If
            
            If fnPruefeUhrzeit(Text3.Text) = 0 Then
            
            Else
                Text3.SetFocus
                Screen.MousePointer = 0
                anzeige "rot", "richtiges Zeitformat wird erwartet", Label15
                Exit Sub
            End If
                
            SpeicherPause czeit, Text3.Text
            ZeigePausenZ
        
        Case 6
            Frame4.Visible = False
            Me.Refresh
            ZeigeZeiten Label1.Caption, byteMonat, lJahr
        Case 7
            Frame4.Visible = True
            ZeigePausenZ
        Case 8
            DruckZeiten

            If NewTableSuchenDBKombi("DRU_ZEI", gdBase) = True Then
                reportbildschirm "", "zZEN52"
            End If
    
            
        Case 9
            'löschen
            dDat = DateValue(Label13.Caption)
            DELstempelUTEXT Label1.Caption, dDat
            
            fülleazu
            
            Text5.Text = ""
        Case 10
        
            If Text5.Text <> "" Then
                If fnPruefeUhrzeit(Text5.Text) = 0 Then
                    czeit = Text5.Text
                Else
                    Text5.SetFocus
                    Screen.MousePointer = 0
                    anzeige "rot", "richtiges Zeitformat wird erwartet", Label15
                    Exit Sub
                End If
            End If
            
            'einfügen
            dDat = DateValue(Label13.Caption)
            
            SpeicherstempelUTEXT Label1.Caption, dDat, Combo1.Text, czeit
            
            
            inAZEITU
            fülleazu
            
            Text5.Text = ""
            
        Case 11 'Korrigieren
        
            If List3.ListIndex < 0 Then
                anzeige "rot", "Bitte einen Eintrag in der Liste markieren!", Label15
                List3.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            llf = CLng(Trim(Right(List3.list(List3.ListIndex), 5)))
            Delstempel llf
        
            Command1_Click 1
            
        Case 12 'Export
            DruckZeiten
            
            
            If NewTableSuchenDBKombi("DRU_ZEI", gdBase) = True Then
                ExportCSV Label2.Caption, Label3.Caption, Label4.Caption
            End If
            
            
            
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub ExportCSV(sBedienername As String, sMonat As String, sJahr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cPfad           As String
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim iRet            As Integer
    Dim rsrs            As Recordset
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    Dim i               As Integer

   
    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label15
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    sSQL = " Select "
    sSQL = sSQL & " TAG "
    sSQL = sSQL & ", TAGNAME  "
    sSQL = sSQL & ", VON  "
    sSQL = sSQL & ", BIS  "
    sSQL = sSQL & ", UNTERB  "
    sSQL = sSQL & ", PAUSENA  "
    sSQL = sSQL & ", GESAMT  "
    sSQL = sSQL & ", UTEXT  "
    sSQL = sSQL & " from DRU_ZEI "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    

        sAusgabedatname = sBedienername & "_" & sMonat & "_" & sJahr & ".csv"

        cPfad1 = gcDBPfad
        If Right$(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If

        cdatei = cPfad1 & "BOX\" & sAusgabedatname
        cPfad = cPfad1 & "BOX"
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        cSatz = "Tag;TAGNAME;VON;BIS;UNTERB;PAUSENDAUER;ARBEITSZEIT;BEMERKUNG" & Chr$(13) & Chr$(10)

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            For i = 0 To 7
                If Not IsNull(rsrs.Fields(i)) Then

                    If i > 0 Then
                        cSatz = cSatz & ";" & rsrs.Fields(i)
                    Else
                        cSatz = rsrs.Fields(i)
                    End If
                Else
                    If i > 0 Then
                        cSatz = cSatz & ";"
                    Else
                        cSatz = ""
                    End If
                End If
            Next i
        
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Datendrin("DRU_ZEI", gdBase) Then
        iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbYes Then
            gcBestellEmail.Attachment1 = cdatei
            Screen.MousePointer = 0
            frmWKL129.Show 1
        Else
            MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
        End If
        anzeige "normal", "", Label15
    Else
        anzeige "rot", "Keine Daten zum Export vorhanden.", Label15
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ExportCSV"
        Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub

Private Sub Command2_Click()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    If byteMonat = 1 Then
        byteMonat = 12
        lJahr = lJahr - 1
    Else
        byteMonat = byteMonat - 1
    End If
    
    Label3.Caption = MonthName(byteMonat)
    Label4.Caption = lJahr
    
    SETZEWEEKDAYS byteMonat, lJahr
    ZeigeZeiten Label1.Caption, byteMonat, lJahr
    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command3_Click()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    If byteMonat = 12 Then
        byteMonat = 1
        lJahr = lJahr + 1
    Else
        byteMonat = byteMonat + 1
    End If
    
    Label3.Caption = MonthName(byteMonat)
    Label4.Caption = lJahr
    
    SETZEWEEKDAYS byteMonat, lJahr
    ZeigeZeiten Label1.Caption, byteMonat, lJahr
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Frame2.Visible = True
    
    LeseDatenBeaZeiten List3, Label1.Caption, byteMonat, lJahr, Index + 1
    ZeigeUtextinCombo Index + 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command7_Click()
On Error GoTo LOKAL_ERROR
    
    Dim ByteH       As Byte
    Dim ByteMin     As Byte
hier:
    If Text1.Text = "" And Option1(0).Value = True Then
        Text1.Text = "08:00"
    ElseIf Text1.Text = "" And Option1(1).Value = True Then
        Text1.Text = "18:00"
    Else
    
        If fnPruefeUhrzeit(Text1.Text) <> 0 Then
            Text1.Text = ""
            GoTo hier
        End If
        
        ByteH = Left(Text1.Text, 2)
        ByteMin = Right(Text1.Text, 2)
        
        If ByteMin >= 59 Then
            If ByteH >= 23 Then
                ByteH = 0
                ByteMin = 0
            Else
                ByteH = ByteH + 1
                ByteMin = 0
            End If
        Else
            ByteMin = ByteMin + 1
        End If
        
        
        
        Text1.Text = Format(CStr(ByteH), "00") & ":" & Format(CStr(ByteMin), "00")
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command8_Click()
On Error GoTo LOKAL_ERROR
    
    Dim ByteH       As Byte
    Dim ByteMin     As Byte

hier:
    If Text1.Text = "" And Option1(0).Value = True Then
        Text1.Text = "08:00"
    ElseIf Text1.Text = "" And Option1(1).Value = True Then
        Text1.Text = "18:00"
    Else
    
        If fnPruefeUhrzeit(Text1.Text) <> 0 Then
            Text1.Text = ""
            GoTo hier
        End If
        
        ByteH = Left(Text1.Text, 2)
        ByteMin = Right(Text1.Text, 2)
        
        If ByteMin = 0 Then
            If ByteH = 0 Then
                ByteH = 23
                ByteMin = 59
            Else
                ByteH = ByteH - 1
                ByteMin = 59
            End If
        Else
            ByteMin = ByteMin - 1
        End If
        
        
        
        Text1.Text = Format(CStr(ByteH), "00") & ":" & Format(CStr(ByteMin), "00")
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Timer1_Timer()
    On Error GoTo LOKAL_ERROR
    
    If iZaehler > 5 Then
        Timer1.Interval = 50
    ElseIf iZaehler > 20 Then
        Timer1.Interval = 5
    ElseIf iZaehler > 40 Then
        Timer1.Interval = 1
    Else
        Timer1.Interval = 100
    End If
    
    iZaehler = iZaehler + 1
    
    Command7_Click
    
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Timer2_Timer()
    On Error GoTo LOKAL_ERROR
    
    If iZaehler > 5 Then
        Timer2.Interval = 50
    ElseIf iZaehler > 20 Then
        Timer2.Interval = 5
    ElseIf iZaehler > 40 Then
        Timer2.Interval = 1
    Else
        Timer2.Interval = 100
    End If
    
    iZaehler = iZaehler + 1
    Command8_Click
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer2_Timer"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR
    Dim cLBSatz As String
    
    Screen.MousePointer = 11
    
    WKL14Positionieren
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    LeseDatenWKL14 List2, List1
    
    byteMonat = Month(Now) - 1
    If byteMonat = 0 Then
        byteMonat = 12
        lJahr = Year(Now) - 1
    Else
        lJahr = Year(Now)
    End If
    
    Label3.Caption = MonthName(byteMonat)
    Label4.Caption = lJahr
    
    SETZEWEEKDAYS byteMonat, lJahr
    
   
    iZaehler = 0
    Timer1.Enabled = False
    Timer2.Enabled = False
    
    fülleazu
    
    cLBSatz = List2.list(0)
    MoveDaten2DialogWKL14 Trim$(Mid(cLBSatz, 1, 3))
    ZeigeZeiten Label1.Caption, byteMonat, lJahr
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fülleazu()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cSatz As String

    Combo1.Clear
    Combo1.AddItem "bitte auswählen"
    Combo1.Text = "bitte auswählen"
    
    sSQL = "Select * from AZEITU"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!utext) Then
                cSatz = rsrs!utext
                Combo1.AddItem cSatz
            End If
            cSatz = ""
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close


    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fülleazu"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub inAZEITU()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    If Trim(Combo1.Text) <> "bitte auswählen" Then
        sSQL = "Delete from AZEITU where UTEXT = '" & Trim(Combo1.Text) & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Select * from AZEITU where UTEXT = '" & Trim(Combo1.Text) & "'"
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If rsrs.EOF Then
            rsrs.AddNew
            rsrs!utext = Trim(Combo1.Text)
            rsrs.Update
        End If
        rsrs.Close
    End If


    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "inAZEITU"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub LeseDatenWKL14(Listx As ListBox, listx1 As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cLBSatz As String
    Dim dWert As Double
    
    listx1.Clear
    listx1.AddItem "Nr. Bediener-Name"
    
    Listx.Clear
    cSQL = "Select * from BEDNAME where BEDNU <> 99 order by BEDNU"
'    cSQL = "Select * from BEDNAME where BEDNU <> 99 and ( SYNSTATUS is null or SYNSTATUS = 'E' or SYNSTATUS = 'A' ) order by BEDNU"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
                dWert = rsrs!BEDNU
            Else
                dWert = 0
            End If
            cFeld = Format$(dWert, "##0")
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLBSatz = cFeld & " "
            
            If Not IsNull(rsrs!bedname) Then
                cFeld = rsrs!bedname
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(35 - Len(cFeld)) & " "
            cLBSatz = cLBSatz & cFeld
    
            Listx.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "LeseDatenWKL14"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseDatenBeaZeiten(Listx As ListBox, sbed As String, byteMon As Byte, lja As Long, iDay As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cLBSatz As String
    Dim cDatum As String
    
    cDatum = CStr(iDay) & "." & CStr(byteMon) & "." & CStr(lja)
    
    Label13.Caption = ""
    Label13.Caption = Format(cDatum, "DD.MM.YY")

    Text1.Text = ""
    Listx.Clear
    
    sSQL = "select * from stempel where bednu = " & CLng(sbed)
    sSQL = sSQL & " and  month(datum) = " & byteMon
    sSQL = sSQL & " and  year(datum) = " & lja
    sSQL = sSQL & " and  day(datum) = " & iDay
    sSQL = sSQL & " order by zeit "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!zeit) Then
                cFeld = rsrs!zeit
            Else
                cFeld = 0
            End If
            cFeld = Format$(TimeValue(cFeld), "HH:MM")
            
            cLBSatz = cFeld & " "
            
            If Not IsNull(rsrs!art) Then
                cFeld = rsrs!art
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cLBSatz = cLBSatz & cFeld
            
            
            If Not IsNull(rsrs!lLFNR) Then
                cFeld = rsrs!lLFNR
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & Space(60) & cFeld
    
            Listx.AddItem cLBSatz
            
            Label13.Caption = Format(rsrs!Datum, "DD.MM.YY")
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    anzeige "normal", "", Label15
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "LeseDatenBeaZeiten"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Delstempel(llf)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
   
    sSQL = "delete from stempel where llfnr = " & llf
   
    gdBase.Execute sSQL, dbFailOnError
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "delstempel"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Speicherstempel(sbed As String, sbename As String, dateD As Date, cART As String, czeit As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
   
    
    sSQL = "Insert into stempel (bednu,bedname,datum,zeit,art) values  "
    sSQL = sSQL & " ( " & CLng(sbed)
    sSQL = sSQL & " , '" & sbename & "' "
    sSQL = sSQL & " , '" & dateD & "' "
    sSQL = sSQL & " , '" & czeit & "' "
    sSQL = sSQL & " , '" & cART & "' "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "Speicherstempel"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherstempelUTEXT(sbed As String, dateD As Date, sUtext As String, sGesamtZ As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Trim(Combo1.Text) <> "bitte auswählen" Then
    
        sSQL = "Delete from STEMPELAZU where bednu = "
        sSQL = sSQL & "  " & CLng(sbed)
        sSQL = sSQL & " and datum = " & CLng(dateD)
        gdBase.Execute sSQL, dbFailOnError
        
        If sGesamtZ = "" Then
            sSQL = "Insert into STEMPELAZU (bednu,datum,UTEXT) values  "
            sSQL = sSQL & " ( " & CLng(sbed)
            sSQL = sSQL & " , '" & dateD & "' "
            sSQL = sSQL & " , '" & sUtext & "' "
            sSQL = sSQL & " ) "
            gdBase.Execute sSQL, dbFailOnError
        Else
            sSQL = "Insert into STEMPELAZU (bednu,datum,UTEXT,GESAMT) values  "
            sSQL = sSQL & " ( " & CLng(sbed)
            sSQL = sSQL & " , '" & dateD & "' "
            sSQL = sSQL & " , '" & sUtext & "' "
            sSQL = sSQL & " , '" & sGesamtZ & "' "
            sSQL = sSQL & " ) "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "SpeicherstempelUTEXT"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DELstempelUTEXT(sbed As String, dateD As Date)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Delete from STEMPELAZU where bednu = "
    sSQL = sSQL & "  " & CLng(sbed)
    sSQL = sSQL & " and datum = " & CLng(dateD)
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "DELstempelUTEXT"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherPause(czeit As String, cKrit As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Insert into Pausenz (PauseNZ,Krit) values  "
    sSQL = sSQL & " ( '" & czeit & "' "
    sSQL = sSQL & " , '" & cKrit & "' ) "
    gdBase.Execute sSQL, dbFailOnError
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "SpeicherPause"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DelPause(llf As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
   
    sSQL = "delete from PAUSENZ where llfnr = " & llf
   
    gdBase.Execute sSQL, dbFailOnError
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "DelPause"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SETZEWEEKDAYS(byteMon As Byte, lja As Long)
    On Error GoTo LOKAL_ERROR
    Dim i As Integer
    Dim stringdate As String
    
    For i = 1 To 31
        
        stringdate = CStr(i) & "." & CStr(byteMon) & "." & CStr(lja)
        If IsDate(stringdate) Then
            Label5(i - 1).Visible = True
            Label6(i - 1).Visible = True
            Label6(i - 1).Caption = Left(WeekdayName(Weekday(DateValue(stringdate), vbMonday)), 2)
            Label7(i - 1).Visible = True
            Label8(i - 1).Visible = True
            Label9(i - 1).Visible = True
            Text2(i - 1).Visible = True
            Label11(i - 1).Visible = True
            Command4(i - 1).Visible = True
        Else
            Label5(i - 1).Visible = False
            Label6(i - 1).Visible = False
            Label7(i - 1).Visible = False
            Label8(i - 1).Visible = False
            Label9(i - 1).Visible = False
            Text2(i - 1).Visible = False
            Label11(i - 1).Visible = False
            Command4(i - 1).Visible = False
            
        End If
    Next i
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "SETZEWEEKDAYS"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WKL14Positionieren()
    On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 3000
    Frame1.Left = 120
    Frame1.Height = 4935
    Frame1.Width = 11655
    
    Frame2.Top = 3000
    Frame2.Left = 120
    Frame2.Height = 4935
    Frame2.Width = 11655
    
    Frame4.Top = 3000
    Frame4.Left = 120
    Frame4.Height = 4935
    Frame4.Width = 11655
    
    
   
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL14Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub List2_Click()
On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    
    cLBSatz = List2.list(List2.ListIndex)
    MoveDaten2DialogWKL14 Trim$(Mid(cLBSatz, 1, 3))
    ZeigeZeiten Label1.Caption, byteMonat, lJahr

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List3_Click()
On Error GoTo LOKAL_ERROR

    zeigsatz
    
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List3_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigsatz()
On Error GoTo LOKAL_ERROR

    Dim cLBSatz As String
    
    cLBSatz = List3.list(List3.ListIndex)
    Text1.Text = Left(cLBSatz, 5)
    
    If Mid(cLBSatz, 7, 1) = "k" Then
        Option1(0).Value = True
    Else
        Option1(1).Value = True
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigsatz"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List4_Click()
On Error GoTo LOKAL_ERROR

    zeigsatz4
    
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List4_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigsatz4()
On Error GoTo LOKAL_ERROR

    Dim cLBSatz As String
    
    cLBSatz = List4.list(List4.ListIndex)
    Text3.Text = Mid(cLBSatz, 7, 5)
    
    Text4.Text = Mid(cLBSatz, 20, 5)

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigsatz4"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MoveDaten2DialogWKL14(sbed As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL  As String
    Dim rsrs As Recordset
    
    sSQL = "select * from bedname where bednu = " & CLng(sbed)
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BEDNU) Then
            Label1.Caption = rsrs!BEDNU
        End If
        
        If Not IsNull(rsrs!bedname) Then
            Label2.Caption = rsrs!bedname
        Else
            Label2.Caption = ""
        End If
        
        
    End If
    rsrs.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MoveDaten2DialogWKL14"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeZeiten(sbed As String, bytemoni As Byte, lja As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim rsmax           As Recordset
    Dim rsu             As Recordset
    Dim cART            As String
    Dim UZeitgeht       As Date
    Dim UZeitkommt      As Date
    Dim UZeitDiff       As Date
    Dim lMax            As Long
    Dim uPauseGesamt    As Date
    
    
    loeschNEW "stemptemp", gdBase
    
    sSQL = "select * into stemptemp from stempel where bednu = " & CLng(sbed)
    sSQL = sSQL & " and  month(datum) = " & bytemoni
    sSQL = sSQL & " and  year(datum) = " & lja
    gdBase.Execute sSQL, dbFailOnError
    
    For i = 1 To 31
hier:
        Label9(i - 1).Caption = "-"
        Label8(i - 1).Caption = "-"
        Label7(i - 1).Caption = "-"
        Text2(i - 1).Text = "-"
        Label7(i - 1).ForeColor = glS1
        Label8(i - 1).ForeColor = glS1
        Label9(i - 1).ForeColor = glS1
        
        lMax = 0
        sSQL = "Select * from stemptemp where "
        sSQL = sSQL & " day(datum) = " & i
        sSQL = sSQL & " order by zeit "
        Set rsmax = gdBase.OpenRecordset(sSQL)
        If Not rsmax.EOF Then
            rsmax.MoveLast
            lMax = rsmax.RecordCount
        End If
        rsmax.Close
        
        
        If lMax > 0 Then ' ja min 1 tagesStempel
        
            sSQL = "Select min(zeit) as minzeit  from stemptemp where "
            sSQL = sSQL & "  art = 'kommt' "
            sSQL = sSQL & " and  day(datum) = " & i
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!minzeit) Then
                    Label7(i - 1).Caption = Format$(TimeValue(rsrs!minzeit), "HH:MM")
                Else
                    Label7(i - 1).ForeColor = vbRed
                    Label7(i - 1).Caption = "fehlt"
                End If
            End If
            rsrs.Close
            
            
            sSQL = "Select max(zeit) as maxzeit  from stemptemp where "
            sSQL = sSQL & "  art = 'geht' "
            sSQL = sSQL & " and  day(datum) = " & i
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!maxzeit) Then
                    Label8(i - 1).Caption = Format$(TimeValue(rsrs!maxzeit), "HH:MM")
                Else
                    Label8(i - 1).ForeColor = vbRed
                    Label8(i - 1).Caption = "fehlt"
                End If
            End If
            rsrs.Close
        End If
        
        'unterbrechung ja /nein
        
        If lMax > 2 Then ' ja Unterbrechung
            j = lMax Mod 2
            If j = 0 Then
            
                sSQL = "Select * from stemptemp where "
                sSQL = sSQL & "  day(datum) = " & i
                sSQL = sSQL & " order by zeit "
                Set rsu = gdBase.OpenRecordset(sSQL)
                If Not rsu.EOF Then
                    rsu.MoveFirst
                    

                    k = 2
                    Do While Not k = lMax
                    
                    rsu.MoveNext
                    
                    If Not IsNull(rsu!art) Then
                        cART = rsu!art
                    End If
                    
                    If cART = "geht" Then
                        If Not IsNull(rsu!zeit) Then
                            UZeitgeht = Format$(TimeValue(rsu!zeit), "HH:MM")
                        Else
                            UZeitgeht = 0
                        End If
                    Else
                        Label9(i - 1).ForeColor = vbRed
                        Label9(i - 1).Caption = "Fehler"
                        If i = 31 Then
                            Exit For
                        Else
                            i = i + 1
                            GoTo hier:
                        End If
                    End If
                    
                    
                    
                    rsu.MoveNext
                    
                    If Not IsNull(rsu!art) Then
                        cART = rsu!art
                    End If
                    
                    If cART = "kommt" Then
                        If Not IsNull(rsu!zeit) Then
                            UZeitkommt = Format$(TimeValue(rsu!zeit), "HH:MM")
                        Else
                            UZeitkommt = 0
                        End If
                    Else
                        Label9(i - 1).ForeColor = vbRed
                        Label9(i - 1).Caption = "Fehler"
                        If i = 31 Then
                            Exit For
                        Else
                            i = i + 1
                            GoTo hier:
                        End If
                    
                    End If
                    
                    UZeitDiff = UZeitkommt - UZeitgeht
                    uPauseGesamt = uPauseGesamt + UZeitDiff
                    k = k + 2
                    Loop
                    
                    Label9(i - 1).Caption = Format$(uPauseGesamt, "HH:MM")
                    uPauseGesamt = 0
                        
                End If
                rsu.Close
            
            Else
                Label9(i - 1).ForeColor = vbRed
                Label9(i - 1).Caption = "Fehler"
            End If
        End If
       
    Next i
    
    anzeige "normal", "", Label15
    
    ZeigePausenundEnd
    ZeigeUtext
    Label21.Caption = ermGesamtIstZeit
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeZeiten"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigePausenZ()
    On Error GoTo LOKAL_ERROR
    
    
    Dim sSQL            As String
    Dim cSatz           As String
    Dim rsrs            As Recordset
    
    Text3.Text = ""
    Text4.Text = ""
    List4.Clear
    
    
    sSQL = "select *  from Pausenz "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
   
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        If Not IsNull(rsrs!krit) Then
            cSatz = " über " & Format$(TimeValue(rsrs!krit), "HH:MM") & " h "
    
    
            If Not IsNull(rsrs!pausenz) Then
                cSatz = cSatz & Space(5) & Format$(TimeValue(rsrs!pausenz), "HH:MM") & " min"
                
                If Not IsNull(rsrs!lLFNR) Then
                    cSatz = cSatz & Space(50) & rsrs!lLFNR
                End If
                List4.AddItem cSatz
            End If
        End If
        
        rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close
    
    anzeige "normal", "", Label15
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigePausenZ"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DruckZeiten()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim i               As Integer
    Dim j               As Integer
    
    Dim dateVon         As Date
    Dim dateBis         As Date
    Dim dateUnt         As Date
    Dim datePau         As Date
    Dim dateGes         As Date
    Dim cDatum          As String
    
    loeschNEW "DRU_ZEI", gdBase
    
    For i = 1 To 31
        If Label7(i - 1).Caption = "fehlt" And Label7(i - 1).Visible = True Then
            anzeige "rot", "Bitte die Fehler erst korrigieren...", Label15
            Exit Sub
        End If
        If Label8(i - 1).Caption = "fehlt" And Label8(i - 1).Visible = True Then
            anzeige "rot", "Bitte die Fehler erst korrigieren...", Label15
            Exit Sub
        End If
        If Label9(i - 1).Caption = "Fehler" And Label9(i - 1).Visible = True Then
            anzeige "rot", "Bitte die Fehler erst korrigieren...", Label15
            Exit Sub
        End If
    Next i
    
    anzeige "normal", "Druckvorschau wird vorbereitet...", Label15
    
    
    CreateTable "DRU_ZEI", gdBase
    
    For i = 1 To 31
        If Label7(i - 1).Visible = False Then
            Exit For
        End If
        
        If Label7(i - 1).Caption <> "-" And Label7(i - 1).Visible = True Then
            dateVon = Label7(i - 1).Caption
        Else
            dateVon = 0
        End If
        If Label8(i - 1).Caption <> "-" And Label8(i - 1).Visible = True Then
            dateBis = Label8(i - 1).Caption
        Else
            dateBis = 0
        End If
    
        If Label9(i - 1).Caption <> "-" And Label9(i - 1).Visible = True Then
            dateUnt = Label9(i - 1).Caption
        Else
            dateUnt = 0
        End If
        
        If Text2(i - 1).Text <> "-" And Text2(i - 1).Visible = True Then
            datePau = Text2(i - 1).Text
        Else
            datePau = 0
        End If
        
        If Label11(i - 1).Caption <> "-" And Label11(i - 1).Visible = True Then
            dateGes = Label11(i - 1).Caption
        Else
            dateGes = 0
        End If
        
        For j = 1 To 12
            If gcMonat(j) = Label3.Caption Then
                cDatum = i & "." & j & "." & Label4.Caption
                Exit For
            End If
        Next j
        
        If dateGes = "00:00" Then
        
            sSQL = "Insert into DRU_ZEI (TAG,TAGNAME,VON,BIS,UNTERB,PAUSENA,Datum) values  "
            sSQL = sSQL & " ( " & Label5(i - 1).Caption
            sSQL = sSQL & " , '" & Label6(i - 1).Caption & "' "
            sSQL = sSQL & " , '" & dateVon & "' "
            sSQL = sSQL & " , '" & dateBis & "' "
            sSQL = sSQL & " , '" & dateUnt & "' "
            sSQL = sSQL & " , '" & datePau & "' "
            sSQL = sSQL & " , '" & cDatum & "' "
            sSQL = sSQL & " ) "
            gdBase.Execute sSQL, dbFailOnError
            
        Else
        
            sSQL = "Insert into DRU_ZEI (TAG,TAGNAME,VON,BIS,UNTERB,PAUSENA,GESAMT,Datum) values  "
            sSQL = sSQL & " ( " & Label5(i - 1).Caption
            sSQL = sSQL & " , '" & Label6(i - 1).Caption & "' "
            sSQL = sSQL & " , '" & dateVon & "' "
            sSQL = sSQL & " , '" & dateBis & "' "
            sSQL = sSQL & " , '" & dateUnt & "' "
            sSQL = sSQL & " , '" & datePau & "' "
            sSQL = sSQL & " , '" & dateGes & "' "
            sSQL = sSQL & " , '" & cDatum & "' "
            sSQL = sSQL & " ) "
            gdBase.Execute sSQL, dbFailOnError
            
        End If
    Next i
    
    sSQL = "update DRU_ZEI "
    sSQL = sSQL & " set Bednu = " & Label1.Caption
    sSQL = sSQL & " , Nachname = '" & Label2.Caption & "' "
    sSQL = sSQL & " , GGESAMT =  '" & Label21.Caption & "' "
    sSQL = sSQL & " , MONAT = '" & Label3.Caption & "' "
    sSQL = sSQL & " , JAHR =  " & Label4.Caption
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "update DRU_ZEI inner join STEMPELAZU on DRU_ZEI.DATUM = STEMPELAZU.DATUM "
    sSQL = sSQL & " and DRU_ZEI.Bednu = STEMPELAZU.BEDNU "
    sSQL = sSQL & " set DRU_ZEI.UTEXT = STEMPELAZU.UTEXT "
    sSQL = sSQL & " , DRU_ZEI.GESAMT = STEMPELAZU.GESAMT "
    gdBase.Execute sSQL, dbFailOnError
    
    

    anzeige "normal", "", Label15
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckZeiten"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermGesamtIstZeit() As String
    On Error GoTo LOKAL_ERROR
    
    ermGesamtIstZeit = 0
    Dim i               As Integer
    Dim lMinuten        As Long
    Dim lStunden        As Long
    
    For i = 1 To 31
        If Label11(i - 1).Caption <> "-" And Label11(i - 1).Visible = True Then
            lMinuten = lMinuten + CLng(Right(Label11(i - 1).Caption, 2))
            lStunden = lStunden + CLng(Left(Label11(i - 1).Caption, 2))
        End If
    Next i
    
    lStunden = lStunden + Fix((lMinuten / 60))
    lMinuten = lMinuten - (Fix((lMinuten / 60)) * 60)
    
    ermGesamtIstZeit = CStr(lStunden) & "," & Format(CStr(lMinuten), "00")
    

    anzeige "normal", "", Label15
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermGesamtIstZeit"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermanzpausen() As PauseA
On Error GoTo LOKAL_ERROR

Dim sSQL            As String
Dim rs              As Recordset
Dim lanzpausen      As Integer
Dim j               As Integer

sSQL = "Select * from Pausenz"
Set rs = gdBase.OpenRecordset(sSQL)
If Not rs.EOF Then
    rs.MoveLast
    lanzpausen = rs.RecordCount
End If
rs.Close

If lanzpausen = 0 Then

Else

    ReDim Ermpausen(0 To lanzpausen - 1)
        
    For j = 0 To lanzpausen - 1
        Ermpausen(j).Pausenkrit = "00:00"
        Ermpausen(j).Pausenlaenge = "00:00"
    Next j
    
    j = 0
    sSQL = "Select * from Pausenz order by krit desc "
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            If Not IsNull(rs!krit) Then
                Ermpausen(j).Pausenkrit = Format$(TimeValue(rs!krit), "HH:MM")
                If Not IsNull(rs!pausenz) Then
                Ermpausen(j).Pausenlaenge = Format$(TimeValue(rs!pausenz), "HH:MM")
                j = j + 1
                End If
            
            End If
            
        
        rs.MoveNext
        Loop
    End If
    rs.Close
End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermanzpausen"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function ermAnzahlpausen() As Integer
On Error GoTo LOKAL_ERROR

Dim sSQL            As String
Dim rs              As Recordset

ermAnzahlpausen = 0

sSQL = "Select * from Pausenz"
Set rs = gdBase.OpenRecordset(sSQL)
If Not rs.EOF Then
    rs.MoveLast
    ermAnzahlpausen = rs.RecordCount
End If
rs.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermAnzahlpausen"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub ZeigePausenundEnd()
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim UZeitgeht       As Date
    Dim UZeitkommt      As Date
    Dim UZeitDiff       As Date
    Dim UZeitGesamt     As Date
    Dim UZeitZW         As Date

    ermanzpausen
    
    
    For i = 1 To 31
    
hier:

        Label11(i - 1).Caption = "-"
        Text2(i - 1).Text = "-"
        Label11(i - 1).ForeColor = glS1
        
        If Label9(i - 1).Caption = "Fehler" Then
            If i = 31 Then
                Exit For
            Else
            i = i + 1
                GoTo hier
            End If
        
        End If
        
        If fnPruefeUhrzeit(Label7(i - 1).Caption) = 0 Then
            UZeitkommt = Label7(i - 1).Caption
        Else
            UZeitkommt = "0"
        End If
        
        If fnPruefeUhrzeit(Label8(i - 1).Caption) = 0 Then
            UZeitgeht = Label8(i - 1).Caption
        Else
            UZeitgeht = "0"
        End If
        
        If fnPruefeUhrzeit(Label9(i - 1).Caption) = 0 Then
            UZeitDiff = Label9(i - 1).Caption
        Else
            UZeitDiff = "0"
        End If
        
        If UZeitkommt <> "0" And UZeitgeht <> "0" Then
        
            UZeitZW = UZeitgeht - UZeitkommt - UZeitDiff
            
            Dim Anzpause As Integer
            
            Anzpause = ermAnzahlpausen
            
            For j = 0 To Anzpause - 1
                If UZeitZW > Ermpausen(j).Pausenkrit Then
                
                    UZeitGesamt = UZeitZW - Ermpausen(j).Pausenlaenge
                    Text2(i - 1).Text = Format$(Ermpausen(j).Pausenlaenge, "HH:MM")
                    Exit For
                Else
                    UZeitGesamt = UZeitZW
                
                End If
            
            Next j
            
            Label11(i - 1).Caption = Format$(UZeitGesamt, "HH:MM")
            
        End If
        
    Next i
    
    anzeige "normal", "", Label15
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigePausenundEnd"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeUtext()
    On Error GoTo LOKAL_ERROR
    
    Dim i               As Integer
    Dim j               As Integer
    Dim cDatum          As String
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
   
    For i = 1 To 31
        For j = 1 To 12
            If gcMonat(j) = Label3.Caption Then
                cDatum = i & "." & j & "." & Label4.Caption
                Exit For
            End If
        Next j
        
        Label24(i - 1).Caption = ""
        Label24(i - 1).ToolTipText = ""
        Label24(i - 1).BackColor = glH1
'        Label11(i - 1).Caption = "-"
            
        If IsDate(cDatum) Then
            sSQL = "Select * from STEMPELAZU where datum = " & CLng(DateValue(cDatum))
            sSQL = sSQL & " and bednu =  " & CLng(Label1.Caption)
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!utext) Then
                    Label24(i - 1).Caption = Left(rsrs!utext, 1)
                    Label24(i - 1).ToolTipText = rsrs!utext
                    Label24(i - 1).BackColor = glWarn
                End If
                
                If Not IsNull(rsrs!gesamt) Then
                    If Label11(i - 1).Caption = "-" Then
                        Label11(i - 1).Caption = Format$(rsrs!gesamt, "HH:MM")
                    End If
                End If
                
            
            End If
            rsrs.Close
        End If
    Next i
    
    
    anzeige "normal", "", Label15
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeUtext"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeUtextinCombo(iTag As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim j               As Integer
    Dim cDatum          As String
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
   
    
    For j = 1 To 12
        If gcMonat(j) = Label3.Caption Then
            cDatum = iTag & "." & j & "." & Label4.Caption
            Exit For
        End If
    Next j
    If IsDate(cDatum) Then
        sSQL = "Select * from STEMPELAZU where datum = " & CLng(DateValue(cDatum))
        sSQL = sSQL & " and bednu =  " & CLng(Label1.Caption)
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!utext) Then
                Combo1.Text = rsrs!utext
            End If
            
            If Not IsNull(rsrs!gesamt) Then
                Text5.Text = Format$(rsrs!gesamt, "HH:MM")
            End If
            
        
        End If
        rsrs.Close
    End If
    
    anzeige "normal", "", Label15
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeUtextinCombo"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus()
On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_lostFocus"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Text1_GotFocus()
On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = glSelBack1
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
