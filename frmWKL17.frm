VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKL17 
   BackColor       =   &H00C0C000&
   Caption         =   "Lieferantendaten - Bearbeitung"
   ClientHeight    =   8595
   ClientLeft      =   1155
   ClientTop       =   1815
   ClientWidth     =   11880
   Icon            =   "frmWKL17.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox txtStatus 
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
      Left            =   11040
      TabIndex        =   329
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin sevCommand3.Command Command4 
      Height          =   345
      Index           =   11
      Left            =   11280
      TabIndex        =   311
      Top             =   240
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
   Begin sevCommand3.Command Command4 
      Height          =   345
      Index           =   6
      Left            =   10200
      TabIndex        =   310
      Top             =   240
      Width           =   945
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
      Caption         =   "Spalten?"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'Kein
      Height          =   615
      Left            =   -360
      TabIndex        =   263
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
      Begin sevCommand3.Command Command4 
         Height          =   345
         Index           =   5
         Left            =   9720
         TabIndex        =   264
         Top             =   6960
         Width           =   1545
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "A-Wert"
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
         Left            =   240
         TabIndex        =   309
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "zu erreichender Auftragswert beim Lieferanten"
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
         Left            =   3120
         TabIndex        =   308
         Top             =   720
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "LUG"
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
         Left            =   240
         TabIndex        =   307
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Lagerumschlagsgeschwindigkeit (Mittelwert aller Artikel)"
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
         Left            =   3120
         TabIndex        =   306
         Top             =   960
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "LAGER(SEK)"
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
         Left            =   240
         TabIndex        =   305
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "derzeitiger Lagerwert zum Schnitteinkaufspreis"
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
         Left            =   3120
         TabIndex        =   304
         Top             =   1200
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "EINKAUF akt Jahr"
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
         Index           =   14
         Left            =   240
         TabIndex        =   303
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Einkaufsumsatz im aktuellen Jahr beim Lieferanten zum Schnitteinkaufspreis"
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
         Left            =   3120
         TabIndex        =   302
         Top             =   1440
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "EINKAUF vor Jahr"
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
         Left            =   240
         TabIndex        =   301
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "UMSATZ Br akt Jahr"
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
         Left            =   240
         TabIndex        =   300
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bruttoumsatz im aktuellen Jahr"
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
         Index           =   19
         Left            =   3120
         TabIndex        =   299
         Top             =   1920
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "UMSATZ Br vor Jahr"
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
         Index           =   20
         Left            =   240
         TabIndex        =   298
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bruttoumsatz im Vorjahr"
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
         Index           =   21
         Left            =   3120
         TabIndex        =   297
         Top             =   2160
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "UMSATZ SEK akt Jahr"
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
         Index           =   22
         Left            =   240
         TabIndex        =   296
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Umsatz im aktuellen Jahr zum Schnitteinkaufspreis"
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
         Index           =   23
         Left            =   3120
         TabIndex        =   295
         Top             =   2400
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "UMSATZ SEK vor Jahr"
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
         Index           =   24
         Left            =   240
         TabIndex        =   294
         Top             =   2640
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Umsatz im Vorjahr zum Schnitteinkaufspreis"
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
         Index           =   25
         Left            =   3120
         TabIndex        =   293
         Top             =   2640
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "UMS Br l. 12M"
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
         Index           =   26
         Left            =   240
         TabIndex        =   292
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bruttoumsatz letzten 12 abgeschlossenen Monate"
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
         Index           =   27
         Left            =   3120
         TabIndex        =   291
         Top             =   2880
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "UMS Br l. 12M VJZR"
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
         Index           =   28
         Left            =   240
         TabIndex        =   290
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bruttoumsatz letzten 12 abgeschlossenen Monate des Vorjahres"
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
         Index           =   29
         Left            =   3120
         TabIndex        =   289
         Top             =   3120
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "UMS SEK l. 12M"
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
         Index           =   30
         Left            =   240
         TabIndex        =   288
         Top             =   3360
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Umsatz zum Schnitteinkaufspreis der letzten 12 abgeschlossenen Monate"
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
         Index           =   31
         Left            =   3120
         TabIndex        =   287
         Top             =   3360
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "UMS SEK l. 12M VJZR"
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
         Index           =   32
         Left            =   240
         TabIndex        =   286
         Top             =   3600
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Umsatz zum Schnitteinkaufspreis der letzten 12 abgeschlossenen Monate des Vorjahres"
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
         Index           =   33
         Left            =   3120
         TabIndex        =   285
         Top             =   3600
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "DIFF UMS BR 12M €"
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
         Index           =   34
         Left            =   240
         TabIndex        =   284
         Top             =   3840
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Differenz in € zwischen 'UMS Br l. 12M' und 'UMS Br l. 12M VJZR'"
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
         Index           =   35
         Left            =   3120
         TabIndex        =   283
         Top             =   3840
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "DIFF UMS BR 12M %"
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
         Index           =   36
         Left            =   240
         TabIndex        =   282
         Top             =   4080
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Differenz in % zwischen 'UMS Br l. 12M' und 'UMS Br l. 12M VJZR'"
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
         Index           =   37
         Left            =   3120
         TabIndex        =   281
         Top             =   4080
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "DIFF UMS SEK 12M €"
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
         Index           =   38
         Left            =   240
         TabIndex        =   280
         Top             =   4320
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Differenz in € zwischen 'UMS SEK l. 12M' und 'UMS SEK l. 12M VJZR'"
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
         Index           =   39
         Left            =   3120
         TabIndex        =   279
         Top             =   4320
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "DIFF UMS SEK 12M %"
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
         Index           =   40
         Left            =   240
         TabIndex        =   278
         Top             =   4560
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Differenz in % zwischen 'UMS SEK l. 12M' und 'UMS SEK l. 12M VJZR'"
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
         Index           =   41
         Left            =   3120
         TabIndex        =   277
         Top             =   4560
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Einkaufsumsatz im Vorjahr beim Lieferanten zum Schnitteinkaufspreis"
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
         Left            =   3120
         TabIndex        =   276
         Top             =   1680
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Spaltenköpfe der Tabelle/Erläuterung"
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
         Left            =   240
         TabIndex        =   275
         Top             =   240
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Penneranteil in Prozent gemessen am Schnitteinkaufswert"
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
         Index           =   43
         Left            =   3120
         TabIndex        =   274
         Top             =   5760
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Panteil SEK in %"
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
         Index           =   44
         Left            =   240
         TabIndex        =   273
         Top             =   5760
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Penneranteil in Prozent gemessen an der Stückzahl"
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
         Index           =   45
         Left            =   3120
         TabIndex        =   272
         Top             =   5520
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Panteil Stück in %"
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
         Index           =   46
         Left            =   240
         TabIndex        =   271
         Top             =   5520
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "zuletzt ermittelte Lagerbestandssumme an Pennern in Stück"
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
         Index           =   47
         Left            =   3120
         TabIndex        =   270
         Top             =   5280
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Penner(Stück)"
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
         Index           =   48
         Left            =   240
         TabIndex        =   269
         Top             =   5280
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "zuletzt ermittelter Lagerwert an Pennern zum Schnitteinkaufspreis"
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
         Index           =   49
         Left            =   3120
         TabIndex        =   268
         Top             =   5040
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Penner(SEK)"
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
         Index           =   50
         Left            =   240
         TabIndex        =   267
         Top             =   5040
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "derzeitige Lagerbestandssumme in Stück"
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
         Index           =   51
         Left            =   3120
         TabIndex        =   266
         Top             =   4800
         Width           =   8175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "LAGER(Stück)"
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
         Index           =   52
         Left            =   240
         TabIndex        =   265
         Top             =   4800
         Width           =   2775
      End
   End
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
      Height          =   3735
      Left            =   0
      TabIndex        =   179
      Top             =   5520
      Visible         =   0   'False
      Width           =   11655
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   36
         Left            =   1560
         TabIndex        =   243
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   0
         Left            =   120
         TabIndex        =   242
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
         Index           =   1
         Left            =   840
         TabIndex        =   241
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
         Index           =   2
         Left            =   1560
         TabIndex        =   240
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
         Index           =   3
         Left            =   2280
         TabIndex        =   239
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
         Index           =   4
         Left            =   3000
         TabIndex        =   238
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
         Index           =   5
         Left            =   3720
         TabIndex        =   237
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
         Index           =   6
         Left            =   4440
         TabIndex        =   236
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
         Index           =   7
         Left            =   5160
         TabIndex        =   235
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
         Index           =   8
         Left            =   5880
         TabIndex        =   234
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
         Index           =   9
         Left            =   6600
         TabIndex        =   233
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
         Index           =   10
         Left            =   7320
         TabIndex        =   232
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
      Begin sevCommand3.Command Command4 
         Height          =   600
         Index           =   2
         Left            =   8040
         TabIndex        =   231
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
         Index           =   11
         Left            =   120
         TabIndex        =   230
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   12
         Left            =   840
         TabIndex        =   229
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
         Index           =   13
         Left            =   1560
         TabIndex        =   228
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
         Index           =   14
         Left            =   2280
         TabIndex        =   227
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
         Index           =   15
         Left            =   3000
         TabIndex        =   226
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
         Index           =   16
         Left            =   3720
         TabIndex        =   225
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
         Index           =   17
         Left            =   4440
         TabIndex        =   224
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
         Index           =   18
         Left            =   5160
         TabIndex        =   223
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
         Index           =   19
         Left            =   5880
         TabIndex        =   222
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
         Index           =   20
         Left            =   6600
         TabIndex        =   221
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
         Index           =   21
         Left            =   7320
         TabIndex        =   220
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
      Begin sevCommand3.Command Command4 
         Height          =   600
         Index           =   3
         Left            =   8040
         TabIndex        =   219
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
         Index           =   22
         Left            =   360
         TabIndex        =   218
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   23
         Left            =   1080
         TabIndex        =   217
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
         Index           =   24
         Left            =   1800
         TabIndex        =   216
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
         Index           =   25
         Left            =   2520
         TabIndex        =   215
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
         Index           =   26
         Left            =   3240
         TabIndex        =   214
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
         Index           =   27
         Left            =   3960
         TabIndex        =   213
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
         Index           =   28
         Left            =   4680
         TabIndex        =   212
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
         Index           =   29
         Left            =   5400
         TabIndex        =   211
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
         Index           =   30
         Left            =   6120
         TabIndex        =   210
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
         Index           =   31
         Left            =   6840
         TabIndex        =   209
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
         Index           =   32
         Left            =   7560
         TabIndex        =   208
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
         Index           =   33
         Left            =   8280
         TabIndex        =   207
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
         Index           =   34
         Left            =   9000
         TabIndex        =   206
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
      Begin sevCommand3.Command Command4 
         Height          =   600
         Index           =   4
         Left            =   9720
         TabIndex        =   205
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
         Index           =   35
         Left            =   840
         TabIndex        =   204
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   37
         Left            =   2280
         TabIndex        =   203
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
         Index           =   38
         Left            =   3000
         TabIndex        =   202
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
         Index           =   39
         Left            =   3720
         TabIndex        =   201
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
         Index           =   40
         Left            =   4440
         TabIndex        =   200
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
         Index           =   41
         Left            =   5160
         TabIndex        =   199
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
         Index           =   42
         Left            =   5880
         TabIndex        =   198
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
         Index           =   43
         Left            =   6600
         TabIndex        =   197
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
         Index           =   44
         Left            =   7320
         TabIndex        =   196
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
         Index           =   45
         Left            =   8040
         TabIndex        =   195
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
         Index           =   46
         Left            =   8760
         TabIndex        =   194
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
      Begin sevCommand3.Command Command4 
         Height          =   600
         Index           =   0
         Left            =   9480
         TabIndex        =   193
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
         Index           =   47
         Left            =   1200
         TabIndex        =   192
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
      Begin sevCommand3.Command Command3 
         Height          =   600
         Index           =   48
         Left            =   1920
         TabIndex        =   191
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
         Index           =   49
         Left            =   2640
         TabIndex        =   190
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
         Index           =   50
         Left            =   3360
         TabIndex        =   189
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
         Index           =   51
         Left            =   4080
         TabIndex        =   188
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
         Index           =   52
         Left            =   4800
         TabIndex        =   187
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
         Index           =   53
         Left            =   5520
         TabIndex        =   186
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
         Index           =   54
         Left            =   6240
         TabIndex        =   185
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
         Index           =   55
         Left            =   6960
         TabIndex        =   184
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
         Index           =   56
         Left            =   7680
         TabIndex        =   183
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
         Index           =   57
         Left            =   8400
         TabIndex        =   182
         Top             =   2400
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1058
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
         Index           =   58
         Left            =   9720
         TabIndex        =   181
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
      Begin sevCommand3.Command Command4 
         Height          =   600
         Index           =   1
         Left            =   10440
         TabIndex        =   180
         Top             =   2400
         Width           =   1215
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
         TabIndex        =   246
         Top             =   120
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
         TabIndex        =   245
         Top             =   360
         Width           =   2175
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
         TabIndex        =   244
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   79
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   3
         Left            =   9960
         TabIndex        =   84
         Top             =   6360
         Width           =   1695
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
      Begin VB.TextBox Text4 
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
         Index           =   0
         Left            =   7920
         MaxLength       =   7
         TabIndex        =   81
         Top             =   7080
         Width           =   1455
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   5
         Left            =   9960
         TabIndex        =   80
         Top             =   6960
         Width           =   1695
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
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Vorjahr"
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
         Index           =   37
         Left            =   1440
         TabIndex        =   345
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   13
         Left            =   2040
         TabIndex        =   344
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   12
         Left            =   2040
         TabIndex        =   343
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   11
         Left            =   2040
         TabIndex        =   342
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   10
         Left            =   2040
         TabIndex        =   341
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   9
         Left            =   2040
         TabIndex        =   340
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   8
         Left            =   2040
         TabIndex        =   339
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   7
         Left            =   2040
         TabIndex        =   338
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   6
         Left            =   2040
         TabIndex        =   337
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   5
         Left            =   2040
         TabIndex        =   336
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   4
         Left            =   2040
         TabIndex        =   335
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   2040
         TabIndex        =   334
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   2040
         TabIndex        =   333
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   2040
         TabIndex        =   332
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Vorjahr"
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
         Index           =   36
         Left            =   120
         TabIndex        =   331
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Zielsetzung"
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
         Index           =   32
         Left            =   9240
         TabIndex        =   178
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   10560
         TabIndex        =   177
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   10560
         TabIndex        =   176
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   10560
         TabIndex        =   175
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   10560
         TabIndex        =   174
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   10560
         TabIndex        =   173
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   10560
         TabIndex        =   172
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   10560
         TabIndex        =   171
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   10560
         TabIndex        =   170
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   10560
         TabIndex        =   169
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   10560
         TabIndex        =   168
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   10560
         TabIndex        =   167
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   10560
         TabIndex        =   166
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Zielsetzung Verkaufsumsatz(EK) für das aktuelle Jahr:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   27
         Left            =   7680
         TabIndex        =   165
         Top             =   6000
         Width           =   2175
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   12
         Left            =   9360
         TabIndex        =   164
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   11
         Left            =   9360
         TabIndex        =   163
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   10
         Left            =   9360
         TabIndex        =   162
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   9
         Left            =   9360
         TabIndex        =   161
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   8
         Left            =   9360
         TabIndex        =   160
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   7
         Left            =   9360
         TabIndex        =   159
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   6
         Left            =   9360
         TabIndex        =   158
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   5
         Left            =   9360
         TabIndex        =   157
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   4
         Left            =   9360
         TabIndex        =   156
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   9360
         TabIndex        =   155
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   9360
         TabIndex        =   154
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   9360
         TabIndex        =   153
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   13
         Left            =   9360
         TabIndex        =   152
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   2
         X1              =   9240
         X2              =   9240
         Y1              =   960
         Y2              =   5640
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Summe"
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
         Left            =   6960
         TabIndex        =   151
         Top             =   5400
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   6840
         X2              =   6840
         Y1              =   960
         Y2              =   5640
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   13
         Left            =   6000
         TabIndex        =   150
         Top             =   5400
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   13
         Left            =   8040
         TabIndex        =   149
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "summe"
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
         Index           =   13
         Left            =   4800
         TabIndex        =   148
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Summe"
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
         TabIndex        =   147
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6000
         TabIndex        =   146
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6000
         TabIndex        =   145
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6000
         TabIndex        =   144
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   4
         Left            =   6000
         TabIndex        =   143
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   5
         Left            =   6000
         TabIndex        =   142
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   6
         Left            =   6000
         TabIndex        =   141
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   7
         Left            =   6000
         TabIndex        =   140
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   8
         Left            =   6000
         TabIndex        =   139
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   9
         Left            =   6000
         TabIndex        =   138
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   10
         Left            =   6000
         TabIndex        =   137
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   11
         Left            =   6000
         TabIndex        =   136
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   12
         Left            =   6000
         TabIndex        =   135
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   8040
         TabIndex        =   134
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   8040
         TabIndex        =   133
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   8040
         TabIndex        =   132
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   4
         Left            =   8040
         TabIndex        =   131
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   5
         Left            =   8040
         TabIndex        =   130
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   6
         Left            =   8040
         TabIndex        =   129
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   7
         Left            =   8040
         TabIndex        =   128
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   8
         Left            =   8040
         TabIndex        =   127
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   9
         Left            =   8040
         TabIndex        =   126
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   10
         Left            =   8040
         TabIndex        =   125
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   11
         Left            =   8040
         TabIndex        =   124
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   12
         Left            =   8040
         TabIndex        =   123
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   4800
         TabIndex        =   122
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   4800
         TabIndex        =   121
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   4800
         TabIndex        =   120
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   4
         Left            =   4800
         TabIndex        =   119
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   5
         Left            =   4800
         TabIndex        =   118
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   6
         Left            =   4800
         TabIndex        =   117
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   7
         Left            =   4800
         TabIndex        =   116
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   8
         Left            =   4800
         TabIndex        =   115
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   9
         Left            =   4800
         TabIndex        =   114
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   10
         Left            =   4800
         TabIndex        =   113
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   11
         Left            =   4800
         TabIndex        =   112
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Index           =   12
         Left            =   4800
         TabIndex        =   111
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6960
         TabIndex        =   110
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6960
         TabIndex        =   109
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6960
         TabIndex        =   108
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6960
         TabIndex        =   107
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6960
         TabIndex        =   106
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6960
         TabIndex        =   105
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6960
         TabIndex        =   104
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6960
         TabIndex        =   103
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6960
         TabIndex        =   102
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6960
         TabIndex        =   101
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6960
         TabIndex        =   100
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   6960
         TabIndex        =   99
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   120
         TabIndex        =   98
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   120
         TabIndex        =   97
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         TabIndex        =   96
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         Left            =   120
         TabIndex        =   95
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         TabIndex        =   94
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         TabIndex        =   93
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         TabIndex        =   92
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         TabIndex        =   91
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         TabIndex        =   90
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         TabIndex        =   89
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         TabIndex        =   88
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "akt. Jahr"
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
         Index           =   31
         Left            =   7680
         TabIndex        =   87
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Vorjahr"
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
         Left            =   5040
         TabIndex        =   86
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Januar"
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
         TabIndex        =   85
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "€"
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
         Index           =   29
         Left            =   9480
         TabIndex        =   83
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Zielsetzung Verkaufsumsatz(EK) für das aktuelle Jahr:"
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
         Index           =   30
         Left            =   1920
         TabIndex        =   82
         Top             =   120
         Width           =   5535
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
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
      ForeColor       =   &H00000000&
      Height          =   5295
      Left            =   0
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   13695
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
         Height          =   360
         Index           =   27
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   46
         Top             =   1080
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
         Height          =   360
         Index           =   26
         Left            =   10680
         MaxLength       =   5
         TabIndex        =   44
         Top             =   720
         Width           =   855
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   2
         Left            =   9840
         TabIndex        =   351
         Top             =   3480
         Width           =   1695
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   25
         Left            =   6720
         MaxLength       =   30
         TabIndex        =   55
         Top             =   3000
         Width           =   1815
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   24
         Left            =   9720
         MaxLength       =   30
         TabIndex        =   56
         Top             =   3000
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8400
         Style           =   2  'Dropdown-Liste
         TabIndex        =   76
         Top             =   5040
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Hauptlieferant"
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
         Left            =   5640
         TabIndex        =   74
         Top             =   3480
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
         Height          =   360
         Index           =   23
         Left            =   6720
         MaxLength       =   100
         TabIndex        =   47
         Top             =   1080
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
         Height          =   360
         Index           =   22
         Left            =   6360
         MaxLength       =   7
         TabIndex        =   70
         Top             =   3960
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   21
         Left            =   9720
         MaxLength       =   30
         TabIndex        =   54
         Top             =   2640
         Width           =   1815
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   20
         Left            =   6720
         MaxLength       =   30
         TabIndex        =   53
         Top             =   2640
         Width           =   1815
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
         Height          =   360
         Index           =   17
         Left            =   6720
         MaxLength       =   100
         TabIndex        =   50
         Top             =   1920
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
         Height          =   360
         Index           =   5
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   35
         Top             =   0
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
         Height          =   360
         Index           =   6
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   36
         Top             =   0
         Width           =   1455
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
         Height          =   360
         Index           =   7
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   39
         Top             =   360
         Width           =   4335
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
         Height          =   360
         Index           =   8
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   41
         Top             =   720
         Width           =   4335
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
         Height          =   360
         Index           =   9
         Left            =   6720
         MaxLength       =   7
         TabIndex        =   37
         Top             =   0
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
         Height          =   360
         Index           =   10
         Left            =   8280
         MaxLength       =   30
         TabIndex        =   38
         Top             =   0
         Width           =   3255
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
         Height          =   360
         Index           =   11
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   45
         Top             =   1080
         Width           =   2295
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
         Height          =   360
         Index           =   12
         Left            =   6720
         MaxLength       =   20
         TabIndex        =   40
         Top             =   360
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
         Height          =   360
         Index           =   13
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   48
         Top             =   1440
         Width           =   4335
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
         Height          =   360
         Index           =   14
         Left            =   6720
         MaxLength       =   15
         TabIndex        =   42
         Top             =   720
         Width           =   1095
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   0
         Left            =   9840
         TabIndex        =   57
         Top             =   3960
         Width           =   1695
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   1
         Left            =   8040
         TabIndex        =   58
         Top             =   3960
         Width           =   1695
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
         Caption         =   "Schließen"
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
         Height          =   1695
         Index           =   15
         Left            =   1200
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   49
         Top             =   1800
         Width           =   4335
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
         Height          =   360
         Index           =   16
         Left            =   8760
         MaxLength       =   10
         TabIndex        =   43
         Top             =   720
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
         Height          =   360
         Index           =   18
         Left            =   6720
         MaxLength       =   30
         TabIndex        =   51
         Top             =   2280
         Width           =   1815
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
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   19
         Left            =   9720
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   52
         Top             =   2280
         Width           =   1815
      End
      Begin sevCommand3.Command Command5 
         Height          =   360
         Index           =   20
         Left            =   5760
         TabIndex        =   356
         ToolTipText     =   "Kalender"
         Top             =   3960
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
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "LEK-Ab:"
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
         Left            =   3600
         TabIndex        =   359
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Dep-Rab:"
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
         Left            =   9720
         TabIndex        =   358
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Einkaufsumsatz:"
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
         Index           =   42
         Left            =   0
         TabIndex        =   354
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Lagerwert "
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
         Index           =   41
         Left            =   2520
         TabIndex        =   353
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Datum"
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
         Index           =   40
         Left            =   4200
         TabIndex        =   352
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "GLN Nr:"
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
         Left            =   5520
         TabIndex        =   348
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "UST-ID:"
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
         Left            =   8520
         TabIndex        =   347
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Bestellrhythmus:"
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
         Left            =   6360
         TabIndex        =   75
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Email:"
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
         Left            =   5520
         TabIndex        =   73
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Korrigieren um:"
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
         Left            =   3960
         TabIndex        =   72
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Lagerwert "
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
         Index           =   222
         Left            =   2520
         TabIndex        =   71
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Einkaufsumsatz:"
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
         Index           =   190
         Left            =   0
         TabIndex        =   69
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "€"
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
         Index           =   212
         Left            =   7560
         TabIndex        =   68
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "Lagerwert "
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
         Index           =   200
         Left            =   2520
         TabIndex        =   67
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Lagereinkaufswert:"
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
         Index           =   180
         Left            =   240
         TabIndex        =   66
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Format:"
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
         Left            =   8520
         TabIndex        =   65
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "KennNr:"
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
         Left            =   5520
         TabIndex        =   64
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Password:"
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
         Left            =   8520
         TabIndex        =   63
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "User:"
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
         Left            =   5520
         TabIndex        =   62
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Adresse:"
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
         Left            =   5520
         TabIndex        =   61
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Einstellungen für die automatische Bestellverarbeitung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   120
         Left            =   5520
         TabIndex        =   60
         Top             =   1560
         Width           =   6015
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Lief.Nr.:"
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
         TabIndex        =   34
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Kürzel:"
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
         Left            =   3120
         TabIndex        =   33
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Lief.Bez.:"
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
         TabIndex        =   32
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Straße:"
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
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "PLZ:"
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
         Left            =   5520
         TabIndex        =   30
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Ort:"
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
         Left            =   7680
         TabIndex        =   29
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Telefon:"
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
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Fax:"
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
         Left            =   5520
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Kurztext:"
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
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Kund.-Nr:"
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
         Left            =   5640
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
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
         Index           =   10
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "A-wert:"
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
         Left            =   7800
         TabIndex        =   23
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame0 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Kein
      Height          =   1095
      Left            =   9960
      TabIndex        =   8
      Top             =   6960
      Width           =   3015
      Begin VB.Frame Frame6 
         Height          =   1335
         Left            =   6840
         TabIndex        =   322
         Top             =   2760
         Width           =   4695
         Begin sevCommand3.Command Command1 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   326
            Top             =   840
            Width           =   1695
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
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3000
            Style           =   2  'Dropdown-Liste
            TabIndex        =   325
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Auswertung"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   324
            Top             =   480
            Value           =   -1  'True
            Width           =   2775
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Übersicht"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   323
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0C000&
         Caption         =   "nur geführte Lieferanten anzeigen"
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
         Left            =   240
         TabIndex        =   313
         Top             =   2760
         Width           =   5415
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0C000&
         Caption         =   "mit Detailzahlen"
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
         Left            =   240
         TabIndex        =   312
         Top             =   3120
         Width           =   5415
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferanten ohne Bestand 'gelb' einfärben"
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
         Left            =   3600
         TabIndex        =   256
         Top             =   1080
         Width           =   4455
      End
      Begin sevCommand3.Command Command1 
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   255
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
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferantennummer"
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
         Index           =   0
         Left            =   2400
         TabIndex        =   253
         Top             =   2040
         Value           =   -1  'True
         Width           =   2415
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
         Height          =   300
         Index           =   1
         Left            =   2400
         TabIndex        =   252
         Top             =   2400
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Ort"
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
         Left            =   4920
         TabIndex        =   251
         Top             =   2040
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Plz"
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
         Left            =   4920
         TabIndex        =   250
         Top             =   2400
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Auftragswert"
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
         Left            =   7200
         TabIndex        =   249
         Top             =   2040
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Strasse"
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
         Left            =   7200
         TabIndex        =   248
         Top             =   2400
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kürzel"
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
         Index           =   6
         Left            =   240
         TabIndex        =   247
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown-Liste
         TabIndex        =   77
         Top             =   1440
         Width           =   3135
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
         Index           =   4
         Left            =   7440
         MaxLength       =   35
         TabIndex        =   4
         Text            =   "1234567890123"
         Top             =   600
         Width           =   4095
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
         Index           =   3
         Left            =   6480
         MaxLength       =   7
         TabIndex        =   3
         Text            =   "1234567890123"
         Top             =   600
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
         Index           =   2
         Left            =   3000
         MaxLength       =   35
         TabIndex        =   2
         Text            =   "1234567890123"
         Top             =   600
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
         Index           =   1
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "1234567890123"
         Top             =   600
         Width           =   1335
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
         Index           =   0
         Left            =   240
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "JOOP SCHLAGMICHTOT45"
         Top             =   600
         Width           =   1455
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   1
         Left            =   7800
         TabIndex        =   6
         Top             =   1440
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
         Caption         =   "&Neu"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   0
         Left            =   6000
         TabIndex        =   5
         Top             =   1440
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
         Caption         =   "S&uchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   2
         Left            =   9600
         TabIndex        =   7
         Top             =   1440
         Width           =   1935
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command98 
         Height          =   360
         Left            =   3600
         TabIndex        =   355
         Top             =   1440
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
         Picture         =   "frmWKL17.frx":0442
         PictureAlign    =   3
         Version3        =   -1  'True
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
         Height          =   615
         Index           =   6
         Left            =   240
         TabIndex        =   314
         Top             =   3480
         Width           =   6495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
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
         Left            =   240
         TabIndex        =   254
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Bestellrhythmus"
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
         Left            =   240
         TabIndex        =   78
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Ort"
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
         Left            =   7440
         TabIndex        =   16
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "PLZ"
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
         Left            =   6480
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Name"
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
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Kürzel"
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
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Nr"
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
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
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
         Index           =   2
         Left            =   10560
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "von"
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
         Left            =   9840
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
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
         Height          =   375
         Index           =   0
         Left            =   8520
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog cdlopen 
      Left            =   11160
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
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
      Height          =   3375
      Left            =   9840
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   346
         Top             =   7120
         Width           =   1415
         _ExtentX        =   2487
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Caption         =   "'gelbe' löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.PictureBox picprogress 
         Height          =   220
         Left            =   6960
         ScaleHeight     =   165
         ScaleWidth      =   4515
         TabIndex        =   330
         Top             =   6480
         Visible         =   0   'False
         Width           =   4575
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   10
         Left            =   6000
         TabIndex        =   328
         Top             =   6720
         Width           =   1295
         _ExtentX        =   2275
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Zielsetzung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   9
         Left            =   1560
         TabIndex        =   327
         Top             =   7120
         Width           =   1415
         _ExtentX        =   2487
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Caption         =   "'rote' löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   8
         Left            =   3000
         TabIndex        =   321
         Top             =   7120
         Width           =   1655
         _ExtentX        =   2910
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Lagerwerte"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   5
         Left            =   4680
         TabIndex        =   320
         Top             =   7120
         Width           =   1295
         _ExtentX        =   2275
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Marken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   6
         Left            =   6000
         TabIndex        =   319
         Top             =   7120
         Width           =   1295
         _ExtentX        =   2275
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Linien"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   7
         Left            =   4680
         TabIndex        =   318
         Top             =   6720
         Width           =   1295
         _ExtentX        =   2275
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Export"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command8 
         Height          =   220
         Left            =   6480
         TabIndex        =   315
         Top             =   6480
         Visible         =   0   'False
         Width           =   375
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
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   4
         Left            =   3000
         TabIndex        =   262
         Top             =   6720
         Width           =   1655
         _ExtentX        =   2910
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Daten holen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   3
         Left            =   7320
         TabIndex        =   258
         Top             =   7120
         Width           =   1295
         _ExtentX        =   2275
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Artikel"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6015
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   10610
         _Version        =   393216
         Cols            =   18
         FixedCols       =   2
         ForeColorSel    =   8454143
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
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   6720
         Width           =   1415
         _ExtentX        =   2487
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   1
         Left            =   9840
         TabIndex        =   20
         Top             =   7000
         Width           =   1695
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
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   19
         Top             =   6720
         Width           =   1415
         _ExtentX        =   2487
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   12
         Left            =   7320
         TabIndex        =   357
         Top             =   6720
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Vergleich"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label3 
         Appearance      =   0  '2D
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   316
         Tag             =   "Shape"
         Top             =   6480
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label5 
         Appearance      =   0  '2D
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   5520
         TabIndex        =   261
         Tag             =   "Shape"
         Top             =   6120
         Width           =   195
      End
      Begin VB.Label Label5 
         Appearance      =   0  '2D
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   260
         Tag             =   "Shape"
         Top             =   6120
         Width           =   195
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "gelbe Lieferanten = z.Z. kein Lagerbestand"
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
         Index           =   34
         Left            =   5760
         TabIndex        =   259
         Top             =   6120
         Width           =   4335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "rote Lieferanten = keine Artikel"
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
         Index           =   33
         Left            =   480
         TabIndex        =   257
         Top             =   6120
         Width           =   4935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         Caption         =   "grüne Lieferanten = offene Kundenbestellungen(noch nicht bestellt)"
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
         Index           =   35
         Left            =   480
         TabIndex        =   317
         Top             =   6480
         Visible         =   0   'False
         Width           =   5895
      End
   End
   Begin VB.Label Label1 
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
      Index           =   53
      Left            =   5400
      TabIndex        =   350
      ToolTipText     =   "derzeitiger Bestand und der Schnitteinkaufswert "
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label1 
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
      Index           =   5
      Left            =   5400
      TabIndex        =   349
      ToolTipText     =   "derzeitiger Bestand und der Schnitteinkaufswert "
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Lieferantendaten"
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
      Left            =   240
      TabIndex        =   59
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmWKL17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerLINR As Byte

Private Sub WKL17Positionieren()
    On Error GoTo LOKAL_ERROR
    
    Frame0.Top = 960
    Frame0.Left = 120
    Frame0.Height = 4215
    Frame0.Width = 11655
    
    Frame1.Top = 960
    Frame1.Left = 120
    Frame1.Height = 8200
    Frame1.Width = 11895
    
    Frame2.Top = 5400
    Frame2.Left = 120
    Frame2.Height = 3135
    Frame2.Width = 11655
    
    Frame3.Top = 960
    Frame3.Left = 120
    Frame3.Height = 7335
    Frame3.Width = 11655
    
    Frame4.Top = 960
    Frame4.Left = 0
    Frame4.Height = 7575
    Frame4.Width = 11895
    
    Frame5.Top = 960
    Frame5.Left = 120
    Frame5.Height = 8200
    Frame5.Width = 11655
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL17Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
    
End Sub
Private Function fnPruefeEingabeDialogWKL17() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim cValid As String
    Dim cZeichen As String
    Dim iCount As Integer
    
    fnPruefeEingabeDialogWKL17 = 0
    
    cValid = "1234567890"
    
    ctmp = Text1(5).Text
    ctmp = Trim$(ctmp)
    
    For iCount = 1 To Len(ctmp)
        cZeichen = Mid(ctmp, iCount, 1)
        If InStr(cValid, cZeichen) = 0 Then
            fnPruefeEingabeDialogWKL17 = 1
            Exit Function
        End If
    Next iCount
            
''    ctmp = Text1(13).Text
''    ctmp = Trim$(ctmp)
''
''    For iCount = 1 To Len(ctmp)
''        cZeichen = Mid(ctmp, iCount, 1)
''        If InStr(cValid, cZeichen) = 0 Then
''            fnPruefeEingabeDialogWKL17 = 2
''            Exit Function
''        End If
''    Next iCount

    ctmp = Text1(16).Text
    ctmp = Trim$(ctmp)
    
    For iCount = 1 To Len(ctmp)
        cZeichen = Mid(ctmp, iCount, 1)
        If InStr(cValid, cZeichen) = 0 Then
            fnPruefeEingabeDialogWKL17 = 3
            Exit Function
        End If
    Next iCount
    
   
            
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeDialogWKL17"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iMonat As Byte
    Dim iJahr As Integer
    Screen.MousePointer = 11
    Select Case Index
        Case Is = 0     'Suchen
            glSelect = 0
            Text1(1).Text = UCase$(Text1(1).Text)
            If SucheLieferantWKL17 Then
                zeige_Grid
                If MSFlexGrid1.Visible = True Then
                    MSFlexGrid1.Col = 1
                    MSFlexGrid1.Row = 2
                    MSFlexGrid1.SetFocus
                    
                End If
            End If
            Text1(5).Locked = True
            
        Case Is = 1     'Neu
            LeereLieferantMaskeWKL17
            HoleMaxLieferantNr
            Frame3.Visible = True
            Frame0.Visible = False
            Text1(5).SetFocus
            giDlgZustand = giNEU
            Text1(5).Locked = False
            
            If gbBILDTAST = False Then
                Frame2.Visible = False
            Else
                Frame2.Visible = True
            End If
            

        Case Is = 2     'Beenden
            Unload frmWKL17
        Case 3
            If Option2(0).Value = True Then
                Drucklief
            ElseIf Option2(1).Value = True Then
            
                iMonat = CByte(Mid$(Combo3.Text, 1, InStr(1, Combo3.Text, "/") - 1))
                iJahr = CInt(Right(Combo3.Text, 4))
                        
                VorMonatsAuswertung iMonat, iJahr
            End If
        Case 6
            Screen.MousePointer = 0
            Text1_KeyUp 0, vbKeyF2, 0
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub VorMonatsAuswertung(imon As Byte, iJahr As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lLinr As Long
    Dim lAnz As Long
    Dim j As Integer
    
    Dim lLagerST As Long

    Dim lAbsatzVM As Long
    Dim dUmsatzVM As Double
    Dim dUmsatzVVM As Double
    Dim lEinkaufVM As Long
    
    Dim lAbsatzaktJ As Long
    Dim dUmsatzaktJ As Double
    Dim dUmsatzVaktJ As Double
    Dim lEinkaufaktJ As Long
    
    Dim dUmsatzabs As Double
    Dim dUmsatzrela As Double
    
    Dim dSummeUmsatzVM As Double
    Dim dSummeUmsatzAKTJ As Double
    
    anzeige "normal", "Daten werden ermittelt...", Label1(6)
    
    loeschNEW "LIVM" & srechnertab, gdBase
    CreateTableT2 "LIVM" & srechnertab, gdBase
   
    cSQL = "Insert Into LIVM" & srechnertab & " Select   "
    cSQL = cSQL & " LINR  "
    cSQL = cSQL & ", LIEFBEZ "
    cSQL = cSQL & " from LISRT where "
    cSQL = cSQL & " ( SYNSTATUS is null or SYNSTATUS = 'E' or SYNSTATUS = 'A' )  "
    cSQL = cSQL & " and Linr < 500000 "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "mit Detailzahlen: es werden Hintergrunddaten zusammengefasst...", Label1(6)

    If UMS_LINRaktuell = False Then
        ErzeugeLinrUmsatz
    End If
    
    anzeige "normal", "Lagerwerte werden ermittelt...", Label1(6)
    
    LagerwerteschreibenLINRJetzt Label1(6)

    cSQL = "Select * from LIVM" & srechnertab
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!linr) Then
                lLinr = rsrs!linr
                
                lAnz = lAnz - 1
                anzeige "normal", "Lieferant: " & rsrs!LIEFBEZ & " noch " & CStr(lAnz) & " Lieferanten ...", Label1(6)

                lAbsatzVM = ermgesAbsatzLinr(imon, iJahr, lLinr)
                dUmsatzVM = ermgesUmsatzLinr(imon, iJahr, lLinr)
                dUmsatzVVM = ermgesUmsatzLinr(imon, iJahr - 1, lLinr)
                lEinkaufVM = EinkaufsStückermittlung(CStr(lLinr), gdBase, iJahr, imon)
                
                lAbsatzaktJ = 0
                dUmsatzaktJ = 0
                dUmsatzVaktJ = 0
                lEinkaufaktJ = 0
                
                If imon = 12 Then
                    lAbsatzaktJ = lAbsatzaktJ + ermgesAbsatzLinr(0, iJahr, lLinr)
                    dUmsatzaktJ = dUmsatzaktJ + ermgesUmsatzLinr(0, iJahr, lLinr)
                    dUmsatzVaktJ = dUmsatzVaktJ + ermgesUmsatzLinr(0, iJahr - 1, lLinr)
                    lEinkaufaktJ = lEinkaufaktJ + EinkaufsStückermittlung(CStr(lLinr), gdBase, iJahr, 0)
                Else
                    For j = 1 To imon
                        lAbsatzaktJ = lAbsatzaktJ + ermgesAbsatzLinr(CByte(j), iJahr, lLinr)
                        dUmsatzaktJ = dUmsatzaktJ + ermgesUmsatzLinr(CByte(j), iJahr, lLinr)
                        dUmsatzVaktJ = dUmsatzVaktJ + ermgesUmsatzLinr(CByte(j), iJahr - 1, lLinr)
                        lEinkaufaktJ = lEinkaufaktJ + EinkaufsStückermittlung(CStr(lLinr), gdBase, iJahr, CByte(j))
                    Next j
                End If
                
                lLagerST = LAGERStückErmittlungJetzt(lLinr)
            Else
                lLagerST = 0
                lAbsatzVM = 0
                dUmsatzVM = 0
                dUmsatzVVM = 0
                
                lEinkaufVM = 0
                
                lAbsatzaktJ = 0
                dUmsatzaktJ = 0
                dUmsatzVaktJ = 0
                lEinkaufaktJ = 0
                
            End If
            
            rsrs.Edit
            rsrs!LAGERST = lLagerST
            
            rsrs!ABSATZVM = lAbsatzVM
            rsrs!UmsatzVM = dUmsatzVM
            rsrs!UmsatzVVM = dUmsatzVVM
            rsrs!EINKAUFVM = lEinkaufVM
            
            rsrs!AbsatzaktJ = lAbsatzaktJ
            rsrs!EINKAUFaktJ = lEinkaufaktJ
            rsrs!Umsatzaktj = dUmsatzaktJ
            rsrs!UmsatzVAKTJ = dUmsatzVaktJ

            dUmsatzabs = 0
            dUmsatzabs = dUmsatzVM - dUmsatzVVM
            dUmsatzrela = 0
            If dUmsatzVM <> 0 Then
                dUmsatzrela = Round(100 * dUmsatzabs / dUmsatzVM, 0)
            End If

            rsrs!UMSATZMRELA = dUmsatzrela
            
            dUmsatzabs = 0
            dUmsatzabs = dUmsatzaktJ - dUmsatzVaktJ
            dUmsatzrela = 0
            If dUmsatzaktJ <> 0 Then
                dUmsatzrela = Round(100 * dUmsatzabs / dUmsatzaktJ, 0)
            End If

            rsrs!UMSATZJRELA = dUmsatzrela

            If lAbsatzVM <> 0 Then
                rsrs!VKPREISPROSTCK = dUmsatzVM / lAbsatzVM
            End If
            
            If lAbsatzVM <> 0 Then
                rsrs!LagerRWM = lLagerST / lAbsatzVM
            End If
            
            rsrs.Update
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    dSummeUmsatzVM = 0
    cSQL = "Select sum(UmsatzVM) as maxi from LIVM" & srechnertab
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dSummeUmsatzVM = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    dSummeUmsatzAKTJ = 0
    cSQL = "Select sum(Umsatzaktj) as maxi from LIVM" & srechnertab
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            dSummeUmsatzAKTJ = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    cSQL = "Select * from LIVM" & srechnertab
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast

        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!linr) Then
                lLinr = rsrs!linr
                
            End If
            
            rsrs.Edit
            
            If dSummeUmsatzVM <> 0 Then
                rsrs!MarktanteilM = 100 * rsrs!UmsatzVM / dSummeUmsatzVM
            End If
            
            If dSummeUmsatzAKTJ <> 0 Then
                rsrs!MarktanteilJ = 100 * rsrs!Umsatzaktj / dSummeUmsatzAKTJ
            End If
            
            rsrs.Update
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    loeschNEW "LITT", gdBase
    
    cSQL = "Select * into LITT from LIVM" & srechnertab
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "LIVM" & srechnertab, gdBase
    CreateTableT2 "LIVM" & srechnertab, gdBase
   
    cSQL = "Insert Into LIVM" & srechnertab & " Select Top 50 UmsatzVM,* "
    cSQL = cSQL & " from LITT where UmsatzVM > 0 order by UmsatzVM desc"
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "LITT", gdBase
    
    'Platzierungen
    
    lAnz = 1
    cSQL = "Select * from LIVM" & srechnertab & " order by Umsatzvm desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!PlatzUmsatzM = lAnz
            lAnz = lAnz + 1
            rsrs.Update
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    lAnz = 1
    cSQL = "Select * from LIVM" & srechnertab & " order by AbsatzVM desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!PlatzSTCKM = lAnz
            lAnz = lAnz + 1
            rsrs.Update
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    lAnz = 1
    cSQL = "Select * from LIVM" & srechnertab & " order by Umsatzaktj desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!PlatzUmsatzJ = lAnz
            lAnz = lAnz + 1
            rsrs.Update
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    lAnz = 1
    cSQL = "Select * from LIVM" & srechnertab & " order by Absatzaktj desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!PlatzSTCKJ = lAnz
            lAnz = lAnz + 1
            rsrs.Update
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    loeschNEW "LIVMPRINT", gdBase
    CreateTableT2 "LIVMPRINT", gdBase
    
    cSQL = "Insert into LIVMPRINT select * from LIVM" & srechnertab
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "LIVM" & srechnertab, gdBase
    
    
    'Kopfdaten
    loeschNEW "LIVMKOPF", gdBase
    CreateTableT2 "LIVMKOPF", gdBase
    
    Dim sdat As String
    Dim sBasis As String
    sdat = MonthName(imon) & " " & iJahr
    sBasis = "1 Geschäft"
    
    cSQL = "Insert into LIVMKOPF (UEBER,Auswertungsdat,Basis) values ('Lieferanten','" & sdat & "','" & sBasis & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "Druckvorschau wird erstellt...", Label1(6)
    reportbildschirm "", "zZEN17c"
    
    
    
    anzeige "normal", "", Label1(6)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "VorMonatsAuswertung"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeereLieferantMaskeWKL17()
    On Error GoTo LOKAL_ERROR
    
    Dim lWert As Long
    
    For lWert = 5 To 27
        Text1(lWert).Text = ""
    Next lWert
    
    Check1.Value = vbUnchecked

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereLieferantMaskeWKL17"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cSuch As String
    Dim cMeld As String
    Dim iRet As Integer
    Dim lLinr As Long
    Dim sSQL As String
    Dim i As Integer
    Dim sdateiname As String
    Dim cdatei As String
    Dim cPfad1 As String
    Dim cPfad As String
    Dim rsrs As Recordset
    
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    
    Select Case Index
        Case Is = 0     'Auswählen
            
            If IsNumeric(gsLinr) Then
                cSuch = gsLinr
            Else
                If MSFlexGrid1.Row < 1 Then
                    Screen.MousePointer = 0
                    MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                    Exit Sub
                End If
                cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
                cSuch = Trim$(cSuch)
                
                If IsNumeric(cSuch) Then
        
                Else
                    Screen.MousePointer = 0
                    MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                    Exit Sub
                End If
            
            End If
            
            Screen.MousePointer = 11
            
            If Val(cSuch) > 0 Then
                HoleDatenWKL17 cSuch
                Me.Refresh
                Screen.MousePointer = 11
                Label4(222).Caption = Format$(Einkaufsumsatzermittlung(cSuch, gdBase, CInt(Year(Now))), "########0.00") & " " & gcWaehrung
                Label4(41).Caption = Format$(Einkaufsumsatzermittlung(cSuch, gdBase, CInt(Year(Now) - 1)), "########0.00") & " " & gcWaehrung
                Me.Refresh
                Screen.MousePointer = 11
                Label4(200).Caption = Format$(LAGEREKermittlung(cSuch), "########0.00") & " " & gcWaehrung
                Screen.MousePointer = 0
            End If
            
            If gbBILDTAST = False Then
                Frame2.Visible = False
            Else
                Frame2.Visible = True
            End If
            
        Case Is = 1     'Schließen
            Frame1.Visible = False
            
            Frame0.Visible = True
            Text1(0).Enabled = True
            Text1(1).Enabled = True
            Text1(2).Enabled = True
            Text1(3).Enabled = True
            Text1(4).Enabled = True
            LeereDialogWKL17
            Text1(0).SetFocus
        Case Is = 2     'Löschen
        
            If MSFlexGrid1.Row < 1 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
            cSuch = Trim$(cSuch)
            
            If IsNumeric(cSuch) Then
    
            Else
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            
          
            
            Screen.MousePointer = 11
            If Val(cSuch) > 0 Then
            
                Dim rsKJ As Recordset
                Dim cSQL As String
                Dim cDatum As String
                Dim cMindat As String
                Dim lMindat As Long
    
                lMindat = DateValue(Now) - 180
                cMindat = Trim$(Str$(lMindat))
                
                cSQL = "Select adate from Kassjour Where LINR = " & cSuch & " and adate > " & cMindat & " order by adate desc "
                Set rsKJ = gdBase.OpenRecordset(cSQL)
                If Not rsKJ.EOF Then
                    rsKJ.MoveFirst
                    
                    If Not IsNull(rsKJ!ADATE) Then
                        cDatum = rsKJ!ADATE
                    End If
                    
                    cMeld = "ACHTUNG!" & vbCrLf & vbCrLf
                    cMeld = cMeld & "Lieferant Nr. " & cSuch & " soll gelöscht werden." & vbCrLf & vbCrLf
                    cMeld = cMeld & "Artikel von diesem Lieferanten sind am " & cDatum & vbCrLf
                    cMeld = cMeld & "das letzte Mal verkauft worden" & vbCrLf & vbCrLf
                    
                    cMeld = cMeld & "Wollen Sie den Lieferanten trotzdem löschen?"
                    iRet = MsgBox(cMeld, vbYesNo + vbQuestion, "LÖSCHEN")
                    If iRet = vbYes Then
                        LoescheLieferantWKL17 cSuch
                        If SucheLieferantWKL17 Then
                            zeige_Grid
                            If MSFlexGrid1.Visible = True Then
                                MSFlexGrid1.Col = 1
                                MSFlexGrid1.Row = 2
                                MSFlexGrid1.SetFocus
                                
                            End If
                        End If
                    End If
                    
                Else
                    cMeld = "ACHTUNG!" & vbCrLf & vbCrLf
                    cMeld = cMeld & "Lieferant Nr. " & cSuch & " soll gelöscht werden." & vbCrLf & vbCrLf
                    cMeld = cMeld & "Das Löschen eines Lieferanten kann zu Unstimmigkeiten" & vbCrLf
                    cMeld = cMeld & "in der Datenbank führen, wenn der Lieferant bereits" & vbCrLf
                    cMeld = cMeld & "Artikel geliefert hat!" & vbCrLf & vbCrLf
                    cMeld = cMeld & "Wollen Sie den Lieferanten trotzdem löschen?"
                    iRet = MsgBox(cMeld, vbYesNo + vbQuestion, "LÖSCHEN")
                    If iRet = vbYes Then
                        LoescheLieferantWKL17 cSuch
                        If SucheLieferantWKL17 Then
                            zeige_Grid
                            If MSFlexGrid1.Visible = True Then
                                MSFlexGrid1.Col = 1
                                MSFlexGrid1.Row = 2
                                MSFlexGrid1.SetFocus
                                
                            End If
                        End If
                    End If
                End If
                
                rsKJ.Close
                
            End If
            
        Case 3
            If MSFlexGrid1.Row < 1 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
            cSuch = Trim$(cSuch)
            
            If IsNumeric(cSuch) Then
    
            Else
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            
            Dim lrow As Long
            Dim lcol As Long
            lrow = MSFlexGrid1.Row
            lcol = MSFlexGrid1.Col
            
            MSFlexGrid1_KeyUp vbKeyF3, 0
            MSFlexGrid1.Col = lcol
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.SetFocus
        Case 4
        
            If MSFlexGrid1.Row < 1 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
            cSuch = Trim$(cSuch)
            
            If IsNumeric(cSuch) Then
    
            Else
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            
            If Val(cSuch) > 0 Then
                glLiNr = Val(cSuch)
                giKissFtpMode = 17 ' spezielle Lieferanten Stammdaten holen
                frmWKL38.Show 1
                
                iRet = (MsgBox("Möchten Sie die Stammdaten sofort einlesen?", vbQuestion + vbYesNo, "Winkiss Frage:"))
                If iRet = vbYes Then
                    frmWKL11.Show 1
                Else
                
                End If
                
            End If
        Case 8
            If MSFlexGrid1.Row < 1 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
            cSuch = Trim$(cSuch)
            
            If IsNumeric(cSuch) Then
    
            Else
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            
            lrow = MSFlexGrid1.Row
            lcol = MSFlexGrid1.Col
            
            If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)) > 0 Then
                gcSuch = "LINR" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
                frmWKL148.Show 1
                Me.Refresh
                gclinr = ""
            End If
            
            MSFlexGrid1.Col = lcol
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.SetFocus
        Case 5
            If MSFlexGrid1.Row < 1 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
            cSuch = Trim$(cSuch)
            
            If IsNumeric(cSuch) Then
    
            Else
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            
            lrow = MSFlexGrid1.Row
            lcol = MSFlexGrid1.Col
            
            If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)) > 0 Then
                gclinr = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
                frmWKL150.Show 1
                Me.Refresh
                gclinr = ""
            End If
            
            MSFlexGrid1.Col = lcol
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.SetFocus
            
        Case 6
            If MSFlexGrid1.Row < 1 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
            cSuch = Trim$(cSuch)
            
            If IsNumeric(cSuch) Then
    
            Else
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            
            lrow = MSFlexGrid1.Row
            lcol = MSFlexGrid1.Col
            
            If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)) > 0 Then
                gclinr = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
                frmWKL149.Show 1
                Me.Refresh
                gclinr = ""
            End If
            
            MSFlexGrid1.Col = lcol
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.SetFocus
        Case 7
            loeschNEW "LiExc", gdBase
                
            gsZSpalte = "Linr"
            gstab = "BEALIEFX"
            frmWKL36.Show 1
            
            'danach Tablay auswerten
            
            FormatGridOverTablay "BEALIEFX"
            
            If byAnzahlSpalten > 0 Then
                sSQL = "Select " & sSpaltenbez(0) & " "
                
                If byAnzahlSpalten > 1 Then
                    For i = 1 To byAnzahlSpalten - 1
                        sSQL = sSQL & " , " & sSpaltenbez(i) & "  "
                    Next i
                End If
            Else
                Exit Sub
            End If
            
            sSQL = sSQL & " into LiExc from LI" & srechnertab
            gdBase.Execute sSQL, dbFailOnError
        
            
            Dim iFileNr As Integer
            Dim sPfad   As String
           
            Dim sAusgabedatname As String
            
            If sdateiname = "kein Betreff" Then
                sAusgabedatname = "Lieferanten" & ".xls"
            Else
                sAusgabedatname = sdateiname & ".xls"
            End If
            
            cdatei = cPfad1 & "BOX\" & sAusgabedatname
            cPfad = cPfad1 & "BOX"

        
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .DialogTitle = "Speichern der Lieferantenexceltabelle"
                .Filter = "Excel - Dateien (*.xls)|*.xls"
                .FileName = cPfad & "\" & sAusgabedatname
                .ShowSave
            End With
        
            sPfad = cdlopen.FileName
            
            If FileExists(sPfad) Then
                iRet = MsgBox("Eine gleichnamige Datei ist schon vorhanden, möchten Sie diese überschreiben?", vbQuestion + vbYesNo, "Winkiss Frage:")
                If iRet = vbNo Then
                    Exit Sub
                Else
                    Kill sPfad
                End If
            Else
            
            End If

            sSQL = "Select * into LiExc IN '" & sPfad & "' 'Excel 8.0;' from LiExc "
            gdBase.Execute sSQL, dbFailOnError

            MsgBox "Diese Datei ist unter (" & sPfad & ") abgespeichert", vbInformation, "Winkiss Information:"
        Case 9 ' alle roten Löschen
        
            Screen.MousePointer = 11
        
            loeschNEW "disLi", gdBase
            sSQL = "Select distinct(linr) into disLi from artlief "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Alter table LISRT add erkannt Text(1) "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update Lisrt set erkannt = 'N' "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update Lisrt set erkannt = 'J' where linr in (Select linr from disli ) "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "delete * from Lisrt where erkannt = 'N' "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Alter table LISRT drop column erkannt "
            gdBase.Execute sSQL, dbFailOnError
            
            
'            sSQL = "delete * from Lisrt where not linr in (Select linr from disli ) "
'            gdBase.Execute sSQL, dbFailOnError
            
            loeschNEW "disLi", gdBase
            
            If SucheLieferantWKL17 Then
                zeige_Grid
                If MSFlexGrid1.Visible = True Then
                    MSFlexGrid1.Col = 1
                    MSFlexGrid1.Row = 2
                    MSFlexGrid1.SetFocus
                    
                End If
            End If
            
            Screen.MousePointer = 0
        Case 10 'Zielsetzung
        
            If MSFlexGrid1.Row < 1 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            cSuch = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
            cSuch = Trim$(cSuch)
            
            If IsNumeric(cSuch) Then
    
            Else
                Screen.MousePointer = 0
                MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            
            lrow = MSFlexGrid1.Row
            lcol = MSFlexGrid1.Col
            
            If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)) > 0 Then
                lLinr = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
                
                liefstatEWErstellen lLinr
                Detaildaten lLinr
                Label4(36).Caption = lLinr
                Frame4.Visible = True
                
            End If
            
            MSFlexGrid1.Col = lcol
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.SetFocus
        Case 11 ' alle gelben Löschen = ohne Bestand
        
            Screen.MousePointer = 11
        
            loeschNEW "disLi", gdBase
            sSQL = "Select * into disLi from artlief "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Alter Table disli add Bestand Long "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "update disLi inner join artikel on disli.artnr = artikel.artnr   "
            sSQL = sSQL & " set disli.bestand = artikel.bestand "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Delete from disLi where bestand = 0 "
            gdBase.Execute sSQL, dbFailOnError
            
            loeschNEW "disLiDel", gdBase
            
            sSQL = "Select distinct(linr) into disLiDel from disLi "
            gdBase.Execute sSQL, dbFailOnError
            
            loeschNEW "LiDel", gdBase
            
            sSQL = "Select * into LiDEl  from Lisrt where linr not in (Select linr from disLiDel ) "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Delete from LiDEl where linr between 500000 and 600000 "
            gdBase.Execute sSQL, dbFailOnError
            
            
            
            Set rsrs = gdBase.OpenRecordset("LIDEL")
            If Not rsrs.EOF Then
                Do While Not rsrs.EOF
                
                    If Not IsNull(rsrs!linr) Then
                        LoescheLieferantWKL17 Trim(rsrs!linr)
                    End If
                rsrs.MoveNext
                Loop
            
            End If
            rsrs.Close
            
            If SucheLieferantWKL17 Then
                zeige_Grid
                If MSFlexGrid1.Visible = True Then
                    MSFlexGrid1.Col = 1
                    MSFlexGrid1.Row = 2
                    MSFlexGrid1.SetFocus

                End If
            End If
            
            loeschNEW "LIDEL", gdBase
            loeschNEW "disLiDel", gdBase
            loeschNEW "disLi", gdBase
            
            Screen.MousePointer = 0
        Case 12

            frmWKL214.Show 1
            Me.Refresh

    End Select
    Screen.MousePointer = 0
    

err:
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LoescheLieferantWKL17(cLinr As String)
    On Error GoTo LOKAL_ERROR
        
    Dim sSQL As String
    

    
''''    cSQL = "Delete from Artlief where LINR = " & cLinr & " "
''''    gdBase.Execute cSQL, dbFailOnError
''''
''''    cSQL = "Delete from Artikel where LINR = " & cLinr & " "
''''    gdBase.Execute cSQL, dbFailOnError

    loeschNEW "artlief_T", gdBase
    
    sSQL = "Select * into artlief_T from artlief where linr = " & cLinr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into artlief_T select "
    
    sSQL = sSQL & " artlief.artnr  "
    sSQL = sSQL & ", artlief.LINR  "
    sSQL = sSQL & ", artlief.LEKPR  "
    sSQL = sSQL & ", artlief.LIBESNR "
    sSQL = sSQL & ", artlief.MINMEN  "
    sSQL = sSQL & ", artlief.SPANNE  "
    sSQL = sSQL & ", artlief.SYNSTATUS  "
    sSQL = sSQL & ", artlief.EXDAT  "
    sSQL = sSQL & ", artlief.RKZ  "
    sSQL = sSQL & " from artlief inner join artlief_t on artlief.artnr = artlief_t.artnr"
    sSQL = sSQL & " Where artlief.linr <> " & cLinr
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ImportDupli", gdBase
    
    sSQL = "select count(artnr) as count ,artnr into ImportDupli from artlief_t group by artnr having count(artnr) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from ImportDupli where artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from Artlief where artnr in (Select artnr from ImportDupli) and artlief.linr = " & cLinr
    gdBase.Execute sSQL, dbFailOnError

    
    loeschNEW "artlief_T", gdBase
    loeschNEW "ImportDupli", gdBase
    
    'Teil 2
    'Check die übergebliebenen Einzelkombinationen auf Verkäufe
    
    sSQL = "Select * into artlief_T from artlief where linr = " & cLinr
    gdBase.Execute sSQL, dbFailOnError
    
'    SpalteAnfuegenNEW "artlief_T", "verkauft", "BIT", gdBase
'
'    sSQL = "Update artlief_T inner join kassjour on artlief_t.artnr = kassjour.artnr"
'    sSQL = sSQL & " set artlief_T.verkauft = True "
'    gdBase.Execute sSQL, dbFailOnError
'
'    sSQL = "delete from artlief_T where verkauft = True "
'    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Create Index artnr on artlief_T (artnr)"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from Artlief where artnr in (Select artnr from artlief_T) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from Artikel where artnr in (Select artnr from artlief_T) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from ARTEAN_K where artnr in (Select artnr from artlief_T) "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "artlief_T", gdBase
    
    
    'Bleib am Ende nichts übrig? dann löschen
    If DatendrinSQL("Select * from Artlief where LINR = " & cLinr, gdBase) = False Then
        sSQL = "Delete from LISRT where LINR = " & cLinr & " "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheLieferantWKL17"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub HoleDatenWKL17(cSuch As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim ctmp As String
    Dim dWert As Double
    
    cSQL = "Select * from LISRT where LINR = " & cSuch & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!linr) Then
            ctmp = rsrs!linr
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(5).Text = ctmp
        
        If Not IsNull(rsrs!Kuerzel) Then
            ctmp = rsrs!Kuerzel
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(6).Text = ctmp
        
        If Not IsNull(rsrs!LIEFBEZ) Then
            ctmp = rsrs!LIEFBEZ
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(7).Text = ctmp
        
        If Not IsNull(rsrs!strasse) Then
            ctmp = rsrs!strasse
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(8).Text = ctmp
        
        If Not IsNull(rsrs!Plz) Then
            ctmp = rsrs!Plz
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(9).Text = ctmp
        
        If Not IsNull(rsrs!STADT) Then
            ctmp = rsrs!STADT
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(10).Text = ctmp
        
        If Not IsNull(rsrs!Tel) Then
            ctmp = rsrs!Tel
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(11).Text = ctmp
        
        If Not IsNull(rsrs!Fax) Then
            ctmp = rsrs!Fax
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(12).Text = ctmp
        
        If Not IsNull(rsrs!KTEXT) Then
            ctmp = rsrs!KTEXT
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(13).Text = ctmp
        
        If Not IsNull(rsrs!Kundnr) Then
            ctmp = rsrs!Kundnr
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(14).Text = ctmp
        
        If Not IsNull(rsrs!NOTIZ) Then
            ctmp = rsrs!NOTIZ
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(15).Text = ctmp
        
        If Not IsNull(rsrs!AWERT) Then
            ctmp = rsrs!AWERT
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(16).Text = ctmp
        
        If Not IsNull(rsrs!adress) Then
            ctmp = rsrs!adress
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(17).Text = ctmp
        
        If Not IsNull(rsrs!bUser) Then
            ctmp = rsrs!bUser
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(18).Text = ctmp
        
        If Not IsNull(rsrs!Pass) Then
            ctmp = rsrs!Pass
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(19).Text = ctmp
        
        If Not IsNull(rsrs!KennNr) Then
            ctmp = rsrs!KennNr
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(20).Text = ctmp
        
        If Not IsNull(rsrs!Format) Then
            ctmp = rsrs!Format
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(21).Text = ctmp
        
        If Not IsNull(rsrs!Email) Then
            ctmp = rsrs!Email
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(23).Text = ctmp
        
        If Not IsNull(rsrs!br) Then
            ctmp = rsrs!br
        Else
            ctmp = "15"
        End If
        ctmp = Trim$(ctmp)
        Select Case ctmp
            Case 1
                Combo1.Text = "Montag,     gerade KW" & Space(100) & "1"
            Case 2
                Combo1.Text = "Dienstag,   gerade KW" & Space(100) & "2"
            Case 3
                Combo1.Text = "Mittwoch,   gerade KW" & Space(100) & "3"
            Case 4
                Combo1.Text = "Donnerstag, gerade KW" & Space(100) & "4"
            Case 5
                Combo1.Text = "Freitag,    gerade KW" & Space(100) & "5"
            Case 6
                Combo1.Text = "Samstag,    gerade KW" & Space(100) & "6"
            Case 7
                Combo1.Text = "Sonntag,    gerade KW" & Space(100) & "7"
                
                
            Case 8
                Combo1.Text = "Montag,     ungerade KW" & Space(100) & "8"
            Case 9
                Combo1.Text = "Dienstag,   ungerade KW" & Space(100) & "9"
            Case 10
                Combo1.Text = "Mittwoch,   ungerade KW" & Space(100) & "10"
            Case 11
                Combo1.Text = "Donnerstag, ungerade KW" & Space(100) & "11"
            Case 12
                Combo1.Text = "Freitag,    ungerade KW" & Space(100) & "12"
            Case 13
                Combo1.Text = "Samstag,    ungerade KW" & Space(100) & "13"
            
            Case 14
                Combo1.Text = "Sonntag,    ungerade KW" & Space(100) & "14"
            Case 15
                Combo1.Text = "kein Bestellrhythmus" & Space(100) & "15"
        End Select
        
        If Not IsNull(rsrs!HL) Then
            If rsrs!HL = True Then
                Check1.Value = vbChecked
            Else
                Check1.Value = Unchecked
            End If
        Else
            Check1.Value = Unchecked
        End If
        
        If Not IsNull(rsrs!GLN) Then
            ctmp = rsrs!GLN
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(25).Text = ctmp
      
        If Not IsNull(rsrs!USTID) Then
            ctmp = rsrs!USTID
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(24).Text = ctmp
        
        If Not IsNull(rsrs!DEPOTRABATT1) Then
            ctmp = rsrs!DEPOTRABATT1
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(26).Text = ctmp
        
        If Not IsNull(rsrs!LEK_ABSCHLAG) Then
            ctmp = rsrs!LEK_ABSCHLAG
        Else
            ctmp = ""
        End If
        ctmp = Trim$(ctmp)
        Text1(27).Text = ctmp

    
        


        Label3(2).Caption = 5
        
        
        Frame3.Visible = True
        Frame1.Visible = False
       
       
        If gsLinr = "" Then
            Text1(5).SetFocus
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleDatenWKL17"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    
    iZielIndex = Label3(2).Caption
    
    If iZielIndex = 0 Or iZielIndex = 5 Then
        If Index >= 11 And Index <= 20 Then
            Text1(iZielIndex).Text = Text1(iZielIndex).Text & Command3(Index).Caption
        End If
    Else
        Text1(iZielIndex).Text = Text1(iZielIndex).Text & Command3(Index).Caption
    End If
    Text1(iZielIndex).SetFocus

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    

End Sub
Private Sub Command3_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    
    iZielIndex = Label3(2).Caption
    If Frame0.Visible Then
        Text1(iZielIndex).BackColor = glSelBack1
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    Dim lcount As Long
    
    Set WshShell = CreateObject("WScript.Shell")
    'WshShell.SendKeys "+{Tab}", True
    
    iZielIndex = Label3(2).Caption
        
    Select Case Index
        Case 0      'CLEAR
            Text1(iZielIndex).Text = ""
            Text1(iZielIndex).SetFocus
        Case 1      'RÜCKG
            If Len(Text1(iZielIndex).Text) > 0 Then
                Text1(iZielIndex).Text = Left(Text1(iZielIndex).Text, Len(Text1(iZielIndex).Text) - 1)
            End If
            Text1(iZielIndex).SetFocus
        Case 2      'BEFORE
            WshShell.SendKeys "+{Tab}", True
            Text1(iZielIndex).SetFocus
        Case 3     'NEXT
            WshShell.SendKeys "{Tab}", True
            Text1(iZielIndex).SetFocus
        Case 4     'switch to UPPER CASE
            SwitchUpperLowerCaseWKL17
            Text1(iZielIndex).SetFocus
        Case 5
            Frame5.Visible = False
        Case 6
            Frame5.Visible = True
        Case 7
            Command5(0).Enabled = True
            Command5(1).Enabled = True
            Command5(2).Enabled = True
            For lcount = 5 To 30
                Text1(lcount).Enabled = True
            Next lcount
        Case 11
            gsHelpstring = "Lieferanten bearbeiten"
            frmWKL110.Show 1
    End Select
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SwitchUpperLowerCaseWKL17()
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
    Fehler.gsFunktion = "SwitchUpperLowerCaseWKL17"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command4_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    
    iZielIndex = Label3(2).Caption
    
    If iZielIndex < 0 Then
        iZielIndex = 0
    End If
    Text1(iZielIndex).BackColor = glSelBack1
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command5_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    Dim i As Integer
    Dim cPos(1 To 5) As Long
    Dim lRet As Long
    Dim lcount As Long
    Dim cLinr As String
    Dim cTxT As String
    Dim sKorr As String
    Dim lLinr As Long
    
    Select Case Index
    
        Case 2  'Drucken
            lLinr = Val(Text1(5).Text)
            If lLinr = 0 Then Exit Sub
            LieferantDrucken lLinr
            
        Case 3  'Details speichern
            lLinr = Label4(36).Caption
            Detailspeichern lLinr
        Case 0  'Speichern
            '//KUERZEL
            cTxT = Text1(6).Text
            cPos(1) = InStr(1, cTxT, "'")
            cPos(2) = InStr(1, cTxT, ";")
            cPos(3) = InStr(1, cTxT, ",")
            cPos(4) = InStr(1, cTxT, "!")
            cPos(5) = InStr(1, cTxT, "*")
                
            For i = 1 To 5
                If cPos(i) <> 0 Then
                    MsgBox " In Eingabefelder dürfen folgende Zeichen : 'Apostroph ;Semikolon ,Koma !Ausrufezeichen *Sternchen  nicht enthalten !", vbOKOnly, "STOP"
                    MsgBox " Bitte in Feld KUERZEL neu eingeben !", vbOKOnly, " Lösung"
                    Text1(6).Text = ""
                End If
            Next
            '//Lief.Bez
            cTxT = Text1(7).Text
            cPos(1) = InStr(1, cTxT, "'")
            cPos(2) = InStr(1, cTxT, ";")
            cPos(3) = InStr(1, cTxT, ",")
            cPos(4) = InStr(1, cTxT, "!")
            cPos(5) = InStr(1, cTxT, "*")
                
            For i = 1 To 5
                If cPos(i) <> 0 Then
                    MsgBox " In Eingabefelder dürfen folgende Zeichen : 'Apostroph ;Semikolon ,Koma !Ausrufezeichen *Sternchen  nicht enthalten !", vbOKOnly, "STOP"
                    MsgBox " Bitte in Feld Lief.Bez neu eingeben !", vbOKOnly, " Lösung"
                    Text1(7).Text = ""
                End If
            Next
            '//End Aenderung
            
            lRet = fnPruefeDialogEingabenWKL17()
            If lRet = 0 Then
                Text1(6).Text = UCase$(Text1(6).Text)
                lcount = fnPruefeEingabeDialogWKL17()
                Select Case lcount
                    Case Is = 0
                    
                        If (Text1(22).Text) <> "" Then
                            sKorr = Text1(22).Text
                            Einkaufsumsatzkorrektur Trim(Text1(5).Text), sKorr, CLng(DateValue(Label4(40).Caption))
                        End If
                        Text1(22).Text = ""
                        SchreibeDatenWKL17
                        
                        If gsLinr <> "" Then
                            Unload Me
                            gsLinr = ""
                            Screen.MousePointer = 0
                            Exit Sub
                        Else
                            Command5_Click 1
                        End If
                        
                    Case Is = 1
                        MsgBox "Das Feld LIEF.NR darf nur Ziffern enthalten!", vbCritical, "STOP!"
                        Text1(5).SetFocus
                        
                    Case Is = 3
                        MsgBox "Das Feld A-WERT darf nur Ziffern enthalten!", vbCritical, "STOP!"
                        Text1(16).SetFocus

                        
                End Select
            Else
                MsgBox "Das Feld enthält ein ungültiges Zeichen!", vbCritical, "STOP!"
                Text1(lRet).SetFocus
            End If
            
        Case 1
            If gsLinr <> "" Then
                Unload Me
                gsLinr = ""
                Screen.MousePointer = 0
                Exit Sub
            End If

            If giDlgZustand = giNEU Then
                Frame0.Visible = True
                Frame3.Visible = False
                Frame1.Visible = False
                
                Text1(0).Enabled = True
                Text1(1).Enabled = True
                Text1(2).Enabled = True
                Text1(3).Enabled = True
                Text1(4).Enabled = True
                Text1(0).SetFocus
            Else
                Frame1.Visible = True
                Frame2.Visible = False
                Frame3.Visible = False
                Command1_Click 0
            End If
            giDlgZustand = giUPD
            Text1(22).Text = ""
        Case 5
            Frame4.Visible = False
        Case 20
            Screen.MousePointer = 0
            Label4(40) = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Function fnPruefeDialogEingabenWKL17() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim iCount As Integer
    Dim ctmp As String
    Dim cZeichen As String
    Dim cValid As String
    
    For lcount = 5 To 16
        ctmp = Text1(lcount).Text
        ctmp = Trim$(ctmp)
        Select Case lcount
            Case 5, 16
                cValid = "1234567890"
                If ctmp <> "" Then
                    For iCount = 1 To Len(ctmp)
                        cZeichen = Mid(ctmp, iCount, 1)
                        If InStr(cValid, cZeichen) = 0 Then
                            fnPruefeDialogEingabenWKL17 = lcount
                            Exit Function
                        End If
                    Next iCount
                End If
            Case 13
                cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
                cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
                cValid = cValid & "+äÄÜüÖöß"
                If ctmp <> "" Then
                    For iCount = 1 To Len(ctmp)
                        cZeichen = Mid(ctmp, iCount, 1)
                        If InStr(cValid, cZeichen) = 0 Then
                            fnPruefeDialogEingabenWKL17 = lcount
                            Exit Function
                        End If
                    Next iCount
                End If
        End Select
        
    Next lcount
    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeDialogEingabenWKL17"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub SchreibeDatenWKL17()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cKey        As String
    Dim dWert       As Double
    Dim ctmp        As String
    Dim cMaxNr      As String
    Dim bHauptlief  As Boolean
    
    bHauptlief = False
    If Check1.Value = vbChecked Then
        bHauptlief = True
    End If
    
    If Trim$(Text1(26).Text) = "" Then Text1(26).Text = "0"
    
    If Trim$(Text1(27).Text) = "" Then Text1(27).Text = "0"
    
    cKey = Text1(5).Text
    cKey = Trim$(cKey)
    
    If cKey = "" Then
        MsgBox "Schreiben nicht möglich, da keine Lieferantennummer vorhanden!", vbCritical, "STOP!"
        Text1(5).SetFocus
        Exit Sub
    End If
    
    cSQL = "Select max(LINR) from LISRT where LINR > 499999 "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs.Fields(0)) Then
            dWert = rsrs.Fields(0)
        Else
            dWert = 499999
        End If
    Else
        dWert = 499999
    End If
    rsrs.Close: Set rsrs = Nothing
    
    dWert = dWert + 1
    cMaxNr = Format$(dWert, "#####0")
    
    cSQL = "Select * from LISRT where LINR = " & cKey & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        If giDlgZustand = giNEU Then
            MsgBox "Lieferantennummer bereits vorhanden!" & vbCrLf & "Bitte andere Lieferantennummer eingeben (nächste freie Nummer = " & cMaxNr & " ).", vbCritical, "DOPPELTE NUMMER"
            Text1(5).SetFocus
            Exit Sub
        End If
            
        rsrs.Edit
        rsrs!SYNStatus = "E"
    Else
        rsrs.AddNew
        rsrs!SYNStatus = "A"
    End If
    
    ctmp = Trim$(Text1(5).Text)
    
    If ctmp = "" Then
        rsrs!linr = Null
    Else
        rsrs!linr = ctmp
    End If
    rsrs!Kuerzel = Trim$(Text1(6).Text)
    rsrs!LIEFBEZ = Trim$(Text1(7).Text)
    rsrs!strasse = Trim$(Text1(8).Text)
    rsrs!Plz = Trim$(Text1(9).Text)
    rsrs!STADT = Trim$(Text1(10).Text)
    rsrs!Tel = Trim$(Text1(11).Text)
    rsrs!Fax = Trim$(Text1(12).Text)
    rsrs!KTEXT = Trim$(Text1(13).Text)
    rsrs!Kundnr = Trim$(Text1(14).Text)
    rsrs!NOTIZ = Trim$(Text1(15).Text)
    rsrs!AWERT = Val(Trim$(Text1(16).Text))
    rsrs!adress = Trim$(Text1(17).Text)
    rsrs!bUser = Trim$(Text1(18).Text)
    rsrs!Pass = Trim$(Text1(19).Text)
    rsrs!KennNr = Trim$(Text1(20).Text)
    rsrs!Format = Trim$(Text1(21).Text)
    rsrs!Email = Trim$(Text1(23).Text)
    rsrs!HL = bHauptlief
    rsrs!br = CByte(Right$(Combo1.Text, 2))
    
    rsrs!GLN = Trim$(Text1(25).Text)
    rsrs!USTID = Trim$(Text1(24).Text)
    
    rsrs!DEPOTRABATT1 = Trim$(Text1(26).Text)
    
    rsrs!LEK_ABSCHLAG = Trim$(Text1(27).Text)
    
    '//Felder LastDate und LastTime
    rsrs!LASTDATE = DateValue(Now)
    rsrs!LASTTIME = TimeValue(Now)
    
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    If bHauptlief Then
        Insertgrolief ctmp
    Else
        delgrolief ctmp
    End If
    
    LeereLieferantMaskeWKL17
    
    Text1(5).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWKL17"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
'    Resume Next
    
End Sub
Private Sub Detaildaten(lLinr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim sSQL As String
    Dim rsrs As Recordset
    
    sSQL = "Select * from LISRT where linr = " & lLinr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Zielek) Then
            Text4(0).Text = rsrs!Zielek
        Else
            Text4(0).Text = "0"
        End If
    Else
        Text4(0).Text = "0"
    End If
    rsrs.Close: Set rsrs = Nothing
    
    For i = 1 To 12
        Label6(i).Caption = MonthName(CLng(i))
        Label7(i).Caption = MonthName(CLng(i))
    Next i
    
    Label4(37).Caption = Year(DateValue(Now)) - 2
    Label4(28).Caption = Year(DateValue(Now)) - 1
    Label4(31).Caption = Year(DateValue(Now))
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Detaildaten"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub LieferantDrucken(lLinr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    loeschNEW "LIEFBLATT", gdBase
    CreateTableT2 "LIEFBLATT", gdBase
    
    cSQL = "Insert into LIEFBLATT select "
    cSQL = cSQL & " LINR  "
    cSQL = cSQL & ", LIEFBEZ "
    cSQL = cSQL & ", KUERZEL "
    cSQL = cSQL & ", PLZ "
    cSQL = cSQL & ", STADT "
    cSQL = cSQL & ", STRASSE "
    cSQL = cSQL & ", TEL "
    cSQL = cSQL & ", FAX "
    cSQL = cSQL & ", KUNDNR  "
    cSQL = cSQL & ", NOTIZ  "
    cSQL = cSQL & ", AWERT "
    cSQL = cSQL & ", KTEXT "
    cSQL = cSQL & ", HL "
    cSQL = cSQL & ", ZIELEK "
    cSQL = cSQL & ", EMAIL "
    cSQL = cSQL & " from LISRT where linr = " & lLinr & " "
    gdBase.Execute cSQL, dbFailOnError
    
    reportbildschirm "dWKL001b", "aWKL17b"
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LieferantDrucken"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Detailspeichern(lLinr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Text4(0).Text <> "" Then
        If IsNumeric(Text4(0).Text) Then
            sSQL = "Update LISRT set ZIELEK = '" & Text4(0).Text & "'"
            sSQL = sSQL & " where Linr = " & lLinr
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Detailspeichern"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function ermUmsatzEK(i As Integer, jahr As String, lLinr As Long) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsEK As Recordset
    
    ermUmsatzEK = "0"
    
    sSQL = "Select Sum(Kassjour.EKPR*kassjour.Menge) as merg "
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " Where "
    sSQL = sSQL & " month(ADATE) = " & i
    sSQL = sSQL & " and Year(ADATE) = " & jahr
    sSQL = sSQL & " and linr = " & lLinr
    
    Set rsEK = gdBase.OpenRecordset(sSQL)
    If Not rsEK.EOF Then
        If Not IsNull(rsEK!merg) Then
            ermUmsatzEK = rsEK!merg
        End If
    End If
    rsEK.Close
    ermUmsatzEK = Format$(ermUmsatzEK, "#######0.00")

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermUmsatzEK"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Sub liefstatEWErstellen(lLinr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim j As Integer
    Dim sSQL As String
    Dim cUEKA As String
    Dim cUEKAvv As String
    Dim cUEKAaJ As String
    Dim dVorjahrsum As Double
    Dim dvVorjahrsum As Double
    Dim dAKTjahrsum As Double
    Dim rsrs As Recordset
    
    dVorjahrsum = 0
    dAKTjahrsum = 0
    loeschNEW "LIEFENT", gdBase
    CreateTable "LIEFENT", gdBase

    picprogress.Visible = True

    j = 1
    For i = 1 To 12
        txtStatus.Text = j
        cUEKA = ermUmsatzEK(i, Year(DateValue(Now)) - 1, lLinr)
        j = j + 2
        txtStatus.Text = j
        cUEKAvv = ermUmsatzEK(i, Year(DateValue(Now)) - 2, lLinr)
        j = j + 2
        txtStatus.Text = j
        cUEKAaJ = ermUmsatzEK(i, Year(DateValue(Now)), lLinr)
        j = j + 2
        txtStatus.Text = j
        sSQL = "Insert into LiefENT (UEKAVV,UEKA,UEKAaj,monat) values ('" & cUEKAvv & "','" & cUEKA & "','" & cUEKAaJ & "'," & i & ") "
        gdBase.Execute sSQL, dbFailOnError

        j = j + 2
        txtStatus.Text = j
    Next i
    
    sSQL = " select * from LIEFENT "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!ueka) Then
                Label8(rsrs!Monat).Caption = Format$(rsrs!ueka, "###,##0.00 €")
            Else
                Label8(rsrs!Monat).Caption = "0.00 €"
            End If
            
            If Not IsNull(rsrs!uekaaj) Then
                Label9(rsrs!Monat).Caption = Format$(rsrs!uekaaj, "###,##0.00 €")
            Else
                Label9(rsrs!Monat).Caption = "0.00 €"
            End If
            
            If Not IsNull(rsrs!uekavv) Then
                Label15(rsrs!Monat).Caption = Format$(rsrs!uekavv, "###,##0.00 €")
            Else
                Label15(rsrs!Monat).Caption = "0.00 €"
            End If
        
        rsrs.MoveNext
        Loop
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    j = j + 2
    txtStatus.Text = j
    
    
    For i = 1 To 12
        dvVorjahrsum = dvVorjahrsum + CDbl(Left$(Label15(i).Caption, Len(Label15(i).Caption) - 1))
        dVorjahrsum = dVorjahrsum + CDbl(Left$(Label8(i).Caption, Len(Label8(i).Caption) - 1))
        dAKTjahrsum = dAKTjahrsum + CDbl(Left$(Label9(i).Caption, Len(Label9(i).Caption) - 1))
        Label15(13).Caption = Format$(dvVorjahrsum, "###,##0.00 €")
        Label8(13).Caption = Format$(dVorjahrsum, "###,##0.00 €")
        Label9(13).Caption = Format$(dAKTjahrsum, "###,##0.00 €")
        
        Label10(13).Caption = "100,00 %"

    Next i
    
    j = j + 2
    txtStatus.Text = j
    
    If dVorjahrsum = 0 Then
        sSQL = "update LiefENT set UEKR = (ueka * 100)/1"
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "update LiefENT set UEKR = (ueka * 100)/" & Val(dVorjahrsum)
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    j = j + 2
    txtStatus.Text = j
    

    
    sSQL = " select * from LIEFENT "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!uekr) Then
                Label10(rsrs!Monat).Caption = Format$((rsrs!uekr / 100), "0.00 %")
                
            Else
                Label10(rsrs!Monat).Caption = "0,00 %"
            End If
            

        rsrs.MoveNext
        Loop
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    j = j + 2
    txtStatus.Text = j
    
    
    picprogress.Visible = False

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "liefstatEWErstellen"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub fuellecombo()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim iMonat As Integer
    Dim iJahr As Integer
    
    iMonat = Month(Now)
    iJahr = Year(Now)
    
    With Combo3
        .Clear
        For i = 1 To 12
        
            If iMonat = 1 Then
                iMonat = 12
                iJahr = iJahr - 1
            Else
                iMonat = iMonat - 1
                iJahr = iJahr
            End If
            
            .AddItem iMonat & "/" & iJahr
            If .Text = "" Then
                .Text = iMonat & "/" & iJahr
            End If
            
        Next i
        
    End With
    With Combo1
    .Clear
    .AddItem "Montag,     gerade KW" & Space(100) & "1"
    .AddItem "Dienstag,   gerade KW" & Space(100) & "2"
    .AddItem "Mittwoch,   gerade KW" & Space(100) & "3"
    .AddItem "Donnerstag, gerade KW" & Space(100) & "4"
    .AddItem "Freitag,    gerade KW" & Space(100) & "5"
    .AddItem "Samstag,    gerade KW" & Space(100) & "6"
    .AddItem "Sonntag,    gerade KW" & Space(100) & "7"
    
    .AddItem "Montag,     ungerade KW" & Space(100) & "8"
    .AddItem "Dienstag,   ungerade KW" & Space(100) & "9"
    .AddItem "Mittwoch,   ungerade KW" & Space(100) & "10"
    .AddItem "Donnerstag, ungerade KW" & Space(100) & "11"
    .AddItem "Freitag,    ungerade KW" & Space(100) & "12"
    .AddItem "Samstag,    ungerade KW" & Space(100) & "13"
    .AddItem "Sonntag,    ungerade KW" & Space(100) & "14"
    .AddItem "kein Bestellrhythmus" & Space(100) & "15"
    
    
    .Text = "Montag,     gerade KW" & Space(100) & "1"
    End With
    With Combo2
    .Clear
    .AddItem "bitte auswählen"
    .AddItem "Montag,     gerade KW" & Space(100) & "1"
    .AddItem "Dienstag,   gerade KW" & Space(100) & "2"
    .AddItem "Mittwoch,   gerade KW" & Space(100) & "3"
    .AddItem "Donnerstag, gerade KW" & Space(100) & "4"
    .AddItem "Freitag,    gerade KW" & Space(100) & "5"
    .AddItem "Samstag,    gerade KW" & Space(100) & "6"
    .AddItem "Sonntag,    gerade KW" & Space(100) & "7"
    
    .AddItem "Montag,     ungerade KW" & Space(100) & "8"
    .AddItem "Dienstag,   ungerade KW" & Space(100) & "9"
    .AddItem "Mittwoch,   ungerade KW" & Space(100) & "10"
    .AddItem "Donnerstag, ungerade KW" & Space(100) & "11"
    .AddItem "Freitag,    ungerade KW" & Space(100) & "12"
    .AddItem "Samstag,    ungerade KW" & Space(100) & "13"
    .AddItem "Sonntag,    ungerade KW" & Space(100) & "14"
    .AddItem "kein Bestellrhythmus" & Space(100) & "15"
    
    
    .Text = "bitte auswählen"
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub


Private Sub Command8_Click()
On Error GoTo LOKAL_ERROR

    If MSFlexGrid1.Row < 1 Then
        Screen.MousePointer = 0
        MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    gckundnr = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
    gckundnr = Trim$(gckundnr)
    
    If IsNumeric(gckundnr) Then

    Else
        Screen.MousePointer = 0
        MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If

    gsARTNR = ""
    frmWKL147.Show 1
    gckundnr = ""
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten. "
        
    Fehlermeldung1
End Sub
Private Sub Command98_Click()
On Error GoTo LOKAL_ERROR
    
    gsZSpalte = "Linr"
    gstab = "BEALIEF"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command98_Click"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
        
    Fehlermeldung1

End Sub
Private Sub FaerbeLINRohneBest(gridx As MSFlexGrid, spaltelinr As Byte)
    On Error GoTo LOKAL_ERROR

    Dim j           As Integer
    Dim i           As Integer
    Dim lfakt       As Long
    Dim sierg       As Single
    Dim sLinr       As String
    
    
    With gridx
        
        For j = 1 To .Rows - 1
            .Row = j
            .Col = spaltelinr
            sLinr = .Text
            If LAGERBestand(sLinr) = 0 Then
                For i = 1 To .Cols - 1
                
                    .Col = i
                    .CellBackColor = vbYellow
                
                Next i
            End If
        Next j

    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "FaerbeLINRohneBest"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FaerbeLINRmitKB(gridx As MSFlexGrid, spaltelinr As Byte)
    On Error GoTo LOKAL_ERROR

    Dim j           As Integer
    Dim i           As Integer
    Dim lfakt       As Long
    Dim sierg       As Single
    Dim sLinr       As String
    
    
    With gridx
        
        For j = 1 To .Rows - 1
            .Row = j
            .Col = spaltelinr
            sLinr = .Text

            If ermoffenKUB(sLinr) Then

                For i = 1 To .Cols - 1
                
                    .Col = i
                    .CellBackColor = vbGreen
                
                Next i
                
                Command8.Visible = True
                Label3(3).Visible = True
                Label4(35).Visible = True
            End If
        Next j
            

    
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "FaerbeLINRmitKB"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FaerbeLINRohneART(gridx As MSFlexGrid, spaltelinr As Byte)
    On Error GoTo LOKAL_ERROR

    Dim j           As Integer
    Dim i           As Integer
    Dim lfakt       As Long
    Dim sierg       As Single
    Dim sLinr       As String
    
    
    With gridx
        
        For j = 1 To .Rows - 1
            .Row = j
            .Col = spaltelinr
            sLinr = .Text
            If ohneArtikel(sLinr) = 0 Then
                For i = 1 To .Cols - 1
                
                    .Col = i
                    .CellBackColor = vbRed
                
                Next i
            End If
        Next j
            

    
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "FaerbeLINRohneArt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeige_Grid()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim j           As Integer
    Dim recAnz      As Recordset
    Dim rsrs        As Recordset
    Dim ctmp        As String
    Dim siFak       As Single
    Dim cArtnr      As String
    Dim iStufe      As Integer
    Dim iRet        As Integer
    
    If Not NewTableSuchenDBKombi("LI" & srechnertab, gdBase) Then
        MsgBox "Keine Lieferanten gefunden!", vbInformation, "Winkiss Hinweis:"
        
        Exit Sub
    End If
    
    Set recAnz = gdBase.OpenRecordset("LI" & srechnertab)
    
    If recAnz.EOF Then
        
        MSFlexGrid1.Visible = False
        MSFlexGrid1.Clear
        Frame1.Visible = False
        Frame0.Visible = True
        MsgBox "Keine Lieferanten gefunden!", vbInformation, "Winkiss Hinweis:"
        
        Exit Sub
    Else
        
    End If
    recAnz.Close: Set recAnz = Nothing
    
    
    Screen.MousePointer = 11

    Tabcheck "BEALIEF"
    
    FormatGridOverTablay "BEALIEF"

    With MSFlexGrid1
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 2
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
        Next j
    
        FuellenMSFlex17
        
        ermittlespalten
        
        .Redraw = False
        
        Tabellenbreiteanpassen MSFlexGrid1, 1.15 * gdTabfak
        
        If Check4.Value = vbChecked Then
            FaerbeLINRohneBest MSFlexGrid1, SpaltennummerLINR
        End If
        
        FaerbeLINRohneART MSFlexGrid1, SpaltennummerLINR
        
        Command8.Visible = False
        Label3(3).Visible = False
        Label4(35).Visible = False
        
        FaerbeLINRmitKB MSFlexGrid1, SpaltennummerLINR
        
        Frame1.Visible = True
        Frame2.Visible = False
        Frame0.Visible = False
        .Visible = True
        .Redraw = True
        .Row = 1
        
        .SetFocus
    
    End With
    
    Me.Refresh
   
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeige_Grid"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
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
Private Sub FuellenMSFlex17()
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
    
'    cOrder = ""
    
    If Option1(0).Value Then
         corder = " order by  Linr "

    ElseIf Option1(1).Value Then
         corder = " order by liefbez "
    ElseIf Option1(2).Value Then
        corder = " order by stadt "
    ElseIf Option1(3).Value Then
        corder = " order by plz"
    ElseIf Option1(4).Value Then
        corder = " order by awert desc "
    ElseIf Option1(5).Value Then
        corder = " order by strasse "
    ElseIf Option1(6).Value Then
        corder = " order by KUERZEL "
    End If
    
    cSQL = "Select * from LI" & srechnertab & corder
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSFlexGrid1
        .Redraw = False
        lrow = 1
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                lrow = lrow + 1
                .Rows = lrow + 1
                .Col = 0
                
                For i = 0 To byAnzahlSpalten - 1
                    .Row = 0
                    .Col = i
                    
                    If sSpaltenname(i) = .Text Then
                        Select Case sSpaltenname(i)
                            Case Is = "Notiz"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = Left(rsrs(sSpaltenbez(i)), 35) & " ..."
                                Else
                                    sWert = ""
                                End If
                                .Row = lrow
                                sWert = SwapStr(sWert, Chr(13), " ")
                                sWert = SwapStr(sWert, Chr(10), " ")
                                .Text = sWert
                                
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
                                
                            Case Is = "LUG", "LAGER(SEK)", "EINKAUF akt Jahr", "EINKAUF vor Jahr", "Penner(SEK)"
                                
            
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "#######0.00")
                                
                            Case Is = "Panteil Stück in %", "Panteil SEK in %"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "#######0.00")
    
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
        .Redraw = True
        .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMSFlex17"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
        
    Fehlermeldung1
    Resume Next
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "LINR"
                SpaltennummerLINR = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    Screen.MousePointer = 11
    
    WKL17Positionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    Command5(20).BackColorFrom = vbWhite
    Command5(20).BackColorTo = vbWhite
    
    giDlgZustand = -1
    
    fuellecombo
    
    If gbGesEKWert_anzeigen Then
        anzeige "normal", "derzeitiger Bestand: " & ermgesbestand & " Artikel", Label1(5)
        anzeige "normal", "Schnitteinkaufswert: " & Format(ermgesSEKwert, "###,##0.00") & " Euro", Label1(53)
    End If
    
    
    Option2(1).Caption = gsPname & " Auswertung"
    
    Label4(40).Caption = DateValue(Now)
    
    anzeige "normal", "", Label1(6)

    If Not SpalteInTabellegefundenNEW("LISRT", "HL", gdBase) Then
        SpalteAnfuegenNEW "LISRT", "HL", "BIT", gdBase
        sSQL = "Update LISRT Set HL = false"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If gbBILDTAST = False Then
        Frame2.Visible = False
    Else
        Frame2.Visible = True
    End If
    
    If gsLinr <> "" Then
        
        Command2_Click 0
        Screen.MousePointer = 0
    Else
        gbNeuerSatz = False
    
        Frame0.Visible = True
        LeereDialogWKL17
        Screen.MousePointer = 0
    End If
    Label4(190).Caption = Label4(190).Caption & " " & Year(Now) & ":"
    Label4(42).Caption = Label4(42).Caption & " " & Year(Now) - 1 & ":"
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub HoleMaxLieferantNr()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lMaxNr As Long
    
    Dim i       As Long
    
    For i = 500000 To 599999
        cSQL = "Select LINR from LISRT where LINR =" & i
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If rsrs.EOF Then
            Exit For
        End If
        rsrs.Close: Set rsrs = Nothing
    Next i
    
    
    Text1(5).Text = Format$(i, "#####0")
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleMaxLieferantNr"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Function SucheLieferantWKL17() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cwhere As String
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim dWert As Double
    Dim iRet As Integer
    Dim lcount As Long
    Dim lLinr As Long
    Dim lAnz As Long
    Dim dMittelwertLUG As Double
    
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
    
    Dim bymonat As Byte
    Dim iJahr As Integer
    
    Dim j As Integer
    Dim iStufe As Integer
    
    SucheLieferantWKL17 = False
    
    anzeige "normal", "Daten werden ermittelt...", Label1(6)
    
    loeschNEW "LI" & srechnertab, gdBase
    CreateTableT2 "LI" & srechnertab, gdBase
    
    iStufe = 1

    Frame1.Visible = False
   
    cSQL = "Insert Into LI" & srechnertab & " Select   "
    cSQL = cSQL & " LINR  "
    cSQL = cSQL & ", LIEFBEZ "
    cSQL = cSQL & ", KUERZEL "
    cSQL = cSQL & ", PLZ "
    cSQL = cSQL & ", STADT "
    cSQL = cSQL & ", STRASSE "
    cSQL = cSQL & ", TEL "
    cSQL = cSQL & ", FAX "
    cSQL = cSQL & ", ZUSATZ "
    cSQL = cSQL & ", KUNDNR  "
    cSQL = cSQL & ", NOTIZ  "
    cSQL = cSQL & ", AWERT "
    cSQL = cSQL & " from LISRT "
    
    cwhere = ""
    cFeld = Text1(0).Text
    cFeld = Trim$(cFeld)
    
    iStufe = 3
    If cFeld <> "" Then
        iStufe = 31
    
        If Not IsNumeric(cFeld) Then
            iStufe = 32
            MsgBox "Korrigieren Sie bitte Ihre letzte Eingabe!", , "Winkiss Hinweis:"
            Text1(0).Text = ""
            Text1(0).SetFocus
            Exit Function
        End If
        iStufe = 34
        If cwhere = "" Then
            iStufe = 35
            cwhere = "where "
        Else
            iStufe = 36
            cwhere = cwhere & "and "
        End If
        iStufe = 37
        cwhere = cwhere & "LINR = " & cFeld & " "
    End If
    
    iStufe = 4
    cFeld = Text1(1).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "KUERZEL like '" & cFeld & "*' "
    End If
    iStufe = 5
    
    cFeld = Text1(2).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "LIEFBEZ like '" & cFeld & "*' "
    End If
    
    iStufe = 6
    cFeld = Text1(3).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "PLZ like '" & cFeld & "*' "
    End If
    
    iStufe = 7
    cFeld = Text1(4).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "STADT like '" & cFeld & "*' "
    End If
    
    If Check3.Value = vbChecked Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & " KUERZEL <> '' "
    End If
    
    cFeld = Right$(Combo2.Text, 2)
    cFeld = Trim$(cFeld)
    If IsNumeric(cFeld) Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "BR = " & cFeld & " "
    End If

    If cwhere = "" Then
        cwhere = "where "
    Else
        cwhere = cwhere & " and "
    End If
    
    cSQL = cSQL & cwhere
'    cSQL = cSQL & " ( SYNSTATUS = 'E' or SYNSTATUS = 'A' )  "
    cSQL = cSQL & " ( SYNSTATUS is null or SYNSTATUS = 'E' or SYNSTATUS = 'A' )  "
    gdBase.Execute cSQL, dbFailOnError
    
    If Check5.Value = vbChecked Then

        anzeige "normal", "mit Detailzahlen: es werden Hintergrunddaten zusammengefasst...", Label1(6)

        If UMS_LINRaktuell = False Then
            ErzeugeLinrUmsatz
        End If

        anzeige "normal", "Lagerwerte werden ermittelt...", Label1(6)

        CheckIndex "ALLARTLU", "linr", "", gdBase

        LagerwerteschreibenLINRJetzt Label1(6)

        cSQL = "Select * from LI" & srechnertab
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveLast
            lAnz = rsrs.RecordCount
            rsrs.MoveFirst
            Do While Not rsrs.EOF

                If Not IsNull(rsrs!linr) Then
                    lLinr = rsrs!linr

                    lAnz = lAnz - 1
                    anzeige "normal", "Lieferant: " & rsrs!LIEFBEZ & " noch " & CStr(lAnz) & " Lieferanten ...", Label1(6)

                    dMittelwertLUG = MittelwertLugaufLINR(lLinr)
                    dEINKaufswert = CDbl(Einkaufsumsatzermittlung(CStr(lLinr), gdBase, CInt(Year(Now))))
                    dEINKaufswertvj = CDbl(Einkaufsumsatzermittlung(CStr(lLinr), gdBase, CInt(Year(Now) - 1)))

                    dUmsBraktJahr = ermgesUmsatzLinr(0, CInt(Year(Now)), lLinr)
                    dUmsBrvorJahr = ermgesUmsatzLinr(0, CInt(Year(Now) - 1), lLinr)

                    dUmsSEKaktJahr = ermgesEKUmsatzLinr(0, CInt(Year(Now)), lLinr)
                    dUmsSEKvorJahr = ermgesEKUmsatzLinr(0, CInt(Year(Now) - 1), lLinr)

                    dUms12M = 0
                    dUms12MVJZR = 0
                    dUmsSEK12M = 0
                    dUmsSEK12MVJZR = 0

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

                        dUms12M = dUms12M + ermgesUmsatzLinr(bymonat, iJahr, lLinr)
                        dUms12MVJZR = dUms12MVJZR + ermgesUmsatzLinr(bymonat, iJahr - 1, lLinr)

                        dUmsSEK12M = dUmsSEK12M + ermgesEKUmsatzLinr(bymonat, iJahr, lLinr)
                        dUmsSEK12MVJZR = dUmsSEK12MVJZR + ermgesEKUmsatzLinr(bymonat, iJahr - 1, lLinr)

                    Next j

                    dLagerwertzumSEK = LAGEREKermittlungJetzt(lLinr)
                    lLagerST = LAGERStückErmittlungJetzt(lLinr)

                    dPennerwertzumSEK = PennerEKermittlungJetzt(lLinr)
                    lPennerST = PennerStückErmittlungJetzt(lLinr)
                Else
                    dMittelwertLUG = 0
                    dEINKaufswert = 0
                    dEINKaufswertvj = 0

                    dUmsBraktJahr = 0
                    dUmsBrvorJahr = 0

                    dUmsSEKaktJahr = 0
                    dUmsSEKvorJahr = 0

                    dLagerwertzumSEK = 0
                    lLagerST = 0

                    dPennerwertzumSEK = 0
                    lPennerST = 0

                    dUms12M = 0
                    dUms12MVJZR = 0
                    dUmsSEK12M = 0
                    dUmsSEK12MVJZR = 0

                End If

                rsrs.Edit
                rsrs!LUG = dMittelwertLUG
                rsrs!LAGERWSEK = dLagerwertzumSEK
                rsrs!LAGERST = lLagerST

                rsrs!PENNERWSEK = dPennerwertzumSEK
                rsrs!PENNERST = lPennerST

                rsrs!EKaktJahr = dEINKaufswert
                rsrs!EKvorJahr = dEINKaufswertvj

                rsrs!UmsBraktJahr = dUmsBraktJahr
                rsrs!UmsBrvorJahr = dUmsBrvorJahr

                rsrs!UmsSEKaktJahr = dUmsSEKaktJahr
                rsrs!UmsSEKvorJahr = dUmsSEKvorJahr

                rsrs!UmsBrakt12M = dUms12M
                rsrs!UmsBrvor12M = dUms12MVJZR

                dUms12MDIFFabs = 0
                dUms12MDIFFabs = dUms12M - dUms12MVJZR

                dUms12MDIFFrela = 0
                If dUms12M <> 0 Then
                    dUms12MDIFFrela = 100 * dUms12MDIFFabs / dUms12M
                End If

                rsrs!UmsSEKakt12 = dUmsSEK12M
                rsrs!UmsSEKvor12 = dUmsSEK12MVJZR

                dUmsSEK12MDIFFabs = 0
                dUmsSEK12MDIFFabs = dUmsSEK12M - dUmsSEK12MVJZR

                dUmsSEK12MDIFFrela = 0
                If dUmsSEK12M <> 0 Then
                    dUmsSEK12MDIFFrela = 100 * dUmsSEK12MDIFFabs / dUmsSEK12M
                End If

                rsrs!UmsBr12MDIFFabs = dUms12MDIFFabs
                rsrs!UmsSEK12MDIFFabs = dUmsSEK12MDIFFabs

                rsrs!UmsBr12MDIFFrela = dUms12MDIFFrela
                rsrs!UmsSEK12MDIFFrela = dUmsSEK12MDIFFrela

                dPennerAnteilSEK = 0
                If dLagerwertzumSEK <> 0 Then
                    dPennerAnteilSEK = 100 * dPennerwertzumSEK / dLagerwertzumSEK
                End If

                dPennerAnteilST = 0
                If lLagerST <> 0 Then
                    dPennerAnteilST = 100 * lPennerST / lLagerST
                End If

                rsrs!PENANTEILST = dPennerAnteilST
                rsrs!PENANTEILSEK = dPennerAnteilSEK

                rsrs.Update

            rsrs.MoveNext
            Loop
        End If
        rsrs.Close
    End If
    
    anzeige "normal", "", Label1(6)
    
    SucheLieferantWKL17 = True

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheLieferantWKL17"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten." & iStufe
    
    Fehlermeldung1
End Function
Private Function fnPruefeEingabeWKL17()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    fnPruefeEingabeWKL17 = 1
    
     
    
    For lcount = 0 To 4
        If Trim$(Text1(lcount).Text) <> "" Then
            fnPruefeEingabeWKL17 = 0
            Exit Function
        End If
    Next lcount
            
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeWKL17"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function ermAnzArt(cLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermAnzArt = 0
    sSQL = "select count(artnr) as maxi from artikel where linr = " & cLinr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermAnzArt = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
         
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermAnzArt"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function ermsumbest(cLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermsumbest = 0
    sSQL = "select sum(bestand) as maxi from artikel where linr = " & cLinr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermsumbest = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
         
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermsumbest"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub Drucklief()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    loeschNEW "LIEFPRINT", gdBase
    CreateTable "LIEFPRINT", gdBase
    
    Screen.MousePointer = 11
    
    sSQL = "Insert into LIEFPRINT select LINR  "
    sSQL = sSQL & ", LIEFBEZ "
    sSQL = sSQL & ", AWERT  "
    sSQL = sSQL & ", ZIELEK "
    sSQL = sSQL & ", FAX "
    sSQL = sSQL & ", KTEXT "
    sSQL = sSQL & ", KUERZEL "
    sSQL = sSQL & ", NOTIZ "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", STADT "
    sSQL = sSQL & ", STRASSE "
    sSQL = sSQL & ", TEL from LISRT "
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("LIEFPRINT")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!linr) Then
                rsrs.Edit
                rsrs!Artikelanz = ermAnzArt(rsrs!linr)
                rsrs!Bestandsum = ermsumbest(rsrs!linr)
                rsrs.Update
            
            End If
        
        rsrs.MoveNext
        Loop
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    loeschNEW "Lieftemp", gdBase
    sSQL = "select * into Lieftemp from  LIEFPRINT"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "LIEFPRINT", gdBase
    CreateTable "LIEFPRINT", gdBase
    sSQL = "Insert into LIEFPRINT select * from Lieftemp "
    
    If Option1(0).Value Then
         sSQL = sSQL & " order by  Linr "
    ElseIf Option1(1).Value Then
         sSQL = sSQL & " order by liefbez "
    ElseIf Option1(2).Value Then
        sSQL = sSQL & " order by stadt "
    ElseIf Option1(3).Value Then
        sSQL = sSQL & " order by plz"
    ElseIf Option1(4).Value Then
        sSQL = sSQL & " order by awert desc "
    ElseIf Option1(5).Value Then
        sSQL = sSQL & " order by strasse "
    ElseIf Option1(6).Value Then
        sSQL = sSQL & " order by KUERZEL "
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    
    reportbildschirm "", "aWKL17a"
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Drucklief"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LeereDialogWKL17()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    
    Label3(2).Caption = "0"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKL17"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "LI" & srechnertab, gdBase
    loeschNEW "LIVM" & srechnertab, gdBase
    loeschNEW "LIVMPRINT", gdBase
    loeschNEW "Lieftemp", gdBase
    loeschNEW "LIEFPRINT", gdBase
    
    LogtoEnd Me
    If gbFrmComeFrom Then
        gfrmComeFrom.Show
        gbFrmComeFrom = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKL17"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    
    Select Case Index
        Case Is = 0
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
        Case Is = 1
        
            cValid = gcUPPER & gcLower & Chr$(8) & "+äÄÜüÖöß" 'kuerzel
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case Is = 2
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case Is = 3
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8)   'plz
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case Is = 4
        
            cValid = gcUPPER & gcLower & Chr$(8) 'ort
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case Is = 22
            cValid = "1234567890,-" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case Is = 26, 27
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case Is = 6
            cZeichen = UCase$(cZeichen)
            KeyAscii = Asc(cZeichen)
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub MSFlexGrid1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    If MSFlexGrid1.Row > 1 Then
        Command2_Click 0
    Else
        sortierenGrid MSFlexGrid1
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case KeyCode
        Case Is = vbKeyReturn
            Command2_Click 0
            
        Case Is = vbKeyF3
            If Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)) > 0 Then
                gcSuch = "LINR" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerLINR)
                frmWKL70.Show 1
                Me.Refresh
                gcSuch = ""
            End If
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    For lcount = 0 To 16
        Text1(lcount).BackColor = vbWhite
    Next lcount
    Text1(26).BackColor = vbWhite
    Text1(27).BackColor = vbWhite
    
    
    
    If Index >= 5 Then
        Label3(1).Caption = Label4(Index - 5).Caption
        Label3(2).Caption = Trim$(Str$(Index))
    Else
        Label3(1).Caption = Label1(Index).Caption
        Label3(2).Caption = Trim$(Str$(Index))
    End If
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = Len(Text1(Index).Text)
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    
    If Index = 0 Then
        If KeyCode = vbKeyF2 Then
            gF2Prompt.cFeld = ""
            gF2Prompt.cWert = ""
            gF2Prompt.cWert2 = ""
            gF2Prompt.cWahl = ""
            gF2Prompt.bMultiple = False
            
            
            gF2Prompt.cFeld = "LINR"
            
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
            End If
            
            If gF2Prompt.cWahl <> "" Then
                Text1(0).Text = gF2Prompt.cWahl
                
            End If
             Text1(0).SetFocus
        
        End If
    End If
    
    If Index <> 15 Then
        If KeyCode = 13 Then
            If Index < 5 Then
                Command1_Click 0
            Else
                Command5_Click 0
            End If
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text4_Change(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim dZielsum    As Double
    Dim i           As Integer
    
    If Len(Text4(0).Text) > 3 Then
        dZielsum = CDbl(Text4(0).Text)
        For i = 1 To 12

            
            Label12(i).Caption = (dZielsum * CDbl(Left$(Label10(i).Caption, Len(Label10(i).Caption) - 1))) / 100
            Label12(i).Caption = Format$(Label12(i).Caption, "#####0.00 €")
            Label12(i).Refresh
            
        Next i

        
        For i = 1 To 12
            If CDbl(Left$(Label9(i).Caption, Len(Label9(i).Caption) - 1)) < CDbl(Left$(Label12(i).Caption, Len(Label12(i).Caption) - 1)) Then
                Label14(i).Caption = "nicht erreicht"
                Label9(i).ForeColor = vbRed
            Else
                Label14(i).Caption = "erreicht"
                Label9(i).ForeColor = vbGreen
            End If
        Next i
        
        Label12(13).Caption = Format$(dZielsum, "#####0.00 €")
        
    Else
        For i = 1 To 12
        
            Label9(i).ForeColor = glS1

            Label12(i).Caption = ""
            Label12(i).Refresh
            
            Label14(i).Caption = ""
            Label14(i).Refresh
            
            Label12(13).Caption = ""
            
        Next i
    
        
    End If
    
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_Change"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    
    Select Case Index
        Case Is = 0
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
    Fehler.gsFunktion = "Text4_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Lieferanten bearbeiten ist ein Fehler aufgetreten."
    
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
