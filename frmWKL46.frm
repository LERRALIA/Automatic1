VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmWKL46 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Inventur"
   ClientHeight    =   8625
   ClientLeft      =   2085
   ClientTop       =   1905
   ClientWidth     =   11910
   FillColor       =   &H00C0C000&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWKL46.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Tag             =   "Inventur"
   Begin VB.Frame Frame23 
      Caption         =   "Frame23"
      Height          =   735
      Left            =   10920
      TabIndex        =   239
      Top             =   8040
      Width           =   975
      Begin VB.TextBox txtAgn 
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
         Left            =   4320
         MaxLength       =   5
         TabIndex        =   256
         Top             =   1680
         Width           =   975
      End
      Begin VB.ListBox List9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5580
         Left            =   120
         TabIndex        =   244
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Frame Frame24 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame13"
         Height          =   975
         Left            =   120
         TabIndex        =   240
         Top             =   3840
         Width           =   3135
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   26
         Left            =   9600
         TabIndex        =   241
         Top             =   6480
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
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   27
         Left            =   9480
         TabIndex        =   242
         Top             =   1080
         Width           =   2295
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command17 
         Height          =   315
         Left            =   5880
         TabIndex        =   257
         Top             =   1200
         Width           =   450
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
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0FF&
         Caption         =   "sgdag"
         Height          =   300
         Left            =   4320
         TabIndex        =   260
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Filter"
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
         Index           =   24
         Left            =   4320
         TabIndex        =   259
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikelgruppe"
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
         Left            =   4320
         TabIndex        =   258
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Wählen Sie ein Datum aus!"
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
         Index           =   19
         Left            =   120
         TabIndex        =   245
         Top             =   960
         Width           =   4335
      End
      Begin VB.Label Label7 
         Caption         =   "verfügbare Bestandshistorien"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   31
         Left            =   120
         TabIndex        =   243
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame21 
      Caption         =   "Frame19"
      Height          =   975
      Left            =   11280
      TabIndex        =   230
      Top             =   7680
      Width           =   2175
      Begin VB.Frame Frame22 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame13"
         Height          =   975
         Left            =   120
         TabIndex        =   233
         Top             =   3840
         Width           =   3135
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   21
         Left            =   9600
         TabIndex        =   232
         Top             =   6480
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
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   25
         Left            =   9480
         TabIndex        =   231
         Top             =   1080
         Width           =   2295
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
         Caption         =   "Export"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label7 
         Caption         =   "Export einer Artikeldatei"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   30
         Left            =   120
         TabIndex        =   234
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame19 
      Caption         =   "Frame19"
      Height          =   495
      Left            =   480
      TabIndex        =   221
      Top             =   4200
      Width           =   975
      Begin sevCommand3.Command Command16 
         Height          =   495
         Left            =   9480
         TabIndex        =   235
         Top             =   1680
         Width           =   2295
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
         Caption         =   "Abgleich"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command15 
         Height          =   495
         Left            =   9480
         TabIndex        =   225
         Top             =   1080
         Width           =   2295
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
         Caption         =   "Abgleich"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   13
         Left            =   9600
         TabIndex        =   223
         Top             =   6480
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
      Begin VB.Frame Frame20 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame13"
         Height          =   975
         Left            =   120
         TabIndex        =   222
         Top             =   3840
         Width           =   3135
      End
      Begin sevCommand3.Command Command18 
         Height          =   495
         Left            =   9480
         TabIndex        =   264
         Top             =   2280
         Width           =   2295
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
         Caption         =   "Abgleich"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "Übereinstimmung wird hier über die EAN vorgenommen, somit sind die Angaben Artnr und Preis unrelevant."
         Height          =   495
         Index           =   7
         Left            =   4560
         TabIndex        =   267
         Top             =   2760
         Width           =   4815
      End
      Begin VB.Label Label16 
         Caption         =   "Spalten: EAN;leer;Artnr;Preis;Menge"
         Height          =   255
         Left            =   4560
         TabIndex        =   266
         Top             =   2520
         Width           =   4815
      End
      Begin VB.Label Label4 
         Caption         =   "csv-Datei, Spalten durch Semikolon getrennt"
         Height          =   255
         Index           =   6
         Left            =   4560
         TabIndex        =   265
         Top             =   2280
         Width           =   4815
      End
      Begin VB.Label Label1 
         Caption         =   "Importdatei der Firma Cosys"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   240
         TabIndex        =   263
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Importdatei der Firma Gresch / PW"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   240
         TabIndex        =   236
         Top             =   1680
         Width           =   5895
      End
      Begin VB.Label Label26 
         Caption         =   "Spalten: ArtnrKiss, Bestand"
         Height          =   255
         Left            =   4560
         TabIndex        =   228
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "Registername: Bestand"
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   227
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Bestandsveränderung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   240
         TabIndex        =   226
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label7 
         Caption         =   "Import einer Bestandsdatei"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   29
         Left            =   120
         TabIndex        =   224
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   720
      Top             =   0
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Frame15"
      Height          =   1215
      Left            =   240
      TabIndex        =   104
      Top             =   5760
      Width           =   735
      Begin sevCommand3.Command Command10 
         Height          =   310
         Left            =   9600
         TabIndex        =   195
         Top             =   1200
         Visible         =   0   'False
         Width           =   450
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
         Caption         =   "F2"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.ListBox List8 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   8520
         Sorted          =   -1  'True
         TabIndex        =   193
         Top             =   1560
         Width           =   3255
      End
      Begin sevCommand3.Command Command11 
         Height          =   255
         Left            =   8520
         TabIndex        =   192
         Top             =   3600
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
         Caption         =   "Leeren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text6 
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
         Left            =   4800
         MaxLength       =   6
         TabIndex        =   190
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text5 
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
         Left            =   6360
         MaxLength       =   6
         TabIndex        =   189
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin sevCommand3.Command Command9 
         Height          =   310
         Left            =   7920
         TabIndex        =   183
         Top             =   2040
         Visible         =   0   'False
         Width           =   450
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
         Caption         =   "F2"
         PictureAlign    =   2
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
         Height          =   315
         Left            =   6600
         MaxLength       =   6
         TabIndex        =   182
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   19
         Left            =   9600
         TabIndex        =   117
         Top             =   4920
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
         Caption         =   "rückgängig"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   18
         Left            =   9600
         TabIndex        =   116
         Top             =   4200
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
         Caption         =   "Übernahme"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   405
         Index           =   17
         Left            =   9600
         TabIndex        =   113
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
         Caption         =   "rückgängig"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame16 
         BorderStyle     =   0  'Kein
         Height          =   1815
         Left            =   240
         TabIndex        =   109
         Top             =   1320
         Width           =   6135
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Rechts ausgerichtet
            Caption         =   "die Bestände der Artikel in einem bestimmten Lagerplatzbereich"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   2
            Left            =   0
            TabIndex        =   191
            Top             =   960
            Width           =   6135
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Rechts ausgerichtet
            Caption         =   "die Bestände der Artikel eines bestimmten Lieferanten"
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
            Left            =   0
            TabIndex        =   111
            Top             =   600
            Width           =   6135
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Rechts ausgerichtet
            Caption         =   "alle Bestände aller Artikel"
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
            Left            =   0
            TabIndex        =   110
            Top             =   0
            Width           =   6135
         End
      End
      Begin sevCommand3.Command Command2 
         Height          =   405
         Index           =   16
         Left            =   9600
         TabIndex        =   108
         Top             =   240
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
         Caption         =   "Ausführen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   15
         Left            =   9600
         TabIndex        =   105
         Top             =   6480
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
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00404000&
         Height          =   255
         Index           =   2
         Left            =   8520
         TabIndex        =   194
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   3720
         TabIndex        =   188
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   2640
         TabIndex        =   187
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "bis"
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
         Index           =   1
         Left            =   6360
         TabIndex        =   186
         Top             =   3120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   4800
         TabIndex        =   185
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferantennumer"
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
         Index           =   0
         Left            =   6600
         TabIndex        =   184
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label7 
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
         Index           =   19
         Left            =   6720
         TabIndex        =   118
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Datenübernahme"
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
         Index           =   17
         Left            =   2040
         TabIndex        =   115
         Top             =   4200
         Width           =   6615
      End
      Begin VB.Label Label7 
         Caption         =   "Schritt 3"
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
         Index           =   16
         Left            =   120
         TabIndex        =   114
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   2
         X1              =   120
         X2              =   11760
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label7 
         Caption         =   "Artikelbestand auf  ""0"" setzen"
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
         Index           =   15
         Left            =   120
         TabIndex        =   112
         Top             =   840
         Width           =   5775
      End
      Begin VB.Label Label7 
         Caption         =   "Bestände auf ""0"" setzen"
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
         Index           =   14
         Left            =   2040
         TabIndex        =   107
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label Label7 
         Caption         =   "Schritt 2"
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
         Index           =   13
         Left            =   120
         TabIndex        =   106
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Frame12"
      Height          =   975
      Left            =   840
      TabIndex        =   74
      Top             =   7680
      Width           =   11055
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   9600
         MaxLength       =   5
         TabIndex        =   196
         Text            =   "Text7"
         Top             =   5280
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
         Index           =   6
         Left            =   6840
         MaxLength       =   13
         TabIndex        =   178
         Top             =   480
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
         Height          =   375
         Index           =   0
         Left            =   8880
         MaxLength       =   3
         TabIndex        =   177
         Top             =   480
         Width           =   1455
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   5
         Left            =   11160
         TabIndex        =   176
         Top             =   2160
         Width           =   375
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
         Caption         =   "c"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List7 
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
         Left            =   10080
         MultiSelect     =   2  'Erweitert
         TabIndex        =   175
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0C000&
         Caption         =   "Liste speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10560
         TabIndex        =   157
         Top             =   5040
         Value           =   1  'Aktiviert
         Width           =   1215
      End
      Begin VB.Frame Frame17 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame17"
         Height          =   1815
         Left            =   9480
         TabIndex        =   150
         Top             =   2640
         Width           =   2175
         Begin VB.OptionButton Option4 
            Caption         =   "Lagerplatz"
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   181
            Tag             =   "lagerp"
            Top             =   1320
            Width           =   1815
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Artikelnummer"
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   155
            Tag             =   "artnr"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Artikelgruppe"
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   153
            Tag             =   "agn"
            Top             =   840
            Width           =   1815
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Bezeichnung"
            Height          =   330
            Index           =   1
            Left            =   120
            TabIndex        =   152
            Tag             =   "bezeich"
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Lieferant, Linie"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   151
            Tag             =   "linr,lpz,bezeich"
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "sortiert nach"
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
            Left            =   0
            TabIndex        =   154
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0C000&
         Caption         =   "nicht geführte Artikel ausblenden"
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
         Left            =   360
         TabIndex        =   149
         Top             =   2640
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
         Height          =   375
         Index           =   4
         Left            =   6960
         MaxLength       =   13
         TabIndex        =   137
         Text            =   "1234567890123"
         Top             =   1440
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
         Height          =   375
         Index           =   3
         Left            =   5760
         MaxLength       =   3
         TabIndex        =   135
         Text            =   "1234567890123"
         Top             =   1440
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
         Height          =   375
         Index           =   1
         Left            =   3720
         MaxLength       =   13
         TabIndex        =   133
         Text            =   "1234567890123"
         Top             =   1440
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
         Height          =   375
         Index           =   2
         Left            =   360
         MaxLength       =   35
         TabIndex        =   132
         Text            =   "JOOP SCHLAGMICHTOT45"
         Top             =   1440
         Width           =   3375
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   4
         Left            =   9600
         TabIndex        =   143
         Top             =   4560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         Caption         =   "Entfernen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command6 
         Height          =   375
         Index           =   3
         Left            =   9600
         TabIndex        =   144
         Top             =   5640
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
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
         Height          =   3840
         Left            =   360
         MultiSelect     =   2  'Erweitert
         TabIndex        =   147
         Top             =   3000
         Visible         =   0   'False
         Width           =   9015
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Index           =   0
         Left            =   9600
         TabIndex        =   142
         Top             =   960
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
         Caption         =   "Suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown-Liste
         TabIndex        =   141
         Top             =   2280
         Width           =   3495
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   139
         Top             =   2280
         Width           =   6015
      End
      Begin sevCommand3.Command Command6 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   131
         Top             =   1920
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
      Begin sevCommand3.Command Command6 
         Height          =   285
         Index           =   1
         Left            =   5760
         TabIndex        =   130
         Top             =   1080
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
      Begin sevCommand3.Command Command2 
         Height          =   405
         Index           =   10
         Left            =   9600
         TabIndex        =   146
         Top             =   6600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   714
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
      Begin VB.CheckBox Check5 
         Caption         =   "mit Strichcode drucken"
         Height          =   255
         Left            =   9600
         TabIndex        =   166
         Top             =   6000
         Width           =   2175
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
         Index           =   14
         Left            =   6840
         TabIndex        =   180
         Top             =   120
         Width           =   1935
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
         Index           =   12
         Left            =   8880
         TabIndex        =   179
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Dateiname"
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
         Left            =   9600
         TabIndex        =   156
         Top             =   5040
         Width           =   975
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
         Index           =   9
         Left            =   6960
         TabIndex        =   148
         Top             =   1080
         Width           =   1695
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
         Index           =   8
         Left            =   6360
         TabIndex        =   145
         Top             =   1080
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
         Index           =   7
         Left            =   3720
         TabIndex        =   140
         Top             =   1080
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
         Index           =   1
         Left            =   360
         TabIndex        =   138
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
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
         Left            =   6360
         TabIndex        =   136
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferant"
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
         Left            =   840
         TabIndex        =   134
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Inventurliste erzeugen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Frame11"
      Height          =   735
      Left            =   960
      TabIndex        =   71
      Top             =   6240
      Width           =   3135
      Begin VB.CheckBox Check11 
         Caption         =   "mit Strichcode drucken"
         Height          =   255
         Left            =   9600
         TabIndex        =   276
         Top             =   4320
         Width           =   2175
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   7
         Left            =   9600
         TabIndex        =   163
         Top             =   3240
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5895
         Left            =   120
         TabIndex        =   162
         Top             =   1080
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   10398
         _Version        =   393216
         ForeColorSel    =   8454143
         FocusRect       =   0
         HighLight       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   6
         Left            =   9600
         TabIndex        =   160
         Top             =   2640
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
         Caption         =   "Laden"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   375
         Index           =   5
         Left            =   8400
         TabIndex        =   159
         Top             =   240
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List6 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   9600
         MultiSelect     =   2  'Erweitert
         TabIndex        =   158
         Top             =   240
         Width           =   2175
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   9
         Left            =   9600
         TabIndex        =   72
         Top             =   6480
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
         Height          =   375
         Index           =   6
         Left            =   9600
         TabIndex        =   275
         Top             =   3840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         Caption         =   "Drucken"
         Enabled         =   0   'False
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Dateiname:"
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
         Left            =   2760
         TabIndex        =   165
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label34 
         BackColor       =   &H00C0C000&
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
         Left            =   4320
         TabIndex        =   164
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   13
         Left            =   1320
         TabIndex        =   161
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Inventur mit vorhandener Inventurliste"
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
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Frame10"
      Height          =   375
      Left            =   7680
      TabIndex        =   68
      Top             =   7920
      Width           =   1215
      Begin VB.CheckBox Check10 
         Caption         =   "Original-EAN"
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
         Left            =   4920
         TabIndex        =   255
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox Text87 
         Height          =   285
         Left            =   6120
         MaxLength       =   3
         TabIndex        =   199
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text85 
         Height          =   285
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   198
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox Check8 
         Caption         =   "mit Druckvorschau"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4920
         TabIndex        =   197
         Top             =   1800
         Width           =   2175
      End
      Begin VB.PictureBox picprogress 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   4635
         TabIndex        =   174
         Top             =   1440
         Width           =   4695
      End
      Begin VB.TextBox txtStatus 
         Height          =   315
         Left            =   5760
         TabIndex        =   173
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin sevCommand3.Command Command2 
         Height          =   285
         Index           =   23
         Left            =   4920
         TabIndex        =   128
         Top             =   1200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
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
         Caption         =   "Einlesen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   315
         Index           =   22
         Left            =   120
         TabIndex        =   122
         Top             =   6240
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3000
         Left            =   120
         TabIndex        =   121
         Top             =   3120
         Width           =   6975
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   20
         Left            =   9600
         TabIndex        =   119
         Top             =   1200
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
         Caption         =   "weiter"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   8
         Left            =   9600
         TabIndex        =   69
         Top             =   6480
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
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0FF&
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
         Height          =   735
         Left            =   120
         TabIndex        =   261
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label Label7 
         Caption         =   $"frmWKL46.frx":0442
         Height          =   735
         Index           =   18
         Left            =   1680
         TabIndex        =   203
         Top             =   6240
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "BedNr."
         Height          =   255
         Index           =   16
         Left            =   6120
         TabIndex        =   201
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Regal"
         Height          =   255
         Index           =   15
         Left            =   4920
         TabIndex        =   200
         Top             =   600
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   915
         Left            =   6240
         MouseIcon       =   "frmWKL46.frx":04EF
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL46.frx":07F9
         Top             =   2160
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   3
         X1              =   7320
         X2              =   7320
         Y1              =   6960
         Y2              =   360
      End
      Begin VB.Label Label7 
         Caption         =   "Ihr Inventurvorgang enthält:"
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
         Index           =   25
         Left            =   120
         TabIndex        =   127
         Top             =   1800
         Width           =   5055
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "verschiedene Artikel:"
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
         Index           =   24
         Left            =   240
         TabIndex        =   126
         Top             =   2280
         Width           =   3495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "mit einem Gesamtbestand:"
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
         Index           =   23
         Left            =   120
         TabIndex        =   125
         Top             =   2640
         Width           =   3615
      End
      Begin VB.Label Label7 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   3840
         TabIndex        =   124
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   3840
         TabIndex        =   123
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "Bestandsaufnahme beendet? Dann weiter zu Schritt 2 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   20
         Left            =   8040
         TabIndex        =   120
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label7 
         Caption         =   "Inventur mit dem MDE - Gerät"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   2535
      Left            =   2040
      TabIndex        =   65
      Top             =   1320
      Width           =   6855
      Begin VB.Frame Frame18 
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
         Left            =   0
         TabIndex        =   207
         Top             =   1800
         Width           =   7215
         Begin sevCommand3.Command Command14 
            Height          =   645
            Index           =   31
            Left            =   6440
            TabIndex        =   218
            Top             =   120
            Width           =   645
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
         Begin sevCommand3.Command Command14 
            Height          =   645
            Index           =   20
            Left            =   120
            TabIndex        =   208
            Top             =   120
            Width           =   600
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
         Begin sevCommand3.Command Command14 
            Height          =   645
            Index           =   21
            Left            =   720
            TabIndex        =   209
            Top             =   120
            Width           =   600
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
         Begin sevCommand3.Command Command14 
            Height          =   645
            Index           =   22
            Left            =   1320
            TabIndex        =   210
            Top             =   120
            Width           =   600
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
         Begin sevCommand3.Command Command14 
            Height          =   645
            Index           =   23
            Left            =   1920
            TabIndex        =   211
            Top             =   120
            Width           =   600
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
         Begin sevCommand3.Command Command14 
            Height          =   645
            Index           =   24
            Left            =   2520
            TabIndex        =   212
            Top             =   120
            Width           =   600
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
         Begin sevCommand3.Command Command14 
            Height          =   645
            Index           =   25
            Left            =   3120
            TabIndex        =   213
            Top             =   120
            Width           =   600
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
         Begin sevCommand3.Command Command14 
            Height          =   645
            Index           =   26
            Left            =   3720
            TabIndex        =   214
            Top             =   120
            Width           =   600
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
         Begin sevCommand3.Command Command14 
            Height          =   645
            Index           =   27
            Left            =   4320
            TabIndex        =   215
            Top             =   120
            Width           =   600
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
         Begin sevCommand3.Command Command14 
            Height          =   645
            Index           =   28
            Left            =   4920
            TabIndex        =   216
            Top             =   120
            Width           =   600
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
         Begin sevCommand3.Command Command14 
            Height          =   645
            Index           =   29
            Left            =   5520
            TabIndex        =   217
            Top             =   120
            Width           =   600
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
         Begin VB.Label Label11 
            Caption         =   "Label11"
            Height          =   135
            Left            =   6120
            TabIndex        =   219
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.CheckBox Check9 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "nach dem Scannen Menge = 1"
         Height          =   375
         Left            =   3960
         TabIndex        =   205
         Top             =   600
         Value           =   1  'Aktiviert
         Width           =   3135
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Artikel aller Lieferanten aufnehmen"
         Height          =   375
         Left            =   3960
         TabIndex        =   129
         Top             =   240
         Value           =   1  'Aktiviert
         Width           =   3135
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   14
         Left            =   9600
         TabIndex        =   103
         Top             =   1200
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
         Caption         =   "weiter"
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
         Height          =   3000
         Left            =   120
         TabIndex        =   101
         Top             =   3600
         Width           =   6975
      End
      Begin sevCommand3.Command Command2 
         Height          =   315
         Index           =   12
         Left            =   5640
         TabIndex        =   100
         Top             =   3240
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   11
         Left            =   4920
         TabIndex        =   94
         Top             =   1200
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   120
         MaxLength       =   4
         TabIndex        =   93
         Text            =   "1"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         Left            =   1440
         MaxLength       =   13
         TabIndex        =   92
         Top             =   1200
         Width           =   3375
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   7
         Left            =   9600
         TabIndex        =   66
         Top             =   6480
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
      Begin sevCommand3.Command Command8 
         Height          =   230
         Left            =   1020
         TabIndex        =   253
         Top             =   1500
         Width           =   375
         _ExtentX        =   661
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
         Height          =   230
         Left            =   1020
         TabIndex        =   254
         Top             =   1200
         Width           =   375
         _ExtentX        =   661
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
      Begin VB.Label lblLeArtikel 
         Caption         =   "letzter gescannter Artikel:"
         Height          =   735
         Left            =   120
         TabIndex        =   262
         Top             =   6720
         Width           =   6975
      End
      Begin VB.Label Label7 
         Caption         =   $"frmWKL46.frx":0DDC
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   28
         Left            =   8520
         TabIndex        =   204
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "Bestandsaufnahme beendet? Dann weiter zu Schritt 2 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   12
         Left            =   8040
         TabIndex        =   102
         Top             =   240
         Width           =   3735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   7320
         X2              =   7320
         Y1              =   6840
         Y2              =   240
      End
      Begin VB.Label Label7 
         Caption         =   "0"
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
         Index           =   11
         Left            =   3960
         TabIndex        =   99
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "0"
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
         Index           =   10
         Left            =   3960
         TabIndex        =   98
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "mit einem Gesamtbestand:"
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
         Index           =   9
         Left            =   0
         TabIndex        =   97
         Top             =   3240
         Width           =   3735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "verschiedene Artikel:"
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
         Index           =   8
         Left            =   120
         TabIndex        =   96
         Top             =   3000
         Width           =   3615
      End
      Begin VB.Label Label7 
         Caption         =   "Ihr Inventurvorgang enthält:"
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
         Left            =   120
         TabIndex        =   95
         Top             =   2640
         Width           =   4575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "Menge"
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
         Left            =   120
         TabIndex        =   91
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "EAN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   90
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Inventur mit Scanner"
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
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   735
      Left            =   240
      TabIndex        =   62
      Top             =   7080
      Width           =   975
      Begin VB.Frame Frame36 
         BorderStyle     =   0  'Kein
         Height          =   500
         Left            =   1560
         TabIndex        =   271
         Top             =   4560
         Visible         =   0   'False
         Width           =   2385
         Begin VB.OptionButton opt1 
            Caption         =   "größten"
            Height          =   285
            Index           =   22
            Left            =   0
            TabIndex        =   273
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton opt1 
            Caption         =   "kleinsten"
            Height          =   285
            Index           =   23
            Left            =   1200
            TabIndex        =   272
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lbl2 
            Caption         =   "bei mehreren ausgehend vom"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   200
            Index           =   31
            Left            =   120
            TabIndex        =   274
            Top             =   80
            Width           =   1695
         End
      End
      Begin VB.CheckBox Check7 
         Caption         =   "mit Bestand drucken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3960
         TabIndex        =   171
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Frame Frame14 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame13"
         Height          =   975
         Left            =   120
         TabIndex        =   86
         Top             =   3840
         Width           =   3135
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Schnitteinkaufspreis"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   225
            Index           =   7
            Left            =   0
            TabIndex        =   88
            Top             =   120
            Width           =   2775
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Listeneinkaufspreis"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   225
            Index           =   4
            Left            =   0
            TabIndex        =   87
            Top             =   480
            Value           =   -1  'True
            Width           =   2775
         End
      End
      Begin VB.CheckBox check1 
         Caption         =   "nur EK - Zahlen "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   82
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame13"
         Height          =   1695
         Left            =   120
         TabIndex        =   76
         Top             =   840
         Width           =   3735
         Begin VB.OptionButton Option1 
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
            ForeColor       =   &H00404000&
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   81
            Top             =   360
            Value           =   -1  'True
            Width           =   2775
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Artikelbezeichnung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   80
            Top             =   720
            Width           =   2775
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
            ForeColor       =   &H00404000&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   79
            Top             =   1080
            Width           =   2775
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C000&
            Caption         =   "Lieferant, LiefBestNr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   78
            Top             =   1440
            Width           =   2775
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Anzeige sortiert nach:"
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
            Index           =   3
            Left            =   0
            TabIndex        =   85
            Top             =   0
            Width           =   2175
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Anzeige sortiert nach:"
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
            Index           =   0
            Left            =   0
            TabIndex        =   77
            Top             =   0
            Width           =   2175
         End
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   6
         Left            =   9600
         TabIndex        =   63
         Top             =   6480
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventurzählliste:"
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
         Index           =   4
         Left            =   3960
         TabIndex        =   172
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Berechnungsgrundlage des Einkaufswertes:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   84
         Top             =   3600
         Width           =   4335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Druckvorgaben:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   83
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Inventureinstellungen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00808000&
      Caption         =   "Inventur-Datei laden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   8
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   8
         Left            =   1920
         TabIndex        =   206
         Top             =   2520
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
         Caption         =   "Export"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   5295
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   2520
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   3
         Left            =   3720
         TabIndex        =   10
         Top             =   2520
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   2
         Left            =   3360
         TabIndex        =   9
         Top             =   480
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
         Caption         =   "Laden"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventur-Datei laden"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      Caption         =   "Inventur-Datei speichern"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Top             =   7680
      Visible         =   0   'False
      Width           =   855
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   5175
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   1
         Left            =   3240
         TabIndex        =   6
         Top             =   2520
         Width           =   2175
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
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   0
         Left            =   960
         TabIndex        =   5
         Top             =   2520
         Width           =   2175
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   960
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "Text"
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventur-Datei speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "INV_"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   1455
      Left            =   7800
      TabIndex        =   23
      Top             =   6240
      Width           =   5895
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   1
         Left            =   9600
         TabIndex        =   61
         Top             =   6480
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   11535
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   5055
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Visible         =   0   'False
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   8916
            _Version        =   393216
            Cols            =   13
            ForeColorSel    =   8454143
            FocusRect       =   0
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lbl1 
            BackStyle       =   0  'Transparent
            Caption         =   "F2: Artikelbearbeitung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   268
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "EK"
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
            Left            =   8760
            TabIndex        =   170
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "Verkaufswert"
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
            Left            =   6360
            TabIndex        =   169
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "76543210,00"
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
            Left            =   7560
            TabIndex        =   35
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "76543210,00"
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
            Left            =   10320
            TabIndex        =   34
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   0
            Width           =   6135
         End
      End
      Begin VB.Frame Frame0 
         BackColor       =   &H00C0C000&
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
         Left            =   480
         TabIndex        =   36
         Top             =   4800
         Visible         =   0   'False
         Width           =   8655
         Begin sevCommand3.Command Command0 
            Height          =   525
            Index           =   12
            Left            =   6420
            TabIndex        =   39
            Top             =   120
            Width           =   525
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
            ToolTip         =   "Links"
            ToolTipTitle    =   "Links"
            ButtonStyle     =   2
            Caption         =   ""
            PictureAlign    =   3
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   525
            Index           =   13
            Left            =   6960
            TabIndex        =   37
            Top             =   120
            Width           =   525
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
            ToolTip         =   "Rauf"
            ToolTipTitle    =   "Rauf"
            ButtonStyle     =   2
            Caption         =   ""
            PictureAlign    =   3
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   525
            Index           =   14
            Left            =   7500
            TabIndex        =   52
            Top             =   120
            Width           =   525
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
            ToolTip         =   "Runter"
            ToolTipTitle    =   "Runter"
            ButtonStyle     =   2
            Caption         =   ""
            PictureAlign    =   3
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   525
            Index           =   15
            Left            =   8040
            TabIndex        =   38
            Top             =   120
            Width           =   525
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
            ToolTip         =   "Rechts"
            ToolTipTitle    =   "Rechts"
            ButtonStyle     =   2
            Caption         =   ""
            PictureAlign    =   3
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   525
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   120
            Width           =   525
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
            Height          =   525
            Index           =   1
            Left            =   600
            TabIndex        =   50
            Top             =   120
            Width           =   525
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
            Height          =   525
            Index           =   2
            Left            =   1080
            TabIndex        =   49
            Top             =   120
            Width           =   525
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
            Height          =   525
            Index           =   3
            Left            =   1560
            TabIndex        =   48
            Top             =   120
            Width           =   525
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
            Height          =   525
            Index           =   4
            Left            =   2040
            TabIndex        =   47
            Top             =   120
            Width           =   525
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
            Height          =   525
            Index           =   5
            Left            =   2520
            TabIndex        =   46
            Top             =   120
            Width           =   525
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
            Height          =   525
            Index           =   6
            Left            =   3000
            TabIndex        =   45
            Top             =   120
            Width           =   525
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
            Height          =   525
            Index           =   7
            Left            =   3480
            TabIndex        =   44
            Top             =   120
            Width           =   525
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
            Height          =   525
            Index           =   8
            Left            =   3960
            TabIndex        =   43
            Top             =   120
            Width           =   525
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
            Height          =   525
            Index           =   9
            Left            =   4440
            TabIndex        =   42
            Top             =   120
            Width           =   525
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
            Height          =   525
            Index           =   10
            Left            =   4920
            TabIndex        =   41
            Top             =   120
            Width           =   525
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
            Caption         =   ","
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command0 
            Height          =   525
            Index           =   11
            Left            =   5400
            TabIndex        =   40
            Top             =   120
            Width           =   525
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
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C000&
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
         Left            =   600
         TabIndex        =   26
         Top             =   5640
         Width           =   8775
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
            Left            =   6480
            TabIndex        =   202
            Top             =   240
            Width           =   1095
         End
         Begin sevCommand3.Command Command4 
            Height          =   525
            Left            =   5520
            TabIndex        =   30
            Top             =   180
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   926
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
            Caption         =   "Druck"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   525
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   180
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   926
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
            Caption         =   "Liste speichern"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   525
            Index           =   4
            Left            =   1110
            TabIndex        =   28
            Top             =   180
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   926
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
            Caption         =   "Liste laden"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   525
            Index           =   5
            Left            =   3015
            TabIndex        =   27
            Top             =   180
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   926
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
            Caption         =   "Neu berechnen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command20 
            Height          =   405
            Index           =   20
            Left            =   8160
            TabIndex        =   237
            ToolTipText     =   "Kalender"
            Top             =   240
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   714
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
         Begin sevCommand3.Command Command12 
            Height          =   195
            Left            =   7680
            TabIndex        =   246
            Top             =   460
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   344
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
         Begin sevCommand3.Command Command13 
            Height          =   195
            Left            =   7680
            TabIndex        =   247
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   344
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
         Begin sevCommand3.Command Command2 
            Height          =   525
            Index           =   28
            Left            =   4215
            TabIndex        =   269
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   926
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
            Caption         =   "aktualisieren"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command2 
            Height          =   525
            Index           =   29
            Left            =   2040
            TabIndex        =   270
            Top             =   180
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   926
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
      End
   End
   Begin sevCommand3.Command Command2 
      Height          =   525
      Index           =   2
      Left            =   9600
      TabIndex        =   25
      Top             =   8040
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   926
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
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   1095
      Left            =   600
      TabIndex        =   21
      Top             =   1080
      Width           =   5655
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "verfügbare Bestandshistorien"
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
         Index           =   8
         Left            =   120
         TabIndex        =   238
         Top             =   5040
         Width           =   6975
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Export einer Artikeldatei"
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
         Left            =   120
         TabIndex        =   229
         Top             =   4560
         Width           =   6975
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Import einer Bestandsdatei"
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
         TabIndex        =   220
         Top             =   4080
         Width           =   6975
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Inventureinstellungen vornehmen"
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
         Left            =   120
         TabIndex        =   59
         Top             =   3600
         Width           =   6975
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Inventurlisten erzeugen"
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
         Index           =   4
         Left            =   120
         TabIndex        =   58
         Top             =   3120
         Width           =   6975
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Inventur mit vorhandener Inventurliste"
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
         Index           =   3
         Left            =   120
         TabIndex        =   57
         Top             =   2640
         Width           =   6975
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Inventur mit MDE - Gerät"
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
         Index           =   2
         Left            =   120
         TabIndex        =   56
         Top             =   2160
         Width           =   6975
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Inventur mit Scanner"
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
         Left            =   120
         TabIndex        =   55
         Top             =   1680
         Width           =   6975
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Inventurberechnungen vornehmen"
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
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   1200
         Value           =   -1  'True
         Width           =   6975
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   0
         Left            =   9600
         TabIndex        =   53
         Top             =   6480
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
         Caption         =   "weiter"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Import einer Bestandsdatei"
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
         Index           =   23
         Left            =   8400
         MouseIcon       =   "frmWKL46.frx":0E8B
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   252
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Export einer Artikeldatei"
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
         Index           =   22
         Left            =   8400
         MouseIcon       =   "frmWKL46.frx":1195
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   251
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Inventur mit MDE - Gerät"
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
         Index           =   21
         Left            =   8400
         MouseIcon       =   "frmWKL46.frx":149F
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   250
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Hilfethemen:"
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
         Left            =   8400
         TabIndex        =   249
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Inventur mit Scanner"
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
         Index           =   20
         Left            =   8400
         MouseIcon       =   "frmWKL46.frx":17A9
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   248
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label15 
         Caption         =   "Wie möchten Sie vorgehen?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
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
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CheckBox Check6 
         Caption         =   "mit Abschlag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7440
         TabIndex        =   168
         Top             =   240
         Width           =   1695
      End
      Begin sevCommand3.Command Command2 
         Height          =   315
         Index           =   24
         Left            =   5520
         TabIndex        =   167
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
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
         Caption         =   "Farbe"
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
         Height          =   315
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   89
         Top             =   240
         Width           =   975
      End
      Begin sevCommand3.Command Command5 
         Height          =   315
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   450
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
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   315
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
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
         Caption         =   "Ermitteln"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label0 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
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
         Index           =   2
         Left            =   2760
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferantennummer"
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label0 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
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
         Index           =   1
         Left            =   11160
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label0 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
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
         Index           =   0
         Left            =   9960
         TabIndex        =   1
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Wie möchten Sie vorgehen?"
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
      Left            =   240
      TabIndex        =   60
      Top             =   8160
      Width           =   9135
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventur"
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
      TabIndex        =   22
      Top             =   0
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmWKL46"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aBreite(0 To 8)    As Integer
Dim bnoData As Boolean
Dim ue As Boolean
Dim brueck2 As Boolean
Dim mdeErr As Boolean
Dim iscan As Integer
Dim SpaltennummerMENGE        As Byte
Dim SpaltennummerArtnr        As Byte
Private Sub PositionierenWKL46()
    On Error GoTo LOKAL_ERROR
    
    With Frame0
        .Height = 855
        .Left = 0
        .Top = 5640
        .Width = 11775
        .BorderStyle = 0
    End With
    
    With Frame6
        .Height = 7095
        .Left = 0
        .Top = 840
        .Width = 11895
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame7
        .Height = 7095
        .Left = 0
        .Top = 840
        .Width = 11895
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame19
        .Height = 7095
        .Left = 0
        .Top = 840
        .Width = 11895
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame21
        .Height = 7095
        .Left = 0
        .Top = 840
        .Width = 11895
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame23
        .Height = 7095
        .Left = 0
        .Top = 840
        .Width = 11895
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame8
        .Height = 7095
        .Left = 0
        .Top = 840
        .Width = 11895
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame9
        .Height = 7095
        .Left = 0
        .Top = 840
        .Width = 11895
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame10
        .Height = 7095
        .Left = 0
        .Top = 840
        .Width = 11895
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame11
        .Height = 7095
        .Left = 0
        .Top = 840
        .Width = 11895
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame12
        .Height = 7095
        .Left = 0
        .Top = 840
        .Width = 11895
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame15
        .Height = 7095
        .Left = 0
        .Top = 840
        .Width = 11895
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame1
        .Height = 615
        .Left = 2520
        .Top = 0
        .Width = 9255
        .BorderStyle = 0
        .Visible = False
    End With
    
    Frame2.Height = 5775
    Frame2.Left = 0
    Frame2.Top = 0
    Frame2.Width = 11895
    Frame2.BorderStyle = 0
    
    With Frame3
        .Height = 855
        .Left = 0
        .Top = 6360
        .Width = 8655
        .BorderStyle = 0
    End With
    
    Frame4.Height = 3135
    Frame4.Left = 3240
    Frame4.Top = 2520
    Frame4.Width = 5655
    Frame4.BackColor = glH2
    Frame4.BorderStyle = 0
    Frame4.Visible = False

    Frame5.Height = 3135
    Frame5.Left = 3240
    Frame5.Top = 2520
    Frame5.Width = 5655
    Frame5.BackColor = glH2
    Frame5.BorderStyle = 0
    Frame5.Visible = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL46"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub BerechneInventurWKL46()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow            As Long
    Dim lcol            As Long
    Dim lRows           As Long
    Dim lCols           As Long
    Dim lartnr          As Long
    Dim lBestand        As Long
    Dim lLinr           As Long
    Dim cLiefBez        As String
    Dim ctmp            As String
    Dim cMwst           As String
    Dim cBezeich        As String
    Dim cSQL            As String
    Dim cPfad           As String
    
    Dim dKVkPr1         As Double
    Dim dVkPr           As Double
    Dim dVkPrNetto      As Double
    Dim dEkpr           As Double
    Dim dVkWert         As Double
    Dim dEkWert         As Double
    Dim dStueckSpanne   As Double
    Dim dSpanne         As Double
    Dim dSummeVkPr      As Double
    Dim dSummeEkPr      As Double
    Dim dKVkPr1Netto    As Double
    Dim rsrs            As Recordset

    If MSFlexGrid1.Visible = False Or MSFlexGrid1.Rows < 2 Then
        Exit Sub
    End If
    
    anzeige "normal", "Inventurdaten werden neu berechnet...", Label6
    
    MSFlexGrid1.Redraw = False
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    invLiteRefresh
        
    lRows = MSFlexGrid1.Rows
    lCols = MSFlexGrid1.Cols
    
    cSQL = "Select * from INVLITE "
    If Option1(0).Value = True Then
        cSQL = cSQL & " order by ARTNR"
    ElseIf Option1(1).Value = True Then
        cSQL = cSQL & " order by BEZEICH"
    ElseIf Option1(2).Value = True Then
        cSQL = cSQL & " order by LINR"
    ElseIf Option1(3).Value = True Then
        cSQL = cSQL & " order by LPZ,LINR, BEZEICH"
    End If
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    For lrow = 1 To lRows - 1
        MSFlexGrid1.Row = lrow
        rsrs.AddNew
        
        lcol = 0
        MSFlexGrid1.Col = lcol
        ctmp = MSFlexGrid1.Text
        lartnr = Val(ctmp)
        rsrs!artnr = lartnr
        
        lcol = 1
        MSFlexGrid1.Col = lcol
        ctmp = MSFlexGrid1.Text
        cBezeich = ctmp
        rsrs!BEZEICH = cBezeich
        
        lcol = 2
        MSFlexGrid1.Col = lcol
        ctmp = MSFlexGrid1.Text
        lBestand = Val(ctmp)
        rsrs!BESTAND = lBestand
        
        lcol = 3
        MSFlexGrid1.Col = lcol
        ctmp = MSFlexGrid1.Text
        ctmp = fnMoveComma2Point$(ctmp)
        dKVkPr1 = Val(ctmp)
        rsrs!KVKPR1 = dKVkPr1
        
        lcol = 4
        MSFlexGrid1.Col = lcol
        ctmp = MSFlexGrid1.Text
        ctmp = fnMoveComma2Point$(ctmp)
        dVkPr = Val(ctmp)
        rsrs!vkpr = dVkPr
        
        lcol = 5
        MSFlexGrid1.Col = lcol
        ctmp = MSFlexGrid1.Text
        ctmp = fnMoveComma2Point$(ctmp)
        dEkpr = Val(ctmp)
        rsrs!ekpr = dEkpr
        
        lcol = 6
        MSFlexGrid1.Col = lcol
        dVkWert = lBestand * dKVkPr1 'dVkPr
        ctmp = Format$(dVkWert, "#####0.00")
        MSFlexGrid1.Text = ctmp
        dSummeVkPr = dSummeVkPr + dVkWert
        rsrs!VKWERT = dVkWert
        
        lcol = 7
        MSFlexGrid1.Col = lcol
        dEkWert = lBestand * dEkpr
        ctmp = Format$(dEkWert, "#####0.00")
        MSFlexGrid1.Text = ctmp
        dSummeEkPr = dSummeEkPr + dEkWert
        rsrs!EKWERT = dEkWert
        cMwst = MSFlexGrid1.TextMatrix(lrow, 13)
        
        lcol = 8
        MSFlexGrid1.Col = lcol
        Select Case cMwst
            Case Is = "V"
               dKVkPr1Netto = (dKVkPr1 / (100 + gdMWStV)) * 100
               dVkPrNetto = (dVkPr / (100 + gdMWStV)) * 100
            Case Is = "E"
               dKVkPr1Netto = (dKVkPr1 / (100 + gdMWStE)) * 100
               dVkPrNetto = (dVkPr / (100 + gdMWStE)) * 100
            Case Is = "O"
               dKVkPr1Netto = (dKVkPr1 / 100) * 100
               dVkPrNetto = (dVkPr / 100) * 100
        End Select
        If dKVkPr1Netto <> 0 Then
            dStueckSpanne = ((dKVkPr1Netto - dEkpr) * 100) / dKVkPr1Netto
        Else
            dStueckSpanne = 0
        End If
        ctmp = Format$(dStueckSpanne, "#####0.00")
        MSFlexGrid1.Text = ctmp
        rsrs!STSPANNE = dStueckSpanne
        
        lcol = 9
        MSFlexGrid1.Col = lcol
        dSpanne = (dKVkPr1Netto * lBestand) - (dEkpr * lBestand)
        ctmp = Format$(dSpanne, "#####0.00")
        MSFlexGrid1.Text = ctmp
        rsrs!SPANNE = dSpanne
        
        lcol = 10
        MSFlexGrid1.Col = lcol
        ctmp = MSFlexGrid1.Text
        lLinr = Val(ctmp)
        rsrs!linr = lLinr
        
        lcol = 11
        MSFlexGrid1.Col = lcol
        ctmp = MSFlexGrid1.Text
        cLiefBez = ctmp
        rsrs!LIEFBEZ = cLiefBez
        
        lcol = 12
        MSFlexGrid1.Col = lcol
        ctmp = MSFlexGrid1.Text
        rsrs!LPZ = ctmp
        
        lcol = 13
        MSFlexGrid1.Col = lcol
        ctmp = MSFlexGrid1.Text
        rsrs!MWST = ctmp
        
        rsrs.Update
    Next lrow
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Delete from INVLITE where bestand < 1"
    gdBase.Execute cSQL, dbFailOnError
    
    Label3(0).Caption = VkwertErmittlung
    Label3(0).Visible = True
    Label3(1).Caption = EkwertErmittlung
    Label3(1).Visible = True
    
    
    MSFlexGrid1.Redraw = True
    
    anzeige "normal", "", Label6
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BerechneInventurWKL46"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub FormatiereMSFlexGrid1WKL46()
    On Error GoTo LOKAL_ERROR
    
    MSFlexGrid1.Cols = 16
    MSFlexGrid1.Rows = 1
    MSFlexGrid1.FixedCols = 1
    
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = 0
    MSFlexGrid1.ColWidth(0) = 800
    MSFlexGrid1.Text = "ArtNr."
    MSFlexGrid1.Col = 1
    MSFlexGrid1.ColWidth(1) = 3500
    MSFlexGrid1.Text = "Bezeichnung"
    MSFlexGrid1.Col = 2
    MSFlexGrid1.ColWidth(2) = 800
    MSFlexGrid1.Text = "Bestand"
    MSFlexGrid1.Col = 3
    MSFlexGrid1.ColWidth(3) = 800
    MSFlexGrid1.Text = "KVKPr1"
    MSFlexGrid1.Col = 4
    MSFlexGrid1.ColWidth(4) = 800
    MSFlexGrid1.Text = "VKPR"
    MSFlexGrid1.Col = 5
    MSFlexGrid1.ColWidth(5) = 800
    MSFlexGrid1.Text = "EKPR"
    MSFlexGrid1.Col = 6
    MSFlexGrid1.ColWidth(6) = 1000
    MSFlexGrid1.Text = "VKWert"
    MSFlexGrid1.Col = 7
    MSFlexGrid1.ColWidth(7) = 1000
    MSFlexGrid1.Text = "EKWert"
    MSFlexGrid1.Col = 8
    MSFlexGrid1.ColWidth(8) = 1200
    MSFlexGrid1.Text = "Netto Spanne"
    MSFlexGrid1.Col = 9
    MSFlexGrid1.ColWidth(9) = 1000
    MSFlexGrid1.Text = "Spannewert"
    MSFlexGrid1.Col = 10
    MSFlexGrid1.ColWidth(10) = 800
    MSFlexGrid1.Text = "LiefNr"
    MSFlexGrid1.Col = 11
    MSFlexGrid1.ColWidth(11) = 3500
    MSFlexGrid1.Text = "Lieferantenname"
    MSFlexGrid1.Col = 12
    MSFlexGrid1.ColWidth(12) = 700
    MSFlexGrid1.Text = "Linie"
    MSFlexGrid1.Col = 13
    MSFlexGrid1.ColWidth(13) = 700
    MSFlexGrid1.Text = "MWSt"
    MSFlexGrid1.Col = 14
    MSFlexGrid1.ColWidth(14) = 700
    MSFlexGrid1.Text = "AWM"
    MSFlexGrid1.Col = 15
    MSFlexGrid1.ColWidth(15) = 1400
    MSFlexGrid1.Text = "LiefBestNr"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatiereMSFlexGrid1WKL46"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LadeInventurDateiWKL46()
    On Error GoTo LOKAL_ERROR
    
    Dim lRet        As Long
    Dim lfail       As Long
    Dim lcount      As Long
    Dim cLBSatz     As String
    Dim cdatei      As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim cPfad       As String
    Dim bDuplikat   As Boolean
    Dim tdTd        As TableDef
    Dim lAnzFelder  As Long
    Dim cFeldName   As String
    Dim bfind       As Boolean
    Dim cTab        As String
    Dim iRet        As Integer
    Dim cSQL        As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cdatei = Label1(3).Caption
    cdatei = Trim$(UCase$(cdatei))
    If cdatei = "" Then
        anzeigeNew "rot", "Bitte eine Inventur-Datei angeben!", Label6
        List1.SetFocus
        Exit Sub
    End If
    
    gdBase.TableDefs.Refresh 'Dabarefresh
    
    cTab = cdatei
    Set tdTd = gdBase.TableDefs(cTab)
    
    bfind = False
    
    lAnzFelder = tdTd.Fields.Count
    For lcount = 0 To lAnzFelder - 1
        cFeldName = tdTd.Fields(lcount).name
        If cFeldName = "KVKPR1" Then
            bfind = True
        End If
        
        If UCase(cFeldName) = "FILIALE" Then
            anzeigeNew "rot", "Diese Datei kann nicht geladen werden.", Label6
            Screen.MousePointer = 0
            Exit Sub
        End If
    Next lcount
    
    If bfind = False Then
        iRet = MsgBox("Diese Inventurdatei kann nicht geladen werden. Sie enthält ein älteres Format" & vbCrLf _
                & "Möchten Sie diese jetzt löschen?", vbQuestion + vbYesNo, "Winkiss Hinweis:")
        If iRet = vbYes Then
        
            loesch cdatei
            NewListeFuellAnfangsbuch "INV_", frmWKL46.List1, gdBase
            Label1(3).Caption = ""
            Screen.MousePointer = 0
            Exit Sub
            
        ElseIf iRet = vbNo Then

            NewListeFuellAnfangsbuch "INV_", frmWKL46.List1, gdBase
            Label1(3).Caption = ""
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    
    invLiteRefresh
    INVLITEcopy cTab
    
    haengan
    haengan1
    haengan2
    MoveDaten2DialogWKL46
    

    Command3_Click 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LadeInventurDateiWKL46"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub INVLITEcopy(sTab As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If SpalteInTabellegefundenNEW(sTab, "LEKPR", gdBase) Then
    
        sSQL = "insert into INVLITE select "
        sSQL = sSQL & " ARTNR "
        sSQL = sSQL & ", BEZEICH "
        sSQL = sSQL & ", BESTAND "
        sSQL = sSQL & ", KVKPR1 "
        sSQL = sSQL & ", VKPR "
        sSQL = sSQL & ", LEKPR as EKPR"
        sSQL = sSQL & ", LINR "
        sSQL = sSQL & ", LIEFBEZ "
        sSQL = sSQL & ", VKWERT "
        sSQL = sSQL & ", EKWERT "
        sSQL = sSQL & ", STSPANNE "
        sSQL = sSQL & ", SPANNE "
        sSQL = sSQL & ", LPZ "
        sSQL = sSQL & ", MWST "
        sSQL = sSQL & " from " & sTab
        gdBase.Execute sSQL, dbFailOnError
    Else
    
        sSQL = "insert into INVLITE select * from " & sTab
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "INVLITEcopy"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Runden_und_Spannen()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String

    anzeige "normal", "Runden und Nettospannen...", Label6

    '*******************************************************
    '* Alle EKPR und VKPR auf zwei Nachkommastellen runden
    '*******************************************************
    cSQL = "Update INVLITE "
    cSQL = cSQL & "set VKPR = (fix(((VKPR * 100) + 0.5))) / 100"
    cSQL = cSQL & ", EKPR = (fix(((EKPR * 100) + 0.5))) / 100"
    cSQL = cSQL & ", KVKPR1 = (fix(((KVKPR1 * 100) + 0.5))) / 100"
    gdBase.Execute cSQL, dbFailOnError
    
    '*******************************************************
    '* VK-Wert und EK-Wert je Artikel berechnen
    '*******************************************************
    cSQL = "Update INVLITE "
    cSQL = cSQL & " set VKWERT = KVKPR1 * BESTAND"
    cSQL = cSQL & ", EKWERT = EKPR * BESTAND"
    cSQL = cSQL & ", STSPANNE = 0"
    cSQL = cSQL & ", SPANNE = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    '**************************
    '* Nettospanne ermitteln
    '**************************
    
    cSQL = "Update INVLITE "
    cSQL = cSQL & " set "
    cSQL = cSQL & " STSPANNE = ((((KVKPR1/(100 + " & gdMWStV & "))* 100) - EKPR )* 100) / ((KVKPR1/(100 + " & gdMWStV & "))* 100)"
    cSQL = cSQL & " where MWST = 'V' "
    cSQL = cSQL & " and KVKPR1 <> 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update INVLITE "
    cSQL = cSQL & " set "
    cSQL = cSQL & " STSPANNE = ((((KVKPR1/(100 + " & gdMWStE & "))* 100) - EKPR )* 100) / ((KVKPR1/(100 + " & gdMWStE & "))* 100)"
    cSQL = cSQL & " where MWST = 'E' "
    cSQL = cSQL & " and KVKPR1 <> 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update INVLITE "
    cSQL = cSQL & " set "
    cSQL = cSQL & " STSPANNE = ((((KVKPR1/100)* 100) - EKPR )* 100) / ((KVKPR1/100)* 100)"
    cSQL = cSQL & " where MWST = 'O' "
    cSQL = cSQL & " and KVKPR1 <> 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    '**************************
    '* Nettowert ermitteln
    '**************************
    cSQL = "Update INVLITE "
    cSQL = cSQL & " set "
    cSQL = cSQL & " SPANNE = (((KVKPR1/(100 + " & gdMWStV & "))*100)* BESTAND ) - (EKPR * BESTAND)"
    cSQL = cSQL & " where MWST = 'V' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update INVLITE "
    cSQL = cSQL & " set "
    cSQL = cSQL & " SPANNE = (((KVKPR1/(100 + " & gdMWStE & "))*100)* BESTAND ) - (EKPR * BESTAND)"
    cSQL = cSQL & " where MWST = 'E' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update INVLITE "
    cSQL = cSQL & " set "
    cSQL = cSQL & " SPANNE = (((KVKPR1/100)*100)* BESTAND ) - (EKPR * BESTAND)"
    cSQL = cSQL & " where MWST = 'O' "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeige "normal", "", Label6

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Runden_und_Spannen"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub INVLITEakualli()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    anzeige "normal", "Inventurdaten werden aktualisiert...", Label6
    
    sSQL = "Update INVLITE inner join Artikel on Invlite.artnr = artikel.artnr "
    sSQL = sSQL & " Set Invlite.mwst = artikel.mwst "
    sSQL = sSQL & ", Invlite.KVKPR1 = artikel.KVKPR1 "
    sSQL = sSQL & ", Invlite.BEZEICH = artikel.BEZEICH "
    sSQL = sSQL & ", Invlite.VKPR = artikel.VKPR "
    gdBase.Execute sSQL, dbFailOnError
    
    If Option1(7).Value = True Then 'Schnittek
        Label3(3).Caption = "Schnitteinkaufswert"
        sSQL = "Update INVLITE inner join Artikel on Invlite.artnr = artikel.artnr "
        sSQL = sSQL & " Set Invlite.EKPR = artikel.EKPR "
        gdBase.Execute sSQL, dbFailOnError
        
        
        
    ElseIf Option1(4).Value = True Then 'Listeneinkaufspreis
    
    
       
        
        loeschNEW "tempArtlief", gdBase
            
        sSQL = "Select Artnr    "
        If opt1(22).Value = True Then
            sSQL = sSQL & ", max(lekpr) as lek"
        Else
            sSQL = sSQL & ", min(lekpr) as lek"
        End If
        sSQL = sSQL & " into tempArtlief  from artlief"
        sSQL = sSQL & " where lekpr > 0 group by artnr "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Update INVLITE inner join tempArtlief on Invlite.artnr = tempArtlief.artnr "
        sSQL = sSQL & " Set Invlite.EKPR = tempArtlief.lek "
        gdBase.Execute sSQL, dbFailOnError
    
    
    End If
    
    
    sSQL = "Update INVLITE Set invlite.ekpr = 0 where ekpr is null  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    anzeige "normal", "", Label6
    
    Runden_und_Spannen
    
    Label3(0).Caption = VkwertErmittlung
    Label3(0).Visible = True
    Label3(1).Caption = EkwertErmittlung
    Label3(1).Visible = True
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "INVLITEakualli"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Quellcopy(sTab As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW sTab, gdBase
    
    sSQL = "select * into " & sTab & " from INVLITE"
    gdBase.Execute sSQL, dbFailOnError
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Quellcopy"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub MoveDaten2DialogWKL46()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz      As Long
    Dim ctmp          As String
    Dim cSQL          As String
    Dim dWert         As Double
    Dim dBestand      As Double
    Dim dVkPr         As Double
    Dim dKVkPr1       As Double
    Dim dEkpr         As Double
    Dim dVkWert       As Double
    Dim dEkWert       As Double
    Dim dStueckSpanne As Double
    Dim dSpanne       As Double
    Dim dSummeEkPr    As Double
    Dim dSummeVkPr    As Double
    Dim iStufe        As Integer
    Dim rsrs          As Recordset
    
    iStufe = 0
    MSFlexGrid1.Redraw = False
    MSFlexGrid1.Visible = False
    MSFlexGrid1.Rows = 1
    
    anzeige "normal", "Tabelle wird vorbereitet...", Label6
    
    cSQL = "Select * from INVLITE "
    
    
    If Option1(0).Value = True Then
        cSQL = cSQL & "order by ARTNR"
    ElseIf Option1(1).Value = True Then
        cSQL = cSQL & "order by BEZEICH"
    ElseIf Option1(2).Value = True Then
        cSQL = cSQL & "order by LINR,lpz"
    ElseIf Option1(3).Value = True Then
        cSQL = cSQL & "order by linr, Libesnr"
    End If
    
    Dim lcount As Long
    Dim j As Integer
            
    
    iStufe = 1
    Set rsrs = gdBase.OpenRecordset(cSQL)
    iStufe = 2
    If Not rsrs.EOF Then
        lAnzSatz = 0
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        MSFlexGrid1.Visible = False
        MSFlexGrid1.Rows = 1
        iStufe = 3
        Do While Not rsrs.EOF
        
            lcount = lcount - 1
            
            j = lcount Mod 1000
            If j = 0 Then
                anzeige "normal", "Noch " & lcount & " Artikel werden vorbereitet...", Label6
            End If
        
        
            lAnzSatz = lAnzSatz + 1
            MSFlexGrid1.Rows = lAnzSatz + 1
            MSFlexGrid1.Row = lAnzSatz
            
            iStufe = 4
            
            SpaltennummerArtnr = 0
            
            If Not IsNull(rsrs!artnr) Then
                ctmp = rsrs!artnr
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 0
            MSFlexGrid1.Text = ctmp
            
            iStufe = 5
            If Not IsNull(rsrs!BEZEICH) Then
                ctmp = rsrs!BEZEICH
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = ctmp
            
            SpaltennummerMENGE = 2
            
            If Not IsNull(rsrs!BESTAND) Then
                dWert = rsrs!BESTAND
            Else
                dWert = 0
            End If
            dBestand = dWert
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = Format$(dWert, "########0")
            
            iStufe = 7
            If Not IsNull(rsrs!KVKPR1) Then
                dWert = rsrs!KVKPR1
            Else
                dWert = 0
            End If
            dKVkPr1 = dWert
            MSFlexGrid1.Col = 3
            MSFlexGrid1.Text = Format$(dWert, "#####0.00")
          
            If Not IsNull(rsrs!vkpr) Then
                dWert = rsrs!vkpr
            Else
                dWert = 0
            End If
            dVkPr = dWert
            MSFlexGrid1.Col = 4
            MSFlexGrid1.Text = Format$(dWert, "#####0.00")
            
            iStufe = 8
            If Not IsNull(rsrs!ekpr) Then
                dWert = rsrs!ekpr
            Else
                dWert = 0
            End If
            dEkpr = dWert
            MSFlexGrid1.Col = 5
            MSFlexGrid1.Text = Format$(dWert, "#####0.00")
            
            iStufe = 9
            If Not IsNull(rsrs!VKWERT) Then
                dWert = rsrs!VKWERT
            Else
                dWert = 0
            End If
            dSummeVkPr = dSummeVkPr + dWert
            dVkWert = dWert
            MSFlexGrid1.Col = 6
            MSFlexGrid1.Text = Format$(dVkWert, "########0.00")
            
            iStufe = 10
            If Not IsNull(rsrs!EKWERT) Then
                dWert = rsrs!EKWERT
            Else
                dWert = 0
            End If
            dSummeEkPr = dSummeEkPr + dWert
            dEkWert = dWert
            MSFlexGrid1.Col = 7
            MSFlexGrid1.Text = Format$(dEkWert, "########0.00")
            
            iStufe = 11
            If Not IsNull(rsrs!STSPANNE) Then
                dWert = rsrs!STSPANNE
            Else
                dWert = 0
            End If
            dStueckSpanne = dWert
            MSFlexGrid1.Col = 8
            MSFlexGrid1.Text = Format$(dStueckSpanne, "###,###,##0.00")
            
            iStufe = 12
            If Not IsNull(rsrs!SPANNE) Then
                dWert = rsrs!SPANNE
            Else
                dWert = 0
            End If
            dSpanne = dWert
            MSFlexGrid1.Col = 9
            MSFlexGrid1.Text = Format$(dSpanne, "###,###,##0.00")
            
            iStufe = 13
            
            If Not IsNull(rsrs!linr) Then
                ctmp = rsrs!linr
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 10
            MSFlexGrid1.Text = ctmp
            
            iStufe = 14
            If Not IsNull(rsrs!LIEFBEZ) Then
                ctmp = rsrs!LIEFBEZ
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 11
            MSFlexGrid1.Text = ctmp
            
            iStufe = 14
            If Not IsNull(rsrs!LPZ) Then
                ctmp = rsrs!LPZ
            Else
                ctmp = "0"
            End If
            MSFlexGrid1.Col = 12
            MSFlexGrid1.Text = ctmp
            
            iStufe = 15
            If Not IsNull(rsrs!MWST) Then
                ctmp = rsrs!MWST
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 13
            MSFlexGrid1.Text = ctmp
            

            If Not IsNull(rsrs!AWM) Then
                ctmp = rsrs!AWM
            Else
                ctmp = ""
            End If
            FaerbenFlex ctmp, MSFlexGrid1, 0, CInt(lAnzSatz)
            MSFlexGrid1.Col = 14
            MSFlexGrid1.Text = ctmp
            
            If Not IsNull(rsrs!LIBESNR) Then
                ctmp = rsrs!LIBESNR
            Else
                ctmp = ""
            End If
            MSFlexGrid1.Col = 15
            MSFlexGrid1.Text = ctmp
            
            
            
            rsrs.MoveNext
        Loop
            
        iStufe = 16
        MSFlexGrid1.Visible = True
        MSFlexGrid1.Redraw = True
        MSFlexGrid1.Enabled = True
        If Frame5.Visible Then
            Frame5.Visible = False
        End If
        
        iStufe = 17
        
        Label3(0).Caption = VkwertErmittlung
        Label3(0).Visible = True
        Label3(1).Caption = EkwertErmittlung
        Label3(1).Visible = True
        
        
        Frame0.Visible = True
        Frame2.Visible = True
        bnoData = False
        
        Tabellenbreiteanpassen MSFlexGrid1, 1.15 * gdTabfak
        
    Else
        Frame0.Visible = False
        Frame2.Visible = False
        
        bnoData = True

    End If
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", "", Label6

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MoveDaten2DialogWKL46"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
    Resume Next
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
Private Sub haengan()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    anzeige "normal", "Artikelfarbe wird aktualisiert...", Label6
    
    If Not SpalteInTabellegefundenNEW("INVLITE", "AWM", gdBase) Then
        SpalteAnfuegenNEW "INVLITE", "AWM", "Text(2)", gdBase
    
        sSQL = "Update INVLITE inner join artikel on INVLITE.artnr = artikel.artnr "
        sSQL = sSQL & " set INVLITE.AWM = artikel.AWM "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    anzeige "normal", "", Label6
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "haengan"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub haengan1()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    anzeige "normal", "Bestellnummer wird aktualisiert...", Label6
    
    If Not SpalteInTabellegefundenNEW("INVLITE", "LIBESNR", gdBase) Then
        SpalteAnfuegenNEW "INVLITE", "LIBESNR", "Text(13)", gdBase
    
        sSQL = "Update INVLITE inner join artlief on INVLITE.artnr = artlief.artnr and INVLITE.linr = artlief.linr"
        sSQL = sSQL & " set INVLITE.LIBESNR = artlief.LIBESNR "
        gdBase.Execute sSQL, dbFailOnError
           
    End If
    
    anzeige "normal", "", Label6
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "haengan1"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub haengan2()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    anzeige "normal", "AGN wird aktualisiert...", Label6
    
    If Not SpalteInTabellegefundenNEW("INVLITE", "agn", gdBase) Then
        SpalteAnfuegenNEW "INVLITE", "agn", "double", gdBase
    
        sSQL = "Update INVLITE inner join artikel on INVLITE.artnr = artikel.artnr "
        sSQL = sSQL & " set INVLITE.agn = artikel.agn "
        gdBase.Execute sSQL, dbFailOnError
           
    End If
    
    anzeige "normal", "", Label6
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "haengan2"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub haengan3()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Update INVLITE inner join artikel on INVLITE.artnr = artikel.artnr "
    sSQL = sSQL & " set INVLITE.ekpr = artikel.lekpr where INVLITE.ekpr = 0 "
    gdBase.Execute sSQL, dbFailOnError
           
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "haengan3"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub haengan4()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    anzeige "normal", "EAN wird aktualisiert...", Label6
    
    If Not SpalteInTabellegefundenNEW("INV_LITE", "ean", gdBase) Then
        SpalteAnfuegenNEW "INV_LITE", "ean", "Text(13)", gdBase
    
        sSQL = "Update INV_LITE inner join artikel on INV_LITE.artnr = artikel.artnr "
        sSQL = sSQL & " set INV_LITE.ean = artikel.ean "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    anzeige "normal", "", Label6
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "haengan4"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub haengan5()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Not SpalteInTabellegefundenNEW("INV_LITE", "LIBESNR", gdBase) Then
        SpalteAnfuegenNEW "INV_LITE", "LIBESNR", "Text(20)", gdBase
    
        sSQL = "Update INV_LITE inner join artikel on INV_LITE.artnr = artikel.artnr "
        sSQL = sSQL & " set INV_LITE.LIBESNR = artikel.LIBESNR "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "haengan5"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function VkwertErmittlungOhneMW() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    VkwertErmittlungOhneMW = "0"
    
    sSQL = "Select sum(VKwert)as sVKWERT from INVLITE where MWST = 'O' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!sVKWERT) Then
            VkwertErmittlungOhneMW = Format$(rsrs!sVKWERT, "#####0.00")
        Else
            VkwertErmittlungOhneMW = "0"
        End If
        
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "VkwertErmittlungOhneMW"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function ListenVerkaufswertErmittlung() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    ListenVerkaufswertErmittlung = "0"
    
    sSQL = "Select sum(Bestand*VKPR)as sVKWERT from INVLITE "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!sVKWERT) Then
            ListenVerkaufswertErmittlung = Format$(rsrs!sVKWERT, "#####0.00")
        Else
            ListenVerkaufswertErmittlung = "0"
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ListenVerkaufswertErmittlung"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function VkwertErmittlungErmMW() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    VkwertErmittlungErmMW = "0"
    
    sSQL = "Select sum(VKwert)as sVKWERT from INVLITE where MWST = 'E'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!sVKWERT) Then
            VkwertErmittlungErmMW = Format$(rsrs!sVKWERT, "#####0.00")
        Else
            VkwertErmittlungErmMW = "0"
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "VkwertErmittlungErmMW"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function VkwertErmittlungVolleMW() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    VkwertErmittlungVolleMW = "0"
    
    sSQL = "Select sum(VKwert)as sVKWERT from INVLITE where MWST = 'V'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!sVKWERT) Then
            VkwertErmittlungVolleMW = Format$(rsrs!sVKWERT, "#####0.00")
        Else
            VkwertErmittlungVolleMW = "0"
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "VkwertErmittlungVolleMW"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function EkwertErmittlungOhneMW() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    EkwertErmittlungOhneMW = "0"
    
    sSQL = "Select sum(EKwert)as sEKWERT from INVLITE where MWST = 'O' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!sEKWERT) Then
            EkwertErmittlungOhneMW = Format$(rsrs!sEKWERT, "#####0.00")
        Else
            EkwertErmittlungOhneMW = "0"
        End If
        
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EkwertErmittlungOhneMW"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function EkwertErmittlungErmMW() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    EkwertErmittlungErmMW = "0"
    
    sSQL = "Select sum(EKWERT)as sEKWERT from INVLITE where MWST = 'E'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!sEKWERT) Then
            EkwertErmittlungErmMW = Format$(rsrs!sEKWERT, "#####0.00")
        Else
            EkwertErmittlungErmMW = "0"
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EkwertErmittlungErmMW"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function EkwertErmittlungVolleMW() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    EkwertErmittlungVolleMW = "0"
    
    sSQL = "Select sum(EKWERT)as sEKWERT from INVLITE where MWST = 'V'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!sEKWERT) Then
            EkwertErmittlungVolleMW = Format$(rsrs!sEKWERT, "#####0.00")
        Else
            EkwertErmittlungVolleMW = "0"
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EkwertErmittlungVolleMW"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function VkwertErmittlung() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    anzeige "normal", "VK-Wert wird ermittelt...", Label6
    
    VkwertErmittlung = ""
    
    sSQL = "Select sum(VKwert)as sVKWERT from INVLITE"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        VkwertErmittlung = Format$(rsrs!sVKWERT, "#######0.00")
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", "", Label6
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "VkwertErmittlung"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function EkwertErmittlung() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    anzeige "normal", "EK-Wert wird ermittelt...", Label6
    
    EkwertErmittlung = ""
    
    sSQL = "Select sum(EKwert)as sEKWERT from INVLITE"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        EkwertErmittlung = Format$(rsrs!sEKWERT, "#######0.00")
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", "", Label6
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EkwertErmittlung"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub SpeichereInventurDateiWKL46()
    On Error GoTo LOKAL_ERROR
    
    Dim lRet        As Long
    Dim lfail       As Long
    Dim lcount      As Long
    Dim cLBSatz     As String
    Dim cdatei      As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim cPfad       As String
    Dim bDuplikat   As Boolean
    Dim cTab        As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cdatei = Text9.Text
    cdatei = Trim$(UCase$(cdatei))
    If cdatei = "" Then
        anzeigeNew "rot", "Bitte einen gültigen Dateinamen angeben!", Label6
        Text9.SetFocus
        Exit Sub
    End If
    
    cdatei = Label1(0).Caption & cdatei
    
    bDuplikat = False
    For lcount = 0 To List2.ListCount - 1
        cLBSatz = List2.list(lcount)
        cLBSatz = Trim$(UCase$(cLBSatz))
        cLBSatz = Trim(Left(cLBSatz, 8))
        If cLBSatz = cdatei Then
            bDuplikat = True
            Exit For
        End If
    Next lcount
    
    If bDuplikat Then
        lRet = MsgBox("Datei mit diesem Namen existiert bereits!" & vbCrLf & "Überschreiben?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If lRet <> vbYes Then
            Text9.Text = ""
            Text9.SetFocus
            Exit Sub
        End If
    End If
    
    cTab = cdatei
    Quellcopy cTab
    anzeigeNew "normal", "Datei unter " & cTab & " gespeichert!", Label6

    Command3_Click 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeichereInventurDateiWKL46"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
    
End Sub



Private Sub Check2_Click()
    On Error GoTo LOKAL_ERROR
    
    If Check2.Value = vbUnchecked Then
        Frame1.Visible = True
        Text2.Text = ""
        Text2.SetFocus
        Command1.Visible = False
        
    ElseIf Check2.Value = vbChecked Then
        Frame1.Visible = False
        Command1.Visible = True
        Text3(0).SetFocus
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check2_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check6_Click()
On Error GoTo LOKAL_ERROR

Dim cSQL As String

    If Check6.Value = vbChecked Then
        If MSFlexGrid1.Visible = True Then
            If MSFlexGrid1.Rows > 1 Then
            
                Screen.MousePointer = 11
                abschlagen
                
                Runden_und_Spannen
                
'                '*******************************************************
'                '* Alle EKPR und VKPR auf zwei Nachkommastellen runden
'                '*******************************************************
'                cSQL = "Update INVLITE "
'                cSQL = cSQL & "set VKPR = (fix(((VKPR * 100) + 0.5))) / 100"
'                cSQL = cSQL & ", EKPR = (fix(((EKPR * 100) + 0.5))) / 100"
'                cSQL = cSQL & ", KVKPR1 = (fix(((KVKPR1 * 100) + 0.5))) / 100"
'                gdBase.Execute cSQL, dbFailOnError
'
'                '*******************************************************
'                '* VK-Wert und EK-Wert je Artikel berechnen
'                '*******************************************************
'                cSQL = "Update INVLITE "
'                cSQL = cSQL & " set VKWERT = KVKPR1 * BESTAND"
'                cSQL = cSQL & ", EKWERT = EKPR * BESTAND"
'                cSQL = cSQL & ", STSPANNE = 0"
'                cSQL = cSQL & ", SPANNE = 0 "
'                gdBase.Execute cSQL, dbFailOnError
'
'                '**************************
'                '* Nettospanne ermitteln
'                '**************************
'
'                cSQL = "Update INVLITE "
'                cSQL = cSQL & " set "
'                cSQL = cSQL & " STSPANNE = ((((KVKPR1/(100 + " & gdMWStV & "))* 100) - EKPR )* 100) / ((KVKPR1/(100 + " & gdMWStV & "))* 100)"
'                cSQL = cSQL & " where MWST = 'V' "
'                cSQL = cSQL & " and KVKPR1 <> 0 "
'                gdBase.Execute cSQL, dbFailOnError
'
'                cSQL = "Update INVLITE "
'                cSQL = cSQL & " set "
'                cSQL = cSQL & " STSPANNE = ((((KVKPR1/(100 + " & gdMWStE & "))* 100) - EKPR )* 100) / ((KVKPR1/(100 + " & gdMWStE & "))* 100)"
'                cSQL = cSQL & " where MWST = 'E' "
'                cSQL = cSQL & " and KVKPR1 <> 0 "
'                gdBase.Execute cSQL, dbFailOnError
'
'                cSQL = "Update INVLITE "
'                cSQL = cSQL & " set "
'                cSQL = cSQL & " STSPANNE = ((((KVKPR1/100)* 100) - EKPR )* 100) / ((KVKPR1/100)* 100)"
'                cSQL = cSQL & " where MWST = 'O' "
'                cSQL = cSQL & " and KVKPR1 <> 0 "
'                gdBase.Execute cSQL, dbFailOnError
'
'                '**************************
'                '* Nettowert ermitteln
'                '**************************
'                cSQL = "Update INVLITE "
'                cSQL = cSQL & " set "
'                cSQL = cSQL & " SPANNE = (((KVKPR1/(100 + " & gdMWStV & "))*100)* BESTAND ) - (EKPR * BESTAND)"
'                cSQL = cSQL & " where MWST = 'V' "
'                gdBase.Execute cSQL, dbFailOnError
'
'                cSQL = "Update INVLITE "
'                cSQL = cSQL & " set "
'                cSQL = cSQL & " SPANNE = (((KVKPR1/(100 + " & gdMWStE & "))*100)* BESTAND ) - (EKPR * BESTAND)"
'                cSQL = cSQL & " where MWST = 'E' "
'                gdBase.Execute cSQL, dbFailOnError
'
'                cSQL = "Update INVLITE "
'                cSQL = cSQL & " set "
'                cSQL = cSQL & " SPANNE = (((KVKPR1/100)*100)* BESTAND ) - (EKPR * BESTAND)"
'                cSQL = cSQL & " where MWST = 'O' "
'                gdBase.Execute cSQL, dbFailOnError
            
                haengan
                haengan1
                
                MoveDaten2DialogWKL46
                
                Screen.MousePointer = 0
                
            End If
        End If
    End If
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check6_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub cmdHelp_Click()
    On Error GoTo LOKAL_ERROR
    
    zeigeHilfe "KISSHELP", Me.Tag & ".doc", gcPfad

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdHelp_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo2_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    AutocompleteCombo KeyCode, Shift, Combo2
    
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
            Combo2.Text = gF2Prompt.cWahl
            
        End If
        Combo2.SetFocus
    
    End If
    
    If KeyCode = vbKeyEscape Then
        Command2_Click 10
    End If
    
    If KeyCode = vbKeyReturn Then
        Command6_Click 0
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo2_LostFocus()
On Error GoTo LOKAL_ERROR
    
    Combo2.BackColor = vbWhite
    If Combo2.Text = "" Then
        Combo3.Clear
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo3_Click()
On Error GoTo LOKAL_ERROR

    Dim clpz As String
    
    clpz = Left(Combo3.Text, 3)
    clpz = Trim(clpz)
    
    List7.Visible = True
    List7.AddItem clpz
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo3_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Combo3_GotFocus()
    On Error GoTo LOKAL_ERROR

    Dim cLinr As String
    
    Combo3.BackColor = glSelBack1


    cLinr = ErmittleLinr(Combo2.Text)
    If cLinr <> "" Then
        LeseLinie Combo3, cLinr
        

    Else
        Combo3.Clear

        Combo2.SetFocus
    End If
    
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo3_LostFocus()
On Error GoTo LOKAL_ERROR
    
    Combo3.BackColor = vbWhite

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo3_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo3_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyEscape Then
        Command2_Click 10
    End If
    
    If KeyCode = vbKeyReturn Then
        Command6_Click 0
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo2_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Combo2.BackColor = glSelBack1
    Combo2.SelStart = Len(Combo2.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim lcol As Long
    Dim ctmp As String
    
    If Frame4.Visible = True Then
        Select Case Index
            Case 0 To 9         'Ziffern
                Text9.Text = Text9.Text & Command0(Index).Caption
            
            Case Is = 11        'Clear
                Text9.Text = ""
                
            Case Else
                'nichts tun
                
        End Select
        Text9.SetFocus
        
    Else
        If Label0(2).Caption = "1" Then
            ctmp = Text2.Text
            If ctmp = "______" Then
                ctmp = ""
            Else
                ctmp = Trim$(Str$(Val(ctmp)))
            End If
            Select Case Index
                Case 0 To 9         'Ziffern
                    ctmp = ctmp & Command0(Index).Caption
                
                Case Is = 11        'Clear
                    ctmp = ""
                    
                Case Else
                    'nichts tun
                    
            End Select
            If Len(ctmp) < 6 Then
                ctmp = ctmp & String$(6 - Len(ctmp), "_")
            End If
            Text2.Text = ctmp
            Text2.SetFocus
            
        Else
            lrow = Val(Label0(0).Caption)
            lcol = Val(Label0(1).Caption)
            
            If lrow < 1 Or lcol < 2 Then
                lrow = 1
                lcol = 2
            End If
            If lrow > MSFlexGrid1.Rows - 1 Or lcol > 5 Then
                lrow = MSFlexGrid1.Rows - 1
                lcol = 5
            End If
            
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.Col = lcol
            
            Select Case Index
                Case 0 To 9         'Ziffern
                    MSFlexGrid1.Text = MSFlexGrid1 & Command0(Index).Caption
                    
                Case Is = 10        'Komma
                    If InStr(MSFlexGrid1.Text, ",") = 0 Then
                        MSFlexGrid1.Text = MSFlexGrid1 & Command0(Index).Caption
                    End If
                    
                Case Is = 11        'Clear
                    MSFlexGrid1.Text = ""
                    
                Case Is = 12        'Links
                    If lcol > 2 Then
                        lcol = lcol - 1
                    End If
                Case Is = 13        'Hoch
                    If lrow > 1 Then
                        lrow = lrow - 1
                    End If
        
                Case Is = 14        'Tief
                    If lrow < MSFlexGrid1.Rows - 1 Then
                        lrow = lrow + 1
                    End If
                Case Is = 15        'Rechts
                    If lcol < 5 Then
                        lcol = lcol + 1
                    End If
            End Select
            
            Label0(0).Caption = Trim$(Str$(lrow))
            Label0(1).Caption = Trim$(Str$(lcol))
            
            MSFlexGrid1.SetFocus
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.Col = lcol
        End If
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Command1_Click()
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    Dim lAnzSatz    As Long
    Dim cLinr       As String
    Dim cSQL        As String
    Dim ctmp        As String
    Dim cPfad       As String
    
    Dim dWert       As Double
    Dim iRet        As Integer
    Dim rsrs        As Recordset
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cLinr = Text2.Text
    If cLinr = "" Then
        cLinr = ""
    Else
        cLinr = Trim$(Str$(Val(cLinr)))
    End If

    anzeigeNew "normal", "Daten werden ermittelt...", Label6

    invLiteRefresh
    
    
    
    '************************************
    '* Alle zutreffenden Artikel einlesen
    '************************************
    If Option1(7).Value = True Then 'Schnittek
        Label3(3).Caption = "Schnitteinkaufswert"
        'EKPR
        cSQL = "Insert into INVLITE "
        cSQL = cSQL & " Select A.ARTNR,A.BEZEICH,A.BESTAND,A.VKPR"
        cSQL = cSQL & " ,A.KVKPR1,A.EKPR "
        cSQL = cSQL & " ,A.LINR, '' as LIEFBEZ, A.LPZ, A.MWST "
        cSQL = cSQL & " from  ARTIKEL A   "
        cSQL = cSQL & " where A.BESTAND > 0   "
        If cLinr <> "" Then
            cSQL = cSQL & "and A.LINR = " & cLinr & " "
        End If
        cSQL = cSQL & " group by  A.ARTNR, A.BEZEICH,A.BESTAND,A.VKPR"
        cSQL = cSQL & " ,A.KVKPR1,A.EKPR  "
        cSQL = cSQL & " ,A.LINR, A.LPZ, A.MWST"
        gdBase.Execute cSQL, dbFailOnError
        
    ElseIf Option1(4).Value = True Then 'Listeneinkaufspreis
        'LEKPR

        loeschNEW "ARTIKEL_INV", gdBase
        cSQL = "select artnr into ARTIKEL_INV from artikel where bestand > 0 "
        gdBase.Execute cSQL, dbFailOnError

        cSQL = "Create Index ARTNR on ARTIKEL_INV(ARTNR)"
        gdBase.Execute cSQL, dbFailOnError
        
        
        
        cSQL = "Insert into INVLITE "
        cSQL = cSQL & " Select ARTNR"
        
        
        If cLinr <> "" Then
            If opt1(22).Value = True Then
                cSQL = cSQL & " ,max(LEKPR) as EKPR "
            Else
                cSQL = cSQL & " ,min(LEKPR) as EKPR "
            End If
        Else
            cSQL = cSQL & " ,0 as EKPR "
        End If
        
        
        
        cSQL = cSQL & " from ARTLIEF "
        If cLinr <> "" Then
            cSQL = cSQL & " where LINR = " & cLinr & " "
            cSQL = cSQL & " and Artnr in (select artnr from ARTIKEL_INV ) "
        Else
            cSQL = cSQL & " where Artnr in (select artnr from ARTIKEL_INV ) "
        End If

        cSQL = cSQL & " group by ARTNR "
        gdBase.Execute cSQL, dbFailOnError
        
        
        
        
        
        
        

'        cSQL = "Insert into INVLITE "
'        cSQL = cSQL & " Select ARTNR"
'
'        If opt1(22).Value = True Then
'            cSQL = cSQL & " ,max(LEKPR) as EKPR "
'        Else
'            cSQL = cSQL & " ,min(LEKPR) as EKPR "
'        End If
'
'        cSQL = cSQL & " from ARTLIEF "
'
'        If cLinr <> "" Then
'            cSQL = cSQL & " where LINR = " & cLinr & " "
'            cSQL = cSQL & " and Artnr in (select artnr from ARTIKEL_INV ) "
'        Else
'            cSQL = cSQL & " where Artnr in (select artnr from ARTIKEL_INV ) "
'        End If
'
'        cSQL = cSQL & " group by ARTNR "
'        gdBase.Execute cSQL, dbFailOnError
'
        
        
        
        loeschNEW "ARTIKEL_INV", gdBase

        cSQL = "Update INVLITE inner join Artikel on INVLITE.artnr = Artikel.artnr "
        cSQL = cSQL & " Set INVLITE.BEZEICH = Artikel.BEZEICH "
        cSQL = cSQL & " , INVLITE.KVKPR1 = Artikel.KVKPR1 "
        cSQL = cSQL & " , INVLITE.VKPR = Artikel.VKPR "
        cSQL = cSQL & " , INVLITE.BESTAND = Artikel.BESTAND "
        cSQL = cSQL & " , INVLITE.MWST = Artikel.MWST "
        cSQL = cSQL & " , INVLITE.LPZ = Artikel.LPZ "
        If cLinr <> "" Then
            cSQL = cSQL & " , INVLITE.LINR =" & cLinr & " "
        Else
'            cSQL = cSQL & " , INVLITE.LINR = Artikel.LINR "
            cSQL = cSQL & " , INVLITE.LINR = 0 "
        End If

        gdBase.Execute cSQL, dbFailOnError
        
        If cLinr = "" Then
            'jetzt den kleinsten/größten EK Lieferanten holen/einfügen
            
            loeschNEW "ARTIKEL_MaxMin_" & srechnertab, gdBase
            
            If opt1(22).Value = True Then
                'max(LEKPR)
                
                cSQL = "select artnr, max(lekpr) as xlekpr into ARTIKEL_MaxMin_" & srechnertab & " from Artlief "
                cSQL = cSQL & " where artnr in(Select artnr from INVLITE )"
                cSQL = cSQL & " and RKZ = 'N' "
                cSQL = cSQL & " group by artnr "
                gdBase.Execute cSQL, dbFailOnError
            Else
                'min(LEKPR)
                
                cSQL = "select artnr, min(lekpr) as xlekpr into ARTIKEL_MaxMin_" & srechnertab & " from Artlief "
                cSQL = cSQL & " where artnr in(Select artnr from INVLITE )"
                cSQL = cSQL & " and RKZ = 'N' "
                cSQL = cSQL & " group by artnr "
                gdBase.Execute cSQL, dbFailOnError
            End If
            
            'Alter table
            cSQL = "Alter Table ARTIKEL_MaxMin_" & srechnertab & " add Linr Long "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update ARTIKEL_MaxMin_" & srechnertab & " inner join Artlief on ARTIKEL_MaxMin_" & srechnertab & ".xlekpr = Artlief.lekpr "
            cSQL = cSQL & " and ARTIKEL_MaxMin_" & srechnertab & ".artnr = Artlief.artnr "
            cSQL = cSQL & " Set ARTIKEL_MaxMin_" & srechnertab & ".LINR = Artlief.LINR "
            cSQL = cSQL & " where RKZ = 'N' "
            gdBase.Execute cSQL, dbFailOnError
        
            cSQL = "Update INVLITE inner join ARTIKEL_MaxMin_" & srechnertab & " on INVLITE.artnr = ARTIKEL_MaxMin_" & srechnertab & ".artnr "
            cSQL = cSQL & " Set INVLITE.ekpr = ARTIKEL_MaxMin_" & srechnertab & ".xlekpr "
            cSQL = cSQL & " , INVLITE.linr = ARTIKEL_MaxMin_" & srechnertab & ".linr "
            gdBase.Execute cSQL, dbFailOnError
            
            
            
            'und zum Schluss alle ohne Linr updaten
            'jetzt den kleinsten/größten EK Lieferanten holen/einfügen
            
            loeschNEW "ARTIKEL_MaxMin_" & srechnertab, gdBase
            
            If opt1(22).Value = True Then
                'max(LEKPR)
                
                cSQL = "select artnr, max(lekpr) as xlekpr into ARTIKEL_MaxMin_" & srechnertab & " from Artlief "
                cSQL = cSQL & " where artnr in(Select artnr from INVLITE )"
                cSQL = cSQL & " group by artnr "
                gdBase.Execute cSQL, dbFailOnError
            Else
                'min(LEKPR)
                
                cSQL = "select artnr, min(lekpr) as xlekpr into ARTIKEL_MaxMin_" & srechnertab & " from Artlief "
                cSQL = cSQL & " where artnr in(Select artnr from INVLITE )"
                cSQL = cSQL & " group by artnr "
                gdBase.Execute cSQL, dbFailOnError
            End If
            
            'Alter table
            cSQL = "Alter Table ARTIKEL_MaxMin_" & srechnertab & " add Linr Long "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update ARTIKEL_MaxMin_" & srechnertab & " inner join Artlief on ARTIKEL_MaxMin_" & srechnertab & ".xlekpr = Artlief.lekpr "
            cSQL = cSQL & " and ARTIKEL_MaxMin_" & srechnertab & ".artnr = Artlief.artnr "
            cSQL = cSQL & " Set ARTIKEL_MaxMin_" & srechnertab & ".LINR = Artlief.LINR "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Update INVLITE inner join ARTIKEL_MaxMin_" & srechnertab & " on INVLITE.artnr = ARTIKEL_MaxMin_" & srechnertab & ".artnr "
            cSQL = cSQL & " Set INVLITE.ekpr = ARTIKEL_MaxMin_" & srechnertab & ".xlekpr "
            cSQL = cSQL & " , INVLITE.linr = ARTIKEL_MaxMin_" & srechnertab & ".linr "
            cSQL = cSQL & " where INVLITE.linr = 0 "
            gdBase.Execute cSQL, dbFailOnError
            
        
        End If
        
        
        
        
        
        
        
        
    End If
    
    cSQL = "Update INVLITE inner join lisrt on invlite.linr = lisrt.linr "
    cSQL = cSQL & " Set invlite.liefbez = lisrt.liefbez "
    gdBase.Execute cSQL, dbFailOnError
   
    cSQL = "Update INVLITE "
    cSQL = cSQL & " Set invlite.ekpr = 0 where ekpr is null  "
    gdBase.Execute cSQL, dbFailOnError
    
    haengan3
    
    haengan2
    
    
    If Check6.Value = vbChecked Then
        abschlagen
    End If
    
    Runden_und_Spannen
   
    haengan
    haengan1
    
    MoveDaten2DialogWKL46
    
    If Not bnoData Then
        MSFlexGrid1.SetFocus
        MSFlexGrid1.Row = 1
        MSFlexGrid1.Col = 2
        If cLinr <> "" Then
            Label5.Caption = "Lieferant: " & cLinr
            Label5.Refresh
        Else
            Label5.Caption = "alle Lieferanten"
            Label5.Refresh
            
        End If
        anzeigeNew "normal", "Fertig! Die Daten sind jetzt ermittelt", Label6
    Else
        anzeigeNew "rot", "Keine Daten gefunden.", Label6

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
        Fehler.gsFunktion = "Command1_Click"
        Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub abschlagen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    anzeige "normal", "Abschlag wird angewendet...", Label6
    
    sSQL = "Update INVLITE inner join agndbf on INVLITE.agn = agndbf.agn "
    sSQL = sSQL & " set INVLITE.EKPR = INVLITE.EKPR - (INVLITE.EKPR * agndbf.invab / 100) "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "", Label6
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "abschlagen"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command10_Click()
On Error GoTo LOKAL_ERROR
    Dim lcount As Long
    
    If Text4.Text = "" Then
        anzeigeNew "rot", "Bitte wählen Sie einen Lieferanten aus!", Label6
        Text4.SetFocus
        Exit Sub
    Else
    
    End If
    
    If IsNumeric(Text4.Text) = False Then
        anzeigeNew "rot", "Bitte wählen Sie einen Lieferanten aus!", Label6
        Text4.SetFocus
        Exit Sub
    Else
    
    End If
    
    gF2Prompt.bMultiple = True
    gF2Prompt.cFeld = "LPZ"
    
    gF2Prompt.cWert = Trim$(Str$(Val(Text4.Text)))
        
    If gF2Prompt.cFeld <> "" Then
        frmWK00a.Show 1
    End If
        
    List8.Clear
    For lcount = 0 To 100
        If gF2Prompt.cArray(lcount) <> "" Then
            List8.Visible = True
            List8.AddItem gF2Prompt.cArray(lcount)
        End If
    Next lcount

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command10_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Command11_Click()
On Error GoTo LOKAL_ERROR
    
        
    List8.Clear
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Command13_Click()
On Error GoTo LOKAL_ERROR
    
    Dim lDat As Long
    If IsDate(Text1(5).Text) = False Then
        Text1(5).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Else
        If IsDate(Text1(5).Text) = True Then
            lDat = CLng(DateValue(Text1(5).Text))
        End If
        lDat = lDat + 1
        Text1(5).Text = Format(lDat, "DD.MM.YYYY")
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command13_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command12_Click()
On Error GoTo LOKAL_ERROR

    Dim lDat As Long

    If IsDate(Text1(5).Text) = False Then
        Text1(5).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Else
        If IsDate(Text1(5).Text) = True Then
            lDat = CLng(DateValue(Text1(5).Text))
        End If
        lDat = lDat - 1
        Text1(5).Text = Format(lDat, "DD.MM.YYYY")
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command12_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Command14_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iFeld As Integer
    Dim cZeichen As String
    
    iFeld = Val(Label11.Caption)
    
    Select Case Index
        Case 20 To 29
            Text3(iFeld).Text = Text3(iFeld).Text & Command14(Index).Caption

        Case Is = 31
            Text3(iFeld).Text = ""
            anzeigeNew "normal", "", Label6
    End Select
    
    Text3(iFeld).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command14_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Command15_Click()
On Error GoTo LOKAL_ERROR

Abgleich_Artikel_Bestand

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command15_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function Is_ArtNr_In(lartnr As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Is_ArtNr_In = False
    
    cSQL = "Select * from Artikel where Artnr = " & lartnr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        Is_ArtNr_In = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Is_ArtNr_In"
    Fehler.gsFehlertext = "Im Programmteil Artikelabgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function pfadseekExcel_Artikel_im() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    Dim sExcelpfad  As String
    
    pfadseekExcel_Artikel_im = False

    sTitle = "Speichern des Pfades"
    
    sFilter = "Excel - Dateien (*.xls)|*.xls"
    
    sOldpfad = ""
    sExcelpfad = pfadaendernplusDatname(sTitle, sFilter, sOldpfad)
    
    If sExcelpfad <> "" Then
        pfadseekExcel_Artikel_im = True
        Label1(0).Caption = sExcelpfad
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "pfadseekExcel_Artikel_im"
    Fehler.gsFehlertext = "Im Programmteil Artikelabgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function pfadseekCSV_Artikel_im() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    Dim sCSVpfad  As String
    
    pfadseekCSV_Artikel_im = False

    sTitle = "Speichern des Pfades"
    
    sFilter = "CSV - Dateien (*.csv)|*.csv"
    
    sOldpfad = ""
    sCSVpfad = pfadaendernplusDatname(sTitle, sFilter, sOldpfad)
    
    If sCSVpfad <> "" Then
        pfadseekCSV_Artikel_im = True
        Label1(0).Caption = sCSVpfad
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "pfadseekCSV_Artikel_im"
    Fehler.gsFehlertext = "Im Programmteil Artikelabgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Abgleich_Artikel_Bestand()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPfad As String
    Dim dbExcel As Database
    Dim lAnzZ As Long
    Dim rsrs As Recordset
    Dim gsExcel50 As String
    Dim cBestand As String
    Dim bnichtgefunden As Boolean
    
    bnichtgefunden = False
    gsExcel50 = "Excel 5.0;"
    
    If pfadseekExcel_Artikel_im = False Then
        anzeige "rot2", "Abbruch durch Benutzer", Label6
        Exit Sub
    End If
    
    Screen.MousePointer = 11

    anzeige "normal", "", Label6
    cPfad = Label1(0).Caption
    
    Set dbExcel = OpenDatabase(cPfad, 0, 0, gsExcel50)

    lAnzZ = 0
    Set rsrs = dbExcel.OpenRecordset("Bestand$")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!ARTNRKISS) Then
                
                cBestand = ""
                
                If Is_ArtNr_In(CLng(rsrs!ARTNRKISS)) Then
                
                    If Not IsNull(rsrs!BESTAND) Then
                        cBestand = rsrs!BESTAND
                    End If
                    
                    sSQL = "Update Artikel set Bestand = '" & cBestand & "'"
                    sSQL = sSQL & " , Lastdate = '" & DateValue(Now) & "' "
                    sSQL = sSQL & " where artnr = " & CLng(rsrs!ARTNRKISS)
                    gdBase.Execute sSQL, dbFailOnError
                    
                    lAnzZ = lAnzZ + 1
                Else
                    'Nicht gefunden schreib Protokoll
                    bnichtgefunden = True
                    schreibeProtokollInventurImport " " & rsrs!ARTNRKISS & " nicht gefunden -> kein Bestandsupdate"
                End If
                
                
                    
            End If
        rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If bnichtgefunden Then
        zeigeHilfeDabapfad "LPROTOK", "Inventur_Import.txt"
    End If
        
    anzeige "normal", lAnzZ & " Artikel wurden abgeglichen.", Label6
    
    Screen.MousePointer = 0

    dbExcel.Close
    
Exit Sub
LOKAL_ERROR:
    
    If err.Number = 3125 Then
        anzeige "rot", "Die Excelliste hat nicht das erwartete Format", Label1(4)
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Abgleich_Artikel_Bestand"
        Fehler.gsFehlertext = "Im Programmteil Artikelabgleich ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Abgleich_Artikel_Bestand_IP()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPfad As String
    Dim rsrs As Recordset
    Dim lAnz As Long
    Dim lGesbestand As Long
    
    If pfadseekCSV_Artikel_im = False Then
        anzeige "rot2", "Abbruch durch Benutzer", Label6
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    anzeige "normal", "", Label6
    cPfad = Label1(0).Caption
    
    IP_INV_EINZEL_Auslesen cPfad, txtStatus, picprogress
    
    Screen.MousePointer = 11
    
    loeschNEW "IP_INV_GROUP", gdBase
    
    sSQL = "Select Artnr, sum(Menge) as SUMMENGE into IP_INV_GROUP from IP_INV "
    sSQL = sSQL & " group by Artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set Bestand = 0 , Lastdate = '" & DateValue(Now) & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel a inner join IP_INV_GROUP i on a.artnr = i.artnr "
    sSQL = sSQL & " set a.bestand = i.summenge , a.Lastdate = '" & DateValue(Now) & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select count(*) as Maxi from Artikel where bestand > 0 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            lAnz = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Select sum(Bestand) as Maxi from Artikel  "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            lGesbestand = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", lAnz & " Artikel mit einem Gesamtbestand von " & lGesbestand & " Stück wurden abgeglichen.", Label6
    
    Screen.MousePointer = 0

    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Abgleich_Artikel_Bestand_IP"
    Fehler.gsFehlertext = "Im Programmteil Artikelabgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Abgleich_Artikel_Bestand_COSYS()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPfad As String
    Dim rsrs As Recordset
    Dim lAnz As Long
    Dim lGesbestand As Long
    
    If pfadseekCSV_Artikel_im = False Then
        anzeige "rot2", "Abbruch durch Benutzer", Label6
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    anzeige "normal", "", Label6
    cPfad = Label1(0).Caption
    
    COSYS_INV_EINZEL_Auslesen cPfad, txtStatus, picprogress
    
    loeschNEW "COSYS_INV_GROUP", gdBase

    'wir arbeiten hier mit EAN
    sSQL = "Update COSYS_INV set artnr = 0  "
    
    'nur ein Test für schäfer
    sSQL = sSQL & " where Val(ean) > 0 "
    
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update COSYS_INV set ean = Val(ean) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update COSYS_INV set ean = '0' & ean  where len(ean) = 11 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    'nach Bereinigung jetzt Artnr holen
    
    sSQL = "Update COSYS_INV i inner join Artikel on "
    sSQL = sSQL & " i.EAN = ARTIKEL.EAN "
    sSQL = sSQL & "Set i.artnr = Artikel.artnr  "
    sSQL = sSQL & " where i.ean  <> '0' "
    sSQL = sSQL & " and i.artnr = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update COSYS_INV i inner join Artikel on "
    sSQL = sSQL & " i.EAN = ARTIKEL.EAN2 "
    sSQL = sSQL & "Set i.artnr = Artikel.artnr  "
    sSQL = sSQL & " where i.ean  <> '0' "
    sSQL = sSQL & " and i.artnr = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update COSYS_INV i inner join Artikel on "
    sSQL = sSQL & " i.EAN = ARTIKEL.EAN3 "
    sSQL = sSQL & "Set i.artnr = Artikel.artnr  "
    sSQL = sSQL & " where i.ean  <> '0' "
    sSQL = sSQL & " and i.artnr = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update COSYS_INV i inner join ARTEAN_K on "
    sSQL = sSQL & " i.EAN = ARTEAN_K.EAN "
    sSQL = sSQL & "Set i.artnr = ARTEAN_K.artnr  "
    sSQL = sSQL & " where i.ean  <> '0' "
    sSQL = sSQL & " and i.artnr = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    'Ende mit EAN
    
    
    
    'am Ende noch eine Besonderheit für Köhler
    'manchmal befinden sich Artnr in der EAN spalte und gleichzeitig bei Artnr eine 0
    
    sSQL = "Update COSYS_INV set artnr = val(ean)  "
    sSQL = sSQL & " where Artnr = 0 and len(ean) <= 6 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select Artnr, sum(Menge) as SUMMENGE into COSYS_INV_GROUP from COSYS_INV "
    sSQL = sSQL & " group by Artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from COSYS_INV_GROUP where artnr = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set Bestand = 0 , Lastdate = '" & DateValue(Now) & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel a inner join COSYS_INV_GROUP i on a.artnr = i.artnr "
    sSQL = sSQL & " set a.bestand = i.summenge , a.Lastdate = '" & DateValue(Now) & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select count(*) as Maxi from Artikel where bestand > 0 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            lAnz = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Select sum(Bestand) as Maxi from Artikel  "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            lGesbestand = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", lAnz & " Artikel mit einem Gesamtbestand von " & lGesbestand & " Stück wurden abgeglichen.", Label6
    
    Screen.MousePointer = 0

    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Abgleich_Artikel_Bestand_COSYS"
    Fehler.gsFehlertext = "Im Programmteil Artikelabgleich ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command16_Click()
On Error GoTo LOKAL_ERROR

Abgleich_Artikel_Bestand_IP

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command16_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command17_Click()
On Error GoTo LOKAL_ERROR

txtAGN_KeyUp vbKeyF2, 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command17_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command18_Click()
On Error GoTo LOKAL_ERROR

Abgleich_Artikel_Bestand_COSYS

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command18_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Dim iRet As Integer
    Dim cLinr As String
    Dim sSQL As String
    
    Screen.MousePointer = 11
    
    Select Case Index
    
        Case 28
            'Daten werden aktualisiert
            INVLITEakualli
    
            haengan
            haengan1
            haengan2
            MoveDaten2DialogWKL46
        Case 19
            rueck2
            rueck1
        Case 17
            rueck1
        Case 18
            UebernahmeInventur
        Case 16
            BestandaufNull
        Case 24 'farbe
            Screen.MousePointer = 0
            frmWKL49.Show 1
        Case 25
            Exportiere_for_MDE
        Case 27
            If List9.ListIndex >= 0 Then
                anzeige "normal", "", Label6
                Screen.MousePointer = 11
                iRet = MsgBox("Möchten Sie die Werte als Lieferantenübersicht angezeigt bekommen? (Nein = Artikelansicht)", vbQuestion + vbYesNo, "Winkiss Frage:")
                If iRet = vbYes Then
                    zeige_Best_Hist_GDPdU Left(List9.list(List9.ListIndex), 8), Label6, False, txtAGN.Text
'                    zeige_Best_Hist List9.list(List9.ListIndex), txtAgn.Text
                Else
                    zeige_Best_Hist_Einzel_GDPdU Left(List9.list(List9.ListIndex), 8), Label6, False, txtAGN.Text
'                    zeige_Best_Hist_Einzel List9.list(List9.ListIndex), txtAGN.Text
                End If
                Screen.MousePointer = 0
            Else
                anzeige "rot", "Wählen Sie bitte ein Datum aus!", Label6
            End If
        
        Case 23
        
            anzeige "normal", "", Label14
            
            sSQL = "Delete from artErrIn "
            gdBase.Execute sSQL, dbFailOnError
            
            MDElesen
            If mdeErr Then
                reportbildschirm "", "aWKL46e" 'Error artikel mde
            End If
        Case 29
            Screen.MousePointer = 0
            frmWKL213.Show 1
        Case 0
            Zeigeauswahlframe
        Case 1 'zurück aus Inventurberechnung
            iRet = MsgBox("Möchten Sie wirklich die Inventur beenden?", vbQuestion + vbYesNo, "Winkiss Frage:")
            If iRet = vbYes Then
                Frame7.Visible = False
                Frame6.Visible = True
                Frame1.Visible = False
                anzeigeNew "normal", "Wie möchten Sie vorgehen?", Label6
            End If
        
        Case 26 'zurück aus best Hist
            Frame23.Visible = False
            Frame6.Visible = True
            anzeigeNew "normal", "Wie möchten Sie vorgehen?", Label6
        Case 6 'zurück aus einstellungen
            Frame8.Visible = False
            Frame6.Visible = True
            anzeigeNew "normal", "Wie möchten Sie vorgehen?", Label6
        Case 13 'zurück aus Import
            Frame19.Visible = False
            Frame6.Visible = True
            anzeigeNew "normal", "Wie möchten Sie vorgehen?", Label6
        Case 21
            'zurück aus Export
            Frame21.Visible = False
            Frame6.Visible = True
            anzeigeNew "normal", "Wie möchten Sie vorgehen?", Label6
            
        Case 7 'zurück aus Inventur Scanner
            Frame9.Visible = False
            Frame6.Visible = True
            anzeigeNew "normal", "Wie möchten Sie vorgehen?", Label6
        Case 8 'zurück aus Inventur MDE
            Frame10.Visible = False
            Frame6.Visible = True
            anzeigeNew "normal", "Wie möchten Sie vorgehen?", Label6
        Case 9 'zurück aus Inventur mit liste
            MSHFlexGrid1.Visible = False
            Label34.Caption = ""
            Frame11.Visible = False
            Frame6.Visible = True
            anzeigeNew "normal", "Wie möchten Sie vorgehen?", Label6
        Case 10 'zurück aus Inventurliste erzeugen
            voreinstellungspeichernE46il
            Frame12.Visible = False
            Frame6.Visible = True
            anzeigeNew "normal", "Wie möchten Sie vorgehen?", Label6
            
        Case 16
            ExportCSV
            
        Case 15 'zurück aus schritt 2
        
            If ue And Not brueck2 Then
                iRet = MsgBox("Haben Sie das Übernahmeprotokoll und die Inventurauswertung ausgedruckt?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
                If iRet = vbNo Then
                    
                    anzeigeNew "normal", "Protokolle werden erstellt...", Label6
                    
                    reportbildschirm "", "aWKL46c" 'übernahme
                    
                    If Modul6.FindFile(gcDBPfad, "aWKL46dS.rpt") Then
                        reportbildschirm "INVENe", "aWKL46dS"
                    Else
                        reportbildschirm "", "aWKL46d" 'Differenz
                    End If
                    
                    Exit Sub
                End If
            End If
            
            Frame15.Visible = False
            
            'mde oder scanner
            
            If iscan = 3 Then ' kommt von scanner
                Frame9.Visible = True
                anzeigeNew "normal", "", Label6
                Frame1.Visible = False
                Command1.Visible = True
                fuellelist List3
                Text3(0).SetFocus
                Label11.Caption = "0"
            ElseIf iscan = 4 Then 'kommt von mde
                Frame10.Visible = True
                anzeigeNew "normal", "", Label6
                Frame1.Visible = False
                Command1.Visible = True
                fuellelist List4
            End If
            
        Case 14 'zu Schritt2 aus scanner
            Frame9.Visible = False
            openschritt2
        Case 20 'zu Schritt2 aus mde
        
            Frame10.Visible = False
            openschritt2
        Case 11
        
            Dim cValid As String
            Dim cFeld As String
            Dim cZeichen As String
            Dim lcount As Long
            Dim bTextSuche As Boolean
            
            anzeigeNew "normal", "", Label6
            
            Screen.MousePointer = 11
            
            cValid = "1234567890"
            
            
            If Left(Text3(0).Text, 1) = "ß" Then
                Text3(0).Text = Right(Text3(0).Text, Len(Text3(0).Text) - 1)
            End If
            
            cFeld = Text3(0).Text
            
            bTextSuche = False
            
            For lcount = 1 To Len(cFeld)
                cZeichen = Mid(cFeld, lcount, 1)
                If InStr(cValid, cZeichen) = 0 Then
                    bTextSuche = True
                    Exit For
                End If
            Next lcount
            
            If bTextSuche Then
                gcSuch = Text3(0).Text
                gsARTNR = ""
                frmWKL70.Show 1
                Me.Refresh
                If gsARTNR <> "" Then
                    Text3(0).Text = gsARTNR
                    gsARTNR = ""
                End If
            End If
            
            Screen.MousePointer = 0
        
            If Text3(1).Text = "" Then Text3(1).Text = "1"
            
            If Check2.Value = vbUnchecked Then 'bestimmter Lieferant
                'und welcher
                cLinr = Text2.Text
        
                If cLinr = "" Then
                    Screen.MousePointer = 0
                    anzeigeNew "rot", "Bitte wählen Sie einen Lieferanten aus!", Label6
                    Text2.SetFocus
                    Exit Sub
                Else
                    cLinr = Trim$(Str$(Val(cLinr)))
                End If
                
                If Len(Text3(0).Text) = 8 And Left(Text3(0).Text, 1) = "2" Then
                    Text3(0).Text = Mid(Text3(0).Text, 2, 6)
                
                End If
                
                Select Case artikelgefundenA2(Text3(0).Text, cLinr)
                
                    Case 1
                        Screen.MousePointer = 0
                        anzeigeNew "rot", "Artikel gehört nicht zu diesem Lieferant.", Label6
                        Exit Sub
                    
                    Case 2
                        'OK
                    Case 3
                        Screen.MousePointer = 0
                        anzeigeNew "rot", "Artikel gehört nicht zu diesem Lieferant.", Label6
                        Exit Sub
                    Case 4
                        'OK
                    
                    Case Else
                        Screen.MousePointer = 0
                        anzeigeNew "rot", "Dieser Artikel wurde nicht erkannt.", Label6
                        Exit Sub
                End Select
                
            End If
            
            If artikelgefunden(Text3(0).Text, CLng(Text3(1).Text)) Then

                fuellelistlf List3
                Label7(10).Caption = ermVart
                Label7(11).Caption = ermGBart
                If Check9.Value = vbChecked Then
                    Text3(1).Text = "1"
                End If
                Text3(0).Text = ""
                Text3(0).SetFocus
            Else
                If artikelgefundenA1(Text3(0).Text) Then
                    speicherINA1 Text3(0).Text, CLng(Text3(1).Text)
                    fuellelist List3
                    Label7(10).Caption = ermVart
                    Label7(11).Caption = ermGBart
                    If Check9.Value = vbChecked Then
                        Text3(1).Text = "1"
                    End If
                    Text3(0).Text = ""
                    Text3(0).SetFocus
                Else
                    Text3(0).SetFocus
                    anzeigeNew "rot", "Dieser Artikel wurde nicht erkannt.", Label6
                End If
            End If
            
        Case 12 'Druck vorschau aus scanner
        
            Text3(0).SetFocus
            loeschNEW "AINV", gdBase
            CreateTable "AINV", gdBase
            
            sSQL = "Insert into AINV select ARTNR,LINR,BEZEICH,BESTAND,MOPREIS,LEKPR from ARTTOINV"
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update AINV inner join lisrt on AINV.linr = lisrt.linr set AINV.LINBEZ = lisrt.liefbez"
            gdBase.Execute sSQL, dbFailOnError
            
            If Modul6.FindFile(gcDBPfad, "aWKL46bS.rpt") Then
                reportbildschirm "INVENe", "aWKL46bS"
            Else
                reportbildschirm "INVENe", "aWKL46b"
            End If
        Case 22 'Druckvorschau aus mde
            
            loeschNEW "AINV", gdBase
            CreateTable "AINV", gdBase
            
            sSQL = "Insert into AINV select ARTNR,LINR,BEZEICH,BESTAND,MOPREIS,LEKPR from ARTTOINV"
            gdBase.Execute sSQL, dbFailOnError
            

            
            
            
            
            
            sSQL = "Update AINV inner join lisrt on AINV.linr = lisrt.linr set AINV.LINBEZ = lisrt.liefbez"
            gdBase.Execute sSQL, dbFailOnError
            
            If Modul6.FindFile(gcDBPfad, "aWKL46bS.rpt") Then
                reportbildschirm "INVENe", "aWKL46bS"
            Else
                reportbildschirm "INVENe", "aWKL46b"
            End If
            
        Case Is = 2         '** Beenden **
            Unload frmWKL46
            
        Case Is = 3         '** Liste speichern **
            loesch "INV_LITE"
            NewListeFuellAnfangsbuch "INV_", frmWKL46.List2, gdBase
            If MSFlexGrid1.Visible = True And MSFlexGrid1.Rows > 1 Then
                BerechneInventurWKL46
                Text9.Text = ""
                Frame4.Visible = True
                Frame1.Enabled = False
                Command1.Enabled = Frame1.Enabled
                Frame2.Enabled = False
                Frame3.Enabled = False
                Command2(1).Enabled = Frame3.Enabled
                Command2(2).Enabled = Frame3.Enabled
                Command2(3).Enabled = Frame3.Enabled
                Command2(4).Enabled = Frame3.Enabled
                Command2(5).Enabled = Frame3.Enabled
                Text9.SetFocus
            Else
                anzeigeNew "rot", "Es sind keine Daten vorhanden!", Label6
                
            End If
        Case Is = 4         '** Liste laden **
            loesch "INV_LITE"
            NewListeFuellAnfangsbuch "INV_", frmWKL46.List1, gdBase
            
            
            Label1(3).Caption = ""
            Frame5.Visible = True
            Frame1.Enabled = False
            Command1.Enabled = Frame1.Enabled
            Frame2.Enabled = False
            Frame3.Enabled = False
            Command4.Enabled = Frame3.Enabled
            Command2(1).Enabled = Frame3.Enabled
            Command2(2).Enabled = Frame3.Enabled
            Command2(3).Enabled = Frame3.Enabled
            Command2(4).Enabled = Frame3.Enabled
            Command2(5).Enabled = Frame3.Enabled
            List1.SetFocus
            
        Case Is = 5         '** Inventur berechnen **
            BerechneInventurWKL46
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub

Private Sub openschritt2()
    On Error GoTo LOKAL_ERROR
    
    anzeigeNew "normal", "Bitte treffen Sie Ihre Auswahl und drücken dann 'Ausführen'!", Label6
            
    Frame15.Visible = True
    Text4.Text = ""
    List8.Clear
    Option3(0).Value = False
    Option3(1).Value = False
    
    ue = False
    brueck2 = False
    
    Command2(19).Enabled = False
    Command2(18).Enabled = False
    Command2(17).Enabled = False
    Command2(16).Enabled = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "openschritt2"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MDElesen()
    On Error GoTo LOKAL_ERROR
    
    Dim rsMDE As Recordset
    
    If MDEeinlesenOhneLinr(Label6, txtStatus, picprogress, frmWKL46) = False Then
        anzeigeNew "rot", "Es konnten keine Daten aus dem MDE - Gerät ausgelesen werden.", Label6
        anzeigeNew "rot2", "Es konnten keine Daten aus dem MDE - Gerät ausgelesen werden.", Label14
    Else
        anzeigeNew "normal", "", Label6
        If Check8.Value = vbChecked Then
            loeschNEW "ARTTEI", gdBase
            CreateTable "ARTTEI", gdBase
        End If
        MdeVerarbeitungInv
        
        If Check8.Value = vbChecked Then
            reportbildschirm "INVENe", "aWKL46u"
        End If
        
    End If
            
        
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MDElesen"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MdeVerarbeitungInv()
    On Error GoTo LOKAL_ERROR
    
    Dim rsMDE As Recordset
    Dim sEAN As String
    Dim lMenge As Long
    Dim lscanfolge As Long
    
    Screen.MousePointer = 11
    mdeErr = False
    lscanfolge = 0
    
    anzeigeNew "normal", "Die Daten aus dem MDE - Gerät werden verarbeitet...", Label6
    
    Set rsMDE = gdBase.OpenRecordset("mdeinh", dbOpenTable)
    
    If Not rsMDE.EOF Then
        rsMDE.MoveFirst
        
        Do While Not rsMDE.EOF
        
            lscanfolge = lscanfolge + 1
            
            If Not IsNull(rsMDE!eancode) Then
                sEAN = Trim(rsMDE!eancode)
                sEAN = checkean(sEAN)
            Else
                sEAN = ""
            End If
            
            If Not IsNull(rsMDE!Menge) Then
                lMenge = Trim(rsMDE!Menge)
            Else
                lMenge = 0
            End If
            
            If artikelgefunden(sEAN, lMenge) Then

                If Check8.Value = vbChecked Then
                    speicherINTemp sEAN, lMenge, lscanfolge
                End If
                speicherErr sEAN, lMenge, lscanfolge, bezis(sEAN)
            Else
                speicherErr sEAN, lMenge, lscanfolge, "unbekannt"
                mdeErr = True
            End If
        rsMDE.MoveNext
        Loop
        
    End If
    rsMDE.Close: Set rsMDE = Nothing
    
    fuellelist List4
    Label7(22).Caption = ermVart
    Label7(21).Caption = ermGBart
    
    anzeigeNew "normal", "Der Einlesevorgang ist beendet.", Label6
   
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MdeVerarbeitungInv"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function artikelgefunden(sEAN As String, lMenge As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsArt       As Recordset
    Dim cPreis      As String
    Dim lLinr       As Long
    Dim dPreis      As Double
    Dim sTemptext   As String
    Dim cArtNr      As String
    
    Dim sorgEAn As String
    sorgEAn = sEAN
    
    artikelgefunden = False
    sEAN = Trim(sEAN)
    If IsNumeric(sEAN) = False Then
        Exit Function
    End If
    
    If sEAN <> "" Then
    
        If Len(sEAN) >= 13 And Left(sEAN, 3) = "419" Then
            dPreis = Val(Mid(sEAN, 9, 4))
            dPreis = dPreis / 100
            lLinr = glZeitungsLinr ' ermLinrInZeitE
            
            If lLinr > 0 Then
                sEAN = ermartnrausLIBESNR(CStr(Val(Mid(sEAN, 4, 5))), lLinr)
                
                If sEAN = "" Then
                    Exit Function
                Else
                    Text3(0).Text = sEAN
                    
                    cSQL = "Update artikel set ekpr =  '" & dPreis & "'"
                    cSQL = cSQL & " ,Lekpr =  '" & dPreis & "'"
                    cSQL = cSQL & " where artnr = " & sEAN
                    gdBase.Execute cSQL, dbFailOnError
                End If
            End If
        End If
        
        If Len(sEAN) >= 13 And Left(sEAN, 3) = "414" Then
            
            dPreis = Val(Mid(sEAN, 9, 4))
            dPreis = dPreis / 100
            lLinr = glZeitungsLinr ' ermLinrInZeitE
            
            If lLinr > 0 Then
                sEAN = ermartnrausLIBESNR(CStr(Val(Mid(sEAN, 4, 5))), lLinr)
                
                If sEAN = "" Then
                    Exit Function
                Else
                    Text3(0).Text = sEAN
                    
                    cSQL = "Update artikel set ekpr =  '" & dPreis & "'"
                    cSQL = cSQL & " ,Lekpr =  '" & dPreis & "'"
                    cSQL = cSQL & " where artnr = " & sEAN
                    gdBase.Execute cSQL, dbFailOnError
                End If
            End If
        End If
    
        If Len(sEAN) = 11 Then
            sEAN = "0" & sEAN
    
            cSQL = "select * from artikel where ean = '" & sEAN & "'"
            cSQL = cSQL & " or ean2 = '" & sEAN & "'"
            cSQL = cSQL & " or ean3 = '" & sEAN & "'"
        ElseIf Len(sEAN) = 8 Then
        
            If Check10.Value = vbChecked Then
                cSQL = "select * from artikel where ean = '" & sEAN & "'"
                cSQL = cSQL & " or ean2 = '" & sEAN & "'"
                cSQL = cSQL & " or ean3 = '" & sEAN & "'"
            Else
                If Left(sEAN, 1) = "2" Then
                    sEAN = Mid$(sEAN, 2, 6)
                    cSQL = "select * from artikel where artnr = " & sEAN
                Else
                    cSQL = "select * from artikel where ean = '" & sEAN & "'"
                    cSQL = cSQL & " or ean2 = '" & sEAN & "'"
                    cSQL = cSQL & " or ean3 = '" & sEAN & "'"
                End If
            End If
        ElseIf Len(sEAN) <= 6 Then
            cSQL = "select * from artikel where artnr = " & sEAN
            
        Else
            cSQL = "select * from artikel where ean = '" & sEAN & "'"
            cSQL = cSQL & " or ean2 = '" & sEAN & "'"
            cSQL = cSQL & " or ean3 = '" & sEAN & "'"
        End If
        
        Set rsArt = gdBase.OpenRecordset(cSQL)
        If Not rsArt.EOF Then
            artikelgefunden = True
            
            cArtNr = rsArt!artnr
            
                
            sTemptext = "letzter Artikel: " & rsArt!BEZEICH & " Menge: " & lMenge & vbCrLf
            sTemptext = sTemptext & "KVK: " & Format(rsArt!KVKPR1, "###,##0.00") & " EAN: " & sorgEAn
                
                
            anzeigeNew "normal", sTemptext, lblLeArtikel
        
        End If
        rsArt.Close: Set rsArt = Nothing
        
        
        
        If artikelgefunden = False Then
        
            cSQL = "select * from artean_k where ean = '" & sEAN & "'"
            Set rsArt = gdBase.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
                artikelgefunden = True
                
                cArtNr = rsArt!artnr
                
                
                
                    
                sTemptext = "letzter Artikel: " & ermBezeichausWGN(rsArt!artnr) & " Menge: " & lMenge & vbCrLf
                sTemptext = sTemptext & "KVK: " & Format(ermKVKPR1(rsArt!artnr), "###,##0.00") & " EAN: " & sorgEAn
                    
                    
                anzeigeNew "normal", sTemptext, lblLeArtikel
            
            End If
            rsArt.Close: Set rsArt = Nothing
        
        End If
    
    End If
    
    If artikelgefunden = True Then
        speicherIN_neu cArtNr, lMenge, sorgEAn
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "artikelgefunden"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function

Private Function artikelgefundenA1(sArt As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsArt       As Recordset
    
    artikelgefundenA1 = False
    sArt = Trim(sArt)
    
    If sArt <> "" Then
    
    
    
        If Len(sArt) = 8 Then
            If Left(sArt, 1) = "2" Then
                sArt = Mid(sArt, 2, 6)
            ElseIf Left(sArt, 1) = "0" Then
                sArt = Mid(sArt, 2, 6)
            Else
                sArt = ""
            End If
        Else
'            sart = ""
        End If
    
        If Len(sArt) < 7 And IsNumeric(sArt) Then
            cSQL = "select * from artikel where artnr = " & sArt
            
            Set rsArt = gdBase.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
                Text3(0).Text = sArt
                artikelgefundenA1 = True
                
                Dim sTemptext As String
                
                sTemptext = "letzter Artikel: " & rsArt!BEZEICH & " Menge: " & Text3(1).Text & vbCrLf
                sTemptext = sTemptext & "KVK: " & Format(rsArt!KVKPR1, "###,##0.00")
                
                
                anzeigeNew "normal", sTemptext, lblLeArtikel
            End If
            rsArt.Close: Set rsArt = Nothing
        End If
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "artikelgefundenA1"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Private Function artikelgefundenA2(sEAN As String, cLinr As String) As Byte
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsArt       As Recordset
    
    artikelgefundenA2 = 9
    sEAN = Trim(sEAN)
    cLinr = Trim(cLinr)
    
    If sEAN <> "" Then
    
    
    
    
    
    
    
        cSQL = "select * from artikel where  ean = '" & sEAN & "'"
        cSQL = cSQL & " or ean2 = '" & sEAN & "'"
        cSQL = cSQL & " or ean3 = '" & sEAN & "'"
        
        
        Set rsArt = gdBase.OpenRecordset(cSQL)
        
        If Not rsArt.EOF Then
            artikelgefundenA2 = 1
        Else
        
'            Dim sEanSeek As String
'            If Len(sEAN) = 8 And Left(sEAN, 1) = "2" Then
'                sEanSeek = Mid(sEAN, 2, 6)
'            Else
'                sEanSeek = sEAN
'            End If
            
            cSQL = "select * from artikel where  artnr = " & sEAN
            rsArt.Close: Set rsArt = Nothing
            Set rsArt = gdBase.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
                artikelgefundenA2 = 3
            End If
            
        End If
        
        If artikelgefundenA2 = 3 Then
           
            cSQL = "select * from artikel where  artnr = " & sEAN
            cSQL = cSQL & "  and  linr = " & cLinr
            rsArt.Close: Set rsArt = Nothing
            Set rsArt = gdBase.OpenRecordset(cSQL)
            
            If Not rsArt.EOF Then
                artikelgefundenA2 = 4
            Else
                
            
            End If
        
        ElseIf artikelgefundenA2 = 1 Then
    
            cSQL = "select * from artikel where ( ean = '" & sEAN & "'"
            cSQL = cSQL & " or ean2 = '" & sEAN & "'"
            cSQL = cSQL & " or ean3 = '" & sEAN & "'"
            cSQL = cSQL & " ) and  linr = " & cLinr
            rsArt.Close: Set rsArt = Nothing
            Set rsArt = gdBase.OpenRecordset(cSQL)
            
            If Not rsArt.EOF Then
                artikelgefundenA2 = 2
            Else
                
            
            End If
        
        End If
        rsArt.Close: Set rsArt = Nothing
    
    End If
    
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "artikelgefundenA2"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermVart() As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsArt       As Recordset
    Dim lvanz       As Long
    
    ermVart = "0"
    
    cSQL = "select artnr from arttoinv group by artnr"
    Set rsArt = gdBase.OpenRecordset(cSQL)
    If Not rsArt.EOF Then
        rsArt.MoveLast
        lvanz = rsArt.RecordCount
        ermVart = CStr(lvanz)
    End If
    rsArt.Close: Set rsArt = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermVart"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermGBart() As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsArt       As Recordset

    ermGBart = "0"
    
    cSQL = "select sum(bestand) as lvanz from arttoinv "
    Set rsArt = gdBase.OpenRecordset(cSQL)
    If Not rsArt.EOF Then
        If Not IsNull(rsArt!lvanz) Then
            ermGBart = rsArt!lvanz
        End If
    End If
    rsArt.Close: Set rsArt = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermGBart"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub speicherIN_neu(sArtnr As String, lMenge As Long, sorgEAn As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    
    sArtnr = Trim(sArtnr)
    
    If sArtnr <> "" Then
    
        cSQL = "Insert into ARTTOINV select artikel.artnr "
        cSQL = cSQL & ", artikel.BEZEICH "
        cSQL = cSQL & ", artikel.LINR "
        cSQL = cSQL & ", artikel.LPZ "
        cSQL = cSQL & ", artikel.AGN "
        
        If Option1(7).Value = True Then
            cSQL = cSQL & ", artikel.EKPR as LEKPR "
            cSQL = cSQL & ", (artikel.EKPR * " & lMenge & ") as MOPREIS "
        ElseIf Option1(4).Value = True Then
        
            cSQL = cSQL & ", artlief.LEKPR "
            cSQL = cSQL & ", (artlief.LEKPR * " & lMenge & ") as MOPREIS "
        End If
        
        cSQL = cSQL & ", artikel.KVKPR1 "
        cSQL = cSQL & ", artikel.VKPR "
        cSQL = cSQL & ", " & lMenge & " as Bestand "
        cSQL = cSQL & ", artlief.LIBESNR "
        cSQL = cSQL & ", '" & DateValue(Now) & "' as LASTDATE "
        cSQL = cSQL & ", '" & TimeValue(Now) & "' as LASTTIME "
        cSQL = cSQL & ", '" & sorgEAn & "' as EAN "
        cSQL = cSQL & " from artikel inner join artlief on artikel.artnr = artlief.artnr "
        cSQL = cSQL & " and artikel.linr = artlief.linr "
        
        cSQL = cSQL & " where artikel.artnr = " & sArtnr
        gdBase.Execute cSQL, dbFailOnError
        
        If Option1(4).Value = True Then
            Dim dLEK As Double
            If opt1(22).Value = True Then
                dLEK = ermLEKPRundWelcher(sArtnr, "MAX")
            Else
                dLEK = ermLEKPRundWelcher(sArtnr, "MIN")
            End If
            
            cSQL = "update ARTTOINV set LEKPR = '" & dLEK & "' where artnr = " & sArtnr
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "update ARTTOINV set MOPREIS = LEKPR * BESTAND where artnr = " & sArtnr
            gdBase.Execute cSQL, dbFailOnError
        End If
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherIN_neu"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherINTemp(sEAN As String, lMenge As Long, lLFNR As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim lRegal      As Long
    Dim iBed        As Integer
    Dim sScancode   As String
    
    sScancode = sEAN
    
    lRegal = Val(Text85.Text)
    iBed = Val(Text87.Text)
    sEAN = Trim(sEAN)
    
    If sEAN <> "" Then
    
        cSQL = "Insert into ARTTEI select artnr "
        cSQL = cSQL & ", BEZEICH "
        cSQL = cSQL & ", LINR "
        cSQL = cSQL & ", LPZ "
        cSQL = cSQL & ", AGN "
        cSQL = cSQL & ", " & lRegal & " as Regal "
        cSQL = cSQL & ", " & iBed & " as Bed "
        
        cSQL = cSQL & ", " & lLFNR & " as LFNR"
        If Option1(7).Value = True Then
            cSQL = cSQL & ", EKPR as LEKPR "
            cSQL = cSQL & ", (EKPR * Bestand) as MOPREIS "
        ElseIf Option1(4).Value = True Then
            cSQL = cSQL & ", LEKPR "
            cSQL = cSQL & ", (LEKPR * Bestand) as MOPREIS "
        End If
        
        
        cSQL = cSQL & ", KVKPR1 "
        cSQL = cSQL & ", VKPR "
        cSQL = cSQL & ", " & lMenge & " as Bestand "
        cSQL = cSQL & ", LIBESNR "
        cSQL = cSQL & ", '" & DateValue(Now) & "' as LASTDATE "
        cSQL = cSQL & ", '" & TimeValue(Now) & "' as LASTTIME "
        
        cSQL = cSQL & ", '" & sScancode & "' as scancode "
        
        
        If Len(sEAN) = 11 Then
            sEAN = "0" & sEAN
    
            cSQL = cSQL & " from artikel where ean = '" & sEAN & "'"
            cSQL = cSQL & " or ean2 = '" & sEAN & "'"
            cSQL = cSQL & " or ean3 = '" & sEAN & "'"
        ElseIf Len(sEAN) = 8 Then
            If Check10.Value = vbChecked Then
                cSQL = cSQL & " from artikel where ean = '" & sEAN & "'"
                cSQL = cSQL & " or ean2 = '" & sEAN & "'"
                cSQL = cSQL & " or ean3 = '" & sEAN & "'"
            Else
                If Left(sEAN, 1) = "2" Then
                    sEAN = Mid$(sEAN, 2, 6)
                    cSQL = cSQL & " from artikel where artnr = " & sEAN
                Else
                    cSQL = cSQL & " from artikel where ean = '" & sEAN & "'"
                    cSQL = cSQL & " or ean2 = '" & sEAN & "'"
                    cSQL = cSQL & " or ean3 = '" & sEAN & "'"
                End If
            End If
        ElseIf Len(sEAN) <= 6 Then
            
            cSQL = cSQL & " from artikel where artnr = " & sEAN
        Else
            cSQL = cSQL & " from artikel where ean = '" & sEAN & "'"
            cSQL = cSQL & " or ean2 = '" & sEAN & "'"
            cSQL = cSQL & " or ean3 = '" & sEAN & "'"
        End If
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherINTemp"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherINA1(sArt As String, lMenge As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    
    sArt = Trim(sArt)
    
    If sArt <> "" Then
        cSQL = "Insert into ARTTOINV select artnr "
        cSQL = cSQL & ", BEZEICH "
        cSQL = cSQL & ", LINR "
        cSQL = cSQL & ", LPZ "
        cSQL = cSQL & ", AGN "
        
        If Option1(7).Value = True Then
            cSQL = cSQL & ", EKPR as LEKPR "
            cSQL = cSQL & ", (EKPR * Bestand) as MOPREIS "
        ElseIf Option1(4).Value = True Then
            cSQL = cSQL & ", LEKPR "
            cSQL = cSQL & ", (LEKPR * Bestand) as MOPREIS "
        End If
        
        cSQL = cSQL & ", KVKPR1 "
        cSQL = cSQL & ", VKPR "
        cSQL = cSQL & ", " & lMenge & " as Bestand "
        cSQL = cSQL & ", LIBESNR "
        cSQL = cSQL & ", '" & DateValue(Now) & "' as LASTDATE "
        cSQL = cSQL & ", '" & TimeValue(Now) & "' as LASTTIME "
        cSQL = cSQL & " from artikel where artnr = " & sArt
        gdBase.Execute cSQL, dbFailOnError
        
        If Option1(4).Value = True Then
            Dim dLEK As Double
            If opt1(22).Value = True Then
                dLEK = ermLEKPRundWelcher(sArt, "MAX")
            Else
                dLEK = ermLEKPRundWelcher(sArt, "MIN")
            End If
            
            cSQL = "update ARTTOINV set LEKPR = '" & dLEK & "' where artnr = " & sArt
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "update ARTTOINV set MOPREIS = LEKPR * BESTAND where artnr = " & sArt
            gdBase.Execute cSQL, dbFailOnError
        End If
        
        
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherINA1"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherErr(sEAN As String, lMenge As Long, lLFNR As Long, cErrart As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    
    sEAN = Trim(sEAN)
    cErrart = SwapStr(cErrart, "'", "")
    
    
    If sEAN <> "" Then
        cSQL = "Insert into artErrIn (ean,menge,lfnr,errArt) values  "
        cSQL = cSQL & " ( '" & sEAN & "'," & lMenge & "," & lLFNR & ",'" & cErrart & "') "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherErr"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub scannerbackanz()
    On Error GoTo LOKAL_ERROR
    
    Label7(10).Caption = "0"
    Label7(11).Caption = "0"

    List3.Clear
    
    Text3(0).SetFocus
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "scannerbackanz"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MDEbackanz()
    On Error GoTo LOKAL_ERROR
    
    Label7(22).Caption = "0"
    Label7(21).Caption = "0"

    List4.Clear
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MDEbackanz"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Zeigeauswahlframe()
    On Error GoTo LOKAL_ERROR
    
    Frame6.Visible = False
    
    If Option2(0).Value = True Then         'Inventurberechnung
        vorbereitungInventberechnung
        
    ElseIf Option2(1).Value = True Then     'Scanner
        vorbereitungScanner

    ElseIf Option2(2).Value = True Then     'Mde
        vorbereitungMDE
        
    ElseIf Option2(3).Value = True Then         'Inventurliste
        vorbereitungInventurliste
        
    ElseIf Option2(4).Value = True Then     'Inventurliste erzeugen
        vorbereitungInventurlisteErzeugen

    ElseIf Option2(5).Value = True Then     'Inventureinstellungen
        vorbereitungInventureinstellungen
    ElseIf Option2(6).Value = True Then     'Import Excel
        vorbereitungImport
    ElseIf Option2(7).Value = True Then     'Export Excel
        vorbereitungExport
    ElseIf Option2(8).Value = True Then     'Export Excel
        vorbereitungBestHist
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zeigeauswahlframe"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub BestandaufNull()
    On Error GoTo LOKAL_ERROR
    
    Dim cLinr   As String
    Dim cSQL    As String
    
    Dim llpzvon As Long
    Dim llpzbis As Long
    Dim lcount As Long
    
    '1.option values?
    
    
    
    If Option3(0).Value Then
        cLinr = "alle"
    ElseIf Option3(1).Value Then
        cLinr = Text4.Text
        
        If cLinr = "" Then
            anzeigeNew "rot", "Bitte wählen Sie einen Lieferanten aus!", Label6
            Text2.SetFocus
            Exit Sub
        Else
            cLinr = Trim$(Str$(Val(cLinr)))
        End If
    ElseIf Option3(2).Value Then
        cLinr = "alle"
        llpzvon = Val(Text6.Text)
        llpzbis = Val(Text5.Text)
    
        If llpzbis = 0 Then
            anzeigeNew "rot", "Bitte geben Sie einen Lagerplatzbereich an", Label6
            Text5.SetFocus
            Exit Sub
        End If
        
        If llpzvon = 0 Then
            anzeigeNew "rot", "Bitte geben Sie einen Lagerplatzbereich an", Label6
            Text6.SetFocus
            Exit Sub
        End If
    Else
        anzeigeNew "rot", "Bitte treffen Sie eine Auswahl!", Label6
        Exit Sub
    End If
    
    anzeigeNew "normal", "Bestände werden auf 0 gesetzt...", Label6
    
    Label7(19).Caption = cLinr
    Label7(19).Refresh
    
    Label7(26).Caption = llpzvon
    Label7(26).Refresh
    
    Label7(27).Caption = llpzbis
    Label7(27).Refresh
    
    Command2(15).Enabled = False
    Command2(2).Enabled = False
    
    'Is Okay
    'sicher der Gesamten Artikel in Btikel
    Screen.MousePointer = 11
    

    loeschNEW "Btikel", gdBase
    cSQL = "Select Artikel.* into Btikel from Artikel "
    If cLinr = "alle" Then
    
    Else
    
    
        cSQL = cSQL & "  inner join artlief on "
        cSQL = cSQL & " artikel.artnr = artlief.artnr "
        cSQL = cSQL & " Where artlief.linr = " & cLinr
   
    
    
    
    
'        cSQL = cSQL & " Where LINR = " & cLinr
        
        If List8.ListCount <> 0 Then
            cSQL = cSQL & " and (Artikel.LPZ = " & Mid(List8.list(0), 1, InStr(1, List8.list(0), " ")) & " "
            For lcount = 1 To List8.ListCount - 1
                cSQL = cSQL & " or Artikel.LPZ = " & Mid(List8.list(lcount), 1, InStr(1, List8.list(lcount), " ")) & " "
            Next lcount
            cSQL = cSQL & ")"
        End If
    End If
    gdBase.Execute cSQL, dbFailOnError
    
    anzeigeNew "normal", "Bestände werden auf 0 gesetzt......", Label6
    
    If Option3(2).Value Then
        cSQL = "Update Artikel inner join lagerplatz on"
        cSQL = cSQL & " artikel.artnr = lagerplatz.artnr set artikel.Bestand = 0 "
        cSQL = cSQL & " where lagerplatz.lagerp between " & llpzvon
        cSQL = cSQL & " and " & llpzbis
    Else
    
        
    
        
        If cLinr = "alle" Then
            cSQL = "Update Artikel set Bestand = 0 "
        Else
            cSQL = "Update Artikel inner join Artlief on Artikel.Artnr = Artlief.Artnr set Artikel.bestand = 0 "
            cSQL = cSQL & " Where Artlief.LINR = " & cLinr
            
            If List8.ListCount <> 0 Then

                cSQL = cSQL & " and (Artikel.LPZ = " & Mid(List8.list(0), 1, InStr(1, List8.list(0), " ")) & " "
                For lcount = 1 To List8.ListCount - 1
                    cSQL = cSQL & " or Artikel.LPZ = " & Mid(List8.list(lcount), 1, InStr(1, List8.list(lcount), " ")) & " "
                Next lcount
                cSQL = cSQL & ")"
            
            End If
        End If
    End If
    gdBase.Execute cSQL, dbFailOnError
    
    'jetzt ist rückgängig aktiv
    Command2(17).Enabled = True
    Command2(16).Enabled = False
    Command2(18).Enabled = True

    anzeigeNew "normal", "Bestände sind auf 0 gesetzt. Drücken Sie jetzt 'Übernahme'!", Label6
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BestandaufNull"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub rueck1()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim cLinr    As String
    Dim lcount As Long
    
    cLinr = Label7(19).Caption
    
    
    Screen.MousePointer = 11
    
    Command2(18).Enabled = False
    
    anzeigeNew "normal", "alte Bestände werden wiederhergestellt...", Label6
    loeschNEW "Ctikel", gdBase
    
    cSQL = "Select * into Ctikel from Btikel "
    gdBase.Execute cSQL, dbFailOnError
    
    If cLinr = "alle" Then
        loeschNEW "Artikel", gdBase
        cSQL = "Select * into Artikel from Btikel"
        gdBase.Execute cSQL, dbFailOnError
    Else
    
        cSQL = "Delete from Artikel "
        cSQL = cSQL & " Where artikel.LINR = " & cLinr
        
        
        
        If List8.ListCount <> 0 Then
            cSQL = cSQL & " and (Artikel.LPZ = " & Mid(List8.list(0), 1, InStr(1, List8.list(0), " ")) & " "
            For lcount = 1 To List8.ListCount - 1
                cSQL = cSQL & " or Artikel.LPZ = " & Mid(List8.list(lcount), 1, InStr(1, List8.list(lcount), " ")) & " "
            Next lcount
            cSQL = cSQL & ")"

            
        
        End If
        
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Insert into artikel Select *  from Btikel"
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    anzeigeNew "normal", "Artikelstrukturen werden wiederhergestellt...", Label6
    IndexArtikel Label6
    
    'jetzt ist rückgängig1 inaktiv
    Command2(17).Enabled = False
    Command2(16).Enabled = True
    Command2(15).Enabled = True
    Command2(2).Enabled = True
    
    
    Screen.MousePointer = 11
    anzeigeNew "normal", "Bitte treffen Sie Ihre Auswahl und drücken dann 'Ausführen'!", Label6

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "rueck1"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub rueck2()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    
    Screen.MousePointer = 11
    brueck2 = True
    Command2(19).Enabled = False
    Command2(18).Enabled = False
    Command2(17).Enabled = False
    Command2(16).Enabled = True
    
    cSQL = "Insert into ARTTOINV select * from BTOINV"
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "BTOINV", gdBase
    
    Screen.MousePointer = 11

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "rueck2"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub UebernahmeInventur()
    On Error GoTo LOKAL_ERROR
    
    Dim cLinr   As String
    Dim cSQL    As String
    Dim lcount As Long
    
    Command2(17).Enabled = False
    cLinr = Label7(19).Caption
    
    If cLinr = "" Then
        anzeigeNew "rot", "Bearbeiten Sie bitte erst Schritt 2!", Label6
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    ue = True
    '1.Update Artikel mit ARTTOINV _ bestand
    loeschNEW "Artkum", gdBase
    
    If cLinr = "alle" Then
        cSQL = "Update ARTTOINV inner join Artikel on ARTTOINV.Artnr = Artikel.Artnr "
        cSQL = cSQL & " set ARTTOINV.LINR = Artikel.LINR "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    cSQL = "Select Artnr,LINR,lpz, sum(bestand)as bestheut into Artkum from ARTTOINV group by artnr,linr ,lpz"
    gdBase.Execute cSQL, dbFailOnError
    
    
    If cLinr = "alle" Then
    
    
        cSQL = "Update Artikel inner join Artkum on Artikel.Artnr = Artkum.Artnr "
        cSQL = cSQL & " set Artikel.Bestand = Artkum.bestheut "
    
    Else
    
        cSQL = "Update Artikel inner join Artkum on Artikel.Artnr = Artkum.Artnr "
        cSQL = cSQL & " set Artikel.Bestand = Artkum.bestheut "
    
        cSQL = cSQL & " Where  Artikel.artnr in (Select artnr from artlief where LINR = " & cLinr & ")"
    
'        cSQL = cSQL & " Where  artikel.LINR = " & cLinr
        
        If List8.ListCount <> 0 Then
            cSQL = cSQL & " and (Artikel.LPZ = " & Mid(List8.list(0), 1, InStr(1, List8.list(0), " ")) & " "
            For lcount = 1 To List8.ListCount - 1
                cSQL = cSQL & " or Artikel.LPZ = " & Mid(List8.list(lcount), 1, InStr(1, List8.list(lcount), " ")) & " "
            Next lcount
            cSQL = cSQL & ")"
        End If
    End If
    gdBase.Execute cSQL, dbFailOnError
 
    loeschNEW "BTOINV", gdBase
    
    
    
    If cLinr = "alle" Then
        cSQL = "Select * into BTOINV from ARTTOINV "
    Else
        cSQL = "Select ARTTOINV.* into BTOINV from ARTTOINV "
        
        
        cSQL = cSQL & "  inner join artlief on "
        cSQL = cSQL & " ARTTOINV.artnr = artlief.artnr "
        cSQL = cSQL & " Where artlief.linr = " & cLinr
        
        
'        cSQL = cSQL & " Where ARTTOINV.LINR = " & cLinr
        If List8.ListCount <> 0 Then
            cSQL = cSQL & " and (ARTTOINV.LPZ = " & Mid(List8.list(0), 1, InStr(1, List8.list(0), " ")) & " "
            For lcount = 1 To List8.ListCount - 1
                cSQL = cSQL & " or ARTTOINV.LPZ = " & Mid(List8.list(lcount), 1, InStr(1, List8.list(lcount), " ")) & " "
            Next lcount
            cSQL = cSQL & ")"
        End If
    End If
    gdBase.Execute cSQL, dbFailOnError

    
    If cLinr = "alle" Then
        cSQL = "Delete from ARTTOINV  "
    Else
        cSQL = "Delete from ARTTOINV  "
        cSQL = cSQL & " Where  artnr in (Select artnr from artlief where LINR = " & cLinr & ") "
'        cSQL = cSQL & " Where ARTTOINV.LINR = " & cLinr
        If List8.ListCount <> 0 Then
            cSQL = cSQL & " and (ARTTOINV.LPZ = " & Mid(List8.list(0), 1, InStr(1, List8.list(0), " ")) & " "
            For lcount = 1 To List8.ListCount - 1
                cSQL = cSQL & " or ARTTOINV.LPZ = " & Mid(List8.list(lcount), 1, InStr(1, List8.list(lcount), " ")) & " "
            Next lcount
            cSQL = cSQL & ")"
        End If
    End If
    gdBase.Execute cSQL, dbFailOnError
    
    anzeigeNew "normal", "Übernahme der Bestände ...", Label6
    
    loeschNEW "DIFFTA", gdBase
    CreateTable "DIFFTA", gdBase
    
    cSQL = "Insert into DIFFTA Select Artnr,BEZEICH,LINR,lpz,bestand as bestsys, 0 as bestheut "
    cSQL = cSQL & ", kvkpr1 "
    cSQL = cSQL & ", 0.0 as LWEKheut "
    cSQL = cSQL & ", 0 as diffbest "
    cSQL = cSQL & ", 0.0 as diffLWEK "
    cSQL = cSQL & ", vkpr "
    
    If Option1(7).Value = True Then
        cSQL = cSQL & ", EKPR as LEKPR "
        cSQL = cSQL & ", (EKPR * Bestand) as LWEKSYS "
    ElseIf Option1(4).Value = True Then
        cSQL = cSQL & ", LEKPR "
        cSQL = cSQL & ", (LEKPR * Bestand) as LWEKSYS "
    End If


    cSQL = cSQL & "  from BTIKEL "
    If cLinr = "alle" Then
    Else
        cSQL = cSQL & " Where BTIKEL.LINR = " & cLinr
        If List8.ListCount <> 0 Then
            cSQL = cSQL & " and (BTIKEL.LPZ = " & Mid(List8.list(0), 1, InStr(1, List8.list(0), " ")) & " "
            For lcount = 1 To List8.ListCount - 1
                cSQL = cSQL & " or BTIKEL.LPZ = " & Mid(List8.list(lcount), 1, InStr(1, List8.list(lcount), " ")) & " "
            Next lcount
            cSQL = cSQL & ")"
        End If
    End If
    gdBase.Execute cSQL, dbFailOnError
    
    anzeigeNew "normal", "Übernahme der Bestände ......", Label6
    
    cSQL = "Update Diffta inner join Artkum on Diffta.Artnr = Artkum.Artnr "
    cSQL = cSQL & " set Diffta.bestheut = Artkum.bestheut "
    If cLinr = "alle" Then
    
    Else
        cSQL = cSQL & " Where Diffta.LINR = " & cLinr
        
        If List8.ListCount <> 0 Then
            cSQL = cSQL & " and (Diffta.LPZ = " & Mid(List8.list(0), 1, InStr(1, List8.list(0), " ")) & " "
            For lcount = 1 To List8.ListCount - 1
                cSQL = cSQL & " or Diffta.LPZ = " & Mid(List8.list(lcount), 1, InStr(1, List8.list(lcount), " ")) & " "
            Next lcount
            cSQL = cSQL & ")"
        End If
          
    End If
    gdBase.Execute cSQL, dbFailOnError
    
    anzeigeNew "normal", "Lagerwerte werden ermittelt...", Label6
    
    cSQL = "Update Diffta set LWEKheut =  lEKPR * bestheut "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeigeNew "normal", "Lagerwerte werden ermittelt......", Label6
    
    cSQL = "Delete from Diffta where bestheut = bestsys "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeigeNew "normal", "Lagerwerte werden ermittelt.........", Label6
    
    cSQL = "Update Diffta set diffbest = bestheut - bestsys "
    gdBase.Execute cSQL, dbFailOnError
    
    anzeigeNew "normal", "Lagerwerte werden ermittelt......", Label6
    
    cSQL = "Update Diffta set diffLWEK = LWEKheut - LWEKsys  "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "AINV", gdBase
    CreateTable "AINV", gdBase
    
    cSQL = "Insert into AINV select ARTNR,LINR,BEZEICH,BESTAND,MOPREIS,LEKPR from BTOINV"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update AINV inner join lisrt on AINV.linr = lisrt.linr set AINV.LINBEZ = lisrt.liefbez"
    gdBase.Execute cSQL, dbFailOnError
    anzeigeNew "normal", "Protokolle werden erstellt...", Label6
    
    
    If Modul6.FindFile(gcDBPfad, "aWKL46cS.rpt") Then
        reportbildschirm "INVENe", "aWKL46cS"
    Else
        reportbildschirm "INVENe", "aWKL46c"
    End If
    
    If Modul6.FindFile(gcDBPfad, "aWKL46dS.rpt") Then
    
        loeschNEW "DIFFDRUCK", gdBase

        cSQL = "Select * into DIFFDRUCK from diffta order by bezeich "
        gdBase.Execute cSQL, dbFailOnError
        
        If Not SpalteInTabellegefundenNEW("DIFFDRUCK", "liefbez", gdBase) Then
            SpalteAnfuegenNEW "DIFFDRUCK", "liefbez", "Text(35)", gdBase
        
            cSQL = "Update DIFFDRUCK inner join LISRT on DIFFDRUCK.linr = lisrt.linr "
            cSQL = cSQL & " set DIFFDRUCK.liefbez = LISRT.liefbez "
            gdBase.Execute cSQL, dbFailOnError
               
        End If
        
        If Not SpalteInTabellegefundenNEW("DIFFDRUCK", "EAN", gdBase) Then
            SpalteAnfuegenNEW "DIFFDRUCK", "EAN", "Text(13)", gdBase
        
            cSQL = "Update DIFFDRUCK inner join Artikel on DIFFDRUCK.artnr = Artikel.artnr "
            cSQL = cSQL & " set DIFFDRUCK.EAN = Artikel.ean "
            gdBase.Execute cSQL, dbFailOnError
        End If
        
        
        
        reportbildschirm "INVENe", "aWKL46dS"
    Else
        loeschNEW "DIFFDRUCK", gdBase

        cSQL = "Select * into DIFFDRUCK from diffta order by bezeich "
        gdBase.Execute cSQL, dbFailOnError
        
        If Not SpalteInTabellegefundenNEW("DIFFDRUCK", "liefbez", gdBase) Then
            SpalteAnfuegenNEW "DIFFDRUCK", "liefbez", "Text(35)", gdBase
        
            cSQL = "Update DIFFDRUCK inner join LISRT on DIFFDRUCK.linr = lisrt.linr "
            cSQL = cSQL & " set DIFFDRUCK.liefbez = LISRT.liefbez "
            gdBase.Execute cSQL, dbFailOnError
               
        End If
        reportbildschirm "INVENe", "aWKL46d"
    End If
    
'    reportbildschirm "", "aWKL46c" 'übernahme
'    reportbildschirm "", "aWKL46d" 'Differenz
    
    anzeigeNew "normal", "Die Übernahme ist beendet.", Label6
    
    Screen.MousePointer = 0
    
    'jetzt ist rückgängig aktiv
    Command2(18).Enabled = False
    Command2(19).Enabled = True
    Command2(15).Enabled = True
    Command2(2).Enabled = True

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UebernahmeInventur"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub vorbereitungInventberechnung()
    On Error GoTo LOKAL_ERROR
    
        Dim i As Integer
        
        Frame7.Visible = True
        Frame1.Visible = True
        Frame0.Visible = False
        Frame2.Visible = False
        
        For i = 12 To 15
            Command0(i).BackColor = vbWhite
            Command0(i).HoverColorFrom = vbWhite
            Command0(i).HoverColorTo = vbWhite
        Next i
        
'        Command1.BackColor = vbWhite
'        Command5.BackColor = vbWhite
        
        invLiteRefresh

        Text2.SetFocus
        
        FormatiereMSFlexGrid1WKL46
        
        anzeigeNew "normal", "", Label6

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungInventberechnung"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub vorbereitungScanner()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Frame9.Visible = True
    Text3(0).SetFocus
    Text3(1).Text = "1"
    anzeigeNew "normal", "", Label6
    iscan = 3
    
    If Not NewTableSuchenDBKombi("ARTTOINV", gdBase) Then
        cSQL = "Select * into ARTTOINV from Artikel where artnr = -1"
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError

        cSQL = "Delete from ARTTOINV "
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
        
        SpalteAnfuegenNEW "ARTTOINV", "lfnr", "autoincrement", gdBase
        
    Else
        Label7(10).Caption = ermVart
        Label7(11).Caption = ermGBart
        
        fuellelist List3
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungScanner"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellelist(lst As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim cFeld   As String
    Dim cLBSatz As String
    Dim lCounter As Long
    
    cSQL = "Select * from ARTTOINV order by lfnr desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    lst.Visible = False
    lst.Clear
    
    If Not rsrs.EOF Then
    
        lCounter = rsrs.RecordCount
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            End If
    
            cLBSatz = cFeld & Space(7 - Len(cFeld))
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If

            cLBSatz = cLBSatz & cFeld & Space(36 - Len(cFeld))

            If Not IsNull(rsrs!BESTAND) Then
                cFeld = rsrs!BESTAND
            Else
                cFeld = ""
            End If

            cLBSatz = cLBSatz & cFeld & Space(6 - Len(cFeld))

            If Not IsNull(rsrs!LASTDATE) Then
                cFeld = rsrs!LASTDATE
            Else
                cFeld = ""
            End If

            cLBSatz = cLBSatz & cFeld & Space(10 - Len(cFeld))

            If Not IsNull(rsrs!LASTTIME) Then
                cFeld = rsrs!LASTTIME
            Else
                cFeld = ""
            End If

            cLBSatz = cLBSatz & cFeld & Space(10 - Len(cFeld))

            If Not IsNull(rsrs!EAN) Then
                cFeld = rsrs!EAN
            Else
                cFeld = ""
            End If

            cLBSatz = cLBSatz & cFeld
            
            lst.AddItem cLBSatz
            
            lCounter = lCounter - 1
            anzeigeNew "normal", "noch " & lCounter & " Artikel", Label6
            
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
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub fuellelistlf(lst As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim cFeld   As String
    Dim cLBSatz As String
    Dim lMax    As Long
    
    cSQL = "Select max(lfnr) as maxi from ARTTOINV "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            lMax = rsrs!maxi
        Else
            lMax = 0
        End If
    Else
        lMax = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    cSQL = "Select * from ARTTOINV where lfnr =" & lMax
    Set rsrs = gdBase.OpenRecordset(cSQL)
    

    If Not rsrs.EOF Then
        

        If Not IsNull(rsrs!artnr) Then
            cFeld = rsrs!artnr
        End If

        cLBSatz = cFeld & Space(7 - Len(cFeld))
        
        If Not IsNull(rsrs!BEZEICH) Then
            cFeld = rsrs!BEZEICH
        Else
            cFeld = ""
        End If
        
        cLBSatz = cLBSatz & cFeld & Space(36 - Len(cFeld))
        
        If Not IsNull(rsrs!BESTAND) Then
            cFeld = rsrs!BESTAND
        Else
            cFeld = ""
        End If
        
        cLBSatz = cLBSatz & cFeld & Space(6 - Len(cFeld))
        
        If Not IsNull(rsrs!LASTDATE) Then
            cFeld = rsrs!LASTDATE
        Else
            cFeld = ""
        End If
        
        cLBSatz = cLBSatz & cFeld & Space(10 - Len(cFeld))
        
        
        If Not IsNull(rsrs!LASTTIME) Then
            cFeld = rsrs!LASTTIME
        Else
            cFeld = ""
        End If
        
        cLBSatz = cLBSatz & cFeld & Space(10 - Len(cFeld))
        
        If Not IsNull(rsrs!EAN) Then
            cFeld = rsrs!EAN
        Else
            cFeld = ""
        End If
        
        cLBSatz = cLBSatz & cFeld
        
        
        lst.AddItem cLBSatz, 0
            

    End If
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellelistlf"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub vorbereitungMDE()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Frame10.Visible = True
    anzeigeNew "normal", "", Label6
    
    loeschNEW "ARTERRIN", gdBase
    CreateTable "ARTERRIN", gdBase
    
    iscan = 4
    
    If Not NewTableSuchenDBKombi("ARTTOINV", gdBase) Then
        cSQL = "Select * into ARTTOINV from Artikel where artnr = -1"
        gdBase.Execute cSQL, dbFailOnError

        cSQL = "Delete from ARTTOINV "
        gdBase.Execute cSQL, dbFailOnError
        
        SpalteAnfuegenNEW "ARTTOINV", "lfnr", "autoincrement", gdBase
        
    Else
        Label7(22).Caption = ermVart
        Label7(21).Caption = ermGBart
        
        fuellelist List4
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungMDE"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub vorbereitungInventurliste()
    On Error GoTo LOKAL_ERROR
    
    Frame11.Visible = True
    anzeigeNew "normal", "", Label6
    
    Command6(6).Enabled = False
    
    New2ListeFuellAnfangsbuch "ILI", frmWKL46.List6, gdBase
            
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungInventurliste"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub vorbereitungInventurlisteErzeugen()
    On Error GoTo LOKAL_ERROR
    
    
    LeereDialogWKL46
    Text7.Text = ""
    LeseLieferanten Combo2, ""
    List5.Clear
    List5.Visible = False
    loeschNEW "li46", gdBase
    Frame12.Visible = True
    
    If NewTableSuchenDBKombi("E46IL", gdBase) Then
        voreinstellungladenE46IL
    End If
    
    anzeigeNew "normal", "", Label6
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungInventurlisteErzeugen"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub vorbereitungInventureinstellungen()
    On Error GoTo LOKAL_ERROR
    
    Frame8.Visible = True
    anzeigeNew "normal", "", Label6
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungInventureinstellungen"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub vorbereitungImport()
    On Error GoTo LOKAL_ERROR
    
    Frame19.Visible = True
    anzeigeNew "normal", "", Label6
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungImport"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub vorbereitungExport()
    On Error GoTo LOKAL_ERROR
    
    Frame21.Visible = True
    anzeigeNew "normal", "", Label6
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungExport"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub vorbereitungBestHist()
    On Error GoTo LOKAL_ERROR
    
'    fuelle_BestDat List9
    fuelle_BestDat_GDPDU List9
    Frame23.Visible = True
    txtAGN.Text = ""
    Label13.Caption = ""
    anzeigeNew "normal", "", Label6
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungBestHist"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
'Private Sub zeige_Best_Hist(cDatum As String, sAGN As String)
'On Error GoTo LOKAL_ERROR
'
'    Dim sSQL    As String
'
'    loeschNEW "Lieflw", gdBase
'    CreateTable "LIEFLW", gdBase
'
'    loeschNEW "ArtTemp", gdBase
'
'    sSQL = "select a.artnr,g.bestand,a.linr,a.ekpr,a.kvkpr1 into arttemp from artikel a inner join glager g on a.artnr = g.artnr "
'    sSQL = sSQL & " where g.Bestand > 0 "
'
'    If sAGN <> "" Then
'         sSQL = sSQL & " and a.agn = " & sAGN & " "
'    End If
'
'    sSQL = sSQL & " and g.datum = " & CLng(DateValue(cDatum)) & ""
'    gdBase.Execute sSQL, dbFailOnError
'
'    sSQL = "Update arttemp inner join artikel on arttemp.artnr = artikel.artnr "
'    sSQL = sSQL & " set arttemp.ekpr = artikel.lekpr where arttemp.ekpr = 0 "
'    gdBase.Execute sSQL, dbFailOnError
'
'
'    sSQL = "INSERT into LIEFLW Select LINR, Sum(arttemp.BESTAND) as BESTAND "
'    sSQL = sSQL & ", Sum(KVKPR1* arttemp.BESTAND) as LagerVK"
'    sSQL = sSQL & ", Sum(EKPR* arttemp.BESTAND) as LagerEK"
'    sSQL = sSQL & " from arttemp "
'    sSQL = sSQL & " Where arttemp.Bestand > 0  "
'
'    sSQL = sSQL & " group BY arttemp.LINR "
'    gdBase.Execute sSQL, dbFailOnError
'
'    loeschNEW "ArtTemp", gdBase
'
'    sSQL = "Update LIEFLW set BGrund = 'Schnitteinkaufswert' "
'    gdBase.Execute sSQL, dbFailOnError
'
'    sSQL = "Update LiefLW inner join lisrt on lieflw.linr = lisrt.linr set lieflw.LIEFBEZ = lisrt.liefbez"
'    gdBase.Execute sSQL, dbFailOnError
'
'    sSQL = "Update LiefLW set auswahl = '" & cDatum & "' "
'    gdBase.Execute sSQL, dbFailOnError
'
'
'    reportbildschirm "", "awkl46j"
'    'awklauh
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "zeige_Best_Hist"
'    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
'Private Sub zeige_Best_Hist_Einzel(cDatum As String, sAGN As String)
'On Error GoTo LOKAL_ERROR
'
'    Dim sSQL    As String
'
''    loeschNEW "Lieflw", gdBase
''    CreateTable "LIEFLW", gdBase
'
'    loeschNEW "ARTHISTE", gdBase
'    CreateTableT2 "ARTHISTE", gdBase
'
'    sSQL = "Insert into ARTHISTE "
'    sSQL = sSQL & " select a.artnr"
'    sSQL = sSQL & " ,a.bezeich"
'    sSQL = sSQL & " ,g.bestand"
'    sSQL = sSQL & " ,a.linr"
'    sSQL = sSQL & " ,a.ekpr"
'    sSQL = sSQL & " ,a.kvkpr1 "
'    sSQL = sSQL & " ,'' as liefbez"
'    sSQL = sSQL & " ,'Schnitteinkaufswert' as BGRUND "
'    sSQL = sSQL & " ,'" & cDatum & "'  as AUSWAHL "
'    sSQL = sSQL & " from artikel a inner join glager g on a.artnr = g.artnr "
'    sSQL = sSQL & " where g.Bestand > 0 "
'
'    If sAGN <> "" Then
'         sSQL = sSQL & " and a.agn = " & sAGN & " "
'    End If
'    sSQL = sSQL & " and g.datum = " & CLng(DateValue(cDatum)) & ""
'    gdBase.Execute sSQL, dbFailOnError
'
'    sSQL = "Update ARTHISTE inner join artikel on ARTHISTE.artnr = artikel.artnr "
'    sSQL = sSQL & " set ARTHISTE.ekpr = artikel.lekpr where ARTHISTE.ekpr = 0 "
'    gdBase.Execute sSQL, dbFailOnError
'
'    sSQL = "Update ARTHISTE inner join lisrt on ARTHISTE.linr = lisrt.linr set ARTHISTE.LIEFBEZ = lisrt.liefbez"
'    gdBase.Execute sSQL, dbFailOnError
'
'    reportbildschirm "", "awkl46k"
'    'awklauh
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "zeige_Best_Hist_Einzel"
'    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
Private Sub fuelle_BestDat_GDPDU(Listx As ListBox)
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim ctemp       As String
    Dim GDPdU_DB    As Database
    Dim cPfad       As String
    
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    Set GDPdU_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsGDPdU_Passwort)
    
    Listx.Visible = False
    Listx.Clear
    
    Dim lMAXDatumUebersicht     As Long
    Dim lMAXDatumGLAGER         As Long
    Dim rsDat                   As DAO.Recordset
    
    If NewTableSuchenDBKombi("GLAGER_GDPdU", GDPdU_DB) Then
    
        CheckIndex "GLAGER_GDPdU", "DATUM", "", GDPdU_DB
        CheckIndex "GLAGER_GDPdU", "BESTAND", "", GDPdU_DB
    
        If NewTableSuchenDBKombi("GLAGER_UEBERSICHT", GDPdU_DB) = False Then
            'dann erstelle eine
            cSQL = "select distinct(datum) as disdatum ,sum(bestand) as mBestand into GLAGER_UEBERSICHT from GLAGER_GDPdU group by datum "
            GDPdU_DB.Execute cSQL, dbFailOnError
        Else
            'füge neue Sätze an
            lMAXDatumUebersicht = 0
            lMAXDatumGLAGER = 0
    
            cSQL = "Select Max(disdatum) as Maxdat from GLAGER_UEBERSICHT"
            Set rsrs = GDPdU_DB.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                lMAXDatumUebersicht = rsrs!Maxdat
            End If
            rsrs.Close: Set rsrs = Nothing
            
            cSQL = "Select Max(datum) as Maxdat from GLAGER_GDPdU"
            Set rsrs = GDPdU_DB.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                lMAXDatumGLAGER = rsrs!Maxdat
            End If
            rsrs.Close: Set rsrs = Nothing
    
            If lMAXDatumGLAGER > lMAXDatumUebersicht Then
                'dann gibt es etwas anzufügen
                cSQL = "Insert into GLAGER_UEBERSICHT select distinct(datum) as disdatum ,sum(bestand) as mBestand "
                cSQL = cSQL & " from GLAGER_GDPdU where datum > " & lMAXDatumUebersicht & " group by datum "
                GDPdU_DB.Execute cSQL, dbFailOnError
                
            End If
        End If
    
        anzeige "", "verfügbare Daten werden ermittelt...", Label6
        
        cSQL = "select  disdatum , mBestand from GLAGER_UEBERSICHT order by disdatum desc"
        Set rsrs = GDPdU_DB.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!disdatum) Then
                    ctemp = Format(rsrs!disdatum, "DD.MM.YY")
                Else
                    ctemp = ""
                End If
                
                If ctemp <> "" Then
                    If Not IsNull(rsrs!mBestand) Then
                        ctemp = ctemp & " (" & rsrs!mBestand & ")"
                    End If

'                    If cboBestHist.Text = "" Then
'                        cboBestHist.Text = ctemp
'                    End If

                    Listx.AddItem ctemp
                End If
                
                rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    Listx.Visible = True
    
    GDPdU_DB.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuelle_BestDat_GDPDU"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



'Private Sub fuelle_BestDat(Listx As ListBox)
'On Error GoTo LOKAL_ERROR
'
'    Dim cSQL    As String
'    Dim rsrs    As Recordset
'    Dim ctemp   As String
'
'    Listx.Visible = False
'    Listx.Clear
'
'    anzeige "", "verfügbare Daten werden ermittelt...", Label6
'
'    cSQL = "select distinct(datum) as disdatum  from glager order by datum desc"
'    Set rsrs = gdBase.OpenRecordset(cSQL)
'    If Not rsrs.EOF Then
'        rsrs.MoveFirst
'        Do While Not rsrs.EOF
'            If Not IsNull(rsrs!disdatum) Then
'                ctemp = rsrs!disdatum
'            Else
'                ctemp = ""
'            End If
'
'            Listx.AddItem ctemp
'
'            rsrs.MoveNext
'        Loop
'    End If
'    rsrs.Close: Set rsrs = Nothing
'
'    Listx.Visible = True
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "fuelle_BestDat"
'    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
Private Sub Command20_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Select Case Index
        Case Is = 20        ' Kalender
            Text1(5).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
        End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command20_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten. "

    Fehlermeldung1
   
End Sub

Private Sub Command3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad   As String
    Dim cdatei  As String
    Dim cdat    As String
    Dim lcount  As Long
    Dim bFound  As Boolean
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0     'SPEICHERN
            SpeichereInventurDateiWKL46
            
        Case Is = 1     'SCHLIESSEN
            anzeigeNew "normal", "", Label6
            If MSFlexGrid1.Visible = True And MSFlexGrid1.Rows > 1 Then
                Frame1.Enabled = True
                Command1.Enabled = Frame1.Enabled
                Frame2.Enabled = True
                Frame3.Enabled = True
                Command4.Enabled = Frame3.Enabled
                Command2(1).Enabled = Frame3.Enabled
                Command2(2).Enabled = Frame3.Enabled
                Command2(3).Enabled = Frame3.Enabled
                Command2(4).Enabled = Frame3.Enabled
                Command2(5).Enabled = Frame3.Enabled
                Frame4.Visible = False
                Frame5.Visible = False
                
                MSFlexGrid1.SetFocus
                MSFlexGrid1.Row = 1
                MSFlexGrid1.Col = 2
            Else
                Frame1.Enabled = True
                Command1.Enabled = Frame1.Enabled
                Frame2.Enabled = True
                Frame3.Enabled = True
                Command4.Enabled = Frame3.Enabled
                Command2(1).Enabled = Frame3.Enabled
                Command2(2).Enabled = Frame3.Enabled
                Command2(3).Enabled = Frame3.Enabled
                Command2(4).Enabled = Frame3.Enabled
                Command2(5).Enabled = Frame3.Enabled
                Frame4.Visible = False
                Frame5.Visible = False
                
            End If
        Case Is = 2     'LADEN
            LadeInventurDateiWKL46
            
        Case Is = 3     'SCHLIESSEN
            Command3_Click 1
        Case Is = 4     'Löschen
        
            cdatei = Label1(3).Caption
            cdatei = Trim$(UCase$(cdatei))
            If cdatei = "" Then
                anzeigeNew "rot", "Bitte eine Inventur-Datei angeben!", Label6
                List1.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            loeschNEW cdatei, gdBase
            anzeigeNew "normal", "Datei " & cdatei & " gelöscht!", Label6
            NewListeFuellAnfangsbuch "INV_", frmWKL46.List1, gdBase
            Label1(3).Caption = ""
            Screen.MousePointer = 0
            Exit Sub
            
        Case 5
        
            Command6(6).Enabled = False
            
            bFound = False
    
            For lcount = 0 To List6.ListCount - 1
                If List6.Selected(lcount) = True Then
                    cdat = "ILI" & Left(List6.list(lcount), 5)
                    cdatei = Left(List6.list(lcount), 5)
                    loeschNEW cdat, gdBase
                    schreibeIProtokoll "Inventurliste: " & cdatei & " gelöscht"
                    bFound = True
                End If
            Next lcount
    
            If bFound Then
                anzeigeNew "normal", "Die Datei/en wurde/n gelöscht!", Label6
                New2ListeFuellAnfangsbuch "ILI", frmWKL46.List6, gdBase
            End If
        Case 6
            bFound = False
            For lcount = 0 To List6.ListCount - 1
                If List6.Selected(lcount) = True Then
                    cdat = "ILI" & Left(List6.list(lcount), 5)
                    cdatei = Left(List6.list(lcount), 5)

                    bFound = True
                End If
            Next lcount
            
            If bFound Then
                anzeigeNew "normal", "Die Datei: " & cdatei & " wird geladen...", Label6
                Dateiladen Trim(cdat), cdatei
            Else
                anzeigeNew "rot", "Bitte markieren Sie eine Datei!", Label6
            End If
        Case 7 'Speichern
            SPEICHERNEWbest
        Case 8     'Excel Export
            cdatei = Label1(3).Caption
            cdatei = Trim$(UCase$(cdatei))
            If cdatei = "" Then
                MsgBox "Bitte eine Inventur-Datei angeben!", vbInformation, "Zentrale Hinweis:"
                List1.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If

            ExcelExport cdatei, gdBase
'            Screen.MousePointer = 0

        End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
     
End Sub
Private Sub SPEICHERNEWbest()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim cArtNr As String
    Dim inewBest As Integer
    Dim ioldBest As Integer
    Dim ctmp As String
    
    loeschNEW "DIFFLIST", gdBase
    CreateTable "DIFFLIST", gdBase

    For lrow = 2 To MSHFlexGrid1.Rows - 1
    
        MSHFlexGrid1.Row = lrow
        
        MSHFlexGrid1.Col = 1
        cArtNr = MSHFlexGrid1.Text
        
        MSHFlexGrid1.Col = 7
        ctmp = MSHFlexGrid1.Text
        ioldBest = CInt(ctmp)
        
        MSHFlexGrid1.Col = 8
        ctmp = MSHFlexGrid1.Text
        If ctmp = "" Then
            inewBest = 0
        Else
            
            inewBest = CInt(ctmp)
        End If
        
        If inewBest <> ioldBest Then
            Bestandsveraenderung cArtNr, CLng(inewBest), "Inventur aus Liste"
            fUelledifferenzliste cArtNr, ioldBest, inewBest
        End If
        
    Next lrow
    
    MSHFlexGrid1.Visible = False
    anzeigeNew "normal", "Bitte wählen Sie eine Datei aus!", Label6
    
    reportbildschirm "", "aWKL46h"
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SPEICHERNEWbest"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fUelledifferenzliste(cArtNr As String, ioldb As Integer, inewb As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Insert into DIFFLIST (ARTNR,Bestsoll,Bestist,datname) values "
    sSQL = sSQL & " ( " & cArtNr & "," & ioldb & "," & inewb & ", '" & Label34.Caption & "')"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DIFFLIST inner join artikel on difflist.artnr = artikel.artnr "
    sSQL = sSQL & " set DIFFLIST.bezeich = artikel.bezeich "
    sSQL = sSQL & " , DIFFLIST.ean = artikel.ean "
    sSQL = sSQL & " , DIFFLIST.LINR = artikel.linr "
    sSQL = sSQL & " , DIFFLIST.LPZ = artikel.lpz "
    sSQL = sSQL & " , DIFFLIST.AGN = artikel.AGN "
    sSQL = sSQL & " , DIFFLIST.VKPR = artikel.VKPR "
    sSQL = sSQL & " , DIFFLIST.KVKPR1 = artikel.KVKPR1 "
    sSQL = sSQL & " , DIFFLIST.EKPR = artikel.EKPR "
    sSQL = sSQL & " , DIFFLIST.LEKPR = artikel.LEKPR "
    sSQL = sSQL & " , DIFFLIST.MWST = artikel.MWST "
    
    gdBase.Execute sSQL, dbFailOnError
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fUelledifferenzliste"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Dateiladen(cdatei As String, clabelname As String)
    On Error GoTo LOKAL_ERROR
    
    erstellegrid
    fuellegrid cdatei, clabelname
    
    anzeigeNew "normal", "mehr Informationen mit Doppelklick auf ArtNr", Label6
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Dateiladen"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub erstellegrid()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim lcol As Long
    Dim i As Integer
    
    With MSHFlexGrid1
        .Redraw = False
        .Visible = False
        .Rows = 2
        .Cols = 9
        .FixedRows = 1
        .FixedCols = 1
        
        .Row = 0
        
        .Col = 0
        .Text = "lfNr."
        
        .Col = 1
        .Text = "ArtNr."
        
        .Col = 2
        .Text = "Artikelbezeichnung"
        
        .Col = 3
        .Text = "Lieferant"
        
        .Col = 4
        .Text = "Linie"
        
        .Col = 5
        .Text = "AGN"
        
        .Col = 6
        .Text = "Kassenpreis"
        
        .Col = 7
        .Text = "Soll Bestand"
        
        .Col = 8
        .Text = "Ist Bestand"
        
        
        
        For i = 0 To 8
            aBreite(i) = TextWidth(.TextMatrix(0, i))
        Next i
    End With
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "erstellegrid"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellegrid(cdat As String, clab As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim i           As Integer
    Dim sSQL        As String
    
    sSQL = "Update " & cdat & " inner join artikel on " & cdat & ".artnr = artikel.artnr "
    sSQL = sSQL & " set " & cdat & ".bestand = artikel.bestand "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    'hier druckdatentab bauen
    
    loeschNEW "li45", gdBase
        
    sSQL = "Select * into li45 from " & cdat & " "
    gdBase.Execute sSQL, dbFailOnError
    
    
    Command6(6).Enabled = True
'    SpalteAnfuegenNEW cdat, "lfnr", "autoincrement", gdBase
    
    'ende hier druckdatentab bauen
    
    
    
    
    sSQL = "Select * from " & cdat & " order by Lfnr "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    With MSHFlexGrid1
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lrow = lrow + 1
            
            .Rows = lrow + 1
            .Row = lrow
            
            If Not IsNull(rsrs!lfnr) Then
                lWert = rsrs!lfnr
            Else
                lWert = 0
            End If
            
            .Col = 0
            .Text = lWert
            
            If Not IsNull(rsrs!artnr) Then
                lWert = rsrs!artnr
            Else
                lWert = 0
            End If
            
            .Col = 1
            .Text = lWert
            
            If Not IsNull(rsrs!BEZEICH) Then
                sWert = rsrs!BEZEICH
            Else
                sWert = ""
            End If
            
            .Col = 2
            .Text = sWert
            
            If Not IsNull(rsrs!linr) Then
                lWert = rsrs!linr
            Else
                lWert = 0
            End If
            
            .Col = 3
            .Text = lWert
            
            If Not IsNull(rsrs!LPZ) Then
                lWert = rsrs!LPZ
            Else
                lWert = 0
            End If
            
            .Col = 4
            .Text = lWert
            
            If Not IsNull(rsrs!AGN) Then
                lWert = rsrs!AGN
            Else
                lWert = 0
            End If
            
            .Col = 5
            .Text = lWert
            
            If Not IsNull(rsrs!KVKPR1) Then
                siWert = rsrs!KVKPR1
            Else
                siWert = 0
            End If
            
            .Col = 6
            .Text = siWert
            
            If Not IsNull(rsrs!BESTAND) Then
                lWert = rsrs!BESTAND
            Else
                lWert = 0
            End If
            
            .Col = 7
            .Text = lWert
            
            .Col = 8
            .CellBackColor = vbGreen
            .Text = lWert
            
            For i = 0 To 8
                If TextWidth(.TextMatrix(lrow, i)) > aBreite(i) Then
                    aBreite(i) = TextWidth(.TextMatrix(lrow, i))
                End If
                

            Next i
            
            rsrs.MoveNext
        Loop
    End If
    
    For i = 0 To 8
        .Col = i
        .ColWidth(i) = aBreite(i) * 1.4
    Next i
    
    rsrs.Close: Set rsrs = Nothing
    
    .RowHeight(1) = 0
    lrow = lrow - 1
    
    .Redraw = True
    .Visible = True
    End With
    
    Label34.Caption = clab
    Label34.Refresh
    
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellegrid"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs        As Recordset
    Dim sSQL        As String
    
    Set rsrs = gdBase.OpenRecordset("INVLITE", dbOpenTable)
    
    If rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        Exit Sub
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Screen.MousePointer = 11
    
    BerechneInventurWKL46
    
    loeschNEW "INV_LITE", gdBase
    sSQL = "Select * into INV_LITE from INVLITE"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "I_LSum", gdBase
    
    sSQL = "create table I_LSum"
    sSQL = sSQL & " ( omw Text(15) "
    sSQL = sSQL & " , VmW Text(15) "
    sSQL = sSQL & " , emW Text(15) ) "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Insert into I_LSum ( omW ,vmW,emW )"
    sSQL = sSQL & " values "
    sSQL = sSQL & " ( '" & VkwertErmittlungOhneMW & "','" & VkwertErmittlungVolleMW & "','" & VkwertErmittlungErmMW & "') "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    'nur für den normal Report
    loeschNEW "I_LSUMEK", gdBase
    
    sSQL = "create table I_LSUMEK"
    sSQL = sSQL & " ( omw Text(15) "
    sSQL = sSQL & " , VmW Text(15) "
    sSQL = sSQL & " , emW Text(15) ) "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Insert into I_LSUMEK ( omW ,vmW,emW )"
    sSQL = sSQL & " values "
    sSQL = sSQL & " ( '" & EkwertErmittlungOhneMW & "','" & EkwertErmittlungVolleMW & "','" & EkwertErmittlungErmMW & "') "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    loeschNEW "I_LKOP", gdBase
    
    sSQL = "create table I_LKOP"
    sSQL = sSQL & " ( INVDAT Text(10),LISTENVKWERT double)"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Insert into I_LKOP ( INVDAT,LISTENVKWERT )"
    sSQL = sSQL & " values "
    sSQL = sSQL & " ( '" & Text1(5).Text & "', '" & ListenVerkaufswertErmittlung & "') "
    gdBase.Execute sSQL, dbFailOnError
    
    haengan4
    
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    
    
    If Check1 = vbChecked Then
        reportbildschirm "INVENe", "aWKL46a"
    Else
        If FileExists(cPfad & "aWKL46S.rpt") Then
            haengan5
            reportbildschirm "INVENe", "aWKL46S"
        Else
            reportbildschirm "INVENe", "aWKL46"
        End If
    End If

    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command5_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    gF2Prompt.cFeld = "LINR"
    gF2Prompt.cWert = ""
    gF2Prompt.cWert2 = ""
    gF2Prompt.cWahl = ""
    
    If gF2Prompt.cFeld <> "" Then
        frmWK00a.Show 1
    End If
    
    If gF2Prompt.cWahl <> "" Then
        ctmp = gF2Prompt.cWahl
        ctmp = ctmp & String$(6 - Len(ctmp), "_")
        Text2.Text = ctmp
    End If
    Text2.SetFocus
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub





Private Sub LeereDialogWKL46()
    On Error GoTo LOKAL_ERROR
    
    Text1(2).Text = ""
    Text1(1).Text = ""
    Combo2.Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKL46"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command6_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    Select Case Index
        Case 0 'suchen
            If SucheArtikelWKL46 Then
                For i = 0 To 3
                    If Option4(i).Value = True Then
                        zeigeartikel Option4(i).Tag
                    Exit For
                    End If
                Next i
            End If
        Case 1      'F2 AGN
            Text1_KeyUp 3, vbKeyF2, 0
        Case 2      'F2 Lieferant
            Combo2_KeyUp vbKeyF2, 0
        Case 3     'drucken
            If List5.ListCount > 0 Then
                If Check4.Value = vbChecked Then
                    If Text7.Text = "" Then
                        anzeigeNew "rot", "Vergeben Sie bitte einen Dateinamen!", Label6
                        Text7.SetFocus
                        Exit Sub
                    Else
                        If gspeichertli46 Then
                            Drucklist46
                            schreibeIProtokoll "Inventurliste: " & Text7.Text & " gespeichert."
                            
                            voreinstellungspeichernE46il
                            vorbereitungInventurlisteErzeugen
                        End If
                    End If
                Else
                    If Text7.Text <> "" Then
                        If gspeichertli46 Then
                            Drucklist46
                            schreibeIProtokoll "Inventurliste: " & Text7.Text & " gespeichert."
                        End If
                    Else
                        Drucklist46
                        schreibeIProtokoll "Eine Inventurliste: ohne Speichern nur Druck - erstellt"
                    End If
                    voreinstellungspeichernE46il
                    vorbereitungInventurlisteErzeugen
                End If
            End If
        Case 4
            EntfernenVonListEin
        Case 5
            List7.Clear
            List7.Visible = False
        Case 6
        
            Drucklist46FERTIG
        
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub EntfernenVonListEin()
    On Error GoTo LOKAL_ERROR

    Dim lcount  As Long
    Dim cArtNr  As String
    Dim cSQL    As String
    Dim bFound  As Boolean
    Dim i       As Integer
    
    bFound = False
    
    For lcount = 0 To List5.ListCount - 1
        If List5.Selected(lcount) = True Then
            cArtNr = Trim$(List5.list(lcount))
            cArtNr = Left(cArtNr, 6)
            cArtNr = Trim$(cArtNr)
            cSQL = "Delete from li46 where ARTNR = " & cArtNr
            schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
            
            bFound = True
        
        End If
    Next lcount
    
    If bFound Then
        For i = 0 To 3
            If Option4(i).Value = True Then
                zeigeartikel Option4(i).Tag
            Exit For
            End If
        Next i
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EntfernenVonListEin"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function gspeichertli46() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cdat    As String
    Dim lRet    As Long
    Dim cSQL    As String
    Dim cdatei  As String

    gspeichertli46 = False
    
    cdat = "ILI" & Trim$(UCase$(Text7.Text))
    cdat = Trim(cdat)
    cdatei = Trim$(UCase$(Text7.Text))
    
    If NewTableSuchenDBKombi(cdat, gdBase) Then
        lRet = MsgBox("Die Datei " & cdatei & " existiert bereits! Überschreiben?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If lRet <> vbYes Then
            Exit Function
        End If
    End If
    
    loeschNEW cdat, gdBase
        
    cSQL = "Select * into " & cdat & " from li45 "
    gdBase.Execute cSQL, dbFailOnError
    
    SpalteAnfuegenNEW cdat, "lfnr", "autoincrement", gdBase
    
    
    anzeigeNew "normal", "Die Datei " & cdatei & " wurde gespeichert.", Label6
    
    gspeichertli46 = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "gspeichertli46"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Drucklist46()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim rsRSQ           As Recordset
    Dim cEAN            As String
    Dim cEANCode        As String
    
    Screen.MousePointer = 11

    loeschNEW "DruLi46", gdBase
    CreateTable "DRULI46", gdBase
    
    Set rsRSQ = gdBase.OpenRecordset("DRULI46")
    
    sSQL = "Select * from li45 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsRSQ.AddNew
            
            rsRSQ!artnr = rsrs!artnr
            rsRSQ!BEZEICH = rsrs!BEZEICH
            rsRSQ!AGN = rsrs!AGN
            rsRSQ!linr = rsrs!linr
            rsRSQ!LPZ = rsrs!LPZ
            rsRSQ!KVKPR1 = rsrs!KVKPR1
            If Check7.Value = vbChecked Then
                rsRSQ!BESTAND = rsrs!BESTAND
            Else
                rsRSQ!BESTAND = Null
            End If
            rsRSQ!LIBESNR = rsrs!LIBESNR
            rsRSQ!ADATE = DateValue(Now)
            rsRSQ!Datname = Text7.Text
            rsRSQ!EAN2 = rsrs!EAN
        
            cEAN = rsrs!artnr
            cEAN = fnMoveArtNr2EAN8(cEAN)
            rsRSQ!EAN = cEAN
                
            cEANCode = fnCodiereEANCode(cEAN)
            rsRSQ!Barcode = cEANCode
            
            rsRSQ.Update
        
        rsrs.MoveNext
        Loop
    End If
    
    rsrs.Close: Set rsrs = Nothing
    rsRSQ.Close
    
    If Check5.Value = vbChecked Then
        reportbildschirm "", "aWKL46f"
    Else

        reportbildschirm "", "aWKL46i"
    End If
    
    Screen.MousePointer = 0
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Drucklist46"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Drucklist46FERTIG()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim rsRSQ           As Recordset
    Dim cEAN            As String
    Dim cEANCode        As String
    
    Screen.MousePointer = 11

    loeschNEW "DruLi46", gdBase
    CreateTable "DRULI46", gdBase

    Set rsRSQ = gdBase.OpenRecordset("DRULI46")
    
    sSQL = "Select * from li45 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsRSQ.AddNew
            
            rsRSQ!artnr = rsrs!artnr
            rsRSQ!BEZEICH = rsrs!BEZEICH
            rsRSQ!AGN = rsrs!AGN
            rsRSQ!linr = rsrs!linr
            rsRSQ!LPZ = rsrs!LPZ
            rsRSQ!KVKPR1 = rsrs!KVKPR1

            rsRSQ!BESTAND = rsrs!BESTAND

            rsRSQ!LIBESNR = rsrs!LIBESNR
            rsRSQ!ADATE = DateValue(Now)
            rsRSQ!Datname = Label34.Caption
            rsRSQ!EAN2 = rsrs!EAN
        
            cEAN = rsrs!artnr
            cEAN = fnMoveArtNr2EAN8(cEAN)
            rsRSQ!EAN = cEAN
                
            cEANCode = fnCodiereEANCode(cEAN)
            rsRSQ!Barcode = cEANCode
            
            rsRSQ.Update
        
        rsrs.MoveNext
        Loop
    End If
    
    rsrs.Close: Set rsrs = Nothing
    rsRSQ.Close
    
    If Check11.Value = vbChecked Then
        reportbildschirm "", "aWKL46f"
    Else

        reportbildschirm "", "aWKL46i"
    End If
    
    Screen.MousePointer = 0
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Drucklist46FERTIG"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function fnPruefeEingabeWKL46()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    fnPruefeEingabeWKL46 = 1
    
    If Trim$(Combo2.Text) <> "" Then 'Liefcombo
        Combo2.Text = Trim(ErmittleLinr(Combo2.Text))
        If Trim$(Combo2.Text) <> "" Then
            fnPruefeEingabeWKL46 = 0
            Exit Function
        End If
    End If
    
    For lcount = 0 To 6
        If lcount = 5 Then
        
        Else
            If Trim$(Text1(lcount).Text) <> "" Then
                fnPruefeEingabeWKL46 = 0
                Exit Function
            End If
        End If
    Next lcount

            
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeWKL46"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function SucheArtikelWKL46() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim cwhere      As String
    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim lcol        As Long
    Dim dWert       As Double
    Dim iRet        As Integer
    Dim i           As Integer
    Dim cEAN        As String
    Dim cArtNr      As String
    Dim cEigNr      As String
    Dim cJoin       As String
    Dim llpzvon     As Long
    Dim llpzbis     As Long
    
    SucheArtikelWKL46 = False
    
    iRet = fnPruefeEingabeWKL46()
    If iRet <> 0 Then
        anzeigeNew "rot", "Bitte mindestens ein Suchkriterium angeben!", Label6
        Text1(1).SetFocus
        Exit Function
    End If
    
    llpzvon = Val(Text1(6).Text)
    llpzbis = Val(Text1(0).Text)
    
    If llpzvon > 0 And llpzbis = 0 Then
        llpzbis = llpzvon
    End If
    
    loeschNEW "li46", gdBase
    
    
    cSQL = "Select B.ARTNR"
    cSQL = cSQL & ", A.BEZEICH"
    cSQL = cSQL & ", A.AGN"
    cSQL = cSQL & ", B.LEKPR"
    cSQL = cSQL & ", A.VKPR"
    cSQL = cSQL & ", A.MWST"
    cSQL = cSQL & ", B.LINR"
    cSQL = cSQL & ", B.LIBESNR"
    cSQL = cSQL & ", A.EAN"
    cSQL = cSQL & ", A.MOPREIS"
    cSQL = cSQL & ", A.RKZ"
    cSQL = cSQL & ", A.LPZ"
    cSQL = cSQL & ", A.NOTIZEN"
    cSQL = cSQL & ", A.BESTAND"
    cSQL = cSQL & ", A.VKMENGE"
    cSQL = cSQL & ", A.VKDATUM"
    cSQL = cSQL & ", B.MINMEN"
    cSQL = cSQL & ", A.EAN2"
    cSQL = cSQL & ", A.EAN3"
    cSQL = cSQL & ", A.INHALT"
    cSQL = cSQL & ", A.INHALTBEZ"
    cSQL = cSQL & ", A.GRUNDPREIS"
    cSQL = cSQL & ", A.MINBEST"
    cSQL = cSQL & ", A.RABATT_OK"
    cSQL = cSQL & ", A.GEFUEHRT"
    cSQL = cSQL & ", A.EKPR"
    cSQL = cSQL & ", A.KVKPR1"
    
    If llpzvon > 0 Then
        cSQL = cSQL & ", c.Lagerp"
    Else
        cSQL = cSQL & ", 0 as Lagerp"
    
    End If
    
    cSQL = cSQL & " into li46 from ARTIKEL A, ARTLIEF B "
    
    If llpzvon > 0 Then
        cSQL = cSQL & " , lagerplatz c where A.artnr = c.artnr"
        cwhere = " and "
    Else
        cwhere = " Where "
    End If
    
    cwhere = cwhere & " ( A.SYNSTATUS = 'E' or A.SYNSTATUS = 'A' or A.SYNSTATUS is null )"
    
    If llpzvon > 0 Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & " c.lagerp between " & llpzvon & " and " & llpzbis & " "
    End If
    
    cFeld = Text1(2).Text
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.BEZEICH like '" & cFeld & "*' "
    End If
    
    cFeld = Text1(1).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cEAN = cFeld
        If Len(cFeld) <= 6 Then
            cArtNr = cFeld
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
            If InStr(cEAN, "*") > 0 Then
                cwhere = cwhere & "A.EAN like '" & cEAN & "' "
            Else
                cwhere = cwhere & "A.EAN = '" & cEAN & "' "
            End If
            If InStr(cEAN, "*") > 0 Then
                cwhere = cwhere & "or A.EAN2 like '" & cEAN & "' "
            Else
                cwhere = cwhere & "or A.EAN2 = '" & cEAN & "' "
            End If
            If InStr(cEAN, "*") > 0 Then
                cwhere = cwhere & "or A.EAN3 like '" & cEAN & "' "
            Else
                cwhere = cwhere & "or A.EAN3 = '" & cEAN & "' "
            End If
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
        
    End If
    
    cFeld = Trim(Combo2.Text)
    
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "B.LINR = " & cFeld & " "
    End If
    
    
    If List7.ListCount = 0 Then
    
    Else
        If cwhere = "" Then
            cwhere = "where ( "
        Else
            cwhere = cwhere & "and ( "
        End If
       
        
        For i = 0 To List7.ListCount - 1
        
            cFeld = Trim(Left(List7.list(i), 3))
            If cFeld <> "" Then
                If i = 0 Then
                    cwhere = cwhere & " A.LPZ = " & cFeld & " "
                Else
                    cwhere = cwhere & "or A.LPZ = " & cFeld & " "
                End If
                
            End If
            
        Next i
        
        cwhere = cwhere & " ) "
        
    End If
    
    
    cFeld = Text1(3).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.AGN = " & cFeld & " "
    End If
    
    cFeld = Text1(4).Text
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "B.LIBESNR like '" & cFeld & "' "
    End If
    
    cJoin = "and A.ARTNR = B.ARTNR "
    
    If Check3.Value = vbChecked Then
        cJoin = cJoin & " and A.GEFUEHRT = 'J' "
    End If
    cSQL = cSQL & cwhere & cJoin
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update li46 inner join Lagerplatz on "
    cSQL = cSQL & " li46.artnr = Lagerplatz.artnr "
    cSQL = cSQL & " set li46.lagerp = Lagerplatz.lagerp "
    gdBase.Execute cSQL, dbFailOnError
    
    DublikateDelLI46
    
    SucheArtikelWKL46 = True

Exit Function
LOKAL_ERROR:
    If err.Number = 30006 Then
        Screen.MousePointer = 0
        anzeigeNew "rot", "Bitte schränken Sie Ihre Suche weiter ein! Es wurden zu viele Artikel ermittelt.", Label6
        Exit Function
    Else
        
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SucheArtikelWKL46"
        Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
   
End Function
Private Sub zeigeartikel(sOrder As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim lAnzSatz    As Long
    
    loeschNEW "li45", gdBase
    
    cSQL = "Select  * into li45 from li46 order by " & sOrder
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from li45 "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    List5.Clear
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzSatz = rsrs.RecordCount
        If lAnzSatz > 5000 Then
            Screen.MousePointer = 0
            anzeigeNew "rot", "Bitte schränken Sie Ihre Suche weiter ein! Es wurden zu viele Artikel ermittelt.", Label6
            Exit Sub
        End If

        rsrs.MoveFirst
        lAnzSatz = 0
        Do While Not rsrs.EOF
            lAnzSatz = lAnzSatz + 1

            If Not IsNull(rsrs!artnr) Then
                cFeld = Trim$(rsrs!artnr) & Space(7 - Len(Trim$(rsrs!artnr)))
            End If
            cLBSatz = cFeld

            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = Trim$(rsrs!BEZEICH) & Space(36 - Len(Trim$(rsrs!BEZEICH)))
            Else
                cFeld = cFeld & Space(36)
            End If

            cLBSatz = cLBSatz & cFeld

            If Not IsNull(rsrs!AGN) Then
                cFeld = Trim$(rsrs!AGN) & Space(6 - Len(Trim$(rsrs!AGN)))
            Else
                cFeld = cFeld & Space(6)
            End If
            cLBSatz = cLBSatz & cFeld
            
            If Not IsNull(rsrs!LPZ) Then
                cFeld = Trim$(rsrs!LPZ) & Space(6 - Len(Trim$(rsrs!LPZ)))
            Else
                cFeld = cFeld & Space(6)
            End If
            cLBSatz = cLBSatz & cFeld

            If Not IsNull(rsrs!linr) Then
                cFeld = Trim$(rsrs!linr) & Space(7 - Len(Trim$(rsrs!linr)))
            Else
                cFeld = cFeld & Space(7)
            End If
            cLBSatz = cLBSatz & cFeld

            If Not IsNull(rsrs!LIBESNR) Then
                cFeld = Trim$(rsrs!LIBESNR) & Space(14 - Len(Trim$(rsrs!LIBESNR)))
            Else
                cFeld = cFeld & Space(14)
            End If
            cLBSatz = cLBSatz & cFeld

            If Not IsNull(rsrs!EAN) Then
                cFeld = Trim$(rsrs!EAN) & Space(14 - Len(Trim$(rsrs!EAN)))
            Else
                cFeld = cFeld & Space(14)
            End If
            cLBSatz = cLBSatz & cFeld
            
            If Not IsNull(rsrs!lagerp) Then
                cFeld = Trim$(rsrs!lagerp) & Space(7 - Len(Trim$(rsrs!lagerp)))
            Else
                cFeld = cFeld & Space(7)
            End If
            cLBSatz = cLBSatz & cFeld

            List5.AddItem cLBSatz

            rsrs.MoveNext
        Loop
        List5.Visible = True
        anzeigeNew "normal", lAnzSatz & " Daten gefunden!", Label6
    Else
        anzeigeNew "rot", "Keine Daten gefunden!", Label6
    End If
    rsrs.Close: Set rsrs = Nothing


Exit Sub
LOKAL_ERROR:
    If err.Number = 30006 Then
        Screen.MousePointer = 0
        anzeigeNew "rot", "Bitte schränken Sie Ihre Suche weiter ein! Es wurden zu viele Artikel ermittelt.", Label6
        Exit Sub
    Else
        
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "zeigeartikel"
        Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub



Private Sub Command7_Click()
On Error GoTo LOKAL_ERROR
    
    Dim iText  As Integer
    iText = CInt(Text3(1).Text)
    If iText = 999 Then
    
    Else
        iText = iText + 1
        Text3(1).Text = CStr(iText)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer1.Enabled = False

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer2.Enabled = False

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command8_Click()
On Error GoTo LOKAL_ERROR
    
    Dim iText  As Integer
    iText = CInt(Text3(1).Text)
    If iText = -999 Then
    
    Else
        iText = iText - 1
        Text3(1).Text = CStr(iText)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command9_Click()
On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    gF2Prompt.cFeld = "LINR"
    gF2Prompt.cWert = ""
    gF2Prompt.cWert2 = ""
    gF2Prompt.cWahl = ""
    
    If gF2Prompt.cFeld <> "" Then
        frmWK00a.Show 1
    End If
    
    If gF2Prompt.cWahl <> "" Then
        ctmp = gF2Prompt.cWahl
        ctmp = ctmp & String$(6 - Len(ctmp), "_")
        Text4.Text = ctmp
    End If
    Text4.SetFocus
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command9_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(20).ForeColor = glS1
    Label1(21).ForeColor = glS1
    Label1(22).ForeColor = glS1
    Label1(23).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame6_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case Is = 20 'Inventur mit Scanner
        
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/123-inventur-leicht-gemacht.html"
        Case Is = 21 'Inventur mit MDE - Gerät
        
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/175-inventur-mit-mde-geraet.html"
        Case Is = 22 'Export einer Artikeldatei
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/hilfe-bei-problemen/44-software-probleme-winkiss/219-daten-fuer-inventurdienstleister-gresch-bereitstellen.html"
        Case Is = 23 'Import einer Bestandsdatei
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/hilfe-bei-problemen/44-software-probleme-winkiss/222-daten-des-inventurdienstleisters-gresch-abgleichen.html"

    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    
    
    If Index = 20 Then
        Label1(20).ForeColor = glLink
    End If
    
    If Index = 21 Then
        Label1(21).ForeColor = glLink
    End If
    
    If Index = 22 Then
        Label1(22).ForeColor = glLink
    End If
    
    If Index = 23 Then
        Label1(23).ForeColor = glLink
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub List1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    
    cLBSatz = List1.list(List1.ListIndex)
    cLBSatz = Trim$(UCase$(cLBSatz))
    cLBSatz = Trim(Left(cLBSatz, 8))
        
    Label1(3).Caption = cLBSatz
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List1_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Modul6.Farbform Me, lblUeberschrift
    PositionierenWKL46
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    
    Frame6.Visible = True
    Option2(2).Caption = Option2(2).Caption & " (" & gsMDEGERAET & ")"
    
    If NewTableSuchenDBKombi("E46", gdBase) Then
        If SpalteInTabellegefundenNEW("E46", "BO16", gdBase) = False Then
            SpalteAnfuegenNEW "E46", "BO16", "BIT", gdBase
            SpalteAnfuegenNEW "E46", "BO17", "BIT", gdBase
        End If
        
        If SpalteInTabellegefundenNEW("E46", "BO18", gdBase) = False Then
            SpalteAnfuegenNEW "E46", "BO18", "BIT", gdBase
        End If
        
        If SpalteInTabellegefundenNEW("E46", "BO19", gdBase) = False Then
            SpalteAnfuegenNEW "E46", "BO19", "BIT", gdBase
        End If
        
        If SpalteInTabellegefundenNEW("E46", "BO20", gdBase) = False Then
            SpalteAnfuegenNEW "E46", "BO20", "BIT", gdBase
        End If
        
        voreinstellungladen
    End If
    
    Text1(5).Text = Format(DateValue(Now), "DD.MM.YYYY")
    
'    Command1.BackColor = vbWhite
'    Command5.BackColor = vbWhite
    Command1.Visible = True
    iscan = 1
    
    Timer1.Enabled = False
    Timer2.Enabled = False
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ExportCSV()
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
    
    anzeige "normal", "Exportdatei wird erstellt...", Label6
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    sSQL = " Select "
    sSQL = sSQL & " EAN  "
    sSQL = sSQL & ", BEZEICH  "
    sSQL = sSQL & ", ARTNR "
    sSQL = sSQL & ", KVKPR1  "
    sSQL = sSQL & " from MDE_EXPORT "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then

        sAusgabedatname = "mde_input.csv"

        cPfad1 = gcDBPfad
        If Right$(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If

        cdatei = cPfad1 & "BOX\" & sAusgabedatname
        cPfad = cPfad1 & "BOX"
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        cSatz = "EAN;BEZEICH;ARTNR;KVKPR" & Chr$(13) & Chr$(10)

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            For i = 0 To 3
                If Not IsNull(rsrs.Fields(i)) Then

                    If i > 0 Then
                        If i = 3 Then
                            If rsrs.Fields(i) = 0 Then
                                cSatz = cSatz & ";"
                            Else
                                cSatz = cSatz & ";" & Format(rsrs.Fields(i), "###,##0.00")
                            End If
                        Else
                            cSatz = cSatz & ";" & rsrs.Fields(i)
                        End If
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
    
    If Datendrin("MDE_EXPORT", gdBase) Then
        iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbYes Then
            gcBestellEmail.Attachment1 = cdatei
            Screen.MousePointer = 0
            frmWKL129.Show 1
        Else
            MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
        End If
        anzeige "normal", "", Label1(8)
    Else
        anzeige "rot", "Keine Daten zum Export vorhanden.", Label1(8)
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
        Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "IP_INV_GROUP", gdBase
    loeschNEW "MDE_EXPORT", gdBase
    loeschNEW "IP_INV", gdBase
    loeschNEW "ARTHISTE", gdBase

    LogtoEnd Me
    voreinstellungspeichern
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern()
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
    Dim bo12 As Integer
    Dim bo13 As Integer
    Dim bo14 As Integer
    Dim bo15 As Integer
    Dim bo16 As Integer
    Dim bo17 As Integer
    Dim bo18 As Integer
    
    Dim bo19 As Integer
    Dim bo20 As Integer
    
    loeschNEW "E46", gdBase
    CreateTable "E46", gdBase
    
    bo0 = Option2(0).Value
    bo1 = Option2(1).Value
    bo2 = Option2(2).Value
    bo3 = Option2(3).Value
    bo4 = Option2(4).Value
    bo5 = Option2(5).Value
    
    If Check1.Value = vbChecked Then
        bo6 = 0
    Else
        bo6 = -1
    End If
    
    bo7 = Option1(0).Value
    bo8 = Option1(1).Value
    bo9 = Option1(2).Value
    bo10 = Option1(3).Value
    
    bo11 = Option1(7).Value 'sek
    bo12 = Option1(4).Value 'lek
    
    bo13 = 1
    bo14 = 1
    
    If Check7.Value = vbChecked Then
        bo15 = 0
    Else
        bo15 = -1
    End If
    
    bo16 = Option2(6).Value
    bo17 = Option2(7).Value
    bo18 = Option2(8).Value
    
    bo19 = opt1(22).Value
    bo20 = opt1(23).Value
    
    sSQL = "Insert into E46 ( bo0,bo1,bo2,bo3,bo4,bo5,bo6,bo7,bo8,bo9,bo10,bo11,bo12,bo13,bo14,bo15,bo16,bo17,bo18,bo19,bo20) "
    sSQL = sSQL & " values (" & bo0 & "," & bo1 & "," & bo2 & "," & bo3 & "," & bo4
    sSQL = sSQL & " , " & bo5 & "," & bo6 & "," & bo7 & "," & bo8 & "," & bo9
    sSQL = sSQL & " , " & bo10 & "," & bo11 & "," & bo12 & "," & bo13 & "," & bo14 & "," & bo15
    sSQL = sSQL & " , " & bo16 & "," & bo17 & "," & bo18 & "," & bo19 & "," & bo20
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladen()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    
    Set rs = gdBase.OpenRecordset("E46")
    If Not rs.EOF Then
        Option2(0).Value = rs!bo0
        Option2(1).Value = rs!bo1
        Option2(2).Value = rs!bo2
        Option2(3).Value = rs!bo3
        Option2(4).Value = rs!bo4
        Option2(5).Value = rs!bo5
        
        If rs!bo6 = True Then
            Check1.Value = vbUnchecked
        Else
            Check1.Value = vbChecked
        End If
        
        Option1(0).Value = rs!bo7
        Option1(1).Value = rs!bo8
        Option1(2).Value = rs!bo9
        Option1(3).Value = rs!bo10
        
        Option1(7).Value = rs!bo11
        Option1(4).Value = rs!bo12
        
        If rs!bo15 = True Then
            Check7.Value = vbUnchecked
        Else
            Check7.Value = vbChecked
        End If
        
        Option2(6).Value = rs!bo16
        Option2(7).Value = rs!bo17
        Option2(8).Value = rs!bo18
        
        opt1(22).Value = rs!bo19
        opt1(23).Value = rs!bo20
    
    End If
    rs.Close: Set rs = Nothing
    
    If Option1(4).Value = True Then
        Frame36.Visible = True
    Else
        Frame36.Visible = False
    End If

     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichernE46il()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    Dim bo0 As Integer
    Dim bo1 As Integer
    Dim bo2 As Integer
    Dim bo3 As Integer
    Dim bo4 As Integer
   
    loeschNEW "E46LI", gdBase
    CreateTable "E46LI", gdBase
    
    bo0 = Option4(0).Value
    bo1 = Option4(1).Value
    bo2 = Option4(2).Value
    bo3 = Option4(3).Value
   
    If Check5.Value = vbChecked Then
        bo4 = 0
    Else
        bo4 = -1
    End If
    
    sSQL = "Insert into E46LI ( bo0,bo1,bo2,bo3,bo4) "
    sSQL = sSQL & " values (" & bo0 & "," & bo1 & "," & bo2 & "," & bo3 & "," & bo4
    sSQL = sSQL & " )"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichernE46LI"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladenE46IL()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    
    Set rs = gdBase.OpenRecordset("E46LI")
    If Not rs.EOF Then
    
        Option4(0).Value = rs!bo0
        Option4(1).Value = rs!bo1
        Option4(2).Value = rs!bo2
        Option4(3).Value = rs!bo3
        
        If rs!bo4 = True Then
            Check5.Value = vbUnchecked
        Else
            Check5.Value = vbChecked
        End If
        
        
    End If
    rs.Close: Set rs = Nothing

     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladenE46IL"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub invLiteRefresh()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    loeschNEW "INVLITE", gdBase
    CreateTable "INVLITE", gdBase
    
    cSQL = "Create Index ARTNR on INVLITE(ARTNR)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index EKPR on INVLITE(EKPR)"
    gdBase.Execute cSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "invLiteRefresh"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub MSFlexGrid1_Click()
    On Error GoTo LOKAL_ERROR
    
    Label0(2).Caption = "0"
    
    Label0(0).Caption = Trim$(Str$(MSFlexGrid1.Row))
    Label0(1).Caption = Trim$(Str$(MSFlexGrid1.Col))
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSHFlexGrid1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Dim cArtNr As String
    
    If MSHFlexGrid1.Row > 1 Then
        anzeigeNew "Normal", "Bestandsverlauf wird ermittelt...", Label6
    
        cArtNr = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
        listeArtikelhistorie "Bestand", cArtNr
        
        anzeigeNew "normal", "mehr Informationen mit Doppelklick auf ArtNr", Label6

    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
    Else
        If MSHFlexGrid1.Col = 8 Then
            If iKeypress = 0 And KeyCode <> vbKeyBack Then
                MSHFlexGrid1.Text = ""
            End If
            
            iKeypress = iKeypress + 1
        End If
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFlexGrid1_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    Dim lcol As Long
    Dim lrow As Long
    
    lcol = MSHFlexGrid1.Col
    lrow = MSHFlexGrid1.Row
    
    cZeichen = Chr$(KeyAscii)
    
    If lcol = 8 Then
        cValid = "1234567890" & Chr$(8)
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        End If
    End If
    
    If KeyAscii <> 0 Then
        MSHFlexGrid1.Row = lrow
        MSHFlexGrid1.Col = lcol
        cValid = MSHFlexGrid1.Text
        If InStr(cValid, ",") > 0 And cZeichen = "," Then
            KeyAscii = 0
        End If
        
        If KeyAscii <> 0 Then
            If KeyAscii <> 8 Then
                cValid = cValid & Chr$(KeyAscii)
            Else
                If Len(cValid) > 0 Then
                    cValid = Left(cValid, Len(cValid) - 1)
                End If
            End If
            MSHFlexGrid1.Text = cValid
        End If
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFlexGrid1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFlexGrid1_LeaveCell()
    On Error GoTo LOKAL_ERROR
    
    iKeypress = 0
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFlexGrid1_LeaveCell"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Option1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    If Option1(4).Value = True Then
        Frame36.Visible = True
    ElseIf Option1(7).Value = True Then
        Frame36.Visible = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    
    
    If Option3(1).Value = True Then
        Text4.Visible = True
        Text4.Text = ""
        Text4.SetFocus
        Command9.Visible = True
        Label9(0).Visible = True
        
        
        
        List8.Visible = True
        List8.Clear
        Command11.Visible = True
        Label9(2).Visible = True
        Command10.Visible = True
        
        
        Text6.Visible = False
        Text5.Visible = False
        Label9(1).Visible = False
        Label10.Visible = False
        
    ElseIf Option3(1).Value = False Then
        Text4.Visible = False
        Text4.Text = ""
        
        Command9.Visible = False
        Label9(0).Visible = False
        
        List8.Visible = False
        List8.Clear
        Command11.Visible = False
        Label9(2).Visible = False
        Command10.Visible = False
    
    End If
    
    If Option3(2).Value = True Then
        Text6.Visible = True
        Text5.Visible = True
        Label9(1).Visible = True
        Label10.Visible = True
        
        
        List8.Visible = False
        List8.Clear
        Command11.Visible = False
        Label9(2).Visible = False
        Command10.Visible = False
        
        Text4.Visible = False
        Text4.Text = ""
        
        Command9.Visible = False
        Label9(0).Visible = False
        Label7(19).Caption = ""
    Else
        Text6.Visible = False
        Text5.Visible = False
        Label9(1).Visible = False
        Label10.Visible = False
        Label7(19).Caption = ""
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option3_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Option4_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR

    If NewTableSuchenDBKombi("li46", gdBase) Then
        zeigeartikel Option4(Index).Tag
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option4_Click"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text2.BackColor = glSelBack1
    Label0(2).Caption = "1"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    If Index = 1 Then

        cValid = "1234567890-" & Chr$(8)
        cZeichen = Chr$(KeyAscii)
    
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        Else
            KeyAscii = Asc(cZeichen)
        End If
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub txtAGN_Change()
On Error GoTo LOKAL_ERROR
    
    If Len(txtAGN.Text) >= 3 Then
        Label13.Caption = Ermittleagntext(txtAGN.Text)
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtAgn_Change"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub txtAgn_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    txtAGN.BackColor = glSelBack1
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtAgn_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        Command6_Click 0
    End If
    
    If KeyCode = vbKeyEscape Then
        Command2_Click 10
    End If
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            Case Is = 1     'ArtNr
                If Combo2.Text = "" Then
                    MsgBox "Bitte einen Lieferanten angeben!", vbCritical, "STOP!"
                    Exit Sub
                End If
                gF2Prompt.cFeld = "ARTNR"
                gF2Prompt.cWert = Trim(ErmittleLinr(Combo2.Text))
            
            Case Is = 3     'AGN
                gF2Prompt.cFeld = "AGN"
                

        End Select
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
            If gF2Prompt.cWahl <> "" Then
                Text1(Index).Text = gF2Prompt.cWahl
            End If
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    If KeyCode = vbKeyReturn Then
        If Frame15.Visible = False Then
            Command1_Click
        End If
    End If
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = "LINR"
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
        End If
        
        If gF2Prompt.cWahl <> "" Then
            ctmp = gF2Prompt.cWahl
            ctmp = ctmp & String$(6 - Len(ctmp), "_")
            Text2.Text = ctmp
        End If
        Text2.SetFocus
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "text2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    sortierenGrid MSFlexGrid1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long

    lrow = MSFlexGrid1.Row
    lcol = MSFlexGrid1.Col
    
    If KeyCode <> vbKeyDown And KeyCode <> vbKeyUp And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft Then
    
        Select Case lcol
            Case Is = SpaltennummerMENGE
        
                If iKeypress = 0 And KeyCode <> vbKeyBack Then
                    If KeyCode = 187 Or KeyCode = 189 Or KeyCode = 107 Or KeyCode = 109 Or KeyCode = vbKeyF2 Or KeyCode = vbKeyF3 Or KeyCode = vbKeyF4 Or KeyCode = vbKeyF5 Or KeyCode = vbKeyF6 Then
                    
                    Else
                        MSFlexGrid1.Row = lrow
                        MSFlexGrid1.Col = lcol
                        MSFlexGrid1.Text = ""
                    
                    End If
                    
                ElseIf iKeypress > 0 And KeyCode = 46 Then
                
                    MSFlexGrid1.Row = lrow
                    MSFlexGrid1.Col = lcol
                    MSFlexGrid1.Text = ""
                
                End If
                iKeypress = iKeypress + 1
            
        End Select
    End If
    
'    If iKeypress = 0 And KeyCode <> vbKeyBack Then
'        MSFlexGrid1.Row = lrow
'        MSFlexGrid1.Col = lcol
'        MSFlexGrid1.Text = ""
'    End If
'    iKeypress = iKeypress + 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    Dim lcol As Long
    Dim lrow As Long
    
    lcol = MSFlexGrid1.Col
    lrow = MSFlexGrid1.Row
    
    cZeichen = Chr$(KeyAscii)
    
    Select Case lcol
        Case Is = 2
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 3 To 5
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
    End Select
    
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
                    cValid = Left(cValid, Len(cValid) - 1)
                End If
            End If
            
            MSFlexGrid1.Text = cValid
            
        End If
        
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
'Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
'    On Error GoTo LOKAL_ERROR
'
'    Dim cZeichen As String
'    Dim cValid As String
'    Dim lcol As Long
'    Dim lrow As Long
'
'    lcol = MSFlexGrid1.Col
'    lrow = MSFlexGrid1.Row
'
''    lbl6(0).Caption = lrow
'
'    cZeichen = Chr$(KeyAscii)
'
'    Select Case lcol
'        Case Is = SpaltennummerMenge
'
'            cValid = "1234567890" & Chr$(8)
'            If InStr(cValid, cZeichen) = 0 Then
'                KeyAscii = 0
'            End If
'
'            If KeyAscii <> 0 Then
'                MSFlexGrid1.Row = lrow
'                MSFlexGrid1.Col = lcol
'                cValid = MSFlexGrid1.Text
'                If InStr(cValid, ",") > 0 And cZeichen = "," Then
'                    KeyAscii = 0
'                End If
'
'                If KeyAscii <> 0 Then
'                    If KeyAscii <> 8 Then
'                        cValid = cValid & Chr$(KeyAscii)
'                    Else
'                        If Len(cValid) > 0 Then
'                            cValid = Left$(cValid, Len(cValid) - 1)
'                        End If
'                    End If
'                    MSFlexGrid1.Text = cValid
'                End If
'            End If
'    End Select
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "MSFlexGrid1_KeyPress"
'    Fehler.gsFehlertext = "Im Programmteil  Bestellung ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    
     Select Case KeyCode
        Case Is = 46    'Del
            MSFlexGrid1.Text = ""
            
        Case Is = vbKeyF2
        
            lrow = MSFlexGrid1.Row
            gsARTNR = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr)
            If gsARTNR <> "" Then
    
                frmWKL10.Show 1
                Me.Refresh
                Screen.MousePointer = 11
                MSFlexGrid1.Col = SpaltennummerMENGE
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.TopRow = lrow
                MSFlexGrid1.SetFocus
                Screen.MousePointer = 0
            End If
            gsARTNR = ""
    
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus(Index As Integer)
On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = Len(Text1(Index).Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer, Index As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    
    Select Case Index
        Case Is = 101, 6, 0, 5, 7
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
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtAGN_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR

    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    cValid = "1234567890" & Chr$(8)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtAgn_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtAgn_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtAGN.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtAgn_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text85_GotFocus()
On Error GoTo LOKAL_ERROR

    Text85.BackColor = glSelBack1
    Text85.SelStart = Len(Text85.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text8_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text85_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    
    
    cValid = "1234567890" & Chr$(8)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
       
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text85_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Text85_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text85.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text85_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text87_GotFocus()
On Error GoTo LOKAL_ERROR

    Text87.BackColor = glSelBack1
    Text87.SelStart = Len(Text87.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text87_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text87_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    
    
    cValid = "1234567890" & Chr$(8)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
       
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text87_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Text87_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text87.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text87_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text3(Index).BackColor = glSelBack1
    Text3(Index).SelStart = 0
    Text3(Index).SelLength = Len(Text3(Index).Text)
    
    Label11.Caption = Trim$(Str$(Index))
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If Index = 0 And (KeyCode = 187 Or KeyCode = 106) Then
        If Len(Text3(0).Text) > 1 Then
            Text3(1).Text = Left(Text3(0).Text, Len(Text3(0).Text) - 1)
            Text3(0).Text = ""
            Text3(0).SetFocus
        End If
    End If
    
    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            Command2_Click 11
        ElseIf Index = 1 Then
            
            Text3(1).Text = ""
            Text3(0).Text = ""
            Command2_Click 11
        End If
    End If
    
    If KeyCode = vbKeyEscape Then
        Command2_Click 7
    End If
    
    If KeyCode = vbKeyRight Then
        Text3(0).SetFocus
    End If
    
    If KeyCode = vbKeyLeft Then
        Text3(1).SetFocus
    End If
    
    If Index = 1 Then
           
        If KeyCode = vbKeyUp Then
            Text3(1).Text = CInt(Text3(1).Text) + 1
        End If
        
        If KeyCode = vbKeyDown Then
            Text3(1).Text = CInt(Text3(1).Text) - 1
        End If
        
    End If
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_LostFocus(Index As Integer)
On Error GoTo LOKAL_ERROR

   Text3(Index).BackColor = vbWhite
   
'    If IsNumeric(Text3(1).Text) Or Text3(1).Text = "" Then
'        Text3(Index).BackColor = vbWhite
'    Else
'        MsgBox "Korrigieren Sie bitte Ihre letzte Eingabe!", , "Winkiss Hinweis:"
'        Text3(1).Text = "1"
'        Text3(1).SetFocus
'        Exit Sub
'    End If
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = "LINR"
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
        End If
        
        If gF2Prompt.cWahl <> "" Then
            ctmp = gF2Prompt.cWahl
            ctmp = ctmp & String$(6 - Len(ctmp), "_")
            Text4.Text = ctmp
        End If
        Text4.SetFocus
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "text4_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    
    
    cValid = gcUPPER & gcLower & gcNUM & Chr$(8)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text7_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    
    Dim cValid As String
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    cValid = gcUPPER & gcLower & "1234567890" & Chr$(8)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text9_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Timer1_Timer()
    On Error GoTo LOKAL_ERROR
    
    Command7_Click
    
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Timer2_Timer()
    On Error GoTo LOKAL_ERROR
    
    Command8_Click
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer2_Timer"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub txtAGN_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        gF2Prompt.cFeld = "AGN"
        frmWK00a.Show 1
        If gF2Prompt.cWahl <> "" Then
            txtAGN.Text = gF2Prompt.cWahl
        End If
                
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtAgn_KeyUp"
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
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Exportiere_for_MDE()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    loeschNEW "MDE_EXPORT", gdBase
    CreateTableT2 "MDE_EXPORT", gdBase
    
    cSQL = "Update ARTIKEL set EAN = '' where EAN is null"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update ARTEAN_K set EAN = '' where EAN is null"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into MDE_EXPORT Select "
    cSQL = cSQL & " EAN "
    cSQL = cSQL & ",'' as  BEZEICH "
    cSQL = cSQL & ", 0 as KVKPR1 "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & " from ARTEAN_K "
    cSQL = cSQL & " where Len(EAN) > 0 "
    cSQL = cSQL & " AND Val(EAN) > 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update MDE_EXPORT a inner join Artikel b on a.artnr = b.artnr Set "
    cSQL = cSQL & " a.BEZEICH = b.BEZEICH "
    cSQL = cSQL & ", a.KVKPR1 = b.KVKPR1 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into MDE_EXPORT Select "
    cSQL = cSQL & " EAN "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", KVKPR1 "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & " from ARTIKEL "
    cSQL = cSQL & " where Len(EAN) > 0 "
    cSQL = cSQL & " AND Val(EAN) > 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into MDE_EXPORT Select "
    cSQL = cSQL & " EAN2 as EAN "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", KVKPR1 "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & " from ARTIKEL "
    cSQL = cSQL & " where Len(EAN2) > 0 "
    cSQL = cSQL & " AND Val(EAN) > 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into MDE_EXPORT Select "
    cSQL = cSQL & " EAN3 as EAN "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", KVKPR1 "
    cSQL = cSQL & ", ARTNR "
    cSQL = cSQL & " from ARTIKEL "
    cSQL = cSQL & " where Len(EAN3) > 0 "
    cSQL = cSQL & " AND Val(EAN) > 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    ExportCSV
    

Exit Sub

LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Exportiere_for_MDE"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

