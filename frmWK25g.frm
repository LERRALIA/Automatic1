VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWK25g 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Arbeitszeit-Listen"
   ClientHeight    =   8625
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWK25g.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Caption         =   "Arbeitszeit nachtragen"
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
      TabIndex        =   32
      Top             =   6600
      Visible         =   0   'False
      Width           =   6015
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "GEHT"
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
         Index           =   1
         Left            =   1200
         TabIndex        =   42
         Top             =   3000
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "KOMMT"
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
         Index           =   0
         Left            =   1200
         TabIndex        =   41
         Top             =   2520
         Value           =   -1  'True
         Width           =   2415
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   615
         Index           =   1
         Left            =   2160
         TabIndex        =   44
         Top             =   4200
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Schließen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   615
         Index           =   0
         Left            =   360
         TabIndex        =   43
         Top             =   4200
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Speichern"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   40
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   33
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Uhrzeit:"
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
         Index           =   6
         Left            =   360
         TabIndex        =   39
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Index           =   5
         Left            =   360
         TabIndex        =   38
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BedNr:"
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
         Index           =   4
         Left            =   360
         TabIndex        =   37
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   36
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   35
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Datum:"
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
         Index           =   3
         Left            =   360
         TabIndex        =   34
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame Frame0 
      BackColor       =   &H00C0C000&
      Height          =   2055
      Left            =   0
      TabIndex        =   30
      Top             =   4560
      Width           =   6015
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   11
         Left            =   3720
         TabIndex        =   17
         Top             =   1080
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   1296
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   10
         Left            =   3000
         TabIndex        =   16
         Top             =   1080
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   9
         Left            =   2280
         TabIndex        =   15
         Top             =   1080
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "9"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   8
         Left            =   1560
         TabIndex        =   14
         Top             =   1080
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   7
         Left            =   840
         TabIndex        =   13
         Top             =   1080
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "7"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   5
         Left            =   3720
         TabIndex        =   11
         Top             =   360
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   1296
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   4
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   3
         Left            =   2280
         TabIndex        =   9
         Top             =   360
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   2
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   1
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label0 
         Caption         =   "Label3"
         Height          =   255
         Left            =   3720
         TabIndex        =   31
         Top             =   2160
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Arbeitzeit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   6000
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
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
         Height          =   6990
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   5655
      End
      Begin VB.ListBox List2 
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
         TabIndex        =   29
         Top             =   360
         Width           =   5655
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   735
         Index           =   2
         Left            =   3960
         TabIndex        =   21
         Top             =   7680
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "Drucken"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   735
         Index           =   1
         Left            =   2040
         TabIndex        =   20
         Top             =   7680
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "Einfügen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   7680
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "Löschen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Mitarbeiter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   6015
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
         Height          =   2220
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   5775
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   2
         Top             =   3360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   1
         Top             =   3360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   2
         Left            =   3960
         TabIndex        =   5
         Top             =   3840
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Schließe Dialog"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   1
         Left            =   2040
         TabIndex        =   4
         Top             =   3840
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Leere Dialog"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   3840
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3413
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Suche Zeiten"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Zeitraum:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "bis"
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
         Left            =   2280
         TabIndex        =   26
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "von"
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
         Left            =   360
         TabIndex        =   25
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   24
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmWK25g"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DisableFrame1WK25g()
    On Error GoTo LOKAL_ERROR
    
    List1.Enabled = False
    MaskEdBox1(0).Enabled = False
    MaskEdBox1(1).Enabled = False
    SSCommand1(0).Enabled = False
    SSCommand1(1).Enabled = False
    SSCommand1(2).Enabled = False
    Frame1.Enabled = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DisableFrame1WK25g"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    
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
Private Sub DruckeArbeitszeitWK25g()
    On Error GoTo LOKAL_ERROR
    
    Dim cDrucker As String
    Dim cTmp As String
    Dim bReturn As Boolean
    Dim lAnz As Long
    Dim lcount As Long
    Dim cTab As String
    Dim cLbSatz As String
    
    cTab = Space$(10)
    
    'Auf Listendrucker umschalten
    
    setzedrucker gcListenDrucker


    Printer.Print
    Printer.Print
    Printer.FontName = "Courier New"
    Printer.Print
    Printer.Print cTab & "AUFSTELLUNG ARBEITSZEIT"
    Printer.Print cTab & "-----------------------"
    Printer.Print
    Printer.Print
    Printer.Print cTab & "BedNr.   : " & Label1(0).Caption
    Printer.Print cTab & "BedName  : " & Label1(1).Caption
    Printer.Print cTab & "Datum Von: " & MaskEdBox1(0).Text
    Printer.Print cTab & "Datum Bis: " & MaskEdBox1(1).Text
    Printer.Print
    Printer.Print
        
    lAnz = List3.ListCount
    
    For lcount = 0 To lAnz - 1
        cLbSatz = List3.list(lcount)
        If InStr(cLbSatz, "--->") > 0 Then
           Printer.Print cTab & cLbSatz
        End If
    Next lcount
    
    Printer.EndDoc
    
    'auf BonDrucker zurückschalten
    
    setzedrucker gcBonDrucker

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckeArbeitszeitWK25g"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub

Private Sub EnableFrame1WK25g()
    On Error GoTo LOKAL_ERROR
    
    List1.Enabled = True
    MaskEdBox1(0).Enabled = True
    MaskEdBox1(1).Enabled = True
    SSCommand1(0).Enabled = True
    SSCommand1(1).Enabled = True
    SSCommand1(2).Enabled = True
    Frame1.Enabled = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EnableFrame1WK25g"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub

Private Sub DisableFrame2WK25g()
    On Error GoTo LOKAL_ERROR
    
    List2.Enabled = False
    List3.Enabled = False
    SSCommand3(0).Enabled = False
    SSCommand3(1).Enabled = False
    SSCommand3(2).Enabled = False
    Frame2.Enabled = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DisableFrame2WK25g"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub

Private Sub EnableFrame2WK25g()
    On Error GoTo LOKAL_ERROR
    
    List2.Enabled = False
    List3.Enabled = True
    SSCommand3(0).Enabled = True
    SSCommand3(1).Enabled = True
    SSCommand3(2).Enabled = True
    Frame2.Enabled = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EnableFrame2WK25g"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub


Private Sub ErfasseArbeitszeitWK25g()
    On Error GoTo LOKAL_ERROR
    
    Label1(2).Caption = Label1(0).Caption
    Label1(3).Caption = Label1(1).Caption
    MaskEdBox1(2).Text = "__.__.____"
    MaskEdBox1(3).Text = "__:__:__"
    
    
    
    DisableFrame1WK25g
    DisableFrame2WK25g
    
    Frame3.Visible = True
    MaskEdBox1(2).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErfasseArbeitszeitWK25g"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
   
End Sub

Private Function fnPruefeEingabeDialogWK25g() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld As String
    Dim cVon As String
    Dim cBis As String
    Dim lVon As Long
    Dim lBis As Long
    
    fnPruefeEingabeDialogWK25g = 0
    
    cFeld = Label1(0).Caption
    cFeld = Trim$(cFeld)
    If cFeld = "" Then
        fnPruefeEingabeDialogWK25g = 1
        Exit Function
    End If
    
    cVon = MaskEdBox1(0).Text
    cBis = MaskEdBox1(1).Text
    
    If cVon = "__.__.____" And cBis <> "__.__.____" Then
        cVon = cBis
    End If
    
    If cVon <> "__.__.____" And cBis = "__.__.____" Then
        cBis = cVon
    End If
    
    MaskEdBox1(0).Text = cVon
    MaskEdBox1(1).Text = cBis
    
    If Not IsDate(cVon) Then
        fnPruefeEingabeDialogWK25g = 2
        Exit Function
    End If
    
    If Not IsDate(cBis) Then
        fnPruefeEingabeDialogWK25g = 3
        Exit Function
    End If
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    
    If lVon > lBis Then
        fnPruefeEingabeDialogWK25g = 4
        Exit Function
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeDialogWK25g"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Function

Private Function fnPruefeEingabeFrame3WK25g() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld As String
    Dim cdatum As String
    Dim czeit As String
    Dim cHH As String
    Dim cMM As String
    Dim cSS As String
    
    fnPruefeEingabeFrame3WK25g = 0
    
    cdatum = MaskEdBox1(2).Text
    
    If Not IsDate(cdatum) Then
        fnPruefeEingabeFrame3WK25g = 1
        Exit Function
    End If
    
    czeit = MaskEdBox1(3).Text
    cHH = Mid(czeit, 1, 2)
    cMM = Mid(czeit, 4, 2)
    cSS = Mid(czeit, 7, 2)
    
    cHH = Trim$(Str$(Val(cHH)))
    cHH = String$(2 - Len(cHH), "0") & cHH
    cMM = Trim$(Str$(Val(cMM)))
    cMM = String$(2 - Len(cMM), "0") & cMM
    cSS = Trim$(Str$(Val(cSS)))
    cSS = String$(2 - Len(cSS), "0") & cSS
    
    If Val(cHH) > 24 Then
        fnPruefeEingabeFrame3WK25g = 2
        Exit Function
    End If
        
    If Val(cMM) > 59 Then
        fnPruefeEingabeFrame3WK25g = 2
        Exit Function
    End If
        
    If Val(cSS) > 59 Then
        fnPruefeEingabeFrame3WK25g = 2
        Exit Function
    End If
        
    czeit = cHH & ":" & cMM & ":" & cSS
    MaskEdBox1(3).Text = czeit
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeDialogWK25g"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Function
Private Sub LeseArbeitszeitWK25g()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lbednu As Long
    Dim lDatVon As Long
    Dim lDatBis As Long
    Dim cFeld As String
    Dim cLbSatz As String
    
    Dim lDatMerker As Long
    Dim dZeit As Double
    Dim dTag As Double
    Dim cart As String
    Dim dWert As Double
    Dim cTmp As String
    
    Dim lMonMerker As Long
    Dim dMonWert As Double
    
    List3.Clear
    
    lbednu = Val(Label1(0).Caption)
    lDatVon = DateValue(MaskEdBox1(0).Text)
    lDatBis = DateValue(MaskEdBox1(1).Text)
    
    cSQL = "Select * from ARBEIT "
    cSQL = cSQL & "where BEDNU = " & Trim$(Str$(lbednu)) & " "
    cSQL = cSQL & "and DATUM >= " & Trim$(Str$(lDatVon)) & " "
    cSQL = cSQL & "and DATUM <= " & Trim$(Str$(lDatBis)) & " "
    cSQL = cSQL & "order by DATUM, ZEIT"

    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        lDatMerker = 0
        lMonMerker = 0
        dMonWert = 0
        Do While Not rsrs.EOF
            cLbSatz = ""
            If Not IsNull(rsrs!Datum) Then
                lDatVon = rsrs!Datum
            Else
                lDatVon = 0
            End If
            'Ändert sich das Tagesdatum?
            If lDatMerker <> lDatVon Then
                If lDatMerker <> 0 Then
                    cFeld = Format$(lDatMerker, "DD.MM.YYYY")
                    cLbSatz = cLbSatz & cFeld & " "
                    cLbSatz = cLbSatz & "-------------> TAG: " & " "
                    dWert = 1440 * dTag
                    cFeld = Trim$(Str$(Fix(dWert / 60)))
                    cFeld = cFeld & ":"
                    
                    cTmp = Trim$(Str$(Abs(dWert Mod 60)))
                    cTmp = String$(2 - Len(cTmp), "0") & cTmp
                    cFeld = cFeld & cTmp
                    
                    cLbSatz = cLbSatz & cFeld
                    List3.AddItem cLbSatz
                    cLbSatz = ""
                    lDatMerker = lDatVon
                    dTag = 0
                    
                    If lMonMerker <> Month(lDatVon) Then
                        dMonWert = dMonWert + dWert
                        cLbSatz = UCase$(gcMonat(lMonMerker))
                        cLbSatz = cLbSatz & Space$(9 - Len(cLbSatz))
                        cLbSatz = cLbSatz & "  "
                        cLbSatz = cLbSatz & "--------------------------> MONAT:" & " "
                        cFeld = Trim$(Str$(Fix(dMonWert / 60)))
                        cFeld = cFeld & ":"
                        
                        cTmp = Trim$(Str$(Abs(dMonWert Mod 60)))
                        cTmp = String$(2 - Len(cTmp), "0") & cTmp
                        cFeld = cFeld & cTmp
                        
                        cLbSatz = cLbSatz & cFeld
                        List3.AddItem cLbSatz
                        cLbSatz = ""
                        lMonMerker = Month(lDatVon)
                        dMonWert = 0
                        
                    Else
                        dMonWert = dMonWert + dWert
                    End If
                Else
                    lDatMerker = lDatVon
                    lMonMerker = Month(lDatVon)
                End If
            End If
            cFeld = Format$(lDatVon, "DD.MM.YYYY")
            cLbSatz = cLbSatz & cFeld & " "
            
            If Not IsNull(rsrs!art) Then
                cFeld = rsrs!art
            Else
                cFeld = ""
            End If
            cart = cFeld
            cFeld = cFeld & Space$(5 - Len(cFeld))
            cLbSatz = cLbSatz & cFeld & " "
            
            If Not IsNull(rsrs!zeit) Then
                cFeld = rsrs!zeit
            Else
                cFeld = "00:00:00"
            End If
            dZeit = TimeValue(cFeld)
            cFeld = cFeld & Space$(8 - Len(cFeld))
            cLbSatz = cLbSatz & cFeld & " "
            
            If cart = "KOMMT" Then
                dTag = dTag - dZeit
            Else
                dTag = dTag + dZeit
            End If
            
            
            List3.AddItem cLbSatz
            
            rsrs.MoveNext
        Loop
        
        'Nachlauf Gruppenwechsel für TAG
        cLbSatz = ""
        cFeld = Format$(lDatMerker, "DD.MM.YYYY")
        cLbSatz = cLbSatz & cFeld & " "
        cLbSatz = cLbSatz & "-------------> TAG: " & " "
        dWert = 1440 * dTag
        
        cFeld = Trim$(Str$(Fix(dWert / 60)))
        cFeld = cFeld & ":"
        
        cTmp = Trim$(Str$(Abs(dWert Mod 60)))
        cTmp = String$(2 - Len(cTmp), "0") & cTmp
        cFeld = cFeld & cTmp
        
        cLbSatz = cLbSatz & cFeld
        
        List3.AddItem cLbSatz
        
        'Nachlauf Gruppenwechsel für MONAT
        dMonWert = dMonWert + dWert
        cLbSatz = UCase$(gcMonat(lMonMerker))
        cLbSatz = cLbSatz & Space$(9 - Len(cLbSatz))
        cLbSatz = cLbSatz & "  "
        cLbSatz = cLbSatz & "--------------------------> MONAT:" & " "
        cFeld = Trim$(Str$(Fix(dMonWert / 60)))
        cFeld = cFeld & ":"
        
        cTmp = Trim$(Str$(Abs(dMonWert Mod 60)))
        cTmp = String$(2 - Len(cTmp), "0") & cTmp
        cFeld = cFeld & cTmp
        
        cLbSatz = cLbSatz & cFeld
        List3.AddItem cLbSatz
        cLbSatz = ""
        lMonMerker = Month(lDatVon)
        dMonWert = 0

        
        
        Frame2.Visible = True
    
    Else
        MsgBox "Keine Daten gefunden.", vbInformation, "Winkiss Hinweis:"
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseArbeitszeitWK25g"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub

Private Sub LeseBedienerWK25g()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cLbSatz As String
    
    List1.Clear
    
    cSQL = "Select * from BEDNAME where BEDNU <> 99 order by BEDNAME, BEDNU"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cLbSatz = ""
            If Not IsNull(rsrs!BEDNU) Then
                cFeld = rsrs!BEDNU
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLbSatz = cLbSatz & cFeld & " "
            
            If Not IsNull(rsrs!BEDNAME) Then
                cFeld = rsrs!BEDNAME
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cLbSatz = cLbSatz & cFeld & " "
            
            List1.AddItem cLbSatz
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseBedienerWK25g"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub

Private Sub LoescheArbeitszeitWK25g()
    On Error GoTo LOKAL_ERROR
    
    Dim cLbSatz As String
    Dim cdatum As String
    Dim cart As String
    Dim cUhrZeit As String
    Dim cbednu As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim ldatum As Long
    
    If List3.ListIndex < 0 Then
        MsgBox "Bitte einen Eintrag in der Liste auswählen!", vbCritical, "STOP!"
        List3.SetFocus
        Exit Sub
    End If
        
    cLbSatz = List3.list(List3.ListIndex)
    cdatum = Mid(cLbSatz, 1, 10)
    cart = Mid(cLbSatz, 12, 5)
    cUhrZeit = Mid(cLbSatz, 18, 8)
    cbednu = Label1(0).Caption
    ldatum = DateValue(cdatum)
    
    cSQL = "Select * from ARBEIT where "
    cSQL = cSQL & "BEDNU = " & cbednu & " "
    cSQL = cSQL & "and DATUM = " & Trim$(Str$(ldatum)) & " "
    cSQL = cSQL & "and ZEIT = '" & cUhrZeit & "' "
    cSQL = cSQL & "and ART = '" & cart & "' "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.delete
        List3.RemoveItem List3.ListIndex
    End If
    rsrs.Close: Set rsrs = Nothing
        
    SSCommand1_Click 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheArbeitszeitWK25g"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
End Sub

Private Sub SchreibeNeueArbeitszeitWK25g()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lbednu As Long
    Dim cbedname As String
    Dim ldatum As Long
    Dim czeit As String
    Dim cart As String
    
    lbednu = Val(Label1(2).Caption)
    cbedname = Label1(3).Caption
    ldatum = DateValue(MaskEdBox1(2).Text)
    czeit = MaskEdBox1(3).Text
    If Option1(0).Value = True Then
        cart = "KOMMT"
    Else
        cart = "GEHT"
    End If
    
    cSQL = "Select * from ARBEIT where "
    cSQL = cSQL & "BEDNU = " & Trim$(Str$(lbednu)) & " "
    cSQL = cSQL & "and BEDNAME = '" & cbedname & "' "
    cSQL = cSQL & "and DATUM = " & Trim$(Str$(ldatum)) & " "
    cSQL = cSQL & "and ZEIT = '" & czeit & "' "
    cSQL = cSQL & "and ART = '" & cart & "' "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        rsrs.AddNew
    Else
        rsrs.Edit
    End If
        
    rsrs!BEDNU = lbednu
    rsrs!BEDNAME = cbedname
    rsrs!Datum = ldatum
    rsrs!zeit = czeit
    rsrs!art = cart
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    SSCommand4_Click 1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeNeueArbeitszeitWK25g"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    positionierenwkl25g
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    Label1(0).Caption = ""
    Label1(1).Caption = ""
    
    LeseBedienerWK25g
    
    List2.AddItem "Datum      Art   Uhrzeit  Tages-Saldo"
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub
Private Sub positionierenwkl25g()
    On Error GoTo LOKAL_ERROR
    
    Frame3.Top = 0
    Frame3.Left = 6000
    Frame3.Height = 8775
    Frame3.Width = 6015
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "positionierenwkl25g"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub

Private Sub List1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cLbSatz As String
    
    cLbSatz = List1.list(List1.ListIndex)
    Label1(0).Caption = Trim$(Mid(cLbSatz, 1, 3))
    Label1(1).Caption = Trim$(Mid(cLbSatz, 4, Len(cLbSatz) - 3))
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List1_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub
Private Sub MaskEdBox1_GotFocus(Index As Integer)
On Error GoTo LOKAL_ERROR

    Label0.Caption = Trim$(Str$(Index))
    MaskEdBox1(Index).BackColor = glSelBack1
    MaskEdBox1(Index).SelStart = 0
    MaskEdBox1(Index).SelLength = Len(MaskEdBox1(Index).Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

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
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub
Private Sub SSCommand1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    Dim iRet As Integer
    
    Select Case Index
        Case Is = 0
            iRet = fnPruefeEingabeDialogWK25g()
            Select Case iRet
                Case Is = 0     'alles okay
                    LeseArbeitszeitWK25g
                    
                Case Is = 1     'kein Bediener
                    MsgBox "Bitte wählen Sie einen Eintrag aus!", vbCritical, "STOP!"
                    List1.SetFocus
                    
                Case Is = 2     'ungültiges VON-Datum
                    MsgBox "Das eingegeben VON-Datum ist ungültig!", vbCritical, "STOP!"
                    MaskEdBox1(0).SetFocus
                
                Case Is = 3     'ungültiges BIS-Datum
                    MsgBox "Das eingegeben BIS-Datum ist ungültig!", vbCritical, "STOP!"
                    MaskEdBox1(1).SetFocus
                
                Case Is = 4     'VON ist größer als BIS
                    MsgBox "Das VON-Datum ist größer als das BIS-Datum!", vbCritical, "STOP!"
                    MaskEdBox1(0).SetFocus
            
            End Select
            
        Case Is = 1
            Label1(0).Caption = ""
            Label1(1).Caption = ""
            MaskEdBox1(0).Text = "__.__.____"
            MaskEdBox1(1).Text = "__.__.____"
            List3.Clear
            Frame2.Visible = False
            List1.SetFocus
            
        Case Is = 2
            Unload frmWK25g
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub


Private Sub SSCommand2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lFeld As Long
    Dim cTmp As String
    Dim cZeichen As String
    Dim iCount As Integer
    
    lFeld = Val(Label0.Caption)
    
    cTmp = MaskEdBox1(lFeld).Text
    
    Select Case Index
        Case 0 To 4, 6 To 10
            cZeichen = SSCommand2(Index).Caption
            For iCount = 1 To Len(cTmp)
                If Mid(cTmp, iCount, 1) = "_" Then
                    Mid(cTmp, iCount, 1) = cZeichen
                    Exit For
                End If
            Next iCount
            MaskEdBox1(lFeld).Text = cTmp
            
        Case Is = 5
            If lFeld < 3 Then
                For iCount = 10 To 1 Step -1
                    cZeichen = Mid(cTmp, iCount, 1)
                    If cZeichen <> "." And cZeichen <> "_" Then
                        If iCount = 3 Or iCount = 6 Then
                            Mid(cTmp, iCount, 1) = "."
                            Exit For
                        Else
                            Mid(cTmp, iCount, 1) = "_"
                            Exit For
                        End If
                    End If
                Next iCount
            Else
                For iCount = 8 To 1 Step -1
                    cZeichen = Mid(cTmp, iCount, 1)
                    If cZeichen <> ":" And cZeichen <> "_" Then
                        If iCount = 3 Or iCount = 6 Then
                            Mid(cTmp, iCount, 1) = ":"
                            Exit For
                        Else
                            Mid(cTmp, iCount, 1) = "_"
                            Exit For
                        End If
                    End If
                Next iCount
            End If
            MaskEdBox1(lFeld).Text = cTmp
            
        Case Is = 11
            MaskEdBox1(lFeld).Text = "__.__.____"
    End Select
    
    MaskEdBox1(lFeld).SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand2_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub


Private Sub SSCommand3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0     'Löschen
            LoescheArbeitszeitWK25g
            
        Case Is = 1     'Einfügen
            ErfasseArbeitszeitWK25g
            
        Case Is = 2     'Drucken
            DruckeArbeitszeitWK25g
            
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand3_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub


Private Sub SSCommand4_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    
    Select Case Index
        Case Is = 0     'Speichern
            iRet = fnPruefeEingabeFrame3WK25g()
            Select Case iRet
                Case Is = 0     'okay
                    SchreibeNeueArbeitszeitWK25g
                    
                Case Is = 1     'Fehler Datum
                    MsgBox "Das Datum ist ungültig!", vbCritical, "STOP!"
                    MaskEdBox1(2).SetFocus
                    
                Case Is = 2     'Fehler Uhrzeit
                    MsgBox "Die Uhrzeit ist ungültig!", vbCritical, "STOP!"
                    MaskEdBox1(3).SetFocus
                
            End Select
        Case Is = 1     'Schließen
            EnableFrame1WK25g
            EnableFrame2WK25g
            Frame3.Visible = False
            SSCommand1_Click 0
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand4_Click"
    Fehler.gsFehlertext = "Im Programmteil Arbeitszeitlisten ist ein Fehler aufgetreten. "

    Fehlermeldung1
    
End Sub


