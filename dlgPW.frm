VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form dlgPW 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   5205
   ClientLeft      =   585
   ClientTop       =   675
   ClientWidth     =   8625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   5205
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      Height          =   2895
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   8655
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   120
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "1"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   1
         Left            =   840
         TabIndex        =   9
         Top             =   120
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "2"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   2
         Left            =   1440
         TabIndex        =   10
         Top             =   120
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "3"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   3
         Left            =   2040
         TabIndex        =   11
         Top             =   120
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "4"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   4
         Left            =   2640
         TabIndex        =   12
         Top             =   120
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "5"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   5
         Left            =   3240
         TabIndex        =   13
         Top             =   120
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "6"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   6
         Left            =   3840
         TabIndex        =   14
         Top             =   120
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "7"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   7
         Left            =   4440
         TabIndex        =   15
         Top             =   120
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "8"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   8
         Left            =   5040
         TabIndex        =   16
         Top             =   120
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "9"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   9
         Left            =   5640
         TabIndex        =   17
         Top             =   120
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "0"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   10
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "Q"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   11
         Left            =   840
         TabIndex        =   19
         Top             =   720
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "W"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   12
         Left            =   1440
         TabIndex        =   20
         Top             =   720
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "E"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   13
         Left            =   2040
         TabIndex        =   21
         Top             =   720
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "R"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   14
         Left            =   2640
         TabIndex        =   22
         Top             =   720
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "T"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   15
         Left            =   3240
         TabIndex        =   23
         Top             =   720
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "Z"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   16
         Left            =   3840
         TabIndex        =   24
         Top             =   720
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "U"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   17
         Left            =   4440
         TabIndex        =   25
         Top             =   720
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "I"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   18
         Left            =   5040
         TabIndex        =   26
         Top             =   720
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "O"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   19
         Left            =   5640
         TabIndex        =   27
         Top             =   720
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "P"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   20
         Left            =   6240
         TabIndex        =   28
         Top             =   720
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "�"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   21
         Left            =   480
         TabIndex        =   29
         Top             =   1320
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "A"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   22
         Left            =   1080
         TabIndex        =   30
         Top             =   1320
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "S"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   23
         Left            =   1680
         TabIndex        =   31
         Top             =   1320
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "D"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   24
         Left            =   2280
         TabIndex        =   32
         Top             =   1320
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "F"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   25
         Left            =   2880
         TabIndex        =   33
         Top             =   1320
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "G"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   26
         Left            =   3480
         TabIndex        =   34
         Top             =   1320
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "H"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   27
         Left            =   4080
         TabIndex        =   35
         Top             =   1320
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "J"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   28
         Left            =   4680
         TabIndex        =   36
         Top             =   1320
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "K"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   29
         Left            =   5280
         TabIndex        =   37
         Top             =   1320
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "L"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   30
         Left            =   5880
         TabIndex        =   38
         Top             =   1320
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "�"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   31
         Left            =   6480
         TabIndex        =   39
         Top             =   1320
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "�"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   32
         Left            =   720
         TabIndex        =   40
         Top             =   1920
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "Y"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   33
         Left            =   1320
         TabIndex        =   41
         Top             =   1920
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "X"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   34
         Left            =   1920
         TabIndex        =   42
         Top             =   1920
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "C"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   35
         Left            =   2520
         TabIndex        =   43
         Top             =   1920
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "V"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   36
         Left            =   3120
         TabIndex        =   44
         Top             =   1920
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "B"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   37
         Left            =   3720
         TabIndex        =   45
         Top             =   1920
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "N"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   38
         Left            =   4320
         TabIndex        =   46
         Top             =   1920
         Width           =   585
         _Version        =   65536
         _ExtentX        =   1032
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "M"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   39
         Left            =   4920
         TabIndex        =   47
         Top             =   1920
         Width           =   1680
         _Version        =   65536
         _ExtentX        =   2963
         _ExtentY        =   1032
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   585
         Index           =   40
         Left            =   6615
         TabIndex        =   48
         Top             =   1920
         Width           =   1785
         _Version        =   65536
         _ExtentX        =   3149
         _ExtentY        =   1032
         _StockProps     =   78
         Caption         =   "LEEREN"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0080FFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   8520
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00808000&
         Caption         =   "(Leertaste)"
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
         Left            =   4920
         TabIndex        =   7
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label0 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   255
         Left            =   4560
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8535
      Begin sevCommand3.Command Command1 
         Height          =   735
         Index           =   1
         Left            =   6000
         TabIndex        =   2
         Top             =   1560
         Width           =   2535
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
         Caption         =   "Abbruch"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   2535
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
         Caption         =   "OK"
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
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   840
         Width           =   8415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         Caption         =   "Passwort:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "dlgPW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_R�ckgabewert As Boolean

Public Property Get Back() As String
    Back = m_R�ckgabewert
End Property
Private Sub Command1_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case index
    
        Case Is = 0
            If Text1.Text = "" Then
                Text1.SetFocus
            Else
                If Text1.Text = "XYC6T349G6" Or Text1.Text = "xyc" Or Text1.Text = "XYC" Then
                    m_R�ckgabewert = True
                Else
                    m_R�ckgabewert = False
                End If
                Unload dlgPW
            End If
        Case Is = 1
            Unload dlgPW
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
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
Private Sub SSCommand2_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
        
    Select Case index
        Case Is = 40
            Text1.Text = ""
            Text1.SelStart = Len(Text1.Text)
            Text1.SetFocus
            
        Case Else
            If Len(Text1.Text) < Text1.MaxLength Then
                Text1.Text = Text1.Text & SSCommand2(index).Caption
                Text1.SelStart = Len(Text1.Text)
                Text1.SetFocus
            End If
    End Select

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand2_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, Nothing
    
    Text1.Text = ""
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = glSelBack1 'glSelBack1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = gcNUM & gcUPPER & Chr$(8)
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
    If KeyAscii <> 0 And KeyAscii <> 8 Then
        If Len(Text1.Text) = Text1.MaxLength - 1 Then
            Command1(0).SetFocus
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = 13 Then
        Command1_Click 0
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
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
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub






