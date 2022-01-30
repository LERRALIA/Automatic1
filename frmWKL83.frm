VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmWKL83 
   Caption         =   "Termine"
   ClientHeight    =   8595
   ClientLeft      =   1155
   ClientTop       =   1815
   ClientWidth     =   11880
   Icon            =   "frmWKL83.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Termine für:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3735
      Left            =   6000
      TabIndex        =   75
      Top             =   1080
      Width           =   5775
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame0 
      BackColor       =   &H00808000&
      Height          =   3375
      Left            =   0
      TabIndex        =   12
      Top             =   4920
      Width           =   11775
      Begin sevCommand3.Command Command1 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "neue Zeile"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   9480
         TabIndex        =   74
         Top             =   1440
         Width           =   2175
      End
      Begin sevCommand3.Command Command1 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "Rückgängig"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   9480
         TabIndex        =   73
         Top             =   840
         Width           =   2175
      End
      Begin sevCommand3.Command Command1 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "Feld leeren"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   9480
         TabIndex        =   72
         Top             =   240
         Width           =   2175
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "a > A"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   58
         Left            =   8880
         TabIndex        =   71
         Top             =   2640
         Width           =   1095
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   " "
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   57
         Left            =   7560
         TabIndex        =   70
         Top             =   2640
         Width           =   1335
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "-"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   56
         Left            =   6840
         TabIndex        =   69
         Top             =   2640
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "."
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   55
         Left            =   6120
         TabIndex        =   68
         Top             =   2640
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   ","
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   54
         Left            =   5400
         TabIndex        =   67
         Top             =   2640
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "m"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   53
         Left            =   4680
         TabIndex        =   66
         Top             =   2640
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "n"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   52
         Left            =   3960
         TabIndex        =   65
         Top             =   2640
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "b"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   51
         Left            =   3240
         TabIndex        =   64
         Top             =   2640
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "v"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   50
         Left            =   2520
         TabIndex        =   63
         Top             =   2640
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "c"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   49
         Left            =   1800
         TabIndex        =   62
         Top             =   2640
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "x"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   48
         Left            =   1080
         TabIndex        =   61
         Top             =   2640
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "y"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   47
         Left            =   480
         TabIndex        =   60
         Top             =   2640
         Width           =   615
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "#"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   46
         Left            =   8640
         TabIndex        =   59
         Top             =   2040
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "ä"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   45
         Left            =   7920
         TabIndex        =   58
         Top             =   2040
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "ö"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   44
         Left            =   7200
         TabIndex        =   57
         Top             =   2040
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "l"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   43
         Left            =   6480
         TabIndex        =   56
         Top             =   2040
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "k"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   42
         Left            =   5760
         TabIndex        =   55
         Top             =   2040
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "j"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   41
         Left            =   5040
         TabIndex        =   54
         Top             =   2040
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "h"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   40
         Left            =   4320
         TabIndex        =   53
         Top             =   2040
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "g"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   39
         Left            =   3600
         TabIndex        =   52
         Top             =   2040
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "f"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   38
         Left            =   2880
         TabIndex        =   51
         Top             =   2040
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "d"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   37
         Left            =   2160
         TabIndex        =   50
         Top             =   2040
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "s"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   36
         Left            =   1440
         TabIndex        =   49
         Top             =   2040
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "a"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   35
         Left            =   720
         TabIndex        =   48
         Top             =   2040
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "+"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   34
         Left            =   8400
         TabIndex        =   47
         Top             =   1440
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "ü"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   33
         Left            =   7680
         TabIndex        =   46
         Top             =   1440
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "p"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   32
         Left            =   6960
         TabIndex        =   45
         Top             =   1440
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "o"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   31
         Left            =   6240
         TabIndex        =   44
         Top             =   1440
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "i"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   30
         Left            =   5520
         TabIndex        =   43
         Top             =   1440
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "u"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   29
         Left            =   4800
         TabIndex        =   42
         Top             =   1440
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "z"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   28
         Left            =   4080
         TabIndex        =   41
         Top             =   1440
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "t"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   27
         Left            =   3360
         TabIndex        =   40
         Top             =   1440
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "r"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   26
         Left            =   2640
         TabIndex        =   39
         Top             =   1440
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "e"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   25
         Left            =   1920
         TabIndex        =   38
         Top             =   1440
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "w"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   24
         Left            =   1200
         TabIndex        =   37
         Top             =   1440
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "q"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   23
         Left            =   480
         TabIndex        =   36
         Top             =   1440
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "ß"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   22
         Left            =   7320
         TabIndex        =   35
         Top             =   840
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "0"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   21
         Left            =   6600
         TabIndex        =   34
         Top             =   840
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "9"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   20
         Left            =   5880
         TabIndex        =   33
         Top             =   840
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "8"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   19
         Left            =   5160
         TabIndex        =   32
         Top             =   840
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "7"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   18
         Left            =   4440
         TabIndex        =   31
         Top             =   840
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "6"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   17
         Left            =   3720
         TabIndex        =   30
         Top             =   840
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "5"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   16
         Left            =   3000
         TabIndex        =   29
         Top             =   840
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "4"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   15
         Left            =   2280
         TabIndex        =   28
         Top             =   840
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "3"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   14
         Left            =   1560
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "2"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   13
         Left            =   840
         TabIndex        =   26
         Top             =   840
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "1"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   12
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "*"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   11
         Left            =   8040
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "?"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   7320
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "="
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   6600
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   ")"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   5880
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "("
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   5160
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "/"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   4440
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "&&"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   3720
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "%"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   3000
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "$"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   2280
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "§"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "´"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin sevCommand3.Command Command0 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "!"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Wochenübersicht"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3735
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   6015
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5953
         _Version        =   393216
         Rows            =   8
         FixedRows       =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
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
      Height          =   1095
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11775
      Begin sevCommand3.Command Command2 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "Schließen"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   9840
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin sevCommand3.Command Command2 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "Löschen"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   7920
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin sevCommand3.Command Command2 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "Speichern"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   6000
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin sevCommand3.Command Command2 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "- Woche"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin sevCommand3.Command Command2 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "+ Woche"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin sevCommand3.Command Command2 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "- Monat"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   3000
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin sevCommand3.Command Command2 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "+ Monat"
         BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   4440
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Montag, der 27.03.2000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frmWKL83"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LeseTagesTermineWKL83()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum As String
    Dim lDatum As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cDatum = Label1.Caption
    cDatum = Right(cDatum, 10)
    
    lDatum = DateValue(cDatum)
    
    Text1.Text = ""
        
    Frame1.Caption = "Notizen für den " & cDatum
    
    cSQL = "Select * from NOTIZEN where DATUM = " & Trim$(Str$(lDatum)) & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!NOTIZ) Then
            Text1.Text = rsrs!NOTIZ
        End If
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseTagesTermineWKL83"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub LoescheNotizenWKL83()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cDatum As String
    Dim lDatum As Long
    
    lDatum = MsgBox("Wollen Sie die Notizen zum aktuell gewählten Datum wirklich löschen?", vbYesNo + vbQuestion, "LÖSCHEN")
    
    If lDatum <> vbYes Then
        Exit Sub
    End If
        
    cDatum = Frame1.Caption
    cDatum = Right(cDatum, 10)
    lDatum = DateValue(cDatum)
    
    cSQL = "Select * from NOTIZEN where DATUM = " & Trim$(Str$(lDatum)) & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.delete
        Text1.Text = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheNotizenWKL83"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub SchreibeNotizenWKL83()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cDatum As String
    Dim lDatum As Long
    Dim iWoTag As Integer
    
    cDatum = Frame1.Caption
    cDatum = Right(cDatum, 10)
    lDatum = DateValue(cDatum)
    
    cSQL = "Select * from NOTIZEN where DATUM = " & Trim$(Str$(lDatum)) & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    
    rsrs!Datum = lDatum
    rsrs!NOTIZ = Text1.Text
    
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    iWoTag = Weekday(lDatum, vbMonday)
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Row = iWoTag - 1
    MSFlexGrid1.Text = Text1.Text
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeNotizenWKL83"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    

End Sub

Private Sub SetzeWochenUebersichtWKL83()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum As String
    Dim lrow As Long
    Dim lDatum As Long
    Dim lStart As Long
    Dim lEnde As Long
    Dim lCount As Long
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim iWoTag As Integer
    Dim iZeile As Integer
    
    MSFlexGrid1.Col = 0
    For lrow = 0 To 6
        MSFlexGrid1.Row = lrow
        Select Case lrow
            Case Is = 0
                MSFlexGrid1.Text = "MO"
            Case Is = 1
                MSFlexGrid1.Text = "DI"
            Case Is = 2
                MSFlexGrid1.Text = "MI"
            Case Is = 3
                MSFlexGrid1.Text = "DO"
            Case Is = 4
                MSFlexGrid1.Text = "FR"
            Case Is = 5
                MSFlexGrid1.Text = "SA"
            Case Is = 6
                MSFlexGrid1.Text = "SO"
        End Select
    Next lrow
        
    cDatum = Label1.Caption
    cDatum = Right(cDatum, 10)
    lDatum = DateValue(cDatum)
    
    iWoTag = Weekday(lDatum, vbMonday)
    
    lStart = lDatum - iWoTag
    lStart = lStart + 1
    
    lEnde = lStart + 6      'nicht 7 !!!
    
    iZeile = 0
    
    For lCount = lStart To lEnde
        
        cSQL = "Select * from NOTIZEN where DATUM = " & Trim$(Str$(lCount)) & " "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        MSFlexGrid1.Col = 1
        MSFlexGrid1.Row = iZeile
        If Not rsrs.EOF Then
            
            If Not IsNull(rsrs!NOTIZ) Then
                MSFlexGrid1.Text = rsrs!NOTIZ
            Else
                MSFlexGrid1.Text = ""
            End If
        Else
            MSFlexGrid1.Text = ""
        End If
        rsrs.Close: Set rsrs = Nothing
        iZeile = iZeile + 1
    Next lCount
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SetzeWochenUebersichtWKL83"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub WocheZurueckWKL83()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum As String
    Dim lDatum As Long
    Dim iWoTag As Integer
    
    cDatum = Label1.Caption
    cDatum = Right(cDatum, 10)
    lDatum = DateValue(cDatum)
    
    lDatum = lDatum - 7
    
    iWoTag = Weekday(lDatum, vbMonday)
    
    cDatum = Format$(lDatum, "DD.MM.YYYY")
    Label1.Caption = gcWochentag(iWoTag) & ", der " & cDatum
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WocheZurueckWKL83"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iCount As Integer
    
    If Index <> 58 Then
        Text1.Text = Text1.Text & Command0(Index).Caption
    Else
        If Command0(Index).Caption = "a > A" Then
            For iCount = 23 To 53
                Command0(iCount).Caption = UCase$(Command0(iCount).Caption)
            Next iCount
            Command0(54).Caption = ";"
            Command0(55).Caption = ":"
            Command0(56).Caption = "_"
            Command0(Index).Caption = "A > a"
        Else
            For iCount = 23 To 53
                Command0(iCount).Caption = LCase$(Command0(iCount).Caption)
            Next iCount
            Command0(54).Caption = ","
            Command0(55).Caption = "."
            Command0(56).Caption = "-"
            Command0(Index).Caption = "a > A"
        End If
    End If
    Text1.SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cTmp As String
    
    Select Case Index
        Case Is = 0
            Text1.Text = ""
        Case Is = 1
            cTmp = Text1.Text
            If Len(cTmp) > 0 Then
                cTmp = Left(cTmp, Len(cTmp) - 1)
            End If
            Text1.Text = cTmp
        Case Is = 2
            cTmp = Text1.Text
            cTmp = cTmp & vbCrLf
            Text1.Text = cTmp
    End Select
    
    Text1.SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    Select Case Index
        Case Is = 0     'eine Woche zurück
            WocheZurueckWKL83
            
        Case Is = 1     'eine Woche vor
            WocheVorWKL83
            
        Case Is = 2     'einen Monat zurück
            MonatZurueckWKL83
            
        Case Is = 3     'einen Monat vor
            MonatVorWKL83
            
        Case Is = 4     'Speichern
            SchreibeNotizenWKL83
            
        Case Is = 5     'Löschen
            LoescheNotizenWKL83
            
        Case Is = 6
            Unload frmWKL83
    End Select
    
    If Index < 4 Then
        SetzeWochenUebersichtWKL83
    
        LeseTagesTermineWKL83
    End If
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MonatVorWKL83()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum As String
    Dim lDatum As Long
    Dim cTT As String
    Dim cMM As String
    Dim cYYYY As String
    Dim iTT As Integer
    Dim iMM As Integer
    Dim iYYYY As Integer
    Dim iWoTag As Integer
    
    cDatum = Label1.Caption
    cDatum = Right(cDatum, 10)
    
    cTT = Mid(cDatum, 1, 2)
    cMM = Mid(cDatum, 4, 2)
    cYYYY = Mid(cDatum, 7, 4)
    
    iTT = Val(cTT)
    iMM = Val(cMM)
    iYYYY = Val(cYYYY)
    iMM = iMM + 1
    If iMM > 12 Then
        iMM = 1
        iYYYY = Val(cYYYY)
        iYYYY = iYYYY + 1
    End If
    
    cMM = Trim$(Str$(iMM))
    cYYYY = Trim$(Str$(iYYYY))
    
    cDatum = cTT & "." & cMM & "." & cYYYY
    
    Do While Not IsDate(cDatum)
        iTT = iTT - 1
        cTT = Trim$(Str$(iTT))
        cDatum = cTT & "." & cMM & "." & cYYYY
    Loop
    
    lDatum = DateValue(cDatum)
    
    iWoTag = Weekday(lDatum, vbMonday)
    
    cDatum = Format$(lDatum, "DD.MM.YYYY")
    Label1.Caption = gcWochentag(iWoTag) & ", der " & cDatum
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MonatVorWKL83"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MonatZurueckWKL83()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum As String
    Dim lDatum As Long
    Dim cTT As String
    Dim cMM As String
    Dim cYYYY As String
    Dim iTT As Integer
    Dim iMM As Integer
    Dim iYYYY As Integer
    Dim iWoTag As Integer
    
    cDatum = Label1.Caption
    cDatum = Right(cDatum, 10)
    
    cTT = Mid(cDatum, 1, 2)
    cMM = Mid(cDatum, 4, 2)
    cYYYY = Mid(cDatum, 7, 4)
    
    iTT = Val(cTT)
    iMM = Val(cMM)
    iYYYY = Val(cYYYY)
    iMM = iMM - 1
    If iMM < 1 Then
        iMM = 12
        iYYYY = Val(cYYYY)
        iYYYY = iYYYY - 1
    End If
    
    cMM = Trim$(Str$(iMM))
    cYYYY = Trim$(Str$(iYYYY))
    
    cDatum = cTT & "." & cMM & "." & cYYYY
    
    Do While Not IsDate(cDatum)
        iTT = iTT - 1
        cTT = Trim$(Str$(iTT))
        cDatum = cTT & "." & cMM & "." & cYYYY
    Loop
    
    lDatum = DateValue(cDatum)
    
    iWoTag = Weekday(lDatum, vbMonday)
    
    cDatum = Format$(lDatum, "DD.MM.YYYY")
    Label1.Caption = gcWochentag(iWoTag) & ", der " & cDatum
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MonatZurueckWKL83"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub WocheVorWKL83()
    On Error GoTo LOKAL_ERROR
    
    Dim cDatum As String
    Dim lDatum As Long
    Dim iWoTag As Integer
    
    cDatum = Label1.Caption
    cDatum = Right(cDatum, 10)
    lDatum = DateValue(cDatum)
    
    lDatum = lDatum + 7
    
    iWoTag = Weekday(lDatum, vbMonday)
    
    cDatum = Format$(lDatum, "DD.MM.YYYY")
    Label1.Caption = gcWochentag(iWoTag) & ", der " & cDatum
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WocheVorWKL83"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim lcol As Long
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    MSFlexGrid1.Rows = 7
    MSFlexGrid1.Cols = 2
    
    For lrow = 0 To 6
        MSFlexGrid1.Row = lrow
        MSFlexGrid1.RowHeight(lrow) = 500
    Next lrow
    
    MSFlexGrid1.Col = 0
    MSFlexGrid1.ColWidth(0) = 500
  
    MSFlexGrid1.Col = 1
    MSFlexGrid1.ColWidth(1) = 8500
    
    HoleTagesDatumWKL83

    SetzeWochenUebersichtWKL83
    
    LeseTagesTermineWKL83
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub


Private Sub HoleTagesDatumWKL83()
    On Error GoTo LOKAL_ERROR
    
    Dim lheute As Long
    Dim cHeute As String
    Dim lWochenTag As String
    Dim cWochenTag As String
    
    lheute = Fix(Now)
    cHeute = Format$(lheute, "DD.MM.YYYY")
    lWochenTag = Weekday(lheute, vbMonday)
    
    Select Case lWochenTag
        Case Is = 1
            cWochenTag = "Montag"
        Case Is = 2
            cWochenTag = "Dienstag"
        Case Is = 3
            cWochenTag = "Mittwoch"
        Case Is = 4
            cWochenTag = "Donnerstag"
        Case Is = 5
            cWochenTag = "Freitag"
        Case Is = 6
            cWochenTag = "Samstag"
        Case Is = 7
            cWochenTag = "Sonntag"
    End Select
    
    Label1.Caption = cWochenTag & ", der " & cHeute
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HoleTagesDatumWKL83"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub MSFlexGrid1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim iWoTag As Integer
    Dim cDatum As String
    Dim lDatum As Long
    Dim iRow As Integer
    Dim Ldiff As Long
    
    iRow = MSFlexGrid1.Row
    iRow = iRow + 1
    
    cDatum = Label1.Caption
    cDatum = Right(cDatum, 10)
    lDatum = DateValue(cDatum)
    
    iWoTag = Weekday(lDatum, vbMonday)
    
    Ldiff = iRow - iWoTag
    
    
    lDatum = lDatum + Ldiff
    
    cDatum = Format$(lDatum, "DD.MM.YYYY")
    
    Label1.Caption = gcWochentag(iRow) & ", der " & cDatum
    
    Frame1.Caption = "Notizen für den " & cDatum
    
    LeseTagesTermineWKL83
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

'Private Sub MSFlexGrid1_EnterCell()
'    If MSFlexGrid1.Row > -1 And MSFlexGrid1.Col > 0 Then
'        MSFlexGrid1.CellBackColor = &HC00000
'        MSFlexGrid1.CellForeColor = vbYellow
'    End If
'
'End Sub


'Private Sub MSFlexGrid1_LeaveCell()
'    If MSFlexGrid1.Row > -1 And MSFlexGrid1.Col > 0 Then
'        MSFlexGrid1.CellBackColor = vbWhite
'        MSFlexGrid1.CellForeColor = vbBlack
'    End If
'
'End Sub


Private Sub Text1_GotFocus()
On Error GoTo LOKAL_ERROR

    Text1.SelStart = Len(Text1.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub


