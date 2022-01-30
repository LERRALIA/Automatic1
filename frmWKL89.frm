VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL89 
   Caption         =   "Spaltenbezeichnungen und ihre Bedeutung"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   Icon            =   "frmWKL89.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   10155
   StartUpPosition =   1  'Fenstermitte
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   0
      Top             =   8760
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
      Caption         =   "Schlieﬂen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   255
      Index           =   0
      Left            =   8400
      TabIndex        =   73
      Top             =   80
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
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
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   34
      Left            =   2880
      TabIndex        =   72
      Top             =   8520
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   33
      Left            =   2880
      TabIndex        =   71
      Top             =   8280
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   32
      Left            =   2880
      TabIndex        =   70
      Top             =   8040
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   34
      Left            =   120
      TabIndex        =   69
      Top             =   8520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   33
      Left            =   120
      TabIndex        =   68
      Top             =   8280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   32
      Left            =   120
      TabIndex        =   67
      Top             =   8040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   2880
      TabIndex        =   66
      Top             =   7800
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   120
      TabIndex        =   65
      Top             =   7800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   120
      TabIndex        =   64
      Top             =   7560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   2880
      TabIndex        =   63
      Top             =   7560
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   2880
      TabIndex        =   62
      Top             =   7320
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   120
      TabIndex        =   61
      Top             =   7320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   2880
      TabIndex        =   60
      Top             =   7080
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   120
      TabIndex        =   59
      Top             =   7080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   2880
      TabIndex        =   58
      Top             =   6840
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   120
      TabIndex        =   57
      Top             =   6840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   2880
      TabIndex        =   56
      Top             =   6600
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   120
      TabIndex        =   55
      Top             =   6600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   2880
      TabIndex        =   54
      Top             =   6360
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   120
      TabIndex        =   53
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   2880
      TabIndex        =   52
      Top             =   6120
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   120
      TabIndex        =   51
      Top             =   6120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   2880
      TabIndex        =   50
      Top             =   5880
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   120
      TabIndex        =   49
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   2880
      TabIndex        =   48
      Top             =   5640
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   120
      TabIndex        =   47
      Top             =   5640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   2880
      TabIndex        =   46
      Top             =   5400
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   120
      TabIndex        =   45
      Top             =   5400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   2880
      TabIndex        =   44
      Top             =   5160
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   120
      TabIndex        =   43
      Top             =   5160
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   2880
      TabIndex        =   42
      Top             =   4920
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   120
      TabIndex        =   41
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   2880
      TabIndex        =   40
      Top             =   4680
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   120
      TabIndex        =   39
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   2880
      TabIndex        =   38
      Top             =   4440
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   37
      Top             =   4440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   2880
      TabIndex        =   36
      Top             =   4200
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   120
      TabIndex        =   35
      Top             =   4200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   2880
      TabIndex        =   34
      Top             =   3960
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   120
      TabIndex        =   33
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   2880
      TabIndex        =   32
      Top             =   3720
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   120
      TabIndex        =   31
      Top             =   3720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   2880
      TabIndex        =   30
      Top             =   3480
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   120
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   2880
      TabIndex        =   28
      Top             =   3240
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   27
      Top             =   3240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   2880
      TabIndex        =   26
      Top             =   3000
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   25
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   2880
      TabIndex        =   24
      Top             =   2760
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   2880
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   2880
      TabIndex        =   20
      Top             =   2280
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   2880
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   2880
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2880
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2880
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   2880
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   2880
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
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
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Spaltenbezeichnungen und ihre Bedeutung"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmWKL89"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
        Case 1
            Unload frmWKL89
        Case 0
            drucken
    End Select
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Spaltenerl‰uterungen auf. "
    Fehlermeldung1
End Sub
Private Sub drucken()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim sSQL As String
    
    loeschNEW "BVOALI", gdBase
    CreateTable "BVOALI", gdBase
    
    For i = 0 To byAnzahlSpalten - 1

        sSQL = "Insert into BVOALI (spalte,alias,stab) values "
        sSQL = sSQL & " ( '" & Label3(i).Caption & "','" & Label4(i).Caption & "','" & Label2.Caption & "')"
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    reportbildschirm "", "aWKL89"
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucken"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Spaltenerl‰uterungen auf. "
    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    
    Modul6.alternativFarbform Me, Label1
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
 
    Label2.Caption = gsARTNR
    Label2.Refresh
    
    gsARTNR = ""
    
    For i = 0 To byAnzahlSpalten - 1
        Label3(i).Caption = sSpaltenname(i)
        Label3(i).Visible = True
        Label3(i).Refresh
        
        Label4(i).Caption = sSpaltenAli(i)
        Label4(i).Visible = True
        Label4(i).Refresh
    Next i

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Spaltenerl‰uterungen auf. "
    Fehlermeldung1
End Sub

