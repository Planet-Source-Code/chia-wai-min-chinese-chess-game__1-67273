VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chinese Chess v1.0"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6255
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrBlink 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   5760
   End
   Begin VB.Frame Frame1 
      Caption         =   "Additional Chess"
      Height          =   1455
      Left            =   960
      TabIndex        =   0
      Tag             =   "0"
      Top             =   7200
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Image picZu 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   2
         Left            =   3720
         Picture         =   "frmMain.frx":0ECA
         Top             =   840
         Width           =   600
      End
      Begin VB.Image picXg 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   2
         Left            =   3120
         Picture         =   "frmMain.frx":139D
         Top             =   840
         Width           =   600
      End
      Begin VB.Image picKg 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   2
         Left            =   2520
         Picture         =   "frmMain.frx":1891
         Top             =   840
         Width           =   600
      End
      Begin VB.Image picSi 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   2
         Left            =   1920
         Picture         =   "frmMain.frx":1D7B
         Top             =   840
         Width           =   600
      End
      Begin VB.Image picPo 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   2
         Left            =   1320
         Picture         =   "frmMain.frx":225B
         Top             =   840
         Width           =   600
      End
      Begin VB.Image picMa 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   2
         Left            =   720
         Picture         =   "frmMain.frx":275F
         Top             =   840
         Width           =   600
      End
      Begin VB.Image picCe 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   2
         Left            =   120
         Picture         =   "frmMain.frx":2C4D
         Top             =   840
         Width           =   600
      End
      Begin VB.Image picZu 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   1
         Left            =   3720
         Picture         =   "frmMain.frx":3129
         Top             =   240
         Width           =   600
      End
      Begin VB.Image picXg 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   1
         Left            =   3120
         Picture         =   "frmMain.frx":35F6
         Top             =   240
         Width           =   600
      End
      Begin VB.Image picKg 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   1
         Left            =   2520
         Picture         =   "frmMain.frx":3AE0
         Top             =   240
         Width           =   600
      End
      Begin VB.Image picSi 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   1
         Left            =   1920
         Picture         =   "frmMain.frx":3FD0
         Top             =   240
         Width           =   600
      End
      Begin VB.Image picPo 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   1
         Left            =   1320
         Picture         =   "frmMain.frx":4450
         Top             =   240
         Width           =   600
      End
      Begin VB.Image picMa 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   1
         Left            =   720
         Picture         =   "frmMain.frx":4953
         Top             =   240
         Width           =   600
      End
      Begin VB.Image picCe 
         DragMode        =   1  'Automatic
         Height          =   600
         Index           =   1
         Left            =   120
         Picture         =   "frmMain.frx":4E1E
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Image picCheck 
      Height          =   7125
      Left            =   -120
      Picture         =   "frmMain.frx":52EE
      Top             =   -360
      Visible         =   0   'False
      Width           =   6450
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   17
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   16
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   15
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   14
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   13
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   12
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   11
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   10
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   9
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   8
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   7
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   6
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   5
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   4
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   3
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   2
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape shpSteps 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      DrawMode        =   2  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   10
      Left            =   360
      Top             =   240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   94
      Left            =   5160
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   95
      Left            =   5160
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   96
      Left            =   5160
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   97
      Left            =   5160
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   98
      Left            =   5160
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   99
      Left            =   5160
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   91
      Left            =   5160
      Top             =   840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   92
      Left            =   5160
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   93
      Left            =   5160
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   83
      Left            =   4560
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   82
      Left            =   4560
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   81
      Left            =   4560
      Top             =   840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   71
      Left            =   3960
      Top             =   840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   61
      Left            =   3360
      Top             =   840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   51
      Left            =   2760
      Top             =   840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   41
      Left            =   2160
      Top             =   840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   31
      Left            =   1560
      Top             =   840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   21
      Left            =   960
      Top             =   840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   11
      Left            =   360
      Top             =   840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   90
      Left            =   5160
      Top             =   240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   89
      Left            =   4560
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   88
      Left            =   4560
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   87
      Left            =   4560
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   86
      Left            =   4560
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   85
      Left            =   4560
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   84
      Left            =   4560
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   80
      Left            =   4560
      Top             =   240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   77
      Left            =   3960
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   73
      Left            =   3960
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   70
      Left            =   3960
      Top             =   240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   69
      Left            =   3360
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   68
      Left            =   3360
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   66
      Left            =   3360
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   65
      Left            =   3360
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   64
      Left            =   3360
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   63
      Left            =   3360
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   62
      Left            =   3360
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   60
      Left            =   3360
      Top             =   240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   57
      Left            =   2760
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   50
      Left            =   2760
      Top             =   240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   40
      Left            =   2160
      Top             =   240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   30
      Left            =   1560
      Top             =   240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   20
      Left            =   960
      Top             =   240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   19
      Left            =   360
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   17
      Left            =   360
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   15
      Left            =   360
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   79
      Left            =   3960
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   78
      Left            =   3960
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   76
      Left            =   3960
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   75
      Left            =   3960
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   74
      Left            =   3960
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   72
      Left            =   3960
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   67
      Left            =   3360
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   59
      Left            =   2760
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   58
      Left            =   2760
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   56
      Left            =   2760
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   55
      Left            =   2760
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   54
      Left            =   2760
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   53
      Left            =   2760
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   52
      Left            =   2760
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   49
      Left            =   2160
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   48
      Left            =   2160
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   47
      Left            =   2160
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   46
      Left            =   2160
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   45
      Left            =   2160
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   44
      Left            =   2160
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   43
      Left            =   2160
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   42
      Left            =   2160
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   39
      Left            =   1560
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   38
      Left            =   1560
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   37
      Left            =   1560
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   36
      Left            =   1560
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   35
      Left            =   1560
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   34
      Left            =   1560
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   33
      Left            =   1560
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   32
      Left            =   1560
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   29
      Left            =   960
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   28
      Left            =   960
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   27
      Left            =   960
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   26
      Left            =   960
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   25
      Left            =   960
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   24
      Left            =   960
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   23
      Left            =   960
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   22
      Left            =   960
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   18
      Left            =   360
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   16
      Left            =   360
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   14
      Left            =   360
      Top             =   2640
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   13
      Left            =   360
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image picBox 
      DragMode        =   1  'Automatic
      Height          =   615
      Index           =   12
      Left            =   360
      Top             =   1440
      Width           =   615
   End
   Begin VB.Shape shpOutline 
      BorderWidth     =   2
      Height          =   5580
      Left            =   600
      Top             =   480
      Width           =   4980
   End
   Begin VB.Image picGame 
      Height          =   735
      Left            =   1080
      Picture         =   "frmMain.frx":D66D
      Top             =   2880
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.Image picLine 
      Height          =   525
      Left            =   960
      Picture         =   "frmMain.frx":E972
      Top             =   3000
      Width           =   4290
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuGameLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuDisplayAvailable 
         Caption         =   "Display Available"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Local Variables
    Dim CurrentTurn As Integer
    Dim CurrentKgIndex(1 To 2) As Integer

'Local Constant
    Private Const BlackSide = "1"
    Private Const OrangeSide = "2"
    Private Const EmptyBox = "000"

Private Function NewGame()
    Dim x As Integer

    'Reset All Tags
        For x = picBox.LBound To picBox.UBound
            picBox(x).Picture = Me.Picture
            picBox(x).Tag = EmptyBox
        Next

    'Put Che (Black)
        picBox(picBox.LBound).Picture = picCe(1).Picture: picBox(picBox.LBound).Tag = "Ce1"
        picBox(picBox.UBound - 9).Picture = picCe(1).Picture: picBox(picBox.UBound - 9).Tag = "Ce1"
        picBox(picBox.LBound + 9).Picture = picCe(2).Picture: picBox(picBox.LBound + 9).Tag = "Ce2"
        picBox(picBox.UBound).Picture = picCe(2).Picture: picBox(picBox.UBound).Tag = "Ce2"
        
    'Put Ma
        picBox(20).Picture = picMa(1).Picture: picBox(20).Tag = "Ma1"
        picBox(80).Picture = picMa(1).Picture: picBox(80).Tag = "Ma1"
        picBox(29).Picture = picMa(2).Picture: picBox(29).Tag = "Ma2"
        picBox(89).Picture = picMa(2).Picture: picBox(89).Tag = "Ma2"
    
    'Put Xiang
        picBox(30).Picture = picXg(1).Picture: picBox(30).Tag = "Xg1"
        picBox(70).Picture = picXg(1).Picture: picBox(70).Tag = "Xg1"
        picBox(39).Picture = picXg(2).Picture: picBox(39).Tag = "Xg2"
        picBox(79).Picture = picXg(2).Picture: picBox(79).Tag = "Xg2"
        
    'Put Shi
        picBox(40).Picture = picSi(1).Picture: picBox(40).Tag = "Si1"
        picBox(60).Picture = picSi(1).Picture: picBox(60).Tag = "Si1"
        picBox(49).Picture = picSi(2).Picture: picBox(49).Tag = "Si2"
        picBox(69).Picture = picSi(2).Picture: picBox(69).Tag = "Si2"
        
    'Put King
        picBox(50).Picture = picKg(1).Picture: picBox(50).Tag = "Kg1"
        picBox(59).Picture = picKg(2).Picture: picBox(59).Tag = "Kg2"
        
    'Put Pao
        picBox(22).Picture = picPo(1).Picture: picBox(22).Tag = "Po1"
        picBox(82).Picture = picPo(1).Picture: picBox(82).Tag = "Po1"
        picBox(27).Picture = picPo(2).Picture: picBox(27).Tag = "Po2"
        picBox(87).Picture = picPo(2).Picture: picBox(87).Tag = "Po2"
        
    'Put Zu
        picBox(13).Picture = picZu(1).Picture: picBox(13).Tag = "Zu1"
        picBox(33).Picture = picZu(1).Picture: picBox(33).Tag = "Zu1"
        picBox(53).Picture = picZu(1).Picture: picBox(53).Tag = "Zu1"
        picBox(73).Picture = picZu(1).Picture: picBox(73).Tag = "Zu1"
        picBox(93).Picture = picZu(1).Picture: picBox(93).Tag = "Zu1"
        picBox(16).Picture = picZu(2).Picture: picBox(16).Tag = "Zu2"
        picBox(36).Picture = picZu(2).Picture: picBox(36).Tag = "Zu2"
        picBox(56).Picture = picZu(2).Picture: picBox(56).Tag = "Zu2"
        picBox(76).Picture = picZu(2).Picture: picBox(76).Tag = "Zu2"
        picBox(96).Picture = picZu(2).Picture: picBox(96).Tag = "Zu2"
        
    'Set Drag Mode
        For x = picBox.LBound To picBox.UBound
            If picBox(x).Picture = Me.Picture Then
                picBox(x).DragMode = 0
            Else
                picBox(x).DragMode = 1
            End If
        Next
        
    'Set Default Kg Index
        CurrentKgIndex(1) = 50
        CurrentKgIndex(2) = 59
        
    'Set Current Turn To Blank
        CurrentTurn = 0
        
    'Set Picture
        picLine.Visible = True
        picGame.Visible = False
        
    'Hide Steps Shape
        Call shpStepsHide
End Function

Private Function DrawLine(FromIndex As Integer, ToIndex As Integer)
    'Draw Line From Box Index Passed
    Me.Line (picBox(FromIndex).Left + picBox(FromIndex).Width / 2, picBox(FromIndex).Top + picBox(FromIndex).Height / 2)-(picBox(ToIndex).Left + picBox(ToIndex).Width / 2, picBox(ToIndex).Top + (picBox(ToIndex).Height / 2))
End Function

Private Function DrawCross(ParamArray lstIndex())
    'Draw Cross For Pao and Zhu
    Dim CenterX As Integer, CenterY As Integer, x As Integer
    For x = LBound(lstIndex) To UBound(lstIndex)
        CenterX = picBox(lstIndex(x)).Left + picBox(lstIndex(x)).Width / 2
        CenterY = picBox(lstIndex(x)).Top + picBox(lstIndex(x)).Height / 2
        
        Me.Line (CenterX - 80, CenterY - 80)-(CenterX - 250, CenterY - 80)
        Me.Line (CenterX - 80, CenterY - 80)-(CenterX - 80, CenterY - 250)
        Me.Line (CenterX + 80, CenterY - 80)-(CenterX + 250, CenterY - 80)
        Me.Line (CenterX + 80, CenterY - 80)-(CenterX + 80, CenterY - 250)
        Me.Line (CenterX - 80, CenterY + 80)-(CenterX - 250, CenterY + 80)
        Me.Line (CenterX - 80, CenterY + 80)-(CenterX - 80, CenterY + 250)
        Me.Line (CenterX + 80, CenterY + 80)-(CenterX + 250, CenterY + 80)
        Me.Line (CenterX + 80, CenterY + 80)-(CenterX + 80, CenterY + 250)
    Next
End Function

Private Function RightDrawCross(ParamArray lstIndex())
    'Draw Half Cross For Zhu Beside Right Outline
    Dim CenterX As Integer, CenterY As Integer, x As Integer
    For x = LBound(lstIndex) To UBound(lstIndex)
        CenterX = picBox(lstIndex(x)).Left + picBox(lstIndex(x)).Width / 2
        CenterY = picBox(lstIndex(x)).Top + picBox(lstIndex(x)).Height / 2
        
        Me.Line (CenterX + 80, CenterY - 80)-(CenterX + 250, CenterY - 80)
        Me.Line (CenterX + 80, CenterY - 80)-(CenterX + 80, CenterY - 250)
        Me.Line (CenterX + 80, CenterY + 80)-(CenterX + 250, CenterY + 80)
        Me.Line (CenterX + 80, CenterY + 80)-(CenterX + 80, CenterY + 250)
    Next
End Function

Private Function LeftDrawCross(ParamArray lstIndex())
    'Draw Half Cross For Zhu Beside Left Outline
    Dim CenterX As Integer, CenterY As Integer, x As Integer
    For x = LBound(lstIndex) To UBound(lstIndex)
        CenterX = picBox(lstIndex(x)).Left + picBox(lstIndex(x)).Width / 2
        CenterY = picBox(lstIndex(x)).Top + picBox(lstIndex(x)).Height / 2
        
        Me.Line (CenterX - 80, CenterY - 80)-(CenterX - 250, CenterY - 80)
        Me.Line (CenterX - 80, CenterY - 80)-(CenterX - 80, CenterY - 250)
        Me.Line (CenterX - 80, CenterY + 80)-(CenterX - 250, CenterY + 80)
        Me.Line (CenterX - 80, CenterY + 80)-(CenterX - 80, CenterY + 250)
    Next
End Function

Private Function DrawBackground()
    Dim x As Integer
    
    'Vertical Line
    DrawLine picBox.LBound, picBox.LBound + 9
    DrawLine picBox.UBound - 9, picBox.UBound
    
    'Half Page Vertical Line
    For x = 2 To 8
        DrawLine Val(x & "0"), Val(x & "0") + 4
        DrawLine Val(x & "0") + 5, Val(x & "0") + 9
    Next
    
    'Horizontal Line
    For x = picBox.LBound To picBox.LBound + 9
         DrawLine x, x + 80
    Next
    
    '45 Degree Draw
    DrawLine 40, 62
    DrawLine 42, 60
    DrawLine 49, 67
    DrawLine 47, 69
    
    'Draw Cross Air
    DrawCross 22, 82, 33, 53, 73, 36, 56, 76, 27, 87
    
    'Half Cross Air
    RightDrawCross 13, 16
    LeftDrawCross 93, 96
    
    'Set Outline Pos
    shpOutline.Left = picBox(picBox.LBound).Left + 240
    shpOutline.Top = picBox(picBox.LBound).Top + 240
End Function

Private Sub Form_Load()
    'Draw Background
    Call DrawBackground
    
    'Start New Game
    Call NewGame
End Sub

Private Function getSide(orgTag As String) As String
    'Return Side
    getSide = Right(orgTag, 1)
End Function

Private Sub mnuAbout_Click()
    MsgBox "Produced by : Chia Wai Min " & vbCrLf & "Email : cwaimin@hotmail.com", vbInformation, "Chinese Chess"
End Sub

Private Sub mnuDisplayAvailable_Click()
    'Display Available Steps
    If mnuDisplayAvailable.Checked = False Then
        mnuDisplayAvailable.Checked = True
    Else
        mnuDisplayAvailable.Checked = False
        Call shpStepsHide
    End If
End Sub

Private Sub shpStepsHide()
    'Hide The Steps Object
    Dim x As Integer
    For x = shpSteps.LBound To shpSteps.UBound
        shpSteps(x).Visible = False
    Next
End Sub

Private Sub mnuGameExit_Click()
    'End The Program
    End
End Sub

Private Sub mnuNewGame_Click()
    'Start New Game
    Call NewGame
End Sub

Private Function chkSteps(Index As Integer)
    'Check Available Steps
    Dim x As Integer, objIndex As Integer
    
    objIndex = objIndex + 1
    If picBox(Index).Tag <> EmptyBox Then
        For x = picBox.LBound To picBox.UBound
            If getSide(picBox(Index).Tag) <> getSide(picBox(x).Tag) Then
                If ValidStep(x, picBox(Index)) Then
                    shpSteps(objIndex).Left = picBox(x).Left
                    shpSteps(objIndex).Top = picBox(x).Top
                    'shpSteps(objIndex).BorderColor = IIf(getSide(picBox(Index).Tag) = BlackSide, vbBlack, vbRed)
                    shpSteps(objIndex).Visible = True
                    objIndex = objIndex + 1
                End If
            End If
        Next
    End If
        
    If objIndex > 0 Then
        For x = objIndex To shpSteps.UBound
            shpSteps(x).Visible = False
        Next
    End If
End Function

Private Sub picBox_DragDrop(Index As Integer, Source As Control, x As Single, Y As Single)
    Dim tmpTag As String, PreKgIndex As String
    If CurrentTurn = 0 Then CurrentTurn = getSide(Source.Tag)
    If CurrentTurn <> getSide(Source.Tag) Then Exit Sub

    'Same Chess Location
    If getSide(picBox(Index).Tag) = getSide(Source.Tag) Then Exit Sub
    
    'Check Valid Step
    If ValidStep(Index, Source) = True Then
        'Keep Destination Index
        tmpTag = picBox(Index).Tag
        picBox(Index).Tag = Source.Tag
        Source.Tag = EmptyBox
        
        'Keep Current King Index For Restore
        If Left(picBox(Index).Tag, 2) = "Kg" Then
            PreKgIndex = CurrentKgIndex(getSide(picBox(Index).Tag))
            CurrentKgIndex(getSide(picBox(Index).Tag)) = Index
        End If
        
        'Check If Under Check By Opposite
        If IsChkMate(Opp(getSide(picBox(Index).Tag))) = True Then
            'Restore King Index
            If Left(picBox(Index).Tag, 2) = "Kg" Then
                CurrentKgIndex(Int(getSide(picBox(Index).Tag))) = PreKgIndex
            End If
        
            'Restore Source Index From Destination
            Source.Tag = picBox(Index).Tag
            picBox(Index).Tag = tmpTag
            tmrBlink.Enabled = True
        Else
            'Switch Player Turn
            CurrentTurn = Opp(getSide(picBox(Index).Tag))
            
            'Set Destination Drag Enable
            picBox(Index).DragMode = 1
            
            'Set Destination Picture
            picBox(Index).Picture = Source.Picture
            
            'Disable Source Drag
            Source.DragMode = 0
            
            'Clear Source Picture
            Source.Picture = Me.Picture
            
            'Disable Player Drag Chess
            Call DragHandle(getSide(picBox(Index).Tag))
            
            'If This Step Is Check The Opposite
            If IsChkMate(getSide(picBox(Index).Tag)) = True Then
                'Do Blink Check Effect
                tmrBlink.Enabled = True
                
                'Check If Opposite If Gamed
                If IsGameOver(Opp(getSide(picBox(Index).Tag))) = True Then Call GameOver
            End If
        End If
    End If
End Sub

Private Function IsGameOver(Side As String) As Boolean
    Dim x As Integer, Y As Integer, tmpTag As String, PreKgIndex As Integer
    
    'Start From First Box
    For x = picBox.LBound To picBox.UBound
        'If Found Own Chess
        If getSide(picBox(x).Tag) = Side Then
        
            'Start Check From First Box
            For Y = picBox.LBound To picBox.UBound
            
                'If The Step Is Valid
                If ValidStep(Y, picBox(x)) = True Then
                    
                    'Store King Index For Restore
                    If Left(picBox(x).Tag, 2) = "Kg" Then
                        PreKgIndex = CurrentKgIndex(Side)
                        CurrentKgIndex(Side) = Y
                    End If
                
                    'Store Source Data For Restore
                    tmpTag = picBox(Y).Tag
                    picBox(Y).Tag = picBox(x).Tag
                    picBox(x).Tag = EmptyBox
                    
                    'If Run This Step Still Under Check
                    If IsChkMate(Opp(Side)) = True Then
                        'Restore Source
                        picBox(x).Tag = picBox(Y).Tag
                        picBox(Y).Tag = tmpTag
                        
                        'Restore King
                        If Left(picBox(x).Tag, 2) = "Kg" Then _
                            CurrentKgIndex(Side) = PreKgIndex
                    Else
                        'Restore Source
                        picBox(x).Tag = picBox(Y).Tag
                        picBox(Y).Tag = tmpTag
                        
                        'Restore King
                        If Left(picBox(x).Tag, 2) = "Kg" Then _
                            CurrentKgIndex(Side) = PreKgIndex
                            
                        'Run This Step Cannot Avoid Check
                        IsGameOver = False
                        Exit Function
                    End If
                End If
            Next
        End If
    Next
    
    'Loop Until Finish Will Result Game Over
    IsGameOver = True
End Function

Private Function GameOver()
    'Game Over Display
    Dim x As Integer
    For x = picBox.LBound To picBox.UBound
        picBox(x).Tag = "000"
        picBox(x).DragMode = 0
    Next
    
    picLine.Visible = False
    picGame.Visible = True
End Function

Private Function DragHandle(Side As String)
    'Disable or Enable Drag
    Dim x As Integer
    For x = picBox.LBound To picBox.UBound
        If getSide(picBox(x).Tag) = Side Then
            picBox(x).DragMode = 0
        ElseIf getSide(picBox(x).Tag) = Opp(Side) Then
            picBox(x).DragMode = 1
        End If
    Next
End Function

Private Function Opp(Side As String) As String
    'Get Opposite Index
    Opp = IIf(Side = "1", "2", "1")
End Function

Private Function IsChkMate(Side As String) As Boolean
    Dim x As Integer
    'Start Check From First Box
    For x = picBox.LBound To picBox.UBound
        'Check If Is Own Chess
        If getSide(picBox(x).Tag) = Side Then
            'Excluded (Xiang and Shi)
            Select Case Left(picBox(x).Tag, 2)
                Case "Xg", "Si"
                Case Else
                    'Check If Own Chess Can Eat The Opposite King Chess
                    If ValidStep(CurrentKgIndex(Opp(Side)), picBox(x)) = True Then
                        'MsgBox CurrentKgIndex(Opp(Side))
                        IsChkMate = True
                        Exit Function
                    End If
            End Select
        End If
    Next
End Function

Private Function ValidStep(Index As Integer, Source As Control) As Boolean
    Select Case Left(Source.Tag, 2)
        Case "Ce" 'Ce Step Validation
            'Horizontal Check
            If picBox(Index).Left = Source.Left Then ValidStep = chkCe(Source.Index, Index, "")
            
            'Vertical Check
            If picBox(Index).Top = Source.Top Then ValidStep = chkCe(Left(Source.Index, 1), Left(Index, 1), Right(Index, 1))
        Case "Ma" 'Ma Step Validation
            ValidStep = chkMa(Index, Source)
        Case "Xg" 'Xiang Step Validation
            ValidStep = chkXg(Index, Source)
        Case "Si" 'Shi Step Validation
            ValidStep = chkSi(Index, Source)
        Case "Kg" 'King Step Validation
            ValidStep = chkKg(Index, Source)
        Case "Po" 'Pao Step Validation
            'Horizontal Check
            If picBox(Index).Left = Source.Left Then ValidStep = chkPo(Index, Source, Source.Index, Index, "")
            
            'Vertical Check
            If picBox(Index).Top = Source.Top Then ValidStep = chkPo(Index, Source, Left(Source.Index, 1), Left(Index, 1), Right(Index, 1))
        Case "Zu" 'Zhu / Bing Step Validation
            ValidStep = chkZu(Index, Source)
    End Select
End Function

Private Function chkCe(StartNum As Integer, EndNum As Integer, RightNum As String) As Boolean
    Dim x As Integer
    For x = StartNum To EndNum Step IIf(StartNum < EndNum, 1, -1)
        If Val(x & RightNum) <> Val(StartNum & RightNum) And Val(x & RightNum) <> Val(EndNum & RightNum) And picBox(Val(x & RightNum)).Tag <> EmptyBox Then Exit Function
        If x = EndNum Then chkCe = True
    Next
End Function

Private Function chkPo(Index As Integer, Source As Control, StartNum As Integer, EndNum As Integer, RightNum As String) As Boolean
    Dim FoundChess As Integer, IsEat As Boolean, x As Integer
    For x = StartNum To EndNum Step IIf(StartNum < EndNum, 1, -1)
        If picBox(Index).Tag <> EmptyBox And getSide(picBox(Index).Tag) <> getSide(Source.Tag) And picBox(Val(x & RightNum)).Tag <> EmptyBox Then
            If Val(x & RightNum) <> Val(StartNum & RightNum) And Val(x & RightNum) <> Val(EndNum & RightNum) Then FoundChess = FoundChess + 1: IsEat = True
        ElseIf Val(x & RightNum) <> Source.Index And picBox(Val(x & RightNum)).Tag <> EmptyBox Then
            Exit Function
        End If
                    
        If x = EndNum Then
            If IsEat = True And FoundChess = 1 Then
                chkPo = True
            ElseIf IsEat = False And FoundChess = 0 And picBox(Index).Tag = EmptyBox Then
                chkPo = True
            End If
        End If
    Next
End Function

Private Function chkSi(Index As Integer, Source As Control) As Boolean
    With Source
        If getSide(.Tag) = BlackSide Then
            Select Case .Index
                Case 40, 42, 60, 62
                    If Index = 51 And getSide(picBox(Index).Tag) <> getSide(.Tag) Then chkSi = True
                Case 51
                    Select Case Index
                        Case 40, 42, 60, 62
                            If getSide(picBox(Index).Tag) <> getSide(.Tag) Then chkSi = True
                    End Select
            End Select
        ElseIf getSide(.Tag) = OrangeSide Then
            Select Case .Index
                Case 49, 47, 69, 67
                    If Index = 58 And getSide(picBox(Index).Tag) <> getSide(.Tag) Then chkSi = True
                Case 58
                    Select Case Index
                        Case 49, 47, 69, 67
                            If getSide(picBox(Index).Tag) <> getSide(.Tag) Then chkSi = True
                    End Select
            End Select
        End If
    End With
End Function

Private Function chkKg(Index As Integer, Source As Control) As Boolean
    Dim x As Integer
    
    'Fei Jiang
    If Left(picBox(Index).Tag, 2) = "Kg" Then
        If picBox(Index).Left = Source.Left Then
            If Abs(Index - Source.Index) >= 5 And Abs(Index - Source.Index) <= 9 Then
                For x = Index To Source.Index Step IIf(Index <= Source.Index, 1, -1)
                    If x <> Index And x <> Source.Index Then
                        If picBox(x).Tag <> EmptyBox Then Exit Function
                    End If
                        
                    If x = Source.Index Then chkKg = True: Exit Function
                Next
            End If
        End If
    End If
    
    Select Case getSide(Source.Tag)
        Case BlackSide
            Select Case Index
                Case 40, 41, 42, 50, 51, 52, 60, 61, 62
                    If Abs(Index - Source.Index) = 10 Or Abs(Index - Source.Index) = 1 Then
                        If getSide(picBox(Index).Tag) <> getSide(Source.Tag) Then chkKg = True
                    End If
            End Select
        Case OrangeSide
            Select Case Index
                Case 47, 48, 49, 57, 58, 59, 67, 68, 69
                    If Abs(Index - Source.Index) = 10 Or Abs(Index - Source.Index) = 1 Then
                        If getSide(picBox(Index).Tag) <> getSide(Source.Tag) Then chkKg = True
                    End If
            End Select
    End Select
End Function

Private Function chkXg(Index As Integer, Source As Control) As Boolean
    If Abs(Index - Source.Index) = 18 Or Abs(Index - Source.Index) = 22 Then
        Select Case getSide(Source.Tag)
            Case BlackSide
                If Val(Right$(Index, 1)) <= 4 Then _
                    If picBox(Source.Index + ((Index - Source.Index) / 2)).Tag = EmptyBox And getSide(picBox(Index).Tag) <> getSide(Source.Tag) Then chkXg = True
            Case OrangeSide
                If Val(Right$(Index, 1)) >= 5 Then _
                    If picBox(Source.Index + ((Index - Source.Index) / 2)).Tag = EmptyBox And getSide(picBox(Index).Tag) <> getSide(Source.Tag) Then chkXg = True
        End Select
    End If
End Function

Private Function chkZu(Index As Integer, Source As Control) As Boolean
    Select Case getSide(Source.Tag)
        Case BlackSide
            Select Case Index - Source.Index
                Case 10, -10
                    If Right(Index, 1) <= 4 And Right(Index, 1) >= 0 Then Exit Function
                    If getSide(picBox(Index).Tag) <> getSide(Source.Tag) Then chkZu = True
                Case 1
                    If getSide(picBox(Index).Tag) <> getSide(Source.Tag) Then chkZu = True
            End Select
        Case OrangeSide
            Select Case Index - Source.Index
                Case 10, -10
                    If Right(Index, 1) >= 5 And Right(Index, 1) <= 9 Then Exit Function
                    If getSide(picBox(Index).Tag) <> getSide(Source.Tag) Then chkZu = True
                Case -1
                    If getSide(picBox(Index).Tag) <> getSide(Source.Tag) Then chkZu = True
            End Select
    End Select
End Function

Private Function chkMa(Index As Integer, Source As Control) As Boolean
    If (Abs(picBox(Index).Top - Source.Top) = 1200 And Abs(picBox(Index).Left - Source.Left) = 600) Or _
        (Abs(picBox(Index).Top - Source.Top) = 600 And Abs(picBox(Index).Left - Source.Left) = 1200) Then
        Select Case Source.Index - Index
            Case 12, -8
                If picBox(Source.Index - 1).Tag = EmptyBox And picBox(Source.Index - 1).Left <> picBox(Index).Left Then
                    If getSide(picBox(Index).Tag) <> getSide(Source.Tag) Then chkMa = True
                End If
            Case 8, -12
                If picBox(Source.Index + 1).Tag = EmptyBox And picBox(Source.Index + 1).Left <> picBox(Index).Left Then
                    If getSide(picBox(Index).Tag) <> getSide(Source.Tag) Then chkMa = True
                End If
            Case 19, 21
                If picBox(Source.Index - 10).Tag = EmptyBox And picBox(Source.Index - 10).Left <> picBox(Index).Left Then
                    If getSide(picBox(Index).Tag) <> getSide(Source.Tag) Then chkMa = True
                End If
            Case -19, -21
                If picBox(Source.Index + 10).Tag = EmptyBox And picBox(Source.Index + 10).Left <> picBox(Index).Left Then
                    If getSide(picBox(Index).Tag) <> getSide(Source.Tag) Then chkMa = True
                End If
        End Select
    End If
End Function

Private Sub picBox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Static PreIndex As Integer
    If mnuDisplayAvailable.Checked = True And PreIndex <> Index Then chkSteps (Index): PreIndex = Index
End Sub

Private Sub tmrBlink_Timer()
    'Jiang Jun Image Blink Effect
    Static Count As Integer
    Count = Count + 1
    If Count <= 6 Then
        picCheck.Visible = IIf(picCheck.Visible = True, False, True)
    Else
        picCheck.Visible = False
        tmrBlink.Enabled = False
        Count = 0
    End If
End Sub


