VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   990
   ClientLeft      =   4125
   ClientTop       =   4530
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   0
      Picture         =   "About.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "andrew.leclercq@orange.net"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "yigitaktan@yahoo.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Coded By Yigit Aktan / M@verick"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "aBOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub




