VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "www.hypersystem.2ya.com"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3000
      TabIndex        =   10
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3000
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Advanced Message Responder"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2880
         TabIndex        =   4
         Text            =   "Hey! How's it going"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Text            =   "Hey!"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "send"
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "send"
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "send"
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "send"
         Height          =   255
         Left            =   2400
         TabIndex        =   21
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "send"
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "If contact types"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "If contact types"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "If contact types"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "If contact types"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "If contact types"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "send"
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "If contact types"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents msn As MsgrObject
Attribute msn.VB_VarHelpID = -1
Private Sub Form_Load()
Set msn = New MsgrObject
End Sub


Private Sub msn_OnTextReceived(ByVal pIMSession As Messenger.IMsgrIMSession, ByVal User As Messenger.IMsgrUser, ByVal bstrMsgHeader As String, ByVal bstrMsgText As String, pfEnableDefault As Boolean)
If bstrMsgText = Text1.Text Then
     User.SendText bstrMsgHeader, Text2.Text, MMSGTYPE_ALL_RESULTS
End If
If bstrMsgText = Text3.Text Then
     User.SendText bstrMsgHeader, Text4.Text, MMSGTYPE_ALL_RESULTS
End If
If bstrMsgText = Text5.Text Then
     User.SendText bstrMsgHeader, Text6.Text, MMSGTYPE_ALL_RESULTS
End If
If bstrMsgText = Text7.Text Then
     User.SendText bstrMsgHeader, Text8.Text, MMSGTYPE_ALL_RESULTS
End If
If bstrMsgText = Text9.Text Then
     User.SendText bstrMsgHeader, Text10.Text, MMSGTYPE_ALL_RESULTS
End If
If bstrMsgText = Text11.Text Then
     User.SendText bstrMsgHeader, Text12.Text, MMSGTYPE_ALL_RESULTS
End If
End Sub
