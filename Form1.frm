VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Click The Button"
   ClientHeight    =   3195
   ClientLeft      =   630
   ClientTop       =   645
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "CENTER THE FORM"
      Height          =   2055
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'you might want to put this into a timer so that if the user's _
desktop changes you can make your form center its self to that _
new change.

'Coded By: Will Hughes

CenterForm Me
End Sub
