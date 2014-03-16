VERSION 5.00
Begin VB.Form frmPupilMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Pupil Menu"
   ClientHeight    =   9645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   12360
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   12360
      TabIndex        =   4
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton cmdResults 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Results!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   4935
   End
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   4935
   End
   Begin VB.CommandButton cmdTest 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Take A Test!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   4935
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11535
   End
End
Attribute VB_Name = "frmPupilMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogout_Click()  'Asks the user are they sure they want to logout

Dim answ As Integer

answ = MsgBox("Are you sure you want to logout?", vbYesNo, "Exit?")

If answ = vbYes Then
    globalUsername = ""

    Unload Me
    frmLogin.Show
    
Else

    'Do nothing

End If

End Sub

Private Sub cmdResults_Click()  'Sends the user to the results form

Unload Me
frmStudentResults.PupilShow

End Sub

Private Sub cmdTest_Click()     'Sends the user to the test form

Unload Me
frmTest.Show

End Sub

Private Sub Form_Load()         'When the form loads it imports the users info and displays a welcome message

frmPupilMenu.DisplayWelcomeMessage

End Sub

Sub DisplayWelcomeMessage() 'takes users info and displays a personal welcome message

lblWelcome.Caption = ("Welcome " & globalFirstName & " " & globalLastName & " Of Year " & globalYear)

End Sub
