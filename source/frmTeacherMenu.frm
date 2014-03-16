VERSION 5.00
Begin VB.Form frmTeacherMenu 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Teacher Menu"
   ClientHeight    =   7470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12375
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   12375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   5655
   End
   Begin VB.CommandButton cmdTests 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add, Edit or Remove Tests"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   5655
   End
   Begin VB.CommandButton cmdResults 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Students Results"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   5655
   End
   Begin VB.CommandButton cmdUsers 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add, Edit or Delete Users"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   5655
   End
   Begin VB.Label lblWelcomeMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "What do you want to do today?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   11895
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
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   12015
   End
End
Attribute VB_Name = "frmTeacherMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogout_Click()   'Asks the user are they sure they want to logout

Dim answ As String

answ = MsgBox("Are you sure you want to logout?", vbYesNo, "Exit?")

If answ = vbYes Then
    globalUsername = ""

    Unload Me
    frmLogin.Show
    
Else

End If

End Sub

Private Sub cmdResults_Click()  'Sends the user to the results form

Unload Me
frmStudentResults.Show

End Sub

Private Sub cmdTests_Click()    'Sends the user to the Add, Edit and Delete Tests form

Unload Me
frmADETests.Show

End Sub

Private Sub cmdUsers_Click()    'Sends the user to the Add, Edit and Delete users form

Unload Me
frmUsers.Show

End Sub

Private Sub Form_Load()         'Imports the users info and displays a welcome message

frmTeacherMenu.DisplayWelcomeMessage

End Sub


Sub DisplayWelcomeMessage()     'Displays a personal welcome message

lblWelcome.Caption = ("Welcome " & globalFirstName & " " & globalLastName & " Of Year " & globalYear)

End Sub
