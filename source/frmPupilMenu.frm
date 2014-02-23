VERSION 5.00
Begin VB.Form frmPupilMenu 
   Caption         =   "Pupil Menu"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
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
      Caption         =   "View Results!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   6120
      TabIndex        =   3
      Top             =   2640
      Width           =   4935
   End
   Begin VB.CommandButton cmdLogout 
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   6120
      TabIndex        =   2
      Top             =   6120
      Width           =   4935
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Take A Test!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   480
      TabIndex        =   1
      Top             =   2640
      Width           =   4935
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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

Dim username As String
Dim password As String
Dim firstName As String
Dim lastName As String
Dim classYear As String
Dim id As String

Private Sub cmdLogout_Click()  'Asks the user are they sure they want to logout

Dim answ As String

answ = MsgBox("Are you sure you want to logout?", vbYesNo, "Exit?")

If answ = vbYes Then
    username = ""
    password = ""

    frmLogin.ShowScreen
    frmLogin.RemovePMenu
Else

End If

End Sub

Private Sub cmdResults_Click()  'Sends the user to the results form

frmStudentResults.PupilShow

End Sub

Private Sub cmdTest_Click()     'Sends the user to the test form

frmTest.Show

End Sub

Private Sub Form_Load()         'When the form loads it imports the users info and displays a welcome message

frmPupilMenu.ImportUserInfo
frmPupilMenu.DisplayWelcomeMessage

End Sub

Sub ImportUserInfo()                'Takes the information from the login screen and stores it in strings

id = frmLogin.txtDBID
username = frmLogin.txtDBUsername
password = frmLogin.txtDBPassword
firstName = frmLogin.txtDBFirstName
lastName = frmLogin.txtDBLastName
classYear = frmLogin.txtDBYear

txtUsername.Text = username
txtID = id

Debug.Print (username + password)
Debug.Print ("Welcome " + firstName + " " + lastName + " Of Year " + classYear)

End Sub

Sub DisplayWelcomeMessage() 'takes users info and displays a personal welcome message

lblWelcome.Caption = ("Welcome " + firstName + " " + lastName + " Of Year " + classYear)

End Sub

Sub CleanUp()   'Removes the login form

frmLogin.Visible = False
frmLogin.Remove

End Sub

Sub Remove()    'Removes current form

Unload Me

End Sub
