VERSION 5.00
Begin VB.Form frmTeacherMenu 
   Caption         =   "Teacher Menu"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   14265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLogout 
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   7440
      TabIndex        =   5
      Top             =   7320
      Width           =   6375
   End
   Begin VB.CommandButton cmdTests 
      Caption         =   "Add, Edit or Remove Tests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   480
      TabIndex        =   4
      Top             =   7320
      Width           =   6375
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "View Students Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   7440
      TabIndex        =   2
      Top             =   3000
      Width           =   6375
   End
   Begin VB.CommandButton cmdUsers 
      Caption         =   "Add, Edit or Delete Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   480
      TabIndex        =   1
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Label lblWelcomeMessage 
      Alignment       =   2  'Center
      Caption         =   "What do you want to do today?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   14055
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
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   13815
   End
End
Attribute VB_Name = "frmTeacherMenu"
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

Private Sub cmdLogout_Click()   'Asks the user are they sure they want to logout

Dim answ As String

answ = MsgBox("Are you sure you want to logout?", vbYesNo, "Exit?")

If answ = vbYes Then
    username = ""
    password = ""

    frmLogin.ShowScreen
    frmLogin.RemoveTMenu
Else

End If

End Sub

Private Sub cmdResults_Click()  'Sends the user to the results form

frmStudentResults.Show

End Sub

Private Sub cmdTests_Click()    'Sends the user to the Add, Edit and Delete Tests form

frmADETests.Show

End Sub

Private Sub cmdUsers_Click()    'Sends the user to the Add, Edit and Delete users form

frmUsers.Show

End Sub

Private Sub Form_Load()         'Imports the users info and displays a welcome message

frmTeacherMenu.ImportUserInfo
frmTeacherMenu.DisplayWelcomeMessage

End Sub

Sub ImportUserInfo()            'Imports user info from the login screen

username = frmLogin.txtDBUsername
password = frmLogin.txtDBPassword
firstName = frmLogin.txtDBFirstName
lastName = frmLogin.txtDBLastName
classYear = frmLogin.txtDBYear

Debug.Print (username + password)
Debug.Print ("Welcome " + firstName + " " + lastName + " Of Year " + classYear)

End Sub

Sub DisplayWelcomeMessage()     'Displays a personal welcome message

lblWelcome.Caption = ("Welcome " + firstName + " " + lastName + " Of Year " + classYear)

End Sub

Sub CleanUp()               'Removes the login screen

frmLogin.Visible = False
frmLogin.Remove

End Sub

Sub Remove()                'Removes the current form

Unload Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)  'When unloaded go to login screen

Unload Me
frmLogin.Show

End Sub
