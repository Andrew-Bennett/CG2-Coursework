VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDBID 
      DataField       =   "ID"
      DataSource      =   "adoUser"
      Height          =   285
      Left            =   1920
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   7200
      Width           =   615
   End
   Begin VB.CommandButton cmdTests 
      Caption         =   "T"
      Height          =   375
      Left            =   7320
      TabIndex        =   17
      Top             =   4200
      Width           =   255
   End
   Begin VB.CommandButton cmdPLogin 
      Caption         =   "P"
      Height          =   375
      Left            =   7320
      TabIndex        =   16
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton cmdTLogin 
      Caption         =   "T"
      Height          =   375
      Left            =   7320
      TabIndex        =   15
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txtDBYear 
      DataField       =   "ClassYear"
      DataSource      =   "adoUser"
      Height          =   285
      Left            =   7200
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   7080
      Width           =   255
   End
   Begin VB.TextBox txtDBLastName 
      DataField       =   "LastName"
      DataSource      =   "adoUser"
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   6960
      Width           =   255
   End
   Begin VB.TextBox txtDBFirstName 
      DataField       =   "FirstName"
      DataSource      =   "adoUser"
      Height          =   285
      Left            =   5640
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   6960
      Width           =   150
   End
   Begin VB.TextBox txtDBPassword 
      DataField       =   "Password"
      DataSource      =   "adoUser"
      Height          =   285
      Left            =   8280
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtDBPermissionLevel 
      DataField       =   "PermissionLevel"
      DataSource      =   "adoUser"
      Height          =   285
      Left            =   9000
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   6960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtDBUsername 
      DataField       =   "Username"
      DataSource      =   "adoUser"
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSAdodcLib.Adodc adoUser 
      Height          =   1815
      Left            =   7680
      Top             =   3240
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3201
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database03.mdb;Mode=Read;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database03.mdb;Mode=Read;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "UserInfo"
      Caption         =   "UserInfo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3960
      TabIndex        =   6
      Top             =   4680
      Width           =   3615
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   3735
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   240
      TabIndex        =   4
      Text            =   "Enter Password"
      Top             =   2880
      Width           =   3255
   End
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Text            =   "Enter Username"
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lblIUser 
      Height          =   735
      Left            =   3600
      TabIndex        =   11
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblTest 
      Height          =   975
      Left            =   3600
      TabIndex        =   10
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblSchoolName 
      Caption         =   "Greenparks School - Spelling Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim username As String
Dim password As String

Dim usernameCorrect As Boolean
Dim passwordCorrect As Boolean

Private Sub Form_Load()         'When the form loads it refreshs the database and turns user&pass correct to false

usernameCorrect = False
passwordCorrect = False
lblIUser.Caption = ""

adoUser.Refresh

End Sub

Private Sub txtPassword_Click()  'When the password box is clicked on it will remove the text, change color and set a password char
If txtPassword = "Enter Password" Then

    txtPassword.Text = ""
    txtPassword.ForeColor = &H80000006
    txtPassword.PasswordChar = "*"
    
End If

If txtUsername.Text = "" Then           'If clicked off it will show enter username again

    txtUsername.ForeColor = &H80000003
    txtUsername.Text = "Enter Username"

End If

End Sub

Private Sub txtUsername_Click() 'When the username box is clicked on it will remove the text and change colour

If txtUsername = "Enter Username" Then
    
    txtUsername.Text = ""
    txtUsername.ForeColor = &H80000006
    
End If

If txtPassword.Text = "" Then   'When clicked off the text will reappear

    txtPassword.ForeColor = &H80000003
    txtPassword.PasswordChar = ""
    txtPassword.Text = "Enter Password"

End If


End Sub

Sub ShowScreen()            'Shows the login screen and removes the splash screen

frmSplashScreen.Remove
frmLogin.Visible = True

End Sub

Private Sub txtUsername_GotFocus()  'If the username is clicked it will change color and remove any text

txtUsername.Text = ""
txtUsername.ForeColor = &H80000006

End Sub

Private Sub txtPassword_GotFocus()  ' if the password is clicked it will set a pass char, change color and remove the text

txtPassword.Text = ""
txtPassword.ForeColor = &H80000006
txtPassword.PasswordChar = "*"

End Sub

Sub Login()

frmLogin.StoreUserPass      'Store the username and password in strings
frmLogin.SearchForUser      'Search for the user in the database

End Sub

Sub SearchForUser()         'Searches for user

frmLogin.LookLoop

End Sub

Sub LoginCheck()    'Checks to see if both the password and the username are correct and if so check permission and login
                                    'Else display incorrect message
If usernameCorrect = True Then
    
    If passwordCorrect = True Then
    
        frmLogin.CheckPermission
    
    Else
    
        frmLogin.DisplayIncorrectMessage
    
    End If

End If

End Sub

Sub CheckPermission()           'Checks to see if the user is either a pupil or teacher and sends them to the relevant form

If txtDBPermissionLevel = "Pupil" Then

    frmPupilMenu.Show
    frmPupilMenu.CleanUp
    
Else

    If txtDBPermissionLevel = "" Then
    
    Else
    
    frmTeacherMenu.Show
    frmTeacherMenu.CleanUp
    
    End If

End If

End Sub

Sub LookLoop()

If txtDBUsername.Text = username Then    'If the username matches the Database username then check if the password matches

    usernameCorrect = True
    frmLogin.LoginCheck

Else
    
    frmLogin.DBNext             'Else go to the next database record

End If

If txtDBPassword.Text = password Then        'If the password matches then check to see if they can login

    passwordCorrect = True
    frmLogin.LoginCheck
    
End If

End Sub

Sub DBNext()

If adoUser.Recordset.EOF = True Then    'If end of file is true then display error message

    frmLogin.DisplayIncorrectMessage
    
Else

    adoUser.Recordset.MoveNext          'Else continue looking for the username
    frmLogin.LookLoop
    
End If
    
End Sub

Sub DisplayIncorrectMessage()   'Displays an error message if either the username or password is wrong

lblTest.Caption = "Incorrect Username or Password"

End Sub

Sub StoreUserPass()         'Stores the username and password to strings

username = txtUsername.Text

password = txtPassword.Text

End Sub
Private Sub cmdExit_Click()         'Asks the user wether they are sure they want to exit

Dim answ As String

answ = MsgBox("Are you sure you want to exit?", vbYesNo, "Exit?")

If answ = vbYes Then

    Unload Me
    
Else

End If

End Sub

Private Sub cmdLogin_Click()    'When the login is clicked refresh database and move to first record before logging in

adoUser.Refresh
adoUser.Recordset.MoveFirst
frmLogin.Login

End Sub

Private Sub cmdPLogin_Click()   'Enters the account details of a pupil for ease of access

txtUsername.Text = "Test"
txtPassword.PasswordChar = "*"
txtPassword.Text = "test"
txtPassword.ForeColor = "&H80000006"

End Sub

Private Sub cmdTests_Click()

frmUsers.Show

End Sub

Private Sub cmdTLogin_Click()   'Enters the account details of a teacher for ease of access

txtUsername.Text = "Teacher"
txtPassword.PasswordChar = "*"
txtPassword.Text = "teacher"
txtPassword.ForeColor = "&H80000006"

End Sub

Sub Remove()            'Remove the login form

Unload Me

End Sub

Sub RemoveTMenu()       'Remove the teacher menu

frmTeacherMenu.Remove

End Sub

Sub RemovePMenu()       'Remove the pupil menu

frmPupilMenu.Remove

End Sub
