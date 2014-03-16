VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7665
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
   ScaleHeight     =   6000
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDBID 
      DataField       =   "ID"
      DataSource      =   "adoUser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   7200
      Width           =   615
   End
   Begin VB.CommandButton cmdPLogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "P"
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdTLogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "T"
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtDBYear 
      DataField       =   "ClassYear"
      DataSource      =   "adoUser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   7080
      Width           =   255
   End
   Begin VB.TextBox txtDBLastName 
      DataField       =   "LastName"
      DataSource      =   "adoUser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   6960
      Width           =   255
   End
   Begin VB.TextBox txtDBFirstName 
      DataField       =   "FirstName"
      DataSource      =   "adoUser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   6960
      Width           =   150
   End
   Begin VB.TextBox txtDBPassword 
      DataField       =   "Password"
      DataSource      =   "adoUser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   3615
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   3735
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
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
         Name            =   "Comic Sans MS"
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
   Begin VB.Label lblIncorrectMessage 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   3600
      TabIndex        =   10
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
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
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
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
      BackStyle       =   0  'Transparent
      Caption         =   "Greenparks School - Spelling Test"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
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

Dim password As String

Private Sub Form_Load()         'When the form loads it refreshs the database and turns user&pass correct to false

adoUser.Refresh

End Sub

Sub Login()

frmLogin.StoreUserInfo      'Store the username and password in strings
frmLogin.SearchForUser      'Search for the user in the database

End Sub

Sub StoreUserInfo() 'Stores the username & password in the text boxes to strings

globalUsername = txtUsername.Text
password = txtPassword.Text

End Sub

Sub SearchForUser() 'Searchs for the user in the database and determines whether or not they can log in

If txtDBUsername.Text = globalUsername Then     'Tests for the correct username & password combo

    If txtDBPassword.Text = password Then
    
        frmLogin.CheckPermissionLevel
        
    End If
    
Else

    If adoUser.Recordset.EOF = True Then    'If at the end of the database file then display incorrect combo message
    
        lblIncorrectMessage.Caption = "Incorrect Username or Password"
    
    Else
    
        adoUser.Recordset.MoveNext      'Else move next and check again
        frmLogin.SearchForUser
        
    End If

End If

End Sub

Sub CheckPermissionLevel()  'Tests for the permission level and sends them to the relevant form

If txtDBPermissionLevel.Text = "Teacher" Then
    
    frmLogin.SetGlobals
    Unload Me
    frmTeacherMenu.Show
    
Else
    
    frmLogin.SetGlobals
    Unload Me
    frmPupilMenu.Show
    
End If

End Sub

Sub SetGlobals()    'Takes their information from the database and stores them in Globals for use later

globalUsername = txtUsername.Text
globalID = txtDBID.Text
globalYear = txtDBYear.Text
globalPermissionLevel = txtDBPermissionLevel.Text
globalFirstName = txtDBFirstName.Text
globalLastName = txtDBLastName.Text

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

Private Sub cmdPLogin_Click()   'Enters the account details of a pupil for ease of access & for testing purposes

txtUsername.Text = "Default_Pupil"
txtPassword.PasswordChar = "*"
txtPassword.Text = "password098"
txtPassword.ForeColor = "&H80000006"

End Sub

Private Sub cmdTLogin_Click()   'Enters the account details of a teacher for ease of access & for testing purposes

txtUsername.Text = "Default_Teacher"
txtPassword.PasswordChar = "*"
txtPassword.Text = "password123"
txtPassword.ForeColor = "&H80000006"

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

Private Sub txtUsername_GotFocus()  'If the username is clicked it will change color and remove any text

txtUsername.Text = ""
txtUsername.ForeColor = &H80000006

End Sub

Private Sub txtPassword_GotFocus()  ' if the password is clicked it will set a pass char, change color and remove the text

txtPassword.Text = ""
txtPassword.ForeColor = &H80000006
txtPassword.PasswordChar = "*"

End Sub
