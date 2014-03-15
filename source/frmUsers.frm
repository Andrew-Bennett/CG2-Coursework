VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmUsers 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Add, Edit & Delete Users"
   ClientHeight    =   9480
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClearDatabase 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear Database"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdPopulate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Populate The Database"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last >>"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next >"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00FFFFFF&
      Caption         =   "< Previous"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<< First"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8040
      Width           =   2055
   End
   Begin VB.TextBox txtPermissionLevel 
      DataField       =   "PermissionLevel"
      DataSource      =   "adoUserInfo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   11
      Top             =   6840
      Width           =   4815
   End
   Begin VB.TextBox txtYear 
      DataField       =   "ClassYear"
      DataSource      =   "adoUserInfo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   10
      Top             =   5520
      Width           =   4815
   End
   Begin VB.TextBox txtLastName 
      DataField       =   "LastName"
      DataSource      =   "adoUserInfo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   9
      Top             =   4200
      Width           =   4815
   End
   Begin VB.TextBox txtFirstName 
      DataField       =   "FirstName"
      DataSource      =   "adoUserInfo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   8
      Top             =   2880
      Width           =   4815
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "Password"
      DataSource      =   "adoUserInfo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   7
      Top             =   1560
      Width           =   4815
   End
   Begin VB.TextBox txtUsername 
      DataField       =   "Username"
      DataSource      =   "adoUserInfo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   6
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc adoUserInfo 
      Height          =   1695
      Left            =   120
      Top             =   7680
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   2990
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
      Orientation     =   1
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database03.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database03.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "UserInfo"
      Caption         =   "User Info"
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
      Height          =   1335
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label lblPermission 
      BackStyle       =   0  'Transparent
      Caption         =   "Permission:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label lblYear 
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label lblLastName 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblFirstName 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim totalRecords As Long
Dim recordCount As Long

Dim newRecordActive As Boolean
Dim LockPL As Boolean

Const cb As Integer = 8

Private Sub cmdAdd_Click()  'Adds a new record to the database, then locks and unlocks certain buttons

cmdCancel.Enabled = True
cmdSave.Enabled = True

newRecordActive = True
adoUserInfo.Recordset.AddNew
frmUsers.LockNDE

End Sub

Sub LockNDE()               'Locks the Add, Delete and Edit buttons

cmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdEdit.Enabled = False

End Sub

Sub ULockNDE()              'Unlocks the add, edit and delete buttons

cmdAdd.Enabled = True
cmdDelete.Enabled = True
cmdEdit.Enabled = True

End Sub

Private Sub cmdCancel_Click()           'Cancels the addition of a new user

txtYear.Text = "0"
cmdSave.Enabled = False
cmdCancel.Enabled = False
adoUserInfo.Recordset.CancelUpdate
adoUserInfo.Recordset.MoveFirst
frmUsers.ULockNDE
newRecordActive = False

End Sub

Private Sub cmdClearDatabase_Click()
Dim answ As String

Do While adoUserInfo.Recordset.recordCount > 0

    adoUserInfo.Recordset.Delete
    adoUserInfo.Recordset.MoveNext
    totalRecords = totalRecords - 1
    
Loop

answ = MsgBox("Database Cleared!")

adoUserInfo.Recordset.AddNew

txtUsername.Text = "Default_Teacher"
txtPassword.Text = "password123"
txtYear.Text = "1"
txtFirstName.Text = "Default"
txtLastName.Text = "Teacher"
txtPermissionLevel.Text = "Teacher"

adoUserInfo.Recordset.Update
adoUserInfo.Recordset.AddNew

txtUsername.Text = "Default_Pupil"
txtPassword.Text = "password098"
txtYear.Text = "1"
txtFirstName.Text = "Default"
txtLastName.Text = "Pupil"
txtPermissionLevel.Text = "Pupil"

adoUserInfo.Recordset.Update

frmUsers.FBLock
frmUsers.LBUnlock

totalRecords = 2
recordCount = 1

adoUserInfo.Recordset.MoveFirst

End Sub

Private Sub cmdDelete_Click()       'Deletes the current record

adoUserInfo.Recordset.Delete
adoUserInfo.Recordset.MoveFirst
totalRecords = totalRecords - 1
recordCount = 1

frmUsers.FBLock
frmUsers.LBUnlock


End Sub

Private Sub cmdExit_Click()     'Asks the user if they want to exit

Dim answ As String
     
If newRecordActive = True Then
    
    answ = MsgBox("Do you want to save?", vbYesNo, "Before you go!")
    
    If answ = vbYes Then
    
        adoUserInfo.Recordset.Update
        
    Else
    
        adoUserInfo.Recordset.CancelUpdate
    
    End If
    
End If

Unload Me

End Sub

Private Sub cmdFirst_Click()        'Moves to the first record in the database

recordCount = 1
adoUserInfo.Recordset.MoveFirst
frmUsers.FBLock

End Sub

Sub FBLock()                        'Locks the first and previous buttons while unlocking the next and last buttons

cmdFirst.Enabled = False
cmdPrevious.Enabled = False
cmdNext.Enabled = True
cmdLast.Enabled = True

End Sub

Sub LBLock()                        'Locks the next and last button while unlocking the first and previous buttons

cmdFirst.Enabled = True
cmdPrevious.Enabled = True
cmdNext.Enabled = False
cmdLast.Enabled = False

End Sub

Sub FBUnlock()                      'Unlocks the fist and previous buttons

cmdFirst.Enabled = True
cmdPrevious.Enabled = True

End Sub

Sub LBUnlock()                      'Unlocks the next and last buttons

cmdNext.Enabled = True
cmdLast.Enabled = True

End Sub

Private Sub cmdLast_Click()         'Moves to the last record

recordCount = totalRecords
adoUserInfo.Recordset.MoveLast
frmUsers.LBLock

End Sub

Private Sub cmdNext_Click()         'Moves to the next record and checks for EOF

If recordCount < totalRecords Then

    recordCount = recordCount + 1
    adoUserInfo.Recordset.MoveNext
    frmUsers.FBUnlock
    
    If recordCount = totalRecords Then
    
        frmUsers.LBLock

    End If
    
End If

End Sub

Private Sub cmdPopulate_Click()

Dim answ As Long
Dim i As Long

answ = InputBox("How many users would you like to add?", "Populate Database")

For i = 1 To answ


    adoUserInfo.Recordset.AddNew
    
    RandomString
    
    txtFirstName.Text = "TEST First Name " & i
    txtLastName.Text = "TEST Last Name " & i
    
    txtUsername.Text = RandomString
    
    RandomString
    
    txtPassword.Text = RandomString
    
    RandomYear
    
    txtYear.Text = RandomYear
    
    RandomLevel
    
    txtPermissionLevel = RandomLevel
    
    adoUserInfo.Recordset.Update
    
    

Next i

frmUsers.FBUnlock
frmUsers.LBLock

End Sub

Function RandomLevel() As String

    RandomLevel = ""
    
    Randomize
    Dim rl As Integer
    
    rl = CInt(Int((6 * Rnd()) + 1))
    
    If rl < 3 Then
    
        RandomLevel = "Teacher"
        
    Else
    
        RandomLevel = "Pupil"
        
    End If

End Function

Function RandomYear() As Integer

    RandomYear = 0
    
    Randomize
    Dim ry As Integer
    
    RandomYear = CInt(Int((6 * Rnd()) + 1))

End Function

Function RandomString() As String
    
    RandomString = ""
    
    Randomize
    Dim rgch As String
    
    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789!£$%^&*()_+-=[]{}#~@;/:?.,><|\`¬"

    Dim i As Long
    For i = 1 To cb
        RandomString = RandomString & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next

End Function

Private Sub cmdPrevious_Click() 'Moves to the previous record and checks for BOF

If recordCount > 1 Then

    recordCount = recordCount - 1
    adoUserInfo.Recordset.MovePrevious
    frmUsers.LBUnlock
    
    If recordCount = 1 Then
    
        frmUsers.FBLock

    End If
    
End If

End Sub

Private Sub cmdSave_Click() 'Checks to see if the infromation is in the right format before saving

Dim PermLevel As String     'If anything is wrong, a message box appears telling them whats wrong

PermLevel = txtPermissionLevel.Text

LockPL = False

If Val(txtYear.Text) > 0 Then

    If PermLevel = "Teacher" Then
        
        adoUserInfo.Recordset.Update
        newRecordActive = False
        
        frmUsers.ULockNDE
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        
        adoUserInfo.Recordset.MoveFirst
        recordCount = 1
        totalRecords = adoUserInfo.Recordset.recordCount
        
    Else
        
        If PermLevel = "Pupil" Then
            
            adoUserInfo.Recordset.Update
            newRecordActive = False
            
            frmUsers.ULockNDE
            cmdSave.Enabled = False
            cmdCancel.Enabled = False
            
            adoUserInfo.Recordset.MoveFirst
            recordCount = 1
            totalRecords = adoUserInfo.Recordset.recordCount
            
        Else
            
            frmUsers.PLError
            
        End If
    End If
        
Else

    Dim answ As String
    
    answ = MsgBox("'Year'should be a Number!", vbOKOnly, "Error!")
    
    txtYear.Text = ""

End If

End Sub

Sub PLError()           'Tells the user that the they have entered the permission level incorrectly

If LockPL = False Then
    
    LockPL = True
    
    Dim answ2 As String
                
    answ2 = MsgBox("'Permission Level' Should be either Pupil or Teacher!", vbOKOnly, "Error!")
    txtPermissionLevel.Text = ""
    
End If

End Sub

Private Sub Form_Load()     'Sets up everything and locks certian buttons

cmdSave.Enabled = False
cmdCancel.Enabled = False

totalRecords = adoUserInfo.Recordset.recordCount
recordCount = 1

cmdFirst.Enabled = False
cmdPrevious.Enabled = False
newRecordActive = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Need to add the code so that it cancels anything thats currently trying to be entered into the database

End Sub

