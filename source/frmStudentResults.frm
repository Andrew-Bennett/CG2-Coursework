VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmStudentResults 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Students Results"
   ClientHeight    =   9630
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   15510
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc adoTests 
      Height          =   735
      Left            =   6120
      Top             =   9840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database03.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database03.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Test"
      Caption         =   "Adodc1"
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
   Begin VB.TextBox txtDBResult_ID 
      DataSource      =   "adoTests"
      Height          =   285
      Left            =   3960
      TabIndex        =   29
      Text            =   "Text2"
      Top             =   10440
      Width           =   1695
   End
   Begin VB.TextBox txtDBTestDate 
      DataField       =   "TestDate"
      DataSource      =   "adoTests"
      Height          =   285
      Left            =   3120
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   10440
      Width           =   615
   End
   Begin VB.TextBox txtPermLevel 
      DataField       =   "PermissionLevel"
      DataSource      =   "adoUser"
      Height          =   285
      Left            =   3120
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   9840
      Width           =   2295
   End
   Begin VB.TextBox txtUser 
      DataField       =   "Username"
      DataSource      =   "adoUser"
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc adoUser 
      Height          =   975
      Left            =   120
      Top             =   9720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1720
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database03.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database03.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "UserInfo"
      Caption         =   "adoUser"
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
   Begin VB.VScrollBar vsbResults 
      Height          =   6855
      Left            =   15120
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Logout"
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8280
      Width           =   2775
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Width           =   2775
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      Width           =   2775
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8280
      Width           =   2775
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   2775
   End
   Begin VB.Frame fraDate 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Results"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      Begin VB.TextBox txtDBID 
         DataField       =   "ID"
         DataSource      =   "adoUser"
         Height          =   375
         Left            =   10200
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cboMonth 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   240
         Width           =   3135
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   14880
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblSelectMonth 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Month:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblMonthAverage 
         Caption         =   "Month Average :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   6480
         Width           =   4215
      End
      Begin VB.Label lblAllTimeAverage 
         Caption         =   "Overall Average:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   7200
         Width           =   4335
      End
   End
   Begin VB.Frame fraResults 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame1"
      Height          =   6975
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   14895
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 - "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   6
         Left            =   7200
         TabIndex        =   27
         Top             =   2640
         Width           =   4335
      End
      Begin VB.Label lblResult 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   6
         Left            =   11640
         TabIndex        =   26
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   14760
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 - "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   4455
      End
      Begin VB.Label lblResult 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   5
         Left            =   4560
         TabIndex        =   24
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 - "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   7200
         TabIndex        =   23
         Top             =   1440
         Width           =   4455
      End
      Begin VB.Label lblResult 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   4
         Left            =   11640
         TabIndex        =   22
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 - "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label lblResult 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   3
         Left            =   4560
         TabIndex        =   20
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 - "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   7200
         TabIndex        =   17
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblResult 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   2
         Left            =   11640
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
      Begin VB.Line Line3 
         X1              =   6840
         X2              =   6840
         Y1              =   240
         Y2              =   6720
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   14640
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblResult 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   4560
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 - "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   32.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Label lblDate 
      Caption         =   "00/00/0000 - "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   1440
      TabIndex        =   19
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Label lblResult 
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   5520
      TabIndex        =   18
      Top             =   3240
      Width           =   2295
   End
End
Attribute VB_Name = "frmStudentResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim currentRecord As Integer
Dim totalRecords As Integer

Dim username As String
Dim id As String

Dim currentNum As Integer

Dim currentlyClear As Boolean

Private Sub cmdNext_Click() 'Moves to the next pupils results --  STILL BEING WORKED ON!

If currentRecord < totalRecords Then
    
    adoUser.Recordset.MoveNext
    currentRecord = currentRecord + 1
    
    If txtPermLevel.Text = "Teacher" Then
    
    adoUser.Recordset.MoveNext
    currentRecord = currentRecord + 1
    
    End If
    
Else

    'Lock the next and last buttons
    
End If

End Sub

Private Sub cboMonth_Click()    'Displays all the tests done in the month selected

currentNum = 1

frmStudentResults.ClearResults

frmStudentResults.FilterResults

End Sub

Sub FilterResults()                         'Checks what month the test date is and filters them accordingly

If cboMonth.Text = "January" Then

Do While adoTests.Recordset.EOF = False

    If Mid$(txtDBTestDate.Text, 4, 2) = "01" Then
        
        frmStudentResults.DisplayDR
        
        
    End If
    
    adoTests.Recordset.MoveNext
    
Loop

End If

If cboMonth.Text = "February" Then

Do While adoTests.Recordset.EOF = False

    If Mid$(txtDBTestDate.Text, 4, 2) = "02" Then
    
        Debug.Print (txtDBTestDate.Text)
        
        frmStudentResults.DisplayDR
        
        
    End If
    
    adoTests.Recordset.MoveNext
    
Loop

End If

If cboMonth.Text = "March" Then

Do While adoTests.Recordset.EOF = False

    If Mid$(txtDBTestDate.Text, 4, 2) = "03" Then
    
        Debug.Print (txtDBTestDate.Text)
        
        frmStudentResults.DisplayDR
        
        
    End If
    
    adoTests.Recordset.MoveNext
    
Loop

End If

If cboMonth.Text = "April" Then

Do While adoTests.Recordset.EOF = False

    If Mid$(txtDBTestDate.Text, 4, 2) = "04" Then
    
        Debug.Print (txtDBTestDate.Text)
        
        frmStudentResults.DisplayDR
        
        
    End If
    
    adoTests.Recordset.MoveNext
    
Loop

End If

If cboMonth.Text = "May" Then

Do While adoTests.Recordset.EOF = False

    If Mid$(txtDBTestDate.Text, 4, 2) = "05" Then
    
        Debug.Print (txtDBTestDate.Text)
        
        frmStudentResults.DisplayDR
        
        
    End If
    
    adoTests.Recordset.MoveNext
    
Loop

End If

If cboMonth.Text = "June" Then

Do While adoTests.Recordset.EOF = False

    If Mid$(txtDBTestDate.Text, 4, 2) = "06" Then
    
        Debug.Print (txtDBTestDate.Text)
        
        frmStudentResults.DisplayDR
        
        
    End If
    
    adoTests.Recordset.MoveNext
    
Loop

End If

If cboMonth.Text = "July" Then

Do While adoTests.Recordset.EOF = False

    If Mid$(txtDBTestDate.Text, 4, 2) = "07" Then
    
        Debug.Print (txtDBTestDate.Text)
        
        frmStudentResults.DisplayDR
        
        
    End If
    
    adoTests.Recordset.MoveNext
    
Loop

End If

If cboMonth.Text = "August" Then

Do While adoTests.Recordset.EOF = False

    If Mid$(txtDBTestDate.Text, 4, 2) = "08" Then
    
        Debug.Print (txtDBTestDate.Text)
        
        frmStudentResults.DisplayDR
        
        
    End If
    
    adoTests.Recordset.MoveNext
    
Loop

End If

If cboMonth.Text = "September" Then

Do While adoTests.Recordset.EOF = False

    If Mid$(txtDBTestDate.Text, 4, 2) = "09" Then
    
        Debug.Print (txtDBTestDate.Text)
        
        frmStudentResults.DisplayDR
        
        
    End If
    
    adoTests.Recordset.MoveNext
    
Loop

End If

If cboMonth.Text = "October" Then

Do While adoTests.Recordset.EOF = False

    If Mid$(txtDBTestDate.Text, 4, 2) = "10" Then
    
        Debug.Print (txtDBTestDate.Text)
        
        frmStudentResults.DisplayDR
        
        
    End If
    
    adoTests.Recordset.MoveNext
    
Loop

End If

If cboMonth.Text = "November" Then

Do While adoTests.Recordset.EOF = False

    If Mid$(txtDBTestDate.Text, 4, 2) = "11" Then
    
        Debug.Print (txtDBTestDate.Text)
        
        frmStudentResults.DisplayDR
        
        
    End If
    
    adoTests.Recordset.MoveNext
    
Loop

End If

If cboMonth.Text = "December" Then

Do While adoTests.Recordset.EOF = False

    If Mid$(txtDBTestDate.Text, 4, 2) = "12" Then
    
        Debug.Print (txtDBTestDate.Text)
        
        frmStudentResults.DisplayDR
        
        
    End If
    
    adoTests.Recordset.MoveNext
    
Loop

End If

End Sub

Sub DisplayDR()     'Displays the test date and the users result

currentlyClear = False
On Error GoTo AddFieldError:

lblDate(currentNum).Caption = txtDBTestDate.Text & " - "
lblDate(currentNum).Visible = True
txtDBResult_ID.DataField = "Result_" & id
lblResult(currentNum).Caption = txtDBResult_ID.Text & "/20"
lblResult(currentNum).Visible = True

currentNum = currentNum + 1

Exit Sub

AddFieldError:

MsgBox ("Error " & Err.Number & " executing statement: " & Err.Description)

Exit Sub

End Sub

Sub PupilShow()         'Loads the form as in pupil mode, so that they can only view their results etc

frmStudentResults.Show
frmStudentResults.Height = 8805

username = frmPupilMenu.txtUsername.Text
id = frmPupilMenu.txtID.Text

frmStudentResults.MoveToActiveUser

End Sub

Sub MoveToActiveUser()      'Moves forward the ado until it reaches the user thats currently using the program

If username = txtUser.Text Then

    'Do nothing
    
Else

    adoUser.Recordset.MoveNext
    
    frmStudentResults.MoveToActiveUser
    
End If

End Sub

Sub ClearResults()      'Clears all the results in the form
Dim num As Integer

adoTests.Recordset.MoveFirst

For num = 1 To 6
lblDate(num).Caption = ""
lblResult(num).Caption = ""
lblDate(num).Visible = False
lblResult(num).Visible = False

Next num

currentlyClear = True

End Sub

Private Sub Form_Load()             'Updates the cbo, clear the results and check wether pupil or teacher

frmStudentResults.UpdateCBOMonth

currentRecord = 1
totalRecords = adoUser.Recordset.recordCount

frmStudentResults.TeacherCheck
frmStudentResults.ClearResults

currentNum = 1

currentlyClear = False

End Sub

Private Sub cmdLogout_Click()   'Logs out the user
 
 Unload Me
 
End Sub

Sub TeacherCheck()      'Moves forward if its a teacher, as teachers wont be taking the tests

If txtPermLevel.Text = "Teacher" Then

adoUser.Recordset.MoveNext
currentRecord = currentRecord + 1

End If

End Sub

Sub UpdateCBOMonth()    'Add all the months in the year to the cbo

cboMonth.Text = "Select Month"

cboMonth.AddItem ("January")
cboMonth.AddItem ("February")
cboMonth.AddItem ("March")
cboMonth.AddItem ("April")
cboMonth.AddItem ("May")
cboMonth.AddItem ("June")
cboMonth.AddItem ("July")
cboMonth.AddItem ("August")
cboMonth.AddItem ("September")
cboMonth.AddItem ("October")
cboMonth.AddItem ("November")
cboMonth.AddItem ("December")

End Sub

Private Sub vsbResults_Scroll() 'Scrolls the results, in case there are lots of test and test results
 
 fraResults.Top = 1200 - vsbResults.Value '1200 is the current top of the frame and needs be to found in form load
 
End Sub
