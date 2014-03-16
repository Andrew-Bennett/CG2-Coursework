VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTest 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Test Time!"
   ClientHeight    =   8445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   7320
      Width           =   3375
   End
   Begin VB.TextBox txtDBResult_ID 
      DataSource      =   "adoTests"
      Height          =   375
      Left            =   9840
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   10200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtDBAnswer 
      DataField       =   "Word10"
      DataSource      =   "adoTests"
      Height          =   285
      Index           =   10
      Left            =   10920
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   9720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtDBAnswer 
      DataField       =   "Word9"
      DataSource      =   "adoTests"
      Height          =   285
      Index           =   9
      Left            =   10800
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   9720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtDBAnswer 
      DataField       =   "Word8"
      DataSource      =   "adoTests"
      Height          =   285
      Index           =   8
      Left            =   10680
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   9720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtDBAnswer 
      DataField       =   "Word7"
      DataSource      =   "adoTests"
      Height          =   285
      Index           =   7
      Left            =   10560
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   9720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtDBAnswer 
      DataField       =   "Word6"
      DataSource      =   "adoTests"
      Height          =   285
      Index           =   6
      Left            =   10440
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   9720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtDBAnswer 
      DataField       =   "Word5"
      DataSource      =   "adoTests"
      Height          =   285
      Index           =   5
      Left            =   10320
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   9720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtDBAnswer 
      DataField       =   "Word4"
      DataSource      =   "adoTests"
      Height          =   285
      Index           =   4
      Left            =   10200
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   9720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtDBAnswer 
      DataField       =   "Word3"
      DataSource      =   "adoTests"
      Height          =   285
      Index           =   3
      Left            =   10080
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   9720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtDBAnswer 
      DataField       =   "Word2"
      DataSource      =   "adoTests"
      Height          =   285
      Index           =   2
      Left            =   9960
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   9720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtDBAnswer 
      DataField       =   "Word1"
      DataSource      =   "adoTests"
      Height          =   285
      Index           =   1
      Left            =   9840
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   9720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtDBTestDate 
      DataField       =   "TestDate"
      DataSource      =   "adoTests"
      Height          =   975
      Left            =   9360
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   9720
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSAdodcLib.Adodc adoTests 
      Height          =   975
      Left            =   7440
      Top             =   9720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Submit"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7320
      Width           =   3375
   End
   Begin VB.Frame fraTest 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame1"
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   11295
      Begin VB.TextBox txtAnswer 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   8040
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   5640
         Width           =   3135
      End
      Begin VB.TextBox txtAnswer 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   8040
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   5040
         Width           =   3135
      End
      Begin VB.TextBox txtAnswer 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   8040
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   4440
         Width           =   3135
      End
      Begin VB.TextBox txtAnswer 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   8040
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   3840
         Width           =   3135
      End
      Begin VB.TextBox txtAnswer 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   8040
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox txtAnswer 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   8040
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox txtAnswer 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   8040
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   2040
         Width           =   3135
      End
      Begin VB.TextBox txtAnswer 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   8040
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtAnswer 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   8040
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtAnswer 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   8040
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblDefinition 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Definition"
         DataField       =   "Definition10"
         DataSource      =   "adoTests"
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
         Index           =   10
         Left            =   720
         TabIndex        =   22
         Top             =   5640
         Width           =   7335
      End
      Begin VB.Label lblDefinition 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Definition"
         DataField       =   "Definition9"
         DataSource      =   "adoTests"
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
         Index           =   9
         Left            =   480
         TabIndex        =   21
         Top             =   5040
         Width           =   7335
      End
      Begin VB.Label lblDefinition 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Definition"
         DataField       =   "Definition8"
         DataSource      =   "adoTests"
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
         Index           =   8
         Left            =   480
         TabIndex        =   20
         Top             =   4440
         Width           =   7335
      End
      Begin VB.Label lblDefinition 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Definition"
         DataField       =   "Definition7"
         DataSource      =   "adoTests"
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
         Index           =   7
         Left            =   480
         TabIndex        =   19
         Top             =   3840
         Width           =   7335
      End
      Begin VB.Label lblDefinition 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Definition"
         DataField       =   "Definition6"
         DataSource      =   "adoTests"
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
         Index           =   6
         Left            =   480
         TabIndex        =   18
         Top             =   3240
         Width           =   7335
      End
      Begin VB.Label lblDefinition 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Definition"
         DataField       =   "Definition5"
         DataSource      =   "adoTests"
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
         Index           =   5
         Left            =   480
         TabIndex        =   17
         Top             =   2640
         Width           =   7335
      End
      Begin VB.Label lblDefinition 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Definition"
         DataField       =   "Definition4"
         DataSource      =   "adoTests"
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
         Index           =   4
         Left            =   480
         TabIndex        =   16
         Top             =   2040
         Width           =   7335
      End
      Begin VB.Label lblDefinition 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Definition"
         DataField       =   "Definition3"
         DataSource      =   "adoTests"
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
         Index           =   3
         Left            =   480
         TabIndex        =   15
         Top             =   1440
         Width           =   7335
      End
      Begin VB.Label lblDefinition 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Definition"
         DataField       =   "Definition2"
         DataSource      =   "adoTests"
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
         Index           =   2
         Left            =   480
         TabIndex        =   14
         Top             =   840
         Width           =   7335
      End
      Begin VB.Label lblDefinition 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Definition"
         DataField       =   "Definition1"
         DataSource      =   "adoTests"
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
         Index           =   1
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "10."
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
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   5640
         Width           =   975
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "9."
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
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "8."
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
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "7."
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
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "6."
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
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "5."
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
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "4."
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
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "3."
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
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "2."
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
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "1."
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
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ComboBox cboTestDate 
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
      Left            =   2760
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "Saving doesn't work fully! Errors may occur!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   46
      Top             =   120
      Width           =   5535
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   11400
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblSelectDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Test:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim num As Integer
Dim score As Integer

Dim lockAnswers As Boolean

Dim correctLength As Boolean
Dim correctMid As Boolean
Dim correctMid2 As Boolean

Dim percentCorrect As Integer
Dim lenPercent As Integer

Dim dbLen As Integer
Dim mENum As Integer
Dim lenScore As Integer

Const PERCENTNEEDED As Integer = 0.75   'This is the percentage that the user needs to achieve for a MinorError mark
Const LOWERPERCENTNEEDED As Integer = 0.5

Dim cbMoveFirst As Boolean

Private Sub cmdExit_Click() 'When exit is pressed close the form & show pupil menu

Unload Me
frmPupilMenu.Show

End Sub

Private Sub Form_Load()

cboTestDate.Text = "Select Date"    'Changes the combo list text to select date

frmTest.ClearDA                 'Clears the definitions and text boxes
frmTest.UpdateCBOList           'Adds all the test dates to the combo list

lockAnswers = True              'Locks answers until a test date has been choosen
frmTest.LockA

score = 0

cmdSubmit.Enabled = False       'Disables the submit button until the user has selected a test date

cbMoveFirst = False

End Sub

Private Sub cmdSubmit_Click()       'Calls for the answers to be marked and for a SQL guery to create a new field

score = 0

frmTest.CompareAnswers
frmTest.CheckField

End Sub

Sub CheckField()    'Adds the new field if needed and then stores the result in the database

frmTest.AddField
frmTest.StoreResult

End Sub

Sub StoreResult()   'Refreshes the database incase of a new field addition
On Error GoTo ErrorM

adoTests.Refresh

txtDBResult_ID.DataField = "Result_" & globalID       'Changes the field to the users result field

txtDBResult_ID.Text = score                     'Change and save the users score in the database

adoTests.Recordset.Update

Exit Sub

ErrorM:

MsgBox ("Error " & Err.Number & " " & Err.Description)

Exit Sub


End Sub

Sub AddField()

    Dim conn As ADODB.Connection
    Dim query As String

    Set conn = New ADODB.Connection 'Open an ado connection
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Persist Security Info=False;" & "Data Source=Database03.mdb"
    
    conn.Open

    On Error GoTo AddFieldError         'Add new field in the form of Result_[ID]
    query = "ALTER TABLE " & "Test" & " ADD COLUMN " & "Result_" & globalID & " " & "Text"
    
    If Len(Trim$("0")) >= 0 Then        'Can be used so that a default can be choosen
        
        query = query & " DEFAULT " & ""
        
    End If
    
    conn.Execute query
    conn.Close
    
    Exit Sub

AddFieldError:
    'MsgBox ("Error " & Err.Number & " executing statement: " & query & Err.Description)
    
    conn.Close      'If it already exists just close connection and continue on
    Exit Sub

End Sub

Sub CompareAnswers()    'Checks to see if the users answer matches the databases answers and marks it accordingly

Dim answ As String

For num = 1 To 10

'TO BE ADDED turn both the answers and DB answers into lowercase

txtAnswer(num).ForeColor = &HFF&

If LCase$(txtAnswer(num).Text) = LCase$(txtDBAnswer(num).Text) Then

    score = score + 2
    
    txtAnswer(num).ForeColor = &HFF00&      'It also gives a visual feedback to the user as well as a message box with their results
    
Else

    MinorError                              'Calls for MinorError
    score = score + MinorError

End If
    
Next num

answ = MsgBox("You have scored " & score & " marks!", vbOKOnly, "Score")

End Sub

Function MinorError()

lenScore = 0

' -- CHECK IF THE FIRST LETTER IS CORRECT --

If Mid$(LCase$(txtAnswer(num).Text), 1, 1) = Mid$(LCase$(txtDBAnswer(num).Text), 1, 1) Then

        'Continue marking the word
    
Else

    Exit Function

End If

'-- LENGTH CHECK --
dbLen = Len(txtDBAnswer(num).Text)

If Len(txtAnswer(num).Text) = dbLen Then
    
    correctLength = True

End If

'-- PERCENTAGE CORRECT --

For mENum = 1 To dbLen

    If Mid$(LCase$(txtAnswer(num).Text), mENum, 1) = Mid$(LCase$(txtDBAnswer(num).Text), mENum, 1) Then
    
        lenScore = lenScore + 1
    
    End If
    
Next mENum

lenPercent = lenScore / dbLen

If lenPercent >= PERCENTNEEDED Then

    correctMid = True
    
End If

If lenPercent >= LOWERPERCENTNEEDED Then

    correctMid2 = True

End If

If (correctMid2 = True) And (correctLength = True) Or (correctMid = True) Then

    MinorError = 1
    txtAnswer(num).ForeColor = &HFFFF&
    
End If


End Function

Sub LockA()     'Lock all the text boxes until a test date has been choosen

If lockAnswers = True Then

    For num = 1 To 10
    
        txtAnswer(num).Locked = True
        
    Next num
    
    cmdSubmit.Enabled = False
    
Else

    For num = 1 To 10
    
        txtAnswer(num).Locked = False
    
    Next num
    
    cmdSubmit.Enabled = True
    
End If

End Sub

Sub MoveToDate()        'Moves the ado until the ado matches the cbo selection

If cbMoveFirst = True Then
    cbMoveFirst = False
    adoTests.Recordset.MoveFirst
End If

If cboTestDate.Text = txtDBTestDate.Text Then

Else

    adoTests.Recordset.MoveNext
    frmTest.MoveToDate

End If

lockAnswers = False
frmTest.LockA

End Sub

Sub UpdateCBOList() 'Adds all the test dates to the cbolist

Do While adoTests.Recordset.EOF = False
        
    cboTestDate.AddItem (txtDBTestDate.Text)
    adoTests.Recordset.MoveNext
        
Loop

End Sub

Private Sub cboTestDate_Click() 'Moves the date selected in the cbo

Dim n As Integer

For n = 1 To 10

    txtAnswer(n).Text = ""
    
Next n
 
frmTest.ClearDA

cbMoveFirst = True

frmTest.MoveToDate

End Sub

Sub ClearDA()   'Clears all the text boxes and adds "Your answer!" instead

For num = 1 To 10

    lblDefinition(num).Caption = ""
    txtAnswer(num).Text = "Your answer!"
    txtAnswer(num).ForeColor = &H808080

Next num

End Sub

Private Sub txtAnswer_Click(Index As Integer)   'Removes the Your Answer! if clicked

If txtAnswer(Index) = "Your answer!" Then

    txtAnswer(Index).Text = ""
    txtAnswer(Index).ForeColor = vbBlack
End If

End Sub

Private Sub txtAnswer_GotFocus(Index As Integer)    'Removes the Your Answer! if tabbed into or clicked

If txtAnswer(Index) = "Your answer!" Then

    txtAnswer(Index).Text = ""
    txtAnswer(Index).ForeColor = vbBlack
End If

End Sub
