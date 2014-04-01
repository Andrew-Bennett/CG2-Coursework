VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmADETests 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Add, Edit and Delete Tests"
   ClientHeight    =   10200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16425
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   16425
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTestTheme 
      Height          =   975
      Left            =   16680
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   6000
      Width           =   615
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
      Height          =   1215
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   120
      Width           =   2295
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
      Height          =   1215
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   120
      Width           =   2295
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
      Height          =   1215
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   120
      Width           =   2295
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
      Height          =   1215
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtDBTestDate 
      DataField       =   "TestDate"
      DataSource      =   "adoTests"
      Height          =   855
      Left            =   16560
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   2880
      Width           =   3615
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   8760
      Width           =   2295
   End
   Begin VB.TextBox txtDefiniton 
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
      Height          =   615
      Index           =   10
      Left            =   6120
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   8040
      Width           =   10215
   End
   Begin VB.TextBox txtDefiniton 
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
      Height          =   615
      Index           =   9
      Left            =   6120
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   7320
      Width           =   10215
   End
   Begin VB.TextBox txtDefiniton 
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
      Height          =   615
      Index           =   8
      Left            =   6120
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   6600
      Width           =   10215
   End
   Begin VB.TextBox txtDefiniton 
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
      Height          =   615
      Index           =   7
      Left            =   6120
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   5880
      Width           =   10215
   End
   Begin VB.TextBox txtDefiniton 
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
      Height          =   615
      Index           =   6
      Left            =   6120
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   5160
      Width           =   10215
   End
   Begin VB.TextBox txtDefiniton 
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
      Height          =   615
      Index           =   5
      Left            =   6120
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   4440
      Width           =   10215
   End
   Begin VB.TextBox txtDefiniton 
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
      Height          =   615
      Index           =   4
      Left            =   6120
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3720
      Width           =   10215
   End
   Begin VB.TextBox txtDefiniton 
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
      Height          =   615
      Index           =   3
      Left            =   6120
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3000
      Width           =   10215
   End
   Begin VB.TextBox txtDefiniton 
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
      Height          =   615
      Index           =   2
      Left            =   6120
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2280
      Width           =   10215
   End
   Begin VB.TextBox txtDefiniton 
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
      Height          =   615
      Index           =   1
      Left            =   6120
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1560
      Width           =   10215
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
      Height          =   1335
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8760
      Width           =   2295
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
      Height          =   1335
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8760
      Width           =   2295
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
      Height          =   1335
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   8760
      Width           =   2295
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
      Height          =   1335
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8760
      Width           =   2295
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word10"
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
      Height          =   615
      Index           =   10
      Left            =   1800
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   8040
      Width           =   2175
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word9"
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
      Height          =   615
      Index           =   9
      Left            =   1800
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   7320
      Width           =   2175
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word8"
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
      Height          =   615
      Index           =   8
      Left            =   1800
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   6600
      Width           =   2175
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word7"
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
      Height          =   615
      Index           =   7
      Left            =   1800
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word6"
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
      Height          =   615
      Index           =   6
      Left            =   1800
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word5"
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
      Height          =   615
      Index           =   5
      Left            =   1800
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word4"
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
      Height          =   615
      Index           =   4
      Left            =   1800
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word3"
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
      Height          =   615
      Index           =   3
      Left            =   1800
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word2"
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
      Height          =   615
      Index           =   2
      Left            =   1800
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word1"
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
      Height          =   615
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtDBDef1 
      DataField       =   "Definition1"
      DataSource      =   "adoTests"
      Height          =   1095
      Left            =   11880
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   12000
      Width           =   2175
   End
   Begin VB.TextBox txtDBWord1 
      DataField       =   "Word1"
      DataSource      =   "adoTests"
      Height          =   1095
      Left            =   11640
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   12000
      Width           =   2175
   End
   Begin VB.TextBox txtDBDate 
      DataField       =   "TestDate"
      DataSource      =   "adoTests"
      Height          =   1095
      Left            =   11400
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   12000
      Width           =   2175
   End
   Begin VB.ComboBox cboDate 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      ItemData        =   "frmADETests.frx":0000
      Left            =   3000
      List            =   "frmADETests.frx":0002
      TabIndex        =   0
      Text            =   "Select Date:"
      Top             =   120
      Width           =   3615
   End
   Begin MSAdodcLib.Adodc adoTests 
      Height          =   1335
      Left            =   16560
      Top             =   4320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2355
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
   Begin VB.Label lblDefinition 
      BackStyle       =   0  'Transparent
      Caption         =   "Definitions -"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   35
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblWord 
      BackStyle       =   0  'Transparent
      Caption         =   "Word 10:"
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
      TabIndex        =   34
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label lblWord 
      BackStyle       =   0  'Transparent
      Caption         =   "Word 9:"
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
      Left            =   120
      TabIndex        =   33
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Label lblWord 
      BackStyle       =   0  'Transparent
      Caption         =   "Word 8:"
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
      Left            =   120
      TabIndex        =   32
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label lblWord 
      BackStyle       =   0  'Transparent
      Caption         =   "Word 7:"
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
      Left            =   120
      TabIndex        =   31
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label lblWord 
      BackStyle       =   0  'Transparent
      Caption         =   "Word 6:"
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
      Left            =   120
      TabIndex        =   30
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label lblWord 
      BackStyle       =   0  'Transparent
      Caption         =   "Word 5:"
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
      Left            =   120
      TabIndex        =   29
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblWord 
      BackStyle       =   0  'Transparent
      Caption         =   "Word 4:"
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
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblWord 
      BackStyle       =   0  'Transparent
      Caption         =   "Word 3:"
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
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblWord 
      BackStyle       =   0  'Transparent
      Caption         =   "Word 2:"
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
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblWord 
      BackStyle       =   0  'Transparent
      Caption         =   "Word 1:"
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
      Left            =   120
      TabIndex        =   25
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblDateOfTest 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Test:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmADETests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim recordCount As Integer
Dim totalRecords As Integer

Dim editMode As Boolean

Private Sub cboDate_Click() 'When the combo list is clicked move to the selected date

adoTests.Recordset.MoveFirst
frmADETests.MoveToDate

End Sub

Sub MoveToDate()                    'Keep moving forward until the correct record is found

If cboDate.Text = txtDBTestDate.Text Then

Else
    
    adoTests.Recordset.MoveNext
    frmADETests.MoveToDate

End If

End Sub

Private Sub cmdAdd_Click()  'Add new mode and unlock the text boxes

adoTests.Recordset.AddNew

frmADETests.UnlockTxtWD

End Sub

Private Sub cmdDelete_Click()   'Delete the current record and move to the first record

adoTests.Recordset.Delete
adoTests.Recordset.MoveFirst
totalRecords = totalRecords - 1
recordCount = 1

frmADETests.FBLock
frmADETests.LBUnlock

End Sub

Private Sub cmdEdit_Click() 'Unlock text boxes and turn editmode on

editMode = True

frmADETests.UnlockTxtWD

End Sub

Private Sub cmdExit_Click()     'Unloads current form

Unload Me
frmTeacherMenu.Show

End Sub

Private Sub cmdFirst_Click()    'Moves to the first record and locks the first and previous buttons

recordCount = 1
adoTests.Recordset.MoveFirst
frmADETests.FBLock

UpdateDate

End Sub

Sub FBLock()        'Lock First and Previous buttons while unlocking the next and last buttons

cmdFirst.Enabled = False
cmdPrevious.Enabled = False
cmdNext.Enabled = True
cmdLast.Enabled = True

End Sub

Sub LBLock()        'Lock the Next and Last buttons while unlocking the first and previous buttons

cmdFirst.Enabled = True
cmdPrevious.Enabled = True
cmdNext.Enabled = False
cmdLast.Enabled = False

End Sub

Sub FBUnlock()  'Unlock the first and previous buttons

cmdFirst.Enabled = True
cmdPrevious.Enabled = True

End Sub

Sub LBUnlock()  'Unlock the last and next buttons

cmdNext.Enabled = True
cmdLast.Enabled = True

End Sub

Private Sub cmdLast_Click() 'Go to the last record and lock the next and last buttons

recordCount = totalRecords
adoTests.Recordset.MoveLast
frmADETests.LBLock

UpdateDate

End Sub

Private Sub cmdNext_Click()             'When going to the next record check to is if the EOF is true

If recordCount < totalRecords Then

    recordCount = recordCount + 1
    adoTests.Recordset.MoveNext
    frmADETests.FBUnlock
    
    UpdateDate
    
    If recordCount = totalRecords Then
    
        frmADETests.LBLock

    End If
    
End If

End Sub

Private Sub cmdPrevious_Click() 'Go the previous record and check to see if the BOF is true

If recordCount > 1 Then

    recordCount = recordCount - 1
    adoTests.Recordset.MovePrevious
    frmADETests.LBUnlock
    
    UpdateDate
    
    If recordCount = 1 Then
    
        frmADETests.FBLock

    End If
    
End If

End Sub

Private Sub cmdSave_Click() 'When save is clicked when editmode is false then ask the user if they wish to set a test date

If editMode = False Then

    Dim answ As String
    
    answ = InputBox("Enter Date Of The Test:", "One more thing!")
    txtDBTestDate.Text = answ
    adoTests.Recordset.UpdateBatch

Else

    adoTests.Recordset.UpdateBatch           'Else update the database

End If

frmADETests.LockTxtWD                   'Lock the text boxes move to the first record, refresh the database and lock the first and previous buttons while unlocking the next and last buttons

adoTests.Recordset.MoveFirst
recordCount = 1
totalRecords = adoTests.Recordset.recordCount

adoTests.Refresh

frmADETests.FBLock
frmADETests.LBUnlock

End Sub

Private Sub Form_Load()

    recordCount = 0                                     'Used the in next / previous locking system
    totalRecords = adoTests.Recordset.recordCount
    
    frmADETests.FBLock              'Lock the first and previous buttons
    
    frmADETests.ComboListUpdate     'Update the combo list
    adoTests.Recordset.MoveFirst
    adoTests.Recordset.MovePrevious
    
    frmADETests.LockTxtWD           'Lock the text boxes
    
    Dim num As Integer              'Clear all the text boxes
    
    For num = 1 To 10

        txtWord(num).Text = ""
        
    Next num
    
End Sub

Sub ComboListUpdate()           'Go through each record and add it to the combo list

Do While adoTests.Recordset.EOF = False
        
    cboDate.AddItem (txtDBDate.Text)
    adoTests.Recordset.MoveNext
        
Loop

End Sub

Sub UpdateDate()    'Change the combo list date to the text box date

cboDate.Text = txtDBTestDate.Text

End Sub

Sub LockTxtWD()     'Lock the text boxes

Dim num As Integer

For num = 1 To 10

    txtWord(num).Locked = True
    txtDefiniton(num).Locked = True

Next num

End Sub

Sub UnlockTxtWD()   'Unlock all the text boxes

Dim num As Integer

For num = 1 To 10

    txtWord(num).Locked = False
    txtDefiniton(num).Locked = False

Next num

End Sub
