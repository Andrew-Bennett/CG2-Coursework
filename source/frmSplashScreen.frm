VERSION 5.00
Begin VB.Form frmSplashScreen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Splash Screen"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr 
      Interval        =   1000
      Left            =   3360
      Top             =   3000
   End
   Begin VB.Image imgBee 
      Height          =   1935
      Left            =   2760
      Picture         =   "frmSplashScreen.frx":0000
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spelling Bee"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   7335
   End
   Begin VB.Label lblClickInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click anywhere to continue"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   6975
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load() 'Centers the image of a bee

    imgBee.Top = (frmSplashScreen.Height / 2) - imgBee.Height / 2   'Centers the image of the bee to the center of the screen no matter what the size of the form is
    imgBee.Left = (frmSplashScreen.Width / 2) - imgBee.Width / 2
    
End Sub

Private Sub Form_Click()            'Responds to input and shows login screen

frmSplashScreen.DisplayLogin

End Sub

Private Sub lblClickInfo_Click()    'Responds to input and shows login screen

frmSplashScreen.DisplayLogin

End Sub

Sub DisplayLogin()                  'Displays the login screen

frmLogin.ShowScreen

End Sub

Private Sub lblTitle_Click()        'Responds to input and shows login screen

frmSplashScreen.DisplayLogin

End Sub

Sub Remove()                       'Removes the form

Unload Me

End Sub

Private Sub tmr_Timer()          'After a certian amount of time it opens up the login screen

frmSplashScreen.DisplayLogin

End Sub
