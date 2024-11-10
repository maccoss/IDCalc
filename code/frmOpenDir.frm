VERSION 5.00
Begin VB.Form frmOpenDir 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Directory"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton DirCancelBtn 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton DirOKbtn 
      Caption         =   "Select Directory"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.DriveListBox DriveListBox 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   2895
   End
   Begin VB.DirListBox DirListBox 
      Height          =   2115
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmOpenDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DirCancelBtn_Click()
    Unload Me
End Sub

Public Sub DirOKbtn_Click()
Dim directory, filename As String
    
RelExfrm.FileList.Clear
RelExfrm.filelblB.Caption = " "
RelExfrm.ratiolblB.Caption = " "
RelExfrm.bkgrdlblB.Caption = " "
RelExfrm.peptidelblB.Caption = " "
RelExfrm.locuslblB.Caption = " "
    
    
    directory = DirListBox.List(DirListBox.ListIndex)
    ChDir (directory)
     
     RelExfrm.dirlbl.Caption = directory & "\"
     exten = "*chr"
     temp = directory & "\" & exten
     
     
     filename = Dir$(temp, vbDirectory)
         
     Do While Len(filename)
                  
         RelExfrm.FileList.AddItem (filename)
         
         filename = Dir$
     Loop
        
         
    Unload Me
End Sub

Private Sub DriveListBox_Change()
    On Error GoTo Error_Handler
    
    DirListBox.Path = DriveListBox.Drive
       
Error_Handler:
If Err Then MsgBox "Drive not available", 64, "Notice"
   
    
End Sub
