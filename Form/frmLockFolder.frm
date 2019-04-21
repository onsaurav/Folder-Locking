VERSION 5.00
Begin VB.Form frmLockFolder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ONS Lock Folder"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   ControlBox      =   0   'False
   Icon            =   "frmLockFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   4575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdUnLock 
      Caption         =   "Unlock"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "Lock"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Folder Locking System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmLockFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title          :   ONSFolderLocking
'Author         :   Self
'URL            :   onsaurav@yahoo.com/onsaurav@gmail.com/onsaurav@hotmail.com
'Description    :   A folder locking System
'Created        :   Saurav Biswas /  May-20-2010
'Modified       :   Saurav Biswas /

Private Sub cmdClose_Click()
        End
End Sub

Private Sub cmdLock_Click()
        'Some string to use for changing the folder behavior
        '".{2559a1f2-21d7-11d4-bdaf-00c04f60b9f0}" Lock
        '".{21EC2020-3AEA-1069-A2DD-08002B30309D}" Control Panal
        '".{2559a1f4-21d7-11d4-bdaf-00c04f60b9f0}" Unknown Type
        '".{645FF040-5081-1018-9F09-00AA002F954E}" Recyclebin
        '".{2559a1f1-21d7-11d4-bdaf-00c04f60b9f0}" Help
        '".{7007ACC7-3202-11b1-AAD2-00805FC1270E}" Network
        On Error GoTo Ext
              
        'Checking the folder already locked or not
        If InStr(strFileName, ".{21EC2020-3AEA-1069-A2DD-08002B30309D}") > 0 Then
           MsgBox "Folder Already locked", vbInformation, vbOK: Exit Sub
        Else
           'Comparing the password
           frmPassword.Show vbModal
           If strPassWord = "" Then MsgBox "Invalid password", vbInformation, vbOK: Exit Sub
           
           'Store the password in a file.
           Dim fn As Integer
           fn = FreeFile
           Open strFileName & "\ONSPSF0000001" For Append As #fn
           Print #fn, strPassWord
           Close #1
           
           'rename the folder with the extention
           Name strFileName As strFileName & ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"
        End If
        strFileName = "": strPassWord = ""
        MsgBox "File lock complete sccessfully", vbInformation, vbOK: Exit Sub
        Dir1.Refresh
        Exit Sub
Ext:
        strFileName = "": strPassWord = ""
        MsgBox "Sorry! Unable to lock the folder", vbInformation, vbOK: Exit Sub
        Err.Clear
End Sub

Private Sub cmdUnLock_Click()
        'Checking the folder already locked or not
        If InStr(strFileName, ".{21EC2020-3AEA-1069-A2DD-08002B30309D}") <= 0 Then
           MsgBox "Folder Already locked", vbInformation, vbOK: Exit Sub
        Else
           'Checking the password
           frmPassword.Show vbModal
           If strPassWord = "" Then MsgBox "Invalid password", vbInformation, vbOK: Exit Sub
           Dim fn As Integer
           Dim Pass As String
           fn = FreeFile
           Open strFileName & "\ONSPSF0000001" For Input As #fn
           Input #fn, Pass
           Close #1
                       
           If Pass <> strPassWord Then MsgBox "Invalid password", vbInformation, vbOK: Exit Sub
           
           'Rename the folder without extention
           Name strFileName As Replace(strFileName, ".{21EC2020-3AEA-1069-A2DD-08002B30309D}", "")
           If Dir(Replace(strFileName, ".{21EC2020-3AEA-1069-A2DD-08002B30309D}", "") & "\ONSPSF0000001") <> "" Then
              'remove the password file
              Kill (Replace(strFileName, ".{21EC2020-3AEA-1069-A2DD-08002B30309D}", "") & "\ONSPSF0000001")
           End If
        End If
        strFileName = "": strPassWord = ""
        MsgBox "File unlock complete sccessfully", vbInformation, vbOK: Exit Sub
        Dir1.Refresh
        Exit Sub
Ext:
        strFileName = "": strPassWord = ""
        MsgBox "Sorry! Unable to unlock the folder", vbInformation, vbOK: Exit Sub
        Err.Clear
End Sub

Private Sub Dir1_Click()
        On Error Resume Next
        strFileName = ""
        strFileName = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
        Dir1.Path = Drive1.Drive
End Sub

