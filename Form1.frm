VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save/Load Bitmap"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRetr 
      Caption         =   "Retrieve"
      Height          =   435
      Left            =   2280
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   435
      Left            =   2280
      TabIndex        =   3
      Top             =   1380
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get Pic"
      Height          =   435
      Left            =   2280
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   2595
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_strPicData As String
Dim m_lngPicLen As Long

Dim m_db As Database

Private Sub cmdGet_Click()

    On Error GoTo EH
    
    With CDlg
        .CancelError = True
        .Filter = "*.bmp|*.bmp"
        .InitDir = App.Path
        .ShowOpen
        '
        ' Load the picture
        Picture1.Picture = LoadPicture(.FileName)
        '
        ' Read pic into memory
        Open .FileName For Binary Access Read As #1
            m_strPicData = Space$(LOF(1))
            Get #1, , m_strPicData
        Close #1
    End With
    Exit Sub
    
EH:
    
End Sub


Private Sub cmdSave_Click()
    
    Dim rs As Recordset
    '
    ' Open database
    Set m_db = OpenDatabase(App.Path & "\photo97.mdb")
    '
    ' Delete all records
    m_db.Execute "DELETE FROM Photos"
    
    Set rs = m_db.OpenRecordset("Photos")
    rs.AddNew
    rs.Fields("Photo").AppendChunk m_strPicData
    rs.Update
    
    rs.Close
    m_db.Close
    
End Sub


Private Sub cmdClear_Click()

    m_strPicData = ""
    Picture1.Picture = LoadPicture(m_strPicData)
    
End Sub


Private Sub cmdRetr_Click()

    Dim rs As Recordset
    Dim TempFile As String
    '
    ' Define temp file
    TempFile = App.Path & "\tmp.bmp"
    '
    ' Open database
    Set m_db = OpenDatabase(App.Path & "\photo97.mdb")
    
    Set rs = m_db.OpenRecordset("Photos")
    m_lngPicLen = rs.Fields("Photo").FieldSize
    '
    ' If there's data in the picture field,
    ' save it to temp file and load it.
    If m_lngPicLen > 0 Then
        '
        ' Extract from database
        m_strPicData = rs.Fields("Photo").GetChunk(0, m_lngPicLen)
        '
        ' Save to temp file
        Open TempFile For Binary As #1
            Put #1, , m_strPicData
        Close #1
        '
        ' Load into picture box
        Picture1.Picture = LoadPicture(TempFile)
        '
        ' Delete temp file
        Kill TempFile
    Else
        Picture1.Picture = LoadPicture("")
    End If
    
    rs.Close
    m_db.Close
    
End Sub


