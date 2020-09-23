VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmExplore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Written By - Kev Heywood (uk) - Visit www.dlcs.fsnet.co.uk"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   Icon            =   "frmExplore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.TreeView tvDir 
      Height          =   3000
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5292
      _Version        =   327682
      Indentation     =   0
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImgLstFolder"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ListView LstV 
      Height          =   3000
      Left            =   3090
      TabIndex        =   4
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5292
      View            =   2
      Arrange         =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   6360
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.FileListBox FileBox 
      Height          =   285
      Left            =   6120
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.DirListBox DirBox 
      Height          =   315
      Left            =   6120
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.DriveListBox DriveBox 
      Height          =   315
      Left            =   6120
      TabIndex        =   0
      Top             =   4200
      Visible         =   0   'False
      Width           =   2145
   End
   Begin ComctlLib.ImageList imglstLV 
      Left            =   7680
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExplore.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExplore.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExplore.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExplore.frx":0640
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImgLstFolder 
      Left            =   7080
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   128
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExplore.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmExplore.frx":0CA4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmExplore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by Kev Heywood (uk) - Visit www.dlcs.fsnet.co.uk
Option Explicit
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal rGetIcon As Long) As Long
Dim fPath As String, rGetIcon As Long

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call GetDriveList
End Sub

Private Sub GetDriveList()
    Dim idx As Integer, dPath As String
    tvDir.Nodes.Clear
    For idx = 0 To DriveBox.ListCount - 1
        dPath = Left(DriveBox.List(idx), 2) & "\"
        tvDir.Nodes.Add , , dPath, DriveBox.List(idx), 1
        tvDir.Nodes.Add dPath, tvwChild, ""
    Next idx
End Sub

Private Sub tvDir_Expand(ByVal Node As ComctlLib.Node)
    On Error GoTo ErrHdler
    Dim idx As Integer, fPos As Integer
    Dim rPath As String, fName As String, fNewPath As String
    MousePointer = 11
    If Node.Child.Text = "" Then
        tvDir.Nodes.Remove Node.Child.Index
        rPath = Node.Key
        DirBox.Path = rPath
        fPos = Len(rPath) + 1
        'Add Folders
        For idx = 0 To DirBox.ListCount - 1
            fName = Mid(DirBox.List(idx), fPos)
            fNewPath = rPath & fName & "\"
            tvDir.Nodes.Add rPath, tvwChild, fNewPath, fName, 1
            DirBox.Path = fNewPath
            If (FileBox.ListCount > 0) Or (DirBox.ListCount > 0) Then
                tvDir.Nodes.Add fNewPath, tvwChild, , ""
                tvDir.Nodes(fNewPath).ExpandedImage = 2
            End If
            DirBox.Path = rPath
        Next idx
    End If
    GoTo ExitSub
ErrHdler:
    'If Drive not ready handle error and re-instate removed item
    tvDir.Nodes.Add Node.Key, tvwChild, , ""
    Resume ExitSub
ExitSub:
    MousePointer = 0
End Sub
Private Sub tvDir_NodeClick(ByVal Node As Node)
    Static tempNK As Integer
    Me.Caption = Node.Key
    tvDir.Nodes(Node.Index).Image = 2
    tvDir.Nodes(Node.Index).ExpandedImage = 2
    If tempNK <> Empty And tempNK <> Node.Index Then
        tvDir.Nodes(tempNK).ExpandedImage = 1
        tvDir.Nodes(tempNK).Image = 1
    End If
    tempNK = Node.Index
    Call AddFiles
End Sub

Private Sub AddFiles()
    On Error Resume Next
    fPath = IIf(Right(DirBox.Path, 1) = "\", DirBox.Path, DirBox.Path & "\")
    If fPath = tvDir.SelectedItem.Key Then Exit Sub
    DirBox.Path = tvDir.SelectedItem.Key
    Dim idx As Integer, sFlderPos As Integer, sloop As Integer
    Dim sFldrName As String
    MousePointer = 11
    LstV.ListItems.Clear
    Set LstV.Icons = Nothing
    Set LstV.SmallIcons = Nothing
    imglstLV.ListImages.Clear
    imglstLV.ImageHeight = 24: imglstLV.ImageWidth = 24
    'Add Any Directory Folders to the ListView
    fPath = IIf(Right(DirBox.Path, 1) = "\", DirBox.Path, DirBox.Path & "\")
    For idx = 0 To DirBox.ListCount - 1
        rGetIcon = ExtractAssociatedIcon(0, DirBox.List(idx) + "\", 1)
        Set picTemp.Picture = Nothing
        DrawIcon picTemp.hdc, 0, 0, rGetIcon
        picTemp.Picture = picTemp.Image
        LstV = imglstLV.ListImages.Add(, , picTemp.Picture)
        DoEvents
    Next idx
    LstV.Icons = imglstLV
    LstV.SmallIcons = imglstLV
    For idx = 0 To DirBox.ListCount - 1
        For sloop = Len(DirBox.List(idx)) To 1 Step -1
            If Mid(DirBox.List(idx), sloop, 1) = "\" Then Exit For
            sFlderPos = sFlderPos + 1
        Next sloop
        sFldrName = Right(DirBox.List(idx), sFlderPos): sFlderPos = 0
        LstV.ListItems.Add , , sFldrName, idx + 1, idx + 1
    Next idx
    'Add the Files contained in the selected directory to the ListView
    fPath = IIf(Right(FileBox.Path, 1) = "\", FileBox.Path, FileBox.Path & "\")
    For idx = 0 To FileBox.ListCount - 1
        rGetIcon = ExtractAssociatedIcon(0, fPath & FileBox.List(idx), 1)
        Set picTemp.Picture = Nothing
        DrawIcon picTemp.hdc, 0, 0, rGetIcon
        picTemp.Picture = picTemp.Image
        LstV = imglstLV.ListImages.Add(, , picTemp.Picture)
        DoEvents
    Next idx
    LstV.Icons = imglstLV
    LstV.SmallIcons = imglstLV
    For idx = DirBox.ListCount To DirBox.ListCount + FileBox.ListCount - 1
        LstV.ListItems.Add , , FileBox.List(idx - DirBox.ListCount), idx + 1, idx + 1
    Next idx
    MousePointer = 0
End Sub

Private Sub DirBox_Change()
    FileBox.Path = DirBox.Path
End Sub

Private Sub DriveBox_Change()
    DirBox.Path = DriveBox.Path
End Sub

