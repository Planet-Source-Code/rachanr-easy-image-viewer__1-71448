VERSION 5.00
Begin VB.Form frmImageViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Easy Image Viewer"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9240
      Top             =   6720
   End
   Begin VB.Frame Frame2 
      Caption         =   "Preview Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   7335
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   5895
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   6855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Path of Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CheckBox Check1 
         Caption         =   "Preview Image"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   6120
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.FileListBox File1 
         Height          =   3015
         Left            =   240
         TabIndex        =   4
         Top             =   3000
         Width           =   3135
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   3135
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   7
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Image Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6720
      Width           =   9135
   End
End
Attribute VB_Name = "frmImageViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Dir1_Change()
'��˹���鹷ҧ���Ѻ file list box
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
'��˹���鹷ҧ���Ѻ directory
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
'��С�Ȥ�ҵ��������Ѻ����鹷ҧ�ͧ����Ҿ
Dim imgPath As String
'��Ǩ�ͺ�����������ҧ�����������
'�Ѻ������Ǣͧ����������� len
'Len(<�����ŷ���ͧ��ùѺ�ӹǹ����ѡ������Ẻ string & number>)
If Len(Dir1.Path) = 3 Then
    '�����������ҧ��������ͧ��� \
    imgPath = Dir1.Path & File1.FileName
Else
    '��������������������ͷ��� ������ \ ��������ҧ
    '��������Ъ���������
    imgPath = Dir1.Path & "\" & File1.FileName
End If
'��Ǩ�ͺ����ա���礶١�ç check box �������
'����ա���礶١ ��� value = checked ���� 1
'�������ա���礶١ ��� value = unchecked ���� 0
If Check1.Value = Checked Then
    '���������ա���ʴ��Ҿ� image �������� LoadPicture
    'LoadPicture(<��鹷ҧ��Ъ�������ҾẺ���>)
    Image1.Picture = LoadPicture(imgPath)
End If
'�ʴ���鹷ҧ�ͧ�������ͤ�ԡ������ file
Label1.Caption = imgPath
End Sub

Private Sub Form_Load()
'��˹��ٻẺ����ʴ��Ţͧ file list box
File1.Pattern = "*.bmp;*.jpg"
'��˹�����ա���ʴ��ѹ���Ẻ��Ǻ� title �ͧ�����
'�¤���� format ��Ф���� Date (�ʴ��ѹ��͹��)
frmImageViewer.Caption = "Easy Image Viewer" & _
     " : Today is " & Format(Date, "dd mmmm yyyy")
'�ѧ�Ѻ��� directory �ʴ��ŷ����� C:\
Dir1.Path = "c:\"
End Sub

Private Sub Timer1_Timer()
'����ʴ������ Label2 �������� Time (�ʴ�����)
'������������˹� interval ��� timer ��������
Label2.Caption = Format(Time, "HH:MM:SS AM/PM")
End Sub
