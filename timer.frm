VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   12060
   ScaleWidth      =   23700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton preview 
      Caption         =   "Prévisualisation"
      Height          =   495
      Left            =   18360
      TabIndex        =   18
      Top             =   1320
      Width           =   1575
   End
   Begin VB.VScrollBar blue 
      Height          =   855
      Left            =   15120
      Max             =   255
      TabIndex        =   17
      Top             =   720
      Width           =   255
   End
   Begin VB.VScrollBar green 
      Height          =   855
      Left            =   14160
      Max             =   255
      TabIndex        =   16
      Top             =   720
      Width           =   255
   End
   Begin VB.VScrollBar red 
      Height          =   855
      Left            =   13080
      Max             =   255
      TabIndex        =   15
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox init_time 
      Height          =   615
      Left            =   8640
      TabIndex        =   14
      Text            =   "00"
      Top             =   960
      Width           =   855
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   3600
      TabIndex        =   11
      Top             =   7920
      Width           =   4095
   End
   Begin VB.TextBox nom_fichier 
      Height          =   735
      Left            =   8760
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   7800
      Width           =   2895
   End
   Begin VB.CommandButton charger 
      Caption         =   "Charger"
      Height          =   495
      Left            =   21600
      TabIndex        =   8
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton enregistrer 
      Caption         =   "Enregistrer"
      Height          =   495
      Left            =   21600
      TabIndex        =   7
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Valider"
      Height          =   615
      Left            =   18360
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   22920
      Top             =   9960
   End
   Begin VB.Label Label11 
      Caption         =   "Bleu"
      Height          =   495
      Left            =   14640
      TabIndex        =   21
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Vert"
      Height          =   375
      Left            =   13560
      TabIndex        =   20
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "Rouge"
      Height          =   255
      Left            =   12480
      TabIndex        =   19
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   615
      Left            =   14160
      TabIndex        =   13
      Top             =   7920
      Width           =   6615
   End
   Begin VB.Label Label6 
      Caption         =   "Chemin d'accès"
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Nom de votre config"
      Height          =   255
      Left            =   8880
      TabIndex        =   10
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Couleur de police"
      Height          =   255
      Left            =   15960
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Height          =   1215
      Left            =   16560
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Minutes:"
      Height          =   255
      Left            =   7800
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   ":"
      Height          =   375
      Left            =   12480
      TabIndex        =   3
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label seconds_result 
      Caption         =   "00"
      Height          =   495
      Left            =   12840
      TabIndex        =   1
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label minute_result 
      Caption         =   "00"
      Height          =   495
      Left            =   12000
      TabIndex        =   0
      Top             =   4800
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim remaining_time As Integer
Dim font_color
Dim file_path As String
Dim intFile As Integer
Dim intFileBis As Integer
Dim favorite_time As Integer
Dim aux As String
Dim max_time As Integer
Dim total_time As Integer

Private Sub charger_Click()

If nom_fichier.Text = "" Then
    MsgBox "Vous n'avez pas donné de nom à la config à charger!", vbOKCancel
    Exit Sub
End If

intFileBis = FreeFile
Open file_path For Input As #intFileBis

Line Input #intFileBis, font_color
Line Input #intFileBis, aux
remaining_time = Int(aux)

Close
End Sub

Private Sub Dir1_Change()
file_path = Dir1.Path & "\settings_timer_" & nom_fichier.Text & ".txt"
Label7.Caption = file_path
End Sub

Private Sub enregistrer_Click()
If nom_fichier.Text = "" Then
    file_path = Dir1.Path & "\settings_timer" & Int(Rnd() * 10)
End If
favorite_time = Int(init_time.Text)
font_color = RGB(Int(red.Value) Mod 256, Int(green.Value) Mod 256, Int(blue.Value) Mod 256)

Label7.Caption = file_path

intFile = FreeFile
Open file_path For Output As #intFile
Print #intFile, font_color
Print #intFile, favorite_time
Close #intFile
End Sub

Private Sub Command1_Click()

If font_color <> "" Then
    font_color = RGB(Int(red.Value) Mod 256, Int(green.Value) Mod 256, Int(blue.Value) Mod 256)
End If

If remaining_time = 0 Then
    If init_time <> 0 Then
        remaining_time = Int(init_time.Text) * 60
        total_time = remaining_time
    End If
Else
    remaining_time = remaining_time * 60
    total_time = remaining_time
End If
Timer1.Enabled = True

Form1.BackColor = RGB(0, 255, 0)
seconds_result.BackColor = RGB(0, 255, 0)
minute_result.BackColor = RGB(0, 255, 0)
Label1.BackColor = RGB(0, 255, 0)

seconds_result.ForeColor = font_color
minute_result.ForeColor = font_color
Label1.ForeColor = font_color

init_time.Visible = False
red.Visible = False
green.Visible = False
blue.Visible = False
Command1.Visible = False

red.Visible = False
green.Visible = False
blue.Visible = False

Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False

Dir1.Visible = False
charger.Visible = False
enregistrer.Visible = False
nom_fichier.Visible = False
preview.Visible = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    
    Timer1.Enabled = False
    
    init_time.Visible = True
    red.Visible = True
    green.Visible = True
    blue.Visible = True
    Command1.Visible = True

    red.Visible = True
    green.Visible = True
    blue.Visible = True

    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Label9.Visible = True
    Label10.Visible = True
    Label11.Visible = True

    Dir1.Visible = True
    charger.Visible = True
    enregistrer.Visible = True
    nom_fichier.Visible = True
    
    remaining_time = 0

End If

If KeyCode = vbKeyF1 Then
    If Timer1.Enabled = False Then
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If
End If
If KeyCode = vbKeyF5 Then
    Timer1.Enabled = False
  
    If remaining_time <> total_time Then
        remaining_time = Int(total_time)
    End If
    
    Timer1.Enabled = True
    
End If
End Sub

Private Sub nom_fichier_Change()
file_path = Dir1.Path & "\settings_timer_" & nom_fichier.Text & ".txt"
Label7.Caption = Dir1.Path & "\settings_timer_" & nom_fichier.Text & ".txt"
End Sub

Private Sub preview_Click()

font_color = RGB(red.Value, green.Value, blue.Value)

seconds_result.BackColor = RGB(0, 255, 0)
minute_result.BackColor = RGB(0, 255, 0)
seconds_result.ForeColor = font_color
minute_result.ForeColor = font_color

End Sub

Private Sub red_Change()
Label3.BackColor = RGB(Int(red.Value) Mod 256, Int(green.Value) Mod 256, Int(blue.Value) Mod 256)
End Sub

Private Sub green_Change()
Label3.BackColor = RGB(Int(red.Value) Mod 256, Int(green.Value) Mod 256, Int(blue.Value) Mod 256)
End Sub

Private Sub blue_Change()
Label3.BackColor = RGB(Int(red.Value), Int(green.Value), Int(blue.Value) Mod 256)
End Sub

Private Sub Timer1_Timer()
seconds_result.Caption = remaining_time Mod 60
minute_result.Caption = Int(remaining_time / 60)
If remaining_time > 0 Then
    remaining_time = remaining_time - 1
End If

If remaining_time = 0 Then
    seconds_result.Caption = 0
    minute_result.Caption = 0
    Timer1.Enabled = False
End If

End Sub
