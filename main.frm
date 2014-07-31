VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydrostatic pressure calculator for pipes"
   ClientHeight    =   5148
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   6276
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5148
   ScaleWidth      =   6276
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Calc"
      Height          =   492
      Left            =   2880
      TabIndex        =   8
      Top             =   3240
      Width           =   2052
   End
   Begin VB.TextBox TextBox5 
      Height          =   288
      Left            =   2880
      TabIndex        =   7
      Top             =   2880
      Width           =   2052
   End
   Begin VB.TextBox TextBox4 
      Height          =   288
      Left            =   2880
      TabIndex        =   6
      Top             =   2400
      Width           =   2052
   End
   Begin VB.ComboBox ComboBox3 
      Height          =   288
      Left            =   3840
      TabIndex        =   5
      Text            =   "select steel"
      Top             =   1680
      Width           =   2052
   End
   Begin VB.ComboBox ComboBox2 
      Height          =   288
      ItemData        =   "main.frx":0000
      Left            =   3840
      List            =   "main.frx":000A
      TabIndex        =   4
      Text            =   "select requirements"
      Top             =   1200
      Width           =   2052
   End
   Begin VB.TextBox TextBox2 
      Height          =   288
      Left            =   3840
      TabIndex        =   3
      Top             =   720
      Width           =   2052
   End
   Begin VB.ComboBox ComboBox1 
      Height          =   288
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   2052
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   840
      Picture         =   "main.frx":0018
      ScaleHeight     =   492
      ScaleWidth      =   612
      TabIndex        =   1
      Top             =   4440
      Width           =   612
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   120
      Picture         =   "main.frx":149A
      ScaleHeight     =   480
      ScaleWidth      =   624
      TabIndex        =   0
      Top             =   4440
      Width           =   624
   End
   Begin VB.Label Label8 
      Height          =   252
      Left            =   600
      TabIndex        =   16
      Top             =   3360
      Width           =   372
   End
   Begin VB.Label Label7 
      Caption         =   $"main.frx":291C
      Height          =   852
      Left            =   1560
      TabIndex        =   15
      Top             =   4200
      Width           =   4600
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6000
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label6 
      Caption         =   "MPa"
      Height          =   252
      Left            =   5160
      TabIndex        =   14
      Top             =   2880
      Width           =   732
   End
   Begin VB.Label Label5 
      Caption         =   "kgf/sm^2"
      Height          =   252
      Left            =   5160
      TabIndex        =   13
      Top             =   2400
      Width           =   732
   End
   Begin VB.Label Label4 
      Caption         =   "Steel"
      Height          =   252
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1452
   End
   Begin VB.Label Label3 
      Caption         =   "Requirements and standarts"
      Height          =   252
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   2652
   End
   Begin VB.Label Label2 
      Caption         =   "Wall thickness, mm"
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   1932
   End
   Begin VB.Label Label1 
      Caption         =   "Outer Diamater, mm"
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1812
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public language As Double

Private Sub ComboBox2_LostFocus()
If language = 1 Then
If ComboBox2.ListIndex = 0 Then
ComboBox3.Clear
ComboBox3.AddItem "Сталь 10"
ComboBox3.AddItem "Сталь 20"
ComboBox3.AddItem "Сталь 09Г2С"
End If

If ComboBox2.ListIndex = 1 Then
ComboBox3.Clear
ComboBox3.AddItem "Сталь 20А"
ComboBox3.AddItem "Сталь 13ХФА"
End If
End If


If language = 2 Then
If ComboBox2.ListIndex = 0 Then
ComboBox3.Clear
ComboBox3.AddItem "Steel 10"
ComboBox3.AddItem "Steel 20"
ComboBox3.AddItem "Steel 09G2S (A516 US) 09Mn2-Si"
End If

If ComboBox2.ListIndex = 1 Then
ComboBox3.Clear
ComboBox3.AddItem "Steel 20A"
ComboBox3.AddItem "Steel 13Cr-V A"
End If
End If


End Sub

Private Sub Command1_Click()
Dim D As Double
Dim S As Double
'wall thickness
Dim R As Double
Dim P As Double
Dim NPD As Double
Dim STEEL As Double
Dim ERROR As Boolean

ERROR = False

If IsNumeric(ComboBox1.List(ComboBox1.ListIndex)) Then
D = ComboBox1.List(ComboBox1.ListIndex)
Else
If language = 1 Then
ERROR = MsgBox("Ошибка, выберите внешний диаметр!", vbOKOnly, "Ошибка")
End If
If language = 2 Then
ERROR = MsgBox("Error, please select outer diameter!", vbOKOnly, "Error")
End If
End If

If IsNumeric(TextBox2.Text) Then
S = TextBox2.Text
Else
If language = 1 Then
ERROR = MsgBox("Ошибка, введите толщину стенки!", vbOKOnly, "Ошибка")
End If
If language = 2 Then
ERROR = MsgBox("Error, please input wall thickness!", vbOKOnly, "Error")
End If
End If

If language = 1 Then
    If ComboBox2.Text = "Выберите НТД" Then
    ERROR = MsgBox("Ошибка, введите НТД!", vbOKOnly, "Ошибка")
    Else
        If ComboBox2.Text = "ГОСТ" Then
        NPD = 0.4 * 1
        Else
        NPD = 0.8 * 1
        End If
    End If
End If

If language = 2 Then
    If ComboBox2.Text = "select requirements" Then
    ERROR = MsgBox("Error, please select type of requirements!", vbOKOnly, "Error")
    Else
        If ComboBox2.Text = "GOST" Then
        NPD = 0.4 * 1
        Else
        NPD = 0.8 * 1
        End If
    End If
End If

If language = 1 Then
    
        Select Case ComboBox3.Text
        Case "Сталь 10"
        STEEL = 36
        Case "Сталь 20"
        STEEL = 42
        Case "Сталь 09Г2С"
        STEEL = 48
        Case "Сталь 20А"
        STEEL = 69 / 2 'it is about 34.5 but becuase there errors with . , within russian and en-us locales
        Case "Сталь 13ХФА"
        STEEL = 38
        End Select
        
        If STEEL = 0 Then
        ERROR = MsgBox("Ошибка, выберите сталь!", vbOKOnly, "Ошибка")
        End If
End If

If language = 2 Then
   
        Select Case ComboBox3.Text
        Case "Steel 10"
        STEEL = 36
        Case "Steel 20"
        STEEL = 42
        Case "Steel 09G2S (A516 US) 09Mn2-Si"
        STEEL = 48
        Case "Steel 20A"
        STEEL = 69 / 2 'it is about 34.5 but becuase there errors with . , within russian and en-us locales
        Case "Steel 13Cr-V A"
        STEEL = 38
        End Select
        
        If STEEL = 0 Then
        ERROR = MsgBox("Error, please select steel!", vbOKOnly, "Error")
        End If
    
End If


If ERROR = False Then
P = Round(200 * S * 0.875 * NPD * STEEL / (D - S), 2)
TextBox4.Text = P
TextBox5.Text = Round(P * 0.0980665, 2)
End If
End Sub

Private Sub Form_Load()
ComboBox1.AddItem "10"
ComboBox1.AddItem "10,2"
ComboBox1.AddItem "12"
ComboBox1.AddItem "13"
ComboBox1.AddItem "14"
ComboBox1.AddItem "16"
ComboBox1.AddItem "18"
ComboBox1.AddItem "19"
ComboBox1.AddItem "20"
ComboBox1.AddItem "21,3"
ComboBox1.AddItem "22"
ComboBox1.AddItem "24"
ComboBox1.AddItem "25"
ComboBox1.AddItem "26"
ComboBox1.AddItem "27"
ComboBox1.AddItem "28"
ComboBox1.AddItem "30"
ComboBox1.AddItem "32"
ComboBox1.AddItem "33"
ComboBox1.AddItem "33,7"
ComboBox1.AddItem "35"
ComboBox1.AddItem "36"
ComboBox1.AddItem "38"
ComboBox1.AddItem "40"
ComboBox1.AddItem "42"
ComboBox1.AddItem "44.5"
ComboBox1.AddItem "45"
ComboBox1.AddItem "48"
ComboBox1.AddItem "48.3"
ComboBox1.AddItem "51"
ComboBox1.AddItem "53"
ComboBox1.AddItem "54"
ComboBox1.AddItem "57"
ComboBox1.AddItem "60"
ComboBox1.AddItem "63.5"
ComboBox1.AddItem "70"
ComboBox1.AddItem "73"
ComboBox1.AddItem "76"
ComboBox1.AddItem "83"
ComboBox1.AddItem "89"
ComboBox1.AddItem "95"
ComboBox1.AddItem "102"
ComboBox1.AddItem "108"
ComboBox1.AddItem "114"
ComboBox1.AddItem "127"
ComboBox1.AddItem "133"
ComboBox1.AddItem "140"
ComboBox1.AddItem "152"
ComboBox1.AddItem "159"
ComboBox1.AddItem "168"
ComboBox1.AddItem "177.8"
ComboBox1.AddItem "180"
ComboBox1.AddItem "193.7"
ComboBox1.AddItem "219"
ComboBox1.AddItem "244.5"
ComboBox1.AddItem "273"
ComboBox1.AddItem "325"
ComboBox1.AddItem "355.6"
ComboBox1.AddItem "377"
ComboBox1.AddItem "406.4"
ComboBox1.AddItem "426"
ComboBox1.AddItem "478"
ComboBox1.AddItem "530"
ComboBox1.AddItem "630"
ComboBox1.AddItem "720"
ComboBox1.AddItem "820"
ComboBox1.AddItem "920"
ComboBox1.AddItem "1020"
ComboBox1.AddItem "1120"
ComboBox1.AddItem "1220"
ComboBox1.AddItem "1420"

language = 1
Label1.Caption = "Внешний диаметр, мм"
Label2.Caption = "Толщина стенки, мм"
Label3.Caption = "НТД"
Label4.Caption = "Сталь"
Label5.Caption = "кгс/см2"
Label6.Caption = "МПа"
Label7.Caption = "Эта программа бесплатна и распространяется под лицензией GNU GPL v2. Автор Гладышев Антон, gladyshev@yandex.ru , 2014. Исходники доступны github.com/areso/hydrostatic_test"
Command1.Caption = "Рассчитать"
main.Caption = "Расчет гидравлического давления труб"
ComboBox2.Clear
ComboBox2.AddItem ("ГОСТ")
ComboBox2.AddItem ("ТУ")
ComboBox2.Text = "Выберите НТД"
ComboBox3.Text = "Выберите сталь"
End Sub

Private Sub Picture2_Click()
language = 2
Label1.Caption = "Outer Diamater, mm"
Label2.Caption = "Wall thickness, mm"
Label3.Caption = "Requirements and standarts"
Label4.Caption = "Steel"
Label5.Caption = "kgf/sm^2"
Label6.Caption = "MPa"
Label7.Caption = "This programm is free(as beer) and licensed under GNU GPL v2. Author is Gladyshev Anton, gladyshev@yandex.ru , 2014. Sources available github.com/areso/hydrostatic_test"
Command1.Caption = "Calc"
main.Caption = "Hydrostatic pressure calculator for pipes"
ComboBox2.Clear
ComboBox2.AddItem ("GOST")
ComboBox2.AddItem ("Technical Specs")
ComboBox2.Text = "select requirements"
ComboBox3.Text = "select steel"
End Sub

Private Sub Picture1_Click()
language = 1
Label1.Caption = "Внешний диаметр, мм"
Label2.Caption = "Толщина стенки, мм"
Label3.Caption = "НТД"
Label4.Caption = "Сталь"
Label5.Caption = "кгс/см2"
Label6.Caption = "МПа"
Label7.Caption = "Эта программа бесплатна и распространяется под лицензией GNU GPL v2. Автор Гладышев Антон, gladyshev@yandex.ru , 2014. Исходники доступны github.com/areso/hydrostatic_test"
Command1.Caption = "Рассчитать"
main.Caption = "Расчет гидравлического давления труб"
ComboBox2.Clear
ComboBox2.AddItem ("ГОСТ")
ComboBox2.AddItem ("ТУ")
ComboBox2.Text = "Выберите НТД"
ComboBox3.Text = "Выберите сталь"
End Sub
