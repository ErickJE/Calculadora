VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "calculadora"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command20 
      Caption         =   "salir"
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command19 
      Caption         =   "^"
      Height          =   615
      Left            =   3360
      TabIndex        =   19
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Raiz"
      Height          =   615
      Left            =   2640
      TabIndex        =   18
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      Caption         =   "c"
      Height          =   615
      Left            =   3000
      TabIndex        =   17
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      Caption         =   "/"
      Height          =   615
      Left            =   3360
      TabIndex        =   16
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "*"
      Height          =   495
      Left            =   3360
      TabIndex        =   15
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "-"
      Height          =   495
      Left            =   2640
      TabIndex        =   14
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "+"
      Height          =   615
      Left            =   2640
      TabIndex        =   13
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "="
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton Command11 
      Caption         =   "."
      Height          =   495
      Left            =   3000
      TabIndex        =   11
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "1"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtresultado 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim operador1 As Single
Dim operador2 As Single
Dim operacion As Single
Dim resultado As Integer



Private Sub Command1_Click()
txtresultado.Text = txtresultado.Text + "1"

End Sub

Private Sub Command10_Click()
txtresultado.Text = txtresultado.Text + "0"

End Sub

Private Sub Command11_Click()
txtresultado.Text = txtresultado.Text + ","

End Sub

Private Sub Command12_Click()
If operacion = 1 Then
operador2 = txtresultado.Text
resultado = operador1 + operador2
txtresultado.Text = CStr(resultado)
End If

If operacion = 2 Then
operador2 = txtresultado.Text
resultado = operador1 - operador2
txtresultado.Text = CStr(resultado)
End If

If operacion = 3 Then
operador2 = txtresultado.Text
resultado = operador1 * operador2
txtresultado.Text = CStr(resultado)
End If

If operacion = 4 Then
operador2 = txtresultado.Text
resultado = operador1 / operador2
txtresultado.Text = CStr(resultado)
End If

If operacion = 5 Then
operador2 = txtresultado.Text
resultado = operador1 ^ operador2
txtresultado.Text = CStr(resultado)
End If




End Sub

Private Sub Command13_Click()
operacion = 1
operador1 = txtresultado.Text
txtresultado.Text = ""

End Sub

Private Sub Command14_Click()
operacion = 2
operador1 = txtresultado.Text
txtresultado.Text = ""

End Sub

Private Sub Command15_Click()
operacion = 3
operador1 = txtresultado.Text
txtresultado.Text = ""

End Sub

Private Sub Command16_Click()
operacion = 4
operador1 = txtresultado.Text
txtresultado.Text = ""

End Sub

Private Sub Command17_Click()
txtresultado.Text = ""

End Sub

Private Sub Command18_Click()
txtresultado.Text = Math.Sqr(Val(txtresultado.Text))

End Sub

Private Sub Command19_Click()
operacion = 5
operador1 = txtresultado.Text
txtresultado.Text = ""

End Sub

Private Sub Command2_Click()
txtresultado.Text = txtresultado.Text + "2"

End Sub

Private Sub Command20_Click()
End

End Sub

Private Sub Command3_Click()
txtresultado.Text = txtresultado.Text + "3"

End Sub

Private Sub Command4_Click()
txtresultado.Text = txtresultado.Text + "4"

End Sub

Private Sub Command5_Click()
txtresultado.Text = txtresultado.Text + "5"

End Sub

Private Sub Command6_Click()
txtresultado.Text = txtresultado.Text + "6"

End Sub

Private Sub Command7_Click()
txtresultado.Text = txtresultado.Text + "7"

End Sub

Private Sub Command8_Click()
txtresultado.Text = txtresultado.Text + "8"

End Sub

Private Sub Command9_Click()
txtresultado.Text = txtresultado.Text + "9"

End Sub
