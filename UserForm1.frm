VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9480.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Position As Integer
Dim Gesamt As Integer
Dim Falsch As Integer
Dim List(1024) As String
Dim i As Integer


Private Sub UserForm_Initialize()
Gesamt = 0
Falsch = 0
GetRange
Position = 1
Label1.Caption = CStr(Position) & ": " & Sheets("Tabelle1").Range("B" & List(Position)).Value
End Sub

Private Sub CommandButton1_Click()
Gesamt = Gesamt + 1
Label3 = "von: " & CStr(Gesamt)

If Sheets("Tabelle1").Range("A" & List(Position)).Value = TextBox1.Text Then
If Position < i Then
Position = Position + 1
Label1.Caption = CStr(Position) & ": " & Sheets("Tabelle1").Range("B" & List(Position)).Value
TextBox1.Text = ""
TextBox1.BackColor = &H80000005
UserForm1.TextBox1.SetFocus
Else
ende
End If
Else
TextBox1.BackColor = &HFF&
UserForm1.TextBox1.SetFocus
Falsch = Falsch + 1
Label2.Caption = "Falsch: " & CStr(Falsch)
End If
End Sub


Private Sub CommandButton2_Click()
MsgBox (Sheets("Tabelle1").Range("A" & List(Position)).Value)
End Sub

Private Sub SpinButton1_SpinDown()
If Position < i Then
Position = Position + 1
Label1.Caption = CStr(Position) & ": " & Sheets("Tabelle1").Range("B" & List(Position)).Value
End If
End Sub


Private Sub SpinButton1_SpinUp()
If Position <> 1 Then
Position = Position - 1
Label1.Caption = CStr(Position) & ": " & Sheets("Tabelle1").Range("B" & List(Position)).Value
End If
End Sub

Sub GetRange()
   Dim Cell As Range
   Dim Pos As Integer
   Dim str As String
   Dim b As Boolean
   b = False
   
   If TypeOf Selection Is Range Then
     For Each Cell In Selection
     str = Cell.AddressLocal
     Pos = InStr(str, "$A$")
     If Pos > 0 Then
     i = i + 1
     List(i) = Right(str, Len(str) - Pos - 2)
     b = True
     Else
     If b = False Then
     MsgBox ("Bitte Makieren Sie einen gültigen Bereich")
     End
     End If
     End If
     
     
     Next
   End If
End Sub

Sub ende()
If Falsch = O Then
MsgBox ("Fertig, Super alles richtig")
Else
MsgBox ("Fertig, Du hast " & CStr(Falsch) & " von " & CStr(Gesamt) & " Vokabeln Falsch")
End If
End Sub

