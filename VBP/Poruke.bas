Attribute VB_Name = "Poruke"
Declare Function sndPlaySound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ALIAS = &H10000
Public Const SND_ASYNC = &H1
Public Const SND_SYNC = &H0
Public Const SND_NOWAIT = &H2000
Public Const SND_LOOP = &H8


Option Explicit

Sub Red()
MsgBox "Polje nije na redu."
End Sub
Sub Najava()
MsgBox "Nije najavljeno."
End Sub
Sub Ruc()
MsgBox "U ovu kolonu se upisuje samo rucni rezultat."
End Sub
Sub mAplauz()
Dim lRetVal As Long, Ira As String

Ira = Dir$(App.Path)
While Ira <> ""
If UCase$(Ira) = UCase$("APPLAUSE.WAV") Then
lRetVal = sndPlaySound(App.Path & "Applause.wav", SND_ASYNC + SND_NOWAIT)
End If
Ira = Dir$
Wend

End Sub
Public Sub mSTART()
Dim lRetVal As Long
Dim Ira As String
Ira = Dir$(App.Path & "\")
While Ira <> ""
If UCase$(Ira) = UCase$("START.WAV") Then
lRetVal = sndPlaySound(App.Path & "\" & "Start.wav", SND_ASYNC + SND_NOWAIT)
End If
Ira = Dir$
Wend

End Sub

Public Function Muz(Naziv) As String
Dim lRetVal As Long
lRetVal = sndPlaySound(Naziv, SND_ASYNC + SND_NOWAIT + SND_LOOP)
End Function
Sub TheKraj()
Dim lRetVal As Long
Dim Ira As String
Ira = Dir$(App.Path & "\")
While Ira <> ""
If UCase$(Ira) = "STOP.WAV" Then
lRetVal = sndPlaySound(App.Path & "\" & "Stop.wav", SND_ASYNC + SND_NOWAIT)
End If
Ira = Dir$
Wend

End Sub

Sub mTada()
Dim lRetVal As Long
Dim Ira As String
Ira = Dir$(App.Path & "\")
While Ira <> ""
If UCase$(Ira) = "TADA.WAV" Then
lRetVal = sndPlaySound(App.Path & "\" & "Tada.wav", SND_ASYNC + SND_NOWAIT)
End If
Ira = Dir$
Wend

End Sub
Sub mGong()
Dim lRetVal As Long
Dim Ira As String
Ira = Dir$(App.Path & "\")
While Ira <> ""
If UCase$(Ira) = "MWS.WAV" Then
lRetVal = sndPlaySound(App.Path & "\" & "Mws.wav", SND_ASYNC + SND_NOWAIT)
End If
Ira = Dir$
Wend

End Sub
Sub mDobos()
Dim lRetVal As Long
Dim Ira As String
Ira = Dir$(App.Path & "\")
While Ira <> ""
If UCase$(Ira) = "DRUMROLL.WAV" Then
lRetVal = sndPlaySound(App.Path & "\" & "Drumroll.wav", SND_ASYNC + SND_NOWAIT)
End If
Ira = Dir$
Wend

End Sub
