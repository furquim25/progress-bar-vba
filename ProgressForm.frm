VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "Processando..."
   ClientHeight    =   2010
   ClientLeft      =   300
   ClientTop       =   1170
   ClientWidth     =   5640
   OleObjectBlob   =   "ProgressForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public isProgressFormOpen As Boolean
Public progressCount As Integer

Private Sub UserForm_Open(Cancel As Integer)
    isProgressFormOpen = True
    progressCount = 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        resultMsgBox = MsgBox("Cancelar processamento?", vbYesNo, "Confirme")
        
        If resultMsgBox = vbYes Then
            Hide
            isProgressFormOpen = False
        Else
            Cancel = True
        End If
        
    End If
End Sub

Public Sub UpdateProgressBar(currentValue As Variant, maxValue As Variant)
    DoEvents
    Dim maxWidth As Integer
    maxWidth = 246
    With ProgressForm
        .ProgressBar.Width = maxWidth * (currentValue / maxValue)
        .LabelPercent.Caption = Round((currentValue / maxValue) * 100, 0) & "%"
    End With
End Sub

Public Sub OpenProgressForm()
    isProgressFormOpen = True
    progressCount = 0
    With ProgressForm
        .LabelFileName.Caption = "Carregando..."
        .LabelPercent.Caption = 0 & "%"
        .ProgressBar.Width = 0
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show vbModeless 'Macro can still run after showing form
    End With
End Sub
