Attribute VB_Name = "ExampleTask"
Sub ExampleTask()
    'Show progress bar
    Call ProgressForm.OpenProgressForm
    'Updates Progress Bar based on task example
    For i = 1 To 100000000
        If i Mod 1000000 = 0 Then
            'Update progress bar
            Call ProgressForm.UpdateProgressBar(i, 100000000)
            'Checks if progress form is closed
            If ProgressForm.isProgressFormOpen = False Then
                Unload ProgressForm
                Exit Sub
            End If
            Debug.Print (i)
        End If
    Next
End Sub

