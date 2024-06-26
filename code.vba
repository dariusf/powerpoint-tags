Option Explicit

Sub MyCustomMacro(Button As IRibbonControl)
    MsgBox "Hello from VBA Add-in!"
End Sub

Sub MyCustomMacro1()
    MsgBox "Hello from VBA Add-in!"
End Sub

Sub TagAsDogs()
    Dim oSl As Slide
    For Each oSl In ActiveWindow.Selection.SlideRange
        oSl.Tags.Add "DOG", "Y"
    Next
End Sub

Sub TagAsPonies()
    Dim oSl As Slide
    For Each oSl In ActiveWindow.Selection.SlideRange
        oSl.Tags.Add "PONY", "Y"
    Next
End Sub

Sub DogShow()
' Hide any slide w/o a DOG tag
    Dim oSl As Slide
    For Each oSl In ActivePresentation.Slides
        If oSl.Tags("DOG") <> "Y" Then
            oSl.SlideShowTransition.Hidden = True
        End If
    Next
End Sub

Sub PonyShow()
' Hide any slide w/o a PONY tag
    Dim oSl As Slide
    For Each oSl In ActivePresentation.Slides
        If oSl.Tags("PONY") <> "Y" Then
            oSl.SlideShowTransition.Hidden = True
        End If
    Next
End Sub

Sub DogAndPonyShow()
' Unhide all of the slides
    Dim oSl As Slide
    For Each oSl In ActivePresentation.Slides
        oSl.SlideShowTransition.Hidden = False
    Next
End Sub