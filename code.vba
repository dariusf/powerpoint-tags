Option Explicit

Sub MyCustomMacro(Button As IRibbonControl)
    MsgBox "Hello from VBA Add-in!"
End Sub

Sub MyCustomMacro1()
    MsgBox "Hello from VBA Add-in!"
End Sub

Sub TagAsDogs()
    Dim oSl As slide
    For Each oSl In ActiveWindow.Selection.SlideRange
        oSl.Tags.Add "DOG", "Y"
    Next
End Sub

Sub TagAsPonies()
    Dim oSl As slide
    For Each oSl In ActiveWindow.Selection.SlideRange
        oSl.Tags.Add "PONY", "Y"
    Next
End Sub

Sub DogShow()
' Hide any slide w/o a DOG tag
    Dim oSl As slide
    For Each oSl In ActivePresentation.Slides
        If oSl.Tags("DOG") <> "Y" Then
            oSl.SlideShowTransition.Hidden = True
        End If
    Next
End Sub

Sub PonyShow()
' Hide any slide w/o a PONY tag
    Dim oSl As slide
    For Each oSl In ActivePresentation.Slides
        If oSl.Tags("PONY") <> "Y" Then
            oSl.SlideShowTransition.Hidden = True
        End If
    Next
End Sub

Sub DogAndPonyShow()
' Unhide all of the slides
    Dim oSl As slide
    For Each oSl In ActivePresentation.Slides
        oSl.SlideShowTransition.Hidden = False
    Next
End Sub

Function DeduplicateCommaDelimitedStrings(ByVal inputList As Variant) As String
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    Dim items() As String
    Dim item As Variant
    
    ' Loop through each string in the input list
    For i = LBound(inputList) To UBound(inputList)
        ' Split the string into individual items
        items = Split(inputList(i), ",")
        
        ' Add each item to the dictionary (which automatically handles duplicates)
        For Each item In items
            If Not dict.exists(Trim(item)) Then
                dict.Add Trim(item), Nothing
            End If
        Next item
    Next i
    
    ' Join the keys of the dictionary into a comma-delimited string
    DeduplicateCommaDelimitedStrings = Join(dict.keys, ",")
End Function

Function BuildInputArrayFromSelection() As Variant
    Dim slideRange As slideRange
    Dim slide As slide
    Dim shape As shape
    Dim inputArray() As String
    Dim itemList As String
    Dim i As Integer
    
    ' Initialize the slideRange with the selected slides
    Set slideRange = ActiveWindow.Selection.SlideRange
    
    ' Initialize a temporary collection to hold the comma-delimited strings
    Dim tempCollection As Collection
    Set tempCollection = New Collection
    
    ' Loop through each slide in the slideRange
    For Each slide In slideRange
        tempCollection.Add shape.TextFrame.TextRange.Text
    Next slide
    
    ' Resize the inputArray to match the number of items in the collection
    ReDim inputArray(1 To tempCollection.Count)
    
    ' Copy the items from the collection to the array
    For i = 1 To tempCollection.Count
        inputArray(i) = tempCollection(i)
    Next i
    
    ' Return the array
    BuildInputArrayFromSelection = inputArray
End Function


Sub EditTags()
    Dim oSl As slide
    Dim Tags As String
    If ActiveWindow.Selection.SlideRange.Count > 1 Then
        ' Tags = DeduplicateCommaDelimitedStrings(BuildInputArrayFromSelection())
        ' Tags = InputBox("Tags to set", "Batch editing", Tags)
        Tags = InputBox("Tags to set", "Batch editing")
        If Tags = "" Then
            Exit Sub
        End If
        For Each oSl In ActiveWindow.Selection.SlideRange
            oSl.Tags.Delete "TAGS"
            oSl.Tags.Add "TAGS", Tags
        Next
    Else
        For Each oSl In ActiveWindow.Selection.SlideRange
            ' this only happens once
            Tags = InputBox("Tags to edit", "Editing tags of slide", oSl.Tags("TAGS"))
            If Tags = "" Then
                Exit Sub
            End If
            oSl.Tags.Delete "TAGS"
            oSl.Tags.Add "TAGS", Tags
        Next
    End If
End Sub

Sub ReallySearch(useTags As String)
    Dim searchWord As String
    Dim sourcePresentation As Presentation
    Dim destinationPresentation As Presentation
    Dim slide As slide
    Dim slideIndex As Integer
    Dim slideFound As Boolean
    
    ' Define the word to search for
    If useTags Then
        searchWord = InputBox("Enter the tag to search for:", "Search Tag")
    Else
        searchWord = InputBox("Enter the word to search for:", "Search Word")
    End If
    If searchWord = "" Then
        Exit Sub
    End If
    
    ' Reference the current presentation
    Set sourcePresentation = ActivePresentation
    
    ' Create a new presentation
    Set destinationPresentation = Presentations.Add
    
    ' Initialize slide index for the destination presentation
    slideIndex = 1
    
    ' Loop through each slide in the source presentation
    For Each slide In sourcePresentation.Slides
        slideFound = False
        If useTags Then
            ' TODO split the tags, and check if any match, then update the message to tags
            If InStr(1, slide.Tags("TAGS"), searchWord, vbTextCompare) > 0 Then
                slideFound = True
            End If
        Else
            ' Check each shape on the slide
            Dim shape As shape
            For Each shape In slide.Shapes
                ' Check if the shape contains text
                If shape.HasTextFrame Then
                    If shape.TextFrame.HasText Then
                        ' Check if the text contains the search word
                        If InStr(1, shape.TextFrame.TextRange.Text, searchWord, vbTextCompare) > 0 Then
                            slideFound = True
                            Exit For
                        End If
                    End If
                End If
            Next shape
        End If
        
        ' If the slide contains the search word, copy it to the new presentation
        If slideFound Then
            slide.Copy
            destinationPresentation.Slides.Paste (slideIndex)
            slideIndex = slideIndex + 1
        End If
    Next slide
    
    ' Save the new presentation
    Dim newFileName As String
    newFileName = "CopiedSlides.pptx"

    If newFileName <> "False" Then
        destinationPresentation.SaveAs newFileName
        ' MsgBox "Slides copied and saved successfully!", vbInformation
    Else
        ' MsgBox "Operation cancelled.", vbExclamation
    End If
End Sub

Sub SearchAndCopySlides()
    ReallySearch (False)
End Sub

Sub SearchTagsAndCopySlides()
    ReallySearch (True)
End Sub