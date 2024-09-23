Attribute VB_Name = "CV"
Option Explicit

Function DrawTitle(theTitle As String)
    Selection.Style = ActiveDocument.Styles("Title")
    Selection.TypeText Text:=theTitle
    Selection.TypeParagraph
End Function

Function DrawSubtitle(theSubtitle As String)
    Selection.Style = ActiveDocument.Styles("Heading 1")
    Selection.TypeText Text:=theSubtitle
    Selection.TypeParagraph
End Function

Function DrawNormalText(theText As String)
    Selection.Style = ActiveDocument.Styles("Normal")
    Selection.TypeText Text:=theText
End Function

Function DrawSmallCaps(theSmallCaps As String)
    With Selection.Font
        .Bold = True
        .SmallCaps = True
    End With
    Selection.TypeText Text:=theSmallCaps
End Function

Function DrawTabEntry(theTabEntry As String)
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = InchesToPoints(0.5)
    Selection.ParagraphFormat.TabStops.Add Position:=InchesToPoints(6.5), _
        Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
    Selection.TypeText Text:=vbTab
    Selection.Font.Bold = wdToggle
    With Selection.Font
        .Bold = False
        .Italic = False
        .SmallCaps = False
        .Color = RGB(GREY_VALUE, GREY_VALUE, GREY_VALUE)
    End With
    Selection.TypeText Text:=theTabEntry
    Selection.TypeParagraph
    ' reset to normal
    Selection.Style = ActiveDocument.Styles("Normal")
End Function

Function DrawBulletedList(theBulletEntries() As String)
    With ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = ChrW(61623) ' Bullet character
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = InchesToPoints(0.25)
        .TextPosition = InchesToPoints(0.5)
        .Font.Name = "Symbol"
    End With
    
    ' Apply bullet list formatting to the selection
    Selection.Range.ListFormat.ApplyListTemplateWithLevel _
        ListTemplate:=ListGalleries(wdBulletGallery).ListTemplates(1), _
        ContinuePreviousList:=False, _
        ApplyTo:=wdListApplyToWholeList, _
        DefaultListBehavior:=wdWord10ListBehavior

    ' Insert each bullet entry
    Dim bulletEntry As Variant
    For Each bulletEntry In theBulletEntries
        Selection.TypeText Text:=bulletEntry
        Selection.TypeParagraph
    Next bulletEntry
    
    ' Remove bullet list formatting after the list is complete
    Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
    Selection.TypeParagraph
End Function

Function DrawCompoundBulletedList(theTitles() As String, theDescriptions() As String)
    With ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = ChrW(61623) ' Bullet character
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = InchesToPoints(0.25)
        .TextPosition = InchesToPoints(0.5)
        .Font.Name = "Symbol"
    End With
    
    ' Apply bullet list formatting to the selection
    Selection.Range.ListFormat.ApplyListTemplateWithLevel _
        ListTemplate:=ListGalleries(wdBulletGallery).ListTemplates(1), _
        ContinuePreviousList:=False, _
        ApplyTo:=wdListApplyToWholeList, _
        DefaultListBehavior:=wdWord10ListBehavior
    
    If (UBound(theTitles) - LBound(theTitles)) <> (UBound(theTitles) - LBound(theTitles)) Then
        MsgBox "Internal error in projects; this will result in an array bound error."
    End If
    
    ' Insert each bullet entry
    Dim i As Integer
    For i = LBound(theTitles) To UBound(theTitles)
        With Selection.Font
            .Bold = True
            .SmallCaps = True
        End With
        Selection.TypeText Text:=theTitles(i)
        With Selection.Font
            .Bold = False
            .SmallCaps = False
        End With
        Selection.TypeText Text:=" " & ChrW(8212) & " " & theDescriptions(i)
        Selection.TypeParagraph
    Next i
    
    ' Remove bullet list formatting after the list is complete
    Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
    Selection.TypeParagraph
End Function

Sub DrawCV()
Attribute DrawCV.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
    Dim cvYAML As Object: Set cvYAML = New YAML
    
    ' Q: Why is the business logic so terrible?
    '    Why is every area of the dictionary referenced from the root with loads of ID_SELFs?
    '    Why not assign it to variables to make it shorter?
    
    ' A: Because VBA does not permit setting array length at runtime, and also seems
    '    to crash with some sort of memory error when assigning subdictionaries to their
    '    own variable.
    '
    '    If I could go back to the drawing board, I probably would have drafted a better solution
    '    (for example not using dictionaries and making the YAML class its own thing)
    '    but at this stage of the project the perfect is the enemy of the good.
    
    cvYAML.path = YAML_PATH
    
    Dim cvProps As Scripting.Dictionary: Set cvProps = cvYAML.props
    
    ' read name
    
    DrawTitle CStr(cvProps("name")(ID_SELF))
    
    ' iterate through subheadings of "cv"
    
    Dim i As Integer: i = 0
    Dim j As Integer: j = 0
    
    For i = LBound(cvProps("cv")(ID_SELF)) To UBound(cvProps("cv")(ID_SELF))
        DrawSubtitle CStr(cvProps("cv")(ID_SELF)(i)("title")(ID_SELF))
        ' iterate through entries
        For j = _
            LBound(cvProps("cv")(ID_SELF)(i)("entries")(ID_SELF)) To _
            UBound(cvProps("cv")(ID_SELF)(i)("entries")(ID_SELF))
            
            DrawSmallCaps CStr(cvProps("cv")(ID_SELF)(i)("entries")(ID_SELF)(j)("title")(ID_SELF))
            DrawTabEntry CStr(cvProps("cv")(ID_SELF)(i)("entries")(ID_SELF)(j)("date")(ID_SELF))
            
            If cvProps("cv")(ID_SELF)(i)("entries")(ID_SELF)(j).Exists("role") Then
                DrawNormalText CStr(cvProps("cv")(ID_SELF)(i)("entries")(ID_SELF)(j)("role")(ID_SELF))
                If cvProps("cv")(ID_SELF)(i)("entries")(ID_SELF)(j).Exists("location") Then
                DrawTabEntry CStr(cvProps("cv")(ID_SELF)(i)("entries")(ID_SELF)(j)("location")(ID_SELF))
                End If
            End If
            
            ' iterate through the bullets of the respective entry
            
            Dim k As Integer
            Dim bulletedList() As String
            
            ReDim bulletedList(LBound(cvProps("cv")(ID_SELF)(i)("entries")(ID_SELF)(j)("bullets")(ID_SELF)) To _
                               UBound(cvProps("cv")(ID_SELF)(i)("entries")(ID_SELF)(j)("bullets")(ID_SELF)))
            For k = LBound(cvProps("cv")(ID_SELF)(i)("entries")(ID_SELF)(j)("bullets")(ID_SELF)) To _
                    UBound(cvProps("cv")(ID_SELF)(i)("entries")(ID_SELF)(j)("bullets")(ID_SELF))
                bulletedList(k) = CStr(cvProps("cv")(ID_SELF)(i)("entries")(ID_SELF)(j)("bullets")(ID_SELF)(k)(ID_SELF))
            Next k
            DrawBulletedList bulletedList
       Next j
    Next i
    
    ' iterate through projects
    
    DrawSubtitle "Projects"
    
    Dim l As Integer
    Dim titles() As String
    Dim descriptions() As String
    ReDim titles(LBound(cvProps("projects")(ID_SELF)) To UBound(cvProps("projects")(ID_SELF)))
    ReDim descriptions(LBound(cvProps("projects")(ID_SELF)) To UBound(cvProps("projects")(ID_SELF)))
    
    For l = LBound(cvProps("projects")(ID_SELF)) To UBound(cvProps("projects")(ID_SELF))
        ' Debug.Print JsonConverter.ConvertToJson(cvProps("projects")(ID_SELF)(l), 2)
        titles(l) = CStr(cvProps("projects")(ID_SELF)(l)("title")(ID_SELF))
        descriptions(l) = CStr(cvProps("projects")(ID_SELF)(l)("description")(ID_SELF))
    Next l
    
    DrawCompoundBulletedList titles, descriptions
    
    ' iterate through skills
    
End Sub

Sub JsonTroubleshoot()
    Dim cvYAML As Object: Set cvYAML = New YAML
    cvYAML.path = YAML_PATH
    Dim cvProps As Scripting.Dictionary: Set cvProps = cvYAML.props
    DrawNormalText JsonConverter.ConvertToJson(cvProps, 2)
End Sub
