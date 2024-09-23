Attribute VB_Name = "NewMacros"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    Selection.Style = ActiveDocument.Styles("Title")
    Selection.TypeText Text:="This is a title"
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.Style = ActiveDocument.Styles("Heading 1")
    Selection.TypeText Text:="Heading 1"
    Selection.TypeParagraph
    Selection.Style = ActiveDocument.Styles("No Spacing")
    Selection.TypeText Text:="Whatababab"
    Selection.TypeParagraph
    With ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = ChrW(61623)
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = InchesToPoints(0.25)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.5)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "Symbol"
        End With
        .LinkedStyle = ""
    End With
    ListGalleries(wdBulletGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:= _
        False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListBehavior
    Selection.TypeText Text:="This is a bulletted list created manually"
    Selection.EscapeKey
End Sub
