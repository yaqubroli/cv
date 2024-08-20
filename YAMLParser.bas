Attribute VB_Name = "YAMLParser"
Option Explicit

Const yamlPath As String = "\\Mac\iCloud\Development\cv\cv.yml"

Const selfIdentifier As String = "___self___"
Const typeIdentifier As String = "___type___"
Const errIdentifier As String = "YAML Error in " & yamlPath

Const malformedTypeError As String = "Malformed YAML code at " & yamlPath & " on line "
Const malformedYAMLError As String = "Malformed type error - this is a problem with the internal dictionary"

Function RemoveEmptyStrings(arr() As String) As String()
    Dim tempArray() As String
    Dim i As Integer, j As Integer: j = 0
    ReDim tempArray(LBound(arr) To UBound(arr))
    j = 0
    For i = LBound(arr) To UBound(arr)
        If Len(arr(i)) > 0 Then
            tempArray(j) = arr(i)
            j = j + 1
        End If
    Next i
    ReDim Preserve tempArray(0 To j - 1)
    RemoveEmptyStrings = tempArray
End Function


Function RegexMatch(inputString As String, pattern As String, Optional isGlobal As Boolean = True) As Boolean
    ' checks for regex match without instantiating 80 gazillion objects
    
    ' parameters
    ' isGlobal: whether the regex check is global
    
    Dim Regex As Object: Set Regex = CreateObject("VBScript.RegExp")
    Regex.pattern = pattern
    Regex.Global = isGlobal
    RegexMatch = Regex.Test(inputString)
End Function

Function RegexSplit(inputString As String, pattern As String, Optional onlyFirst As Boolean = False, Optional splitBefore As Boolean = False) As String()
    ' splits array at any pattern that matches a regex
    
    ' parameters
    ' onlyFirst: if true, only splits the first instance of the match, creating an array of length 2
    ' splitBefore: if true, preserves the actual instance of the match
    
    Dim Regex As Object: Set Regex = CreateObject("VBScript.RegExp")
    Dim matches As Object
    Dim match As Object
    Dim splitParts() As String: ReDim splitParts(0 To 0)
    Dim pos As Integer
    Dim lastPos As Integer: lastPos = 1
    Dim i As Integer: i = 0
    
    ' set regex flags
    Regex.Global = True
    Regex.IgnoreCase = False
    Regex.pattern = pattern
    
    Set matches = Regex.Execute(inputString)
    
    ' lastPos = 1
    ' i = 0
    
    For Each match In matches
        pos = match.FirstIndex + 1
        ReDim Preserve splitParts(i)
        splitParts(i) = Mid(inputString, lastPos, pos - lastPos)
        If splitBefore Then
            lastPos = pos
        Else
            lastPos = pos + Len(match.Value)
        End If
        i = i + 1
        If onlyFirst Then Exit For
    Next match
    
    If lastPos <= Len(inputString) Then
        ReDim Preserve splitParts(i)
        splitParts(i) = Mid(inputString, lastPos)
    End If
    
    ' retvrn
    RegexSplit = RemoveEmptyStrings(splitParts)
End Function

Function RegexSubstitute(inputString As String, pattern As String, Optional substitution As String = "")
    ' does what it says on the tin
    Dim Regex As Object: Set Regex = CreateObject("VBScript.RegExp")
    Regex.pattern = pattern
    Regex.IgnoreCase = False
    Regex.Global = True
    RegexSubstitute = Regex.Replace(inputString, substitution)
End Function

' YAML Layer Parser Pseudocode
' ====
' function GetYAMLLayerAsCollection(String fromYAML) {
'    Collection mainDictionary = New Collection();
'    if (fromYAML.containsRegex(/\n[A-Za-z]/)) {
'      // is a dictionary
'      String[] temporaryArray = fromYAML.split(/\n[A-Za-z]/);
'      for each x in temporaryArray {
'        x.splitByFirstInstanceOf(':\n');
'        x[1].replaceAllInstancesOf('
'        mainDictionary.add(x[0], x[1]);
'      }
'    } else if (fromYAML.containsRegex(/\n-/)) {
'      // if array, process the array and return it as "self"
'      String[] temporaryArray = fromYAML.splitBy('\n-');
'      for each x in temporaryArray {
'        x.removeAllInstancesOf('\n- ');
'        x.replaceAllInstancesOf('\n  ', '\n');
'        mainDictionary.add("self", temporaryArray);
'      }
'    } else if (fromYAML.startsWith('"')) {
'      mainDictionary.add("self", removeQuotes(fromYAML));
'    } else {
'      MsgBox("Processing error: neither array, dictionary, nor string");
'    }
' }


Function GetYAMLLayerAsDictionary(fromYAML As String) As Dictionary
    Dim mainDictionary As Dictionary: Set mainDictionary = New Dictionary
    ' create regex objects to test for dict, array, and string
    
    'Dim regEx_dict As Object: Set regEx_dict = CreateObject("VBScript.RegExp")
    'Dim regEx_arry As Object: Set regEx_arry = CreateObject("VBScript.RegExp")
    'Dim regEx_strn As Object: Set regEx_strn = CreateObject("VBScript.RegExp")
    
    'regEx_dict.Global = True:  regEx_dict.Pattern = "\n[A-Za-z]"
    'regEx_arry.Global = True:  regEx_arry.Pattern = "\n-\s"
    'regEx_strn.Global = False: regEx_strn.Pattern = "^\s*""(.*?)""\s*$"
    
    Dim parts() As String
    
    If RegexMatch(fromYAML, "(?:\n|\^)\w+:", True) Then
        ' is a dictionary
        parts = RegexSplit(fromYAML, "\n\w+:", False, True)
        Dim part As Variant ' not sure why it can't be as string but whatever billy gates
        Call mainDictionary.Add(typeIdentifier, "Dictionary") ' identify as dict
        For Each part In parts
            Dim keyValue() As String: keyValue = RegexSplit(CStr(part), ":\s", True)
            ' trim trailing \n from category
            If UBound(keyValue) > 0 Then
                keyValue(0) = RegexSubstitute(keyValue(0), "^\n+")
                ' trim 2 spaces off of each line if they're there
                keyValue(1) = RegexSubstitute(keyValue(1), "^\s{2}")
                keyValue(1) = RegexSubstitute(keyValue(1), "\n\s{2}", vbLf)
                Call mainDictionary.Add(keyValue(0), keyValue(1))
            End If
        Next part
    ElseIf RegexMatch(fromYAML, "^-\s", True) Then
        ' is an array
        Call mainDictionary.Add(typeIdentifier, "Array")
        parts = RegexSplit(fromYAML, "(^|\n)-\s", False)
        Dim i As Integer
        For i = LBound(parts) To UBound(parts)
            parts(i) = RegexSubstitute(parts(i), "\n\s{2}", vbLf)
        Next i
        Call mainDictionary.Add(selfIdentifier, parts)
    ElseIf RegexMatch(fromYAML, "^\s*""(.*?)""\s*$", True) Then
        ' is a string
        Call mainDictionary.Add(typeIdentifier, "String")
        Call mainDictionary.Add(selfIdentifier, RegexSubstitute(fromYAML, """", ""))
    Else
        Call mainDictionary.Add(selfIdentifier, "")
        Debug.Print _
        "Neither array, dictionary, nor string:" & _
        vbCrLf & vbCrLf & fromYAML & vbCrLf & vbCrLf & _
        "Make sure all strings are enclosed in double quotes." ', _
        'vbOKOnly, "YAML Error")
    End If
    
    Set GetYAMLLayerAsDictionary = mainDictionary
End Function

' YAML Traverser Pseudocode
' ===
'
' function TraverseYAML(String fromYAML) {
'   Dictionary mainDictionary = GetYAMLLayerAsDictionary(fromYAML);
'   if mainDictionary.___type___ = "Dictionary" {
'     for each entry in mainDictionary {
'       TraverseYAML(entry)
'     }
'     return mainDictionary;
'   } else if mainDictionary.___type___ = "Array" {
'     for each entry in mainDictionary.___self___ {
'       TraverseYAML(entry)
'     }
'     return mainDictionary;
'   } else if mainDictionary.___type___ = "String" {
'     return mainDictionary;
'   } else {
'     MsgBox("Internal YAML Error")
'   }
' }
Function GetYAMLAsDictionary(fromYAML As String) As Dictionary
    Dim mainDictionary As Object: Set mainDictionary = GetYAMLLayerAsDictionary(fromYAML)
    Dim entry As Variant
    If mainDictionary(typeIdentifier) = "Dictionary" Then
        For Each entry In mainDictionary
            Debug.Print "=== PROCESSING DICTIONARY ENTRY ==="
            Debug.Print entry & " => " & mainDictionary(entry)
            If entry <> typeIdentifier And entry <> selfIdentifier Then
                Set mainDictionary(entry) = GetYAMLAsDictionary(mainDictionary(entry))
            End If
        Next entry
    ElseIf mainDictionary(typeIdentifier) = "Array" Then
        Dim i As Integer
        Dim subArray() As Object
        For i = LBound(mainDictionary(selfIdentifier)) To UBound(mainDictionary(selfIdentifier))
            Debug.Print "=== PROCESSING ARRAY ENTRY ==="
            Debug.Print mainDictionary(selfIdentifier)(i)
            'Set subDictionary = GetYAMLAsDictionary(mainDictionary(selfIdentifier)(i))
            'Set mainDictionary(selfIdentifier)(i) = subDictionary
            ReDim Preserve subArray(i)
            Set subArray(i) = GetYAMLAsDictionary(CStr(mainDictionary(selfIdentifier)(i)))
        Next i
        
        mainDictionary(selfIdentifier) = subArray
    ElseIf mainDictionary(typeIdentifier) <> "String" Then
        Debug.Print malformedTypeError ', vbOKOnly, errIdentifier)
    End If
    Set GetYAMLAsDictionary = mainDictionary
End Function

' YAML Cleaner Pseudocode
' =====
' function YAMLCleaner(Dictionary mainDictionary) {
'   for each entry in mainDictionary {
'     if entry(typeIdentifier) == "Dictionary" {
'       for each secondOrderEntry in entry {
'         YAMLCleaner(secondOrderEntry)
'       }
'     } else if entry(typeIdentifier) == "Array" {
'       for each secondOrderEntry in entry(selfIdentifier) {
'         YAMLCleaner(secondOrderEntry)
'       }
'     }
'     if entry(typeIdentifier) != "Dictionary" {
'       mainDictionary(entry) = mainDictionary(entry)(selfIdentifier)
'     }
'   }
'   return mainDictionary;
' }

Function YAMLCleaner(mainDictionary As Dictionary) As Dictionary
    Dim entry As Variant
    If mainDictionary(typeIdentifier) = "Array" Then ' go through array and yamlclean it
        Dim i As Integer
        Debug.Print JsonConverter.ConvertToJson(mainDictionary)
        For i = LBound(mainDictionary(selfIdentifier)) To UBound(mainDictionary(selfIdentifier))
            
            'If IsObject(mainDictionary(selfIdentifier)(i)) Then
                'Set mainDictionary(selfIdentifier)(i) = YAMLCleaner(mainDictionary(selfIdentifier)(i))
            'Else
            '    Debug.Print "encountered non-object"
            'End If
        Next i
    End If
    If mainDictionary(typeIdentifier) = "Dictionary" Then 'iterate through dict and yamlclena it
        For Each entry In mainDictionary
            If entry <> typeIdentifier Then
                Set mainDictionary(entry) = YAMLCleaner(mainDictionary(entry))
            End If
        Next entry
    End If
    
    For Each entry In mainDictionary
        If mainDictionary(typeIdentifier) = "Dictionary" And mainDictionary(entry)(typeIdentifier) <> "Dictionary" And entry <> typeIdentifier And entry <> selfIdentifier Then
            Debug.Print "processing " & entry & " which is " & mainDictionary(entry)(typeIdentifier)
            If IsObject(mainDictionary(entry)(selfIdentifier)) Then
                Set mainDictionary(entry) = mainDictionary(entry)(selfIdentifier)
            Else
                mainDictionary(entry) = mainDictionary(entry)(selfIdentifier)
            End If
        End If
    Next entry

    ' destroy type identifier?
    Set YAMLCleaner = mainDictionary
End Function

Function GetFileAsString(filePath As String) As String
    ' Dim fileContent As String
    Dim line As String
    Dim fileNumber As Integer
    
    'filePath = "\\Mac\iCloud\Development\cv\cv.yml"
    
    fileNumber = FreeFile()
    
    Open filePath For Input As fileNumber
    
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        GetFileAsString = GetFileAsString & line & vbCrLf
    Loop
End Function

Sub TryFunction()
    Dim fileString As String: fileString = GetFileAsString("\\Mac\iCloud\Development\cv\cv.yml")
    Dim yamlLayer As Object
    Set yamlLayer = GetYAMLLayerAsDictionary(fileString)
    Dim yamlWholeDict As Object: Set yamlWholeDict = GetYAMLAsDictionary(fileString)
    'Debug.Print JsonConverter.ConvertToJson(yamlWholeDict, 2, 2)
    Dim yamlCleanDict As Object: Set yamlCleanDict = YAMLCleaner(yamlWholeDict)
    Debug.Print JsonConverter.ConvertToJson(yamlCleanDict, 2, 2)
End Sub
