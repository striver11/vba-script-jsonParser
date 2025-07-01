Sub ValidateJsonPathsInline()

    Dim jsonStr As String
    Dim jsonFilePath As String
    Dim fileNum As Integer
    Dim cell As Range
    Dim sc As Object
    Dim jsCode As String
    Dim result As String

    ' === Set your JSON file path here ===
    jsonFilePath = "C:\Users\YourName\Documents\sample.json"

    ' === Read the JSON file ===
    On Error GoTo FileReadError
    fileNum = FreeFile
    Open jsonFilePath For Input As #fileNum
    jsonStr = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    On Error GoTo 0

    ' === Set up JavaScript ScriptControl ===
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"

    ' === JavaScript logic to walk JSON path ===
    jsCode = "function check(json, path) {" & vbCrLf & _
             "  try {" & vbCrLf & _
             "    var obj = JSON.parse(json);" & vbCrLf & _
             "    var parts = path.split('.');" & vbCrLf & _
             "    for (var i = 0; i < parts.length; i++) {" & vbCrLf & _
             "      if (parts[i].includes('[')) {" & vbCrLf & _
             "        var key = parts[i].split('[')[0];" & vbCrLf & _
             "        var index = parseInt(parts[i].match(/\[(\d+)\]/)[1]);" & vbCrLf & _
             "        obj = obj[key][index];" & vbCrLf & _
             "      } else {" & vbCrLf & _
             "        obj = obj[parts[i]];" & vbCrLf & _
             "      }" & vbCrLf & _
             "      if (obj === undefined) return 'NOT_FOUND';" & vbCrLf & _
             "    }" & vbCrLf & _
             "    return 'FOUND';" & vbCrLf & _
             "  } catch (e) { return 'ERROR'; }" & vbCrLf & _
             "}"

    sc.AddCode jsCode

    ' === Loop through all JSON paths in column A ===
    For Each cell In Range("A2:A" & Cells(Rows.Count, 1).End(xlUp).Row)
        Dim path As String
        path = Trim(cell.Value)
        If path <> "" Then
            result = sc.Run("check", jsonStr, path)

            With cell
                Select Case result
                    Case "FOUND"
                        .Interior.Color = RGB(144, 238, 144) ' Light green
                    Case "NOT_FOUND"
                        .Interior.Color = RGB(255, 99, 71)   ' Tomato red
                    Case Else
                        .Interior.Color = RGB(255, 165, 0)   ' Orange
                End Select
            End With
        End If
    Next cell

    MsgBox "Validation complete!", vbInformation
    Exit Sub

FileReadError:
    MsgBox "Could not read the JSON file. Please check the path.", vbCritical
End Sub
