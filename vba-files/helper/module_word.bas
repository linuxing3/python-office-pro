' 命名空间
Attribute VB_Name="module_word"

''namespace=vba-files\helper
Sub test_mkdirp()
    '// add declarations
    On Error GoTo catchError
    Dim strRootPath As String
    Dim bolSuccess As Boolean

    strRootPath = "D:\level1\level2\level3"
    bolSuccess = mkdirpf(strRootPath)

    If bolSuccess Then
      Debug.Print "done!"
    ElseIf bolSuccess Then
      Debug.Print "fail!"
    Else
      Debug.Print "unknow error!"
    End If
exitSub:
    Exit Sub
catchError:
    '// add error handling
    GoTo exitSub
End Sub

Function mkdirpf(path As String) As Boolean
    '// add declarations
    On Error GoTo catchError
    mkdirp(path)
    If Dir(path) <> "" Then
      mkdirpf = True
    Else
      mkdirpf = False
    End If
        
exitFunction:
    Exit Function
catchError:
    '// add error handling
    GoTo exitFunction
End Function

Sub mkdirp(strRootPath as String)
    '// add declarations
    On Error GoTo catchError

    Dim arrResult
    Dim strResult As String
    ' // split strRootPath
    arrResult = Split(strRootPath, "\")
    
    For Each i In arrResult 
      strResult = strResult & "\" & i
      If Dir(strResult, vbDirectory) <> "" Then
        MkDir(strResult)
      end If
    Next i

exitSub:
    Exit Sub
catchError:
    '// add error handling
    GoTo exitSub
End Sub

Sub loopFolder()

  Dim strRootPath as String
  Dim strNewFolderPath as String

  Dim varFileShortName As Variant

  ' folder path ends with \ seperator
  strRootPath = CurDir
  strNewFolderPath = "D:\workspace\"

  ' Loop files in a folder
  varFileShortName = Dir(strRootPath)

  ' Start from first file
  Do While varFileShortName <> ""
    Debug.Print(varFileShortName)
    ' new path
    FileCopy strRootPath & varFileShortName, strNewFolderPath & varFileShortName, 
    ' Point to next file
    varFileShortName = Dir
  Loop


End Sub