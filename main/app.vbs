Set FSO = CreateObject("Scripting.FileSystemObject")

Set workDirectory = GetFolder(DIRECTORY)
CleanDirectory(workDirectory)

Function GetFolder (sFolder)
 On Error Resume Next
 Set GetFolder = FSO.GetFolder(sFolder)
 errorMsg(err)
End Function

Sub Delete(sFile)
 On Error Resume Next
 FSO.DeleteFile sFile, True
 errorMsg(err)
End Sub

Sub CleanDirectory(workDirectory)
    for each item in workDirectory.Files
        FileDate = item.DateLastModified
        Age = DateDiff("d",Now,FileDate)
        If Abs(Age)>DAYNUMBER Then
            Delete(item)
        End If
    next
End Sub

Sub errorMsg(err)
if err.number > 0 then
    WScript.Echo "Operation failed. Please, check initial settings and try again."
    WScript.Quit
 end if
End Sub