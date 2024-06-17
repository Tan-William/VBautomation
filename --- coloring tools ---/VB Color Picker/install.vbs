'--install script for vb color picker.

Dim FSO, r, sPath, Pt1, DeskFol, ProgFol, oFol, CPFol, CPPath
Set FSO = CreateObject("Scripting.FileSystemObject")

 r = MsgBox("The VB Color Picker does not need to be installed but this script can put it in Program Files and make a shortcut on the Desktop. Do you want to do that?", 36, "Pseudo-Install")
   If (r = 7) Then WScript.Quit
   
 sPath = WScript.ScriptFullName
   Pt1 = InStrRev(sPath, "\")
   sPath = left(sPath, Pt1) 
   CPPath = sPath & "ColFind.exe"
 
  If (FSO.FileExists(CPPath) = False) Then
    MsgBox "ColFind.exe not found. It must be placed in same folder as this script.", 64, "Pseudo-Install"
    WScript.Quit
  End If  
  
DeskFol = GetFolPath("Desktop")
ProgFol = GetProgFol()
CPFol = ProgFol & "\VB Color Picker"
 If (FSO.FolderExists(CPFol) = False) Then
     Set oFol = FSO.CreateFolder(CPFol)
     Set oFol = Nothing
 End If
 
   FSO.CopyFile CPPath, CPFol & "\Colfind.exe", True
     CPPath = sPath  & "CP Info.txt"
   If (FSO.FileExists(CPPath) = True) Then 
       FSO.CopyFile CPPath, CPFol & "\CP Info.txt", True
   End If    
  MakeLink CPFol & "\ColFind.exe", DeskFol, "VB Color Picker"
   
 Set FSO = Nothing
 WScript.Quit
     


'----------------------------------------
Sub MakeLink(sTarget, sPathLnk, sName)
Dim Sh2, sLnk, ShCut, Pt1
     If (Right(sPathLnk, 1) <> "\") Then sPathLnk = sPathLnk & "\"
  sPathLnk = sPathLnk & sName & ".lnk"
  Set Sh2 = CreateObject("WScript.Shell")
   Set ShCut = Sh2.CreateShortcut(sPathLnk)
      ShCut.WindowStyle = 1
      ShCut.TargetPath = sTarget
      ShCut.IconLocation = sTarget & ", 0" 
         Pt1 = InStrRev(sTarget, "\")
      ShCut.WorkingDirectory = Left(sTarget, Pt1)
      ShCut.Save 
   Set ShCut = Nothing
  Set Sh2 = Nothing   
End Sub

'----------------------------------------
 Function GetFolPath(sFol)   
        Dim s3, Sh2
          On Error Resume Next
      Set Sh2 = CreateObject("WScript.Shell")
        s3 = Sh2.SpecialFolders(sFol)
        GetFolPath = s3   
     Set Sh2 = Nothing  
  End Function
 '----------------------------------- 
  Function GetProgFol()      
      Dim s4, Sh2
         On Error Resume Next
       Set Sh2 = CreateObject("WScript.Shell")
         s4 = Sh2.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir")
         GetProgFol = s4
       Set Sh2 = Nothing
   End Function