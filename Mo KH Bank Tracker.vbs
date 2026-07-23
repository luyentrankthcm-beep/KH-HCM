' Chi Nhan (2026-07-22): "muon co link co dinh de mo nhanh" ban offline --
' file nay lam 1 VIEC DUY NHAT: tu khoi dong server (neu chua chay) o CHE DO
' AN (khong hien cua so den nhu start.bat), roi tu mo trinh duyet vao dung
' dia chi offline. Bam dup file nay la xong, khong can nho link
' localhost:3000 hay mo cua so lenh nua.
Set objShell = CreateObject("WScript.Shell")
strFolder = objFSO_GetParentFolder()

Function objFSO_GetParentFolder()
  Set fso = CreateObject("Scripting.FileSystemObject")
  objFSO_GetParentFolder = fso.GetParentFolderName(WScript.ScriptFullName)
End Function

objShell.CurrentDirectory = strFolder

' Thu khoi dong server -- neu da chay san (cong 3000 dang duoc dung) thi lenh
' nay se tu that bai trong im lang (khong hien loi), khong sao ca.
objShell.Run "cmd /c cd /d """ & strFolder & """ && npm start", 0, False

' Doi vai giay de server kip khoi dong (neu chua chay truoc do)
WScript.Sleep 2500

' Mo trinh duyet mac dinh vao web offline
objShell.Run "http://localhost:3000"
