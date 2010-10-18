'For use in Group Policy in an Active Directory Enviroment

'The MIT License
'=====================================================================================================================
'Copyright (c) 2010 Christopher J. Allison

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in
'all copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'THE SOFTWARE.

'=====================================================================================================================
'===GLOBAL
On Error Resume Next

'==[ Init Varibles ]=='
Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")
Dim winDir, distroDir, i, ii
Dim tasks(2), apps(3)
Dim objNetwork 
Dim strDriveLetter, strRemotePath, strUser, strPassword, strProfile

'==[ Directories ]=='
winDir = "C:\Windows\"
winTaskDir = "C:\Windows\Tasks\"
distroDir = "\\server\share\distro\"
taskDir = "\\server\share\Tasks\"

'==[ Applications ]=='
apps(0) = "psshutdown.exe"
apps(1) = "CCleaner.exe"
apps(2) = "Defraggler.exe"
apps(3) = "subinacl.exe"

'==[ Tasks ]=='
'=These tasks where created on an admin machine using the
'=normal task creatation GUI.
tasks(0) = "psshutdown.job"
tasks(1) = "CCleaner.job"
tasks(2) = "Defraggler.job"

'=====================================================================================================================
'===This has to be done, basically, only to auth to the file server
'===Since this script is run from the machines system account and not from a domain account
'===(because this is used as a startup script)
strDriveLetter = "Z:" 
strRemotePath = "\\server\share" 
strUser = ""
strPassword = ""
strProfile = "false"
Set objNetwork = CreateObject("WScript.Network") 

objNetwork.MapNetworkDrive strDriveLetter, strRemotePath, _
  strProfile, strUser, strPassword
'=====================================================================================================================
'===Copy files from File Server to local Windows directory.

For i = 0 to 3
	'wscript.echo apps(i)
	fso.CopyFile distroDir & apps(i), windir, TRUE
	ws.Sleep 250
Next

For ii = 0 to 2
	fso.CopyFile taskDir & tasks(ii), winTaskDir, TRUE
	ws.Sleep 250
Next

WScript.Sleep 1000
'=====================================================================================================================

ws.Exec "C:\Windows\subinacl.exe /noverbose /file C:\WINDOWS\Tasks\*.* /grant=DOMAIN\adminaccount=F /setowner=DOMAIN\adminaccount"

ws.Exec "schtasks /change /RU DOMAIN\adminaccount /RP pass /TN psshutdown"
ws.Exec "schtasks /change /RU DOMAIN\adminaccount /RP pass /TN CCleaner"
ws.Exec "schtasks /change /RU DOMAIN\adminaccount /RP pass /TN Defraggler"

if err.number<>0 then 
    'wscript.echo err.number, err.description, err.source '''''Uncomment for debugging info
    err.Clear
end if 