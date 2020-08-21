Dim WSS : Set WSS=CreateObject("WScript.Shell")

temp = WSS.Environment("User").Item("TEMP")
MsgBox WSS.ExpandEnvironmentStrings(temp)

Set User =WSS.Environment("User")
User.Item("VBSCL")=5
msgbox User.Item("VBSCL")
User.Remove("VBSCL")