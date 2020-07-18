Option Explicit	

Dim obj : Set obj=CreateObject("wscript.shell")

pro=InputBox

obj.Run "taskkill /F /IM " & pro