</html>
<html>
	<head>
		<style>
		ul {
  			list-style-type: none;
  			margin: 0;
  			padding: 0;
  			width: 200px;
  			background-color: #f1f1f1;
		}

li a {
  display: block;
  color: #000;
  padding: 8px 16px;
  text-decoration: none;
}

/* Change the link color on hover */
li a:hover {
  background-color: #555;
  color: white;
}
</style>
    <Title>Message Box</Title>
<HTA:Application
    ID="MsgBox"
    WINDOWSTATE="normal"
    border="thick"
    borderStyle="raised"
    scrollFlat=yes
    caption=yes
    SYSMENU=yes
    contextMenu=no
    ShowInTaskBar=yes
    MaximizeButton=no
    MinimizeButton=yes/>
	</head>
  <script language="vbscript">
  class getWind
    
  end class
''''''''''''''''''''''''''''''''
    Sub Window_OnLoad()
	    window.resizeTo 300,200
	    window.moveTo (screen.width/2)-1366/2,360/2
    End Sub

    Function yesBtn()
			Set bin=CreateObject("wscript.shell")
			bin.CurrentDirectory = "C:\Users\andregoro\Documents\Video"
			bin.Run "Estudo.exe"
    End Function

    Function noBtn()
        MsgBox "Close"
        window.close
    End Function

  </script>
  <style>
  html,body{
      height: 100%;
      width: 100%;
      margin: 0px;
      padding: 0px;
      font-family: Arial;
      font-size: 10pt;
  }
  #tpArea{
      margin:10px;
      padding-bottom:60px;
  }
  </style>
    <body onload="Window_OnLoad()">
<h2>Vertical Navigation Bar</h2>

<ul>
  <li><a href="#home">Home</a></li>
  <li><a href="#news">News</a></li>
  <li><a href="#contact">Contact</a></li>
  <li><a href="#about">About</a></li>
</ul>
    </body>
</html>