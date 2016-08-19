<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<script type="text/javascript" src="myword.js"> </script>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>new and open</title>
</head>
<body onload="InitEvent();">
     <input type="button" value="新建word" onclick="NewWord();"/>
     <input type="button" value="关闭word" onclick="CloseWord();"/>
     <input type="button" value="新建word2" onclick="addNewWord();"/>
     <input type="button" value="打开服务器文件" onclick="OpenUrlWord('http://localhost:8080/wordforlast/upload/','test.doc');"/>
      <input type="button" value="打开本地文件" onclick="OpenLocalWord('D:\\2003.doc');"/>
      <input type="button" value="保存到本地文件夹" onclick="SaveToLocal('E:\\test.doc');"/>
      <input type="button" value="保存到服务器" onclick="SaveToSever('http://localhost:8080/wordforlast/upload/severtest.doc');"/>
      <object classid="clsid:00460182-9E5E-11d5-B7C8-B8269041DD57" codebase ="dsoframer.ocx" id="oframe" width="100%" height="500px">
		<param name="BorderStyle" value="1"/>
		<param name="TitlebarColor" value="52479" />
		<param name="TitlebarTextColor" value="0" />
		<param name="Menubar" value="1" />
		<param name="Titlebar" value="0" />
		<param name="Menubar" value="0" />
       </object>
		<!-- <script type="text/javascript" language="jscript" for="oframe" event="OnDocumentOpened(str,obj)">
		   	OnDocumentOpened(str, obj);
		</script>
		<script type="text/javascript" language="jscript" for="oframe" event="OnDocumentClosed()">
			OnDocumentClosed();
		</script>
		-->
</body>
</html>