var oframe;
var isOpend=false;
function InitEvent()
{
	oframe=document.getElementById("oframe");
}
function NewWord()
{
	isOpend=true;
	oframe.showdialog(0);
}
function CloseWord()
{
    oframe.close();
}
function addNewWord()
{
	oframe.CreateNew("Word.Document");
}
function OpenUrlWord(hosturl,filename)
{
	var url=hosturl+filename;
	alert(url);
	oframe.Open(url,true);
}
function OpenLocalWord(filename)
{
	oframe.Open(filename,false,"Word.Document");
}
function SaveToLocal(filename)
{
	oframe.Save(filename,true);
}
function SaveToSever(filename)
{
	
}
