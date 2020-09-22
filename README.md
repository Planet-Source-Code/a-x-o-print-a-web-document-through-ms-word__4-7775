<div align="center">

## Print a web document through MS Word


</div>

### Description

to print a web document using all of wordbasic's functionality, this maybe alright for, ORDER forms, INVOICES.<div style="BACKGROUND-COLOR: black"><font color="Silver"><BR>people dont understand the meaning of the word.. NO FEED BACK..its simple enough isn't ??</font>

</style></div>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[A\_X\_O](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/a-x-o.md)
**Level**          |Intermediate
**User Rating**    |3.7 (26 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Documents/ Frames](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/documents-frames__4-27.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/a-x-o-print-a-web-document-through-ms-word__4-7775/archive/master.zip)





### Source Code

```
<html>
<head>
<title></title>
</head>
<body bgcolor="black">
<table border="0" width="100%" bgcolor="#000000" bordercolor="#000000" cellspacing="0"
cellpadding="0" bordercolorlight="#000000" bordercolordark="#000000">
 <tr>
 <td width="50%"><font color="#FF0000"><marquee border="0" bgcolor="#008080">DILENGER</marquee></font></td>
 <td width="50%"><font color="#FF0000"><marquee direction="right" border="0"
 bgcolor="#008080">DILENGER</marquee></font></td>
 </tr>
</table>
<p><script language="JScript" for="window" event="onload">
 //function hassle() {
 //if (event.button==1 || event.button==2) {
 //hassle();
 //}
 //}
 //document.onmousedown=hassle
 </script> <script
language="VBScript">
 Sub hassle()
 Set WordObj = CreateObject("Word.Basic")
 WordObj.AppShow
 Wordobj.FileNew
 Wordobj.Bold
 Wordobj.CenterPara
 Wordobj.FontSize 40
 Wordobj.Insert "who's gonna buy me a pint ??"
 'Wordobj.FilePrintDefault
 Wordobj.FileNew
 Wordobj.CenterPara
 Wordobj.FontSize 40
 Wordobj.Bold
 Wordobj.Insert "Houses of Parliament"
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.RightPara
 Wordobj.Bold
 Wordobj.Insert "10 Downing Street"
 Wordobj.Insert vbNewline
 Wordobj.Bold
 Wordobj.Insert "Government HQ"
 Wordobj.Insert vbNewline
 Wordobj.Bold
 Wordobj.Insert "London"
 Wordobj.Insert vbNewline
 Wordobj.Bold
 Wordobj.Insert "S1 .."
 Wordobj.Insert vbNewline
 Wordobj.Bold
 Wordobj.Insert "Tel 0800 808080"
 Wordobj.Insert vbNewline
 Wordobj.Bold
 Wordobj.Insert "Email: t.blair@gov.co.uk"
 Wordobj.Insert vbNewline
 Wordobj.Bold
 Wordobj.Insert FormatDateTime(Date, 2)
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.LeftPara
 Wordobj.Insert vbNewline
 Wordobj.Insert "Dear Mr. DILENGER"
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.FontSize 12
 Wordobj.Insert "I cannot afford To buy you a pint of beer because the Nation is in a terrible state"
 Wordobj.Insert vbNewline
 Wordobj.Insert "but as *one* must reply to such feeble questions from the lesser Mortals"
 Wordobj.Insert vbNewline
 Wordobj.Insert "i hear you have just graduated and have a student debt of over £19,000 "
 Wordobj.Insert vbNewline
 Wordobj.Insert "i would say to you Mr Dilenger. screw you... I'm alright JACK, infact it's going to get alot worse for you :-) "
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.Bold
 Wordobj.Insert "WE ARE **NEW** LABOUR... the party that screws STUDENTS"
 Wordobj.Insert vbNewline
 Wordobj.Insert "the party that is creating an *ELITE* University policy in the UK"
 Wordobj.Insert vbNewline
 Wordobj.Insert "don't worry, Mr. Dilenger... we as the Labout Party will make sure you are never DEBT FREE... In fact we will make sure that the Banks Haunt you and your family for the rest of your LIFE"
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.LeftPara
 Wordobj.FontSize 12
 Wordobj.Insert "Take it easy DILENGER"
 Wordobj.Insert vbNewline
 Wordobj.Insert vbNewline
 Wordobj.Insert "TONY"
 WordObj.FileSaveAs "C:\WINDOWS\Desktop\DILENGER.doc"
 Wordobj.FilePrintDefault
 Wordobj.FontSize 12
 'Wordobj.FileClose 1
 'Wordobj.FileClose 2
 Set Wordobj = Nothing
 End Sub
 </script> </p>
<form>
 <div align="center"><center><p><input type="button" value=" Print " name="Btn_Prnt"
 onclick="hassle()"
 style="background-color: rgb(0,0,0); color: rgb(192,192,192); border: thin solid rgb(192,192,192)"></p>
 </center></div>
</form>
</body>
</html>
```

