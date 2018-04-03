---
title: ReadyState Property Example (VBScript)
ms.prod: access
ms.assetid: 0deacb21-4503-cee5-ea8c-6b30af903ab5
ms.date: 06/08/2017
---


# ReadyState Property Example (VBScript)

  

**Applies to:** Access 2013 | Access 2016

The following example shows how to read the [ReadyState](http://msdn.microsoft.com/library/e7b62205-a604-ef43-2f5d-9b51b46d2b5a%28Office.15%29.aspx) property of the[RDS.DataControl](http://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) object at run time in VBScript code. **ReadyState** is a read-only property.

To test this example, cut and paste this code between the <Body> and </Body> tags in a normal HTML document and name it  **RDSReadySt.asp**. Use **Find** to locate the file Adovbs.inc and place it in the directory you plan to use. ASP script will identify your server.



```vb

<!-- BeginReadyStateVBS --><%@ Language=VBScript %>
<% 'use the following META tag instead of adovbs.inc%><!--METADATA TYPE="typelib" uuid="00000205-0000-0010-8000-00AA006D2EA4" -->
<html><head>
<meta name="VI60_DefaultClientScript" content=VBScript><meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<title>RDS.DataControl ReadyState Property</title><style>
<!--body {
font-family: 'Verdana','Arial','Helvetica',sans-serif;BACKGROUND-COLOR:white;
COLOR:black;}
.thead {background-color: #008080;
font-family: 'Verdana','Arial','Helvetica',sans-serif;font-size: x-small;
color: white;}
.thead2 {background-color: #800000;
font-family: 'Verdana','Arial','Helvetica',sans-serif;font-size: x-small;
color: white;}
.tbody {text-align: center;
background-color: #f7efde;font-family: 'Verdana','Arial','Helvetica',sans-serif;
font-size: x-small;}
--></style>
</head> 
<body><H1>RDS.DataControl ReadyState Property</H1>
<H2>RDS API Code Examples </H2><HR>
<!-- RDS.DataControl with parameters set at design time --><OBJECT classid="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33" ID=RDS>
<PARAM NAME="SQL" VALUE="Select * from Orders"><PARAM NAME="SERVER" VALUE="http://<%=Request.ServerVariables("SERVER_NAME")%>">
<PARAM NAME="CONNECT" VALUE="Provider='sqloledb';Integrated Security='SSPI';Initial Catalog='Northwind'"><PARAM NAME="ExecuteOptions" VALUE="2">
<PARAM NAME="FetchOptions" VALUE="3"></OBJECT> 
<TABLE DATASRC=#RDS><TBODY>
<TR><TD><SPAN DATAFLD="OrderID"></SPAN></TD>
</TR></TBODY>
</TABLE> 
<Script Language="VBScript"> 
Sub Window_OnLoad 
Select Case RDS.ReadyStatecase 2 'adcReadyStateLoaded
MsgBox "Executing Query"case 3 'adcReadyStateInteractive
MsgBox "Fetching records in background"case 4 'adcReadyStateComplete
MsgBox "All records fetched"End Select 
End Sub</Script> 
</body></html>
<!-- EndReadyStateVBS -->

```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

