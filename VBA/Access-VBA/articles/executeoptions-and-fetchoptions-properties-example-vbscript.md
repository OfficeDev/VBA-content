---
title: ExecuteOptions and FetchOptions Properties Example (VBScript)
ms.prod: access
ms.assetid: 54a1decc-d774-9521-5808-0fcb4294facb
ms.date: 06/08/2017
---


# ExecuteOptions and FetchOptions Properties Example (VBScript)

  

**Applies to:** Access 2013 | Access 2016

The following code shows how to set the [ExecuteOptions](http://msdn.microsoft.com/library/fb244cbd-9a03-9128-1373-694c9061c9da%28Office.15%29.aspx) and[FetchOptions](http://msdn.microsoft.com/library/0d86c5e4-9abc-5c0e-dc04-4183f4c278cc%28Office.15%29.aspx) properties at design time. If left unset, **ExecuteOptions** defaults to **adcExecSync**. This setting indicates that when the **RDS.Refresh** method is called, it will be executed on the current calling thread — that is, synchronously. Cut and paste the following code to Notepad or another text editor and save it as **ExecuteOptionsDesignVBS.asp**.




```sql

<!-- BeginExecuteOptionsDesignVBS --><%@ Language=VBScript %>
<html><head>
<meta name="VI60_DefaultClientScript" content=VBScript><meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<title>Design-time ExecuteOptions and FetchOptions Properties Example</title><style>
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
<body><h2>Design-time <br> ExecuteOptions and FetchOptions Properties Example</h2> 
<OBJECT CLASSID="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33" ID=RDS height=1 width=1><PARAM NAME="SQL" VALUE="SELECT FirstName, LastName FROM Employees ORDER BY LastName">
<PARAM NAME="Connect" VALUE="Provider='sqloledb';Data Source=<%=Request.ServerVariables("SERVER_NAME")%>;Integrated Security='SSPI';Initial Catalog='Northwind'"><PARAM NAME="Server" VALUE="http://<%=Request.ServerVariables("SERVER_NAME")%>">
<PARAM NAME="ExecuteOptions" VALUE="1"><PARAM NAME="FetchOptions" VALUE="3">
</OBJECT> 
<TABLE DATASRC=#RDS><TBODY>
<TR class="thead2"><TH>First Name</TH>
<TH>Last Name</TH></TR>
<TR class="tbody"><TD><SPAN DATAFLD="FirstName"></SPAN></TD>
<TD><SPAN DATAFLD="LastName"></SPAN></TD></TR>
</TBODY></TABLE> 
</body></html>
<!-- EndExecuteOptionsDesignVBS -->
```

The following example shows how to set the  **ExecuteOptions** and **FetchOptions** properties at run time in VBScript code. See the[Refresh](http://msdn.microsoft.com/library/968baa7c-9128-7155-a1eb-d77aedda6601%28Office.15%29.aspx) method for a working example of these properties. Cut and paste the following code to Notepad or another text editor and save it as **ExecuteOptionsRuntimeVBS.asp**.



```sql

<!-- BeginExecuteOptionsRuntimeVBS --><%@ Language=VBScript %>
<html><head>
<meta name="VI60_DefaultClientScript" content=VBScript><meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<title>Run-time ExecuteOptions and FetchOptions Properties Example</title><style>
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
<body><h2>Run-time <br> ExecuteOptions and FetchOptions Properties Example</h2>
<OBJECT CLASSID="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33" ID=RDS height=1 width=1><PARAM NAME="SQL" VALUE="SELECT FirstName, LastName FROM Employees ORDER BY LastName">
<PARAM NAME="Connect" VALUE="Provider='sqloledb';Data Source=<%=Request.ServerVariables("SERVER_NAME")%>;Integrated Security='SSPI';Initial Catalog='Northwind'"><PARAM NAME="Server" VALUE="http://<%=Request.ServerVariables("SERVER_NAME")%>">
</OBJECT> 
<TABLE DATASRC=#RDS><TBODY>
<TR class="thead2"><TH>First Name</TH>
<TH>Last Name</TH></TR>
<TR class="tbody"><TD><SPAN DATAFLD="FirstName"></SPAN></TD>
<TD><SPAN DATAFLD="LastName"></SPAN></TD></TR>
</TBODY></TABLE>
<Script Language="VBScript">Const adcExecSync = 1
Const adcFetchAsynch = 3 
Sub ExecuteHow' set RDS properties at run-time
RDS1.ExecuteOptions = adcExecSyncRDS1.FetchOptions = adcFetchAsynch
RDS.RefreshEnd Sub
</Script></body>
</html><!-- EndExecuteOptionsRuntimeVBS -->

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

