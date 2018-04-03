---
title: Server Property Example (VBScript)
ms.prod: access
ms.assetid: d35ad670-c9bd-de51-0e4d-ebcad0cb2b55
ms.date: 06/08/2017
---


# Server Property Example (VBScript)

  

**Applies to:** Access 2013 | Access 2016

The following code shows how to set the [RDS.DataControl](http://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) parameter at design time and bind it to a data-aware control using the SQLOLEDB provider. Cut and paste this code into a normal ASP document and name it **ServerDesignVBS.asp**. ASP script will identify your server.




```vb

<!-- BeginServerDesignVBS --><%@ Language=VBScript %>
<html><head>
<meta name="VI60_DefaultClientScript" content=VBScript><meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<title>Server Property Example (VBScript)</title><style>
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
<body><h1>Server Property Example (VBScript)</h1> 
<TABLE DATASRC=#RDS><TR>
<TD> <SPAN DATAFLD="FirstName"></SPAN> </TD><TD> <SPAN DATAFLD="LastName"></SPAN> </TD>
<TD> <SPAN DATAFLD="Title"></SPAN> </TD><TD> <SPAN DATAFLD="Type"></SPAN> </TD>
<TD> <SPAN DATAFLD="Email"></SPAN> </TD></TR>
</TABLE><!-- Remote Data Service with Parameters set at Design Time -->
<OBJECT classid="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33"ID=RDS HEIGHT=1 WIDTH=1>
<PARAM NAME="SQL" VALUE="Select * from Employees"><PARAM NAME="SERVER" VALUE="http://<%=Request.ServerVariables("SERVER_NAME")%>">
<PARAM NAME="CONNECT" VALUE="Provider='sqloledb';Integrated Security='SSPI';Initial Catalog='Northwind'"></OBJECT> 
</body></html>
<!-- EndServerDesignVBS -->
```

The following example shows how to set the necessary parameters of  **RDS.DataControl** at run time. To test this example, cut and paste this code into a normal ASP document and name it **ServerRuntimeVBS.asp**. ASP script will identify your server.



```vb

<!-- BeginServerRuntimeVBS --><%@ Language=VBScript %>
<html><head>
<meta name="VI60_DefaultClientScript" content=VBScript><meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<title>Server Property Example (VBScript)</title><style>
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
<body><h1>Server Property Example (VBScript)</h1> 
<H2>RDS API Code Examples</H2> 
<H3>Remote Data Service Server Property Set at Run Time</H3> 
<!-- RDS.DataControl with no parameters set at design time --><OBJECT classid="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33"
ID=RDC HEIGHT=1 WIDTH=1></OBJECT> 
<TABLE DATASRC=#RDC><TR>
<TD> <SPAN DATAFLD="FirstName"></SPAN> </TD><TD> <SPAN DATAFLD="LastName"></SPAN> </TD>
<TD> <SPAN DATAFLD="Title"></SPAN> </TD><TD> <SPAN DATAFLD="Type"></SPAN> </TD>
<TD> <SPAN DATAFLD="Email"></SPAN> </TD></TR>
</TABLE> 
<HR><Input Size=70 Name="txtServer" Value="HTTP://<%= Request.ServerVariables("SERVER_NAME")%>">
<BR><Input Size=70 Name="txtConnect" Value="Provider='sqloledb';Integrated Security='SSPI';Initial Catalog='Northwind'">
<BR><Input Size=70 Name="txtSQL" Value="Select * from Employees">
<HR><INPUT TYPE=BUTTON NAME="Run" VALUE="Run"><BR> 
<Script Language="VBScript"><!--
' Set parameters of RDS.DataControl at Run TimeSub Run_OnClick
RDC.Server = txtServer.ValueRDC.SQL = txtSQL.Value
RDC.Connect = txtConnect.ValueRDC.Refresh
End Sub-->
</Script> 
</body></html>
<!-- EndServerRuntimeVBS -->
```

[DataControl Object (RDS)](http://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) |[Server Property (RDS)](http://msdn.microsoft.com/library/17519dbe-a43a-1d0d-22c1-dc0def2f63ab%28Office.15%29.aspx)
 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

