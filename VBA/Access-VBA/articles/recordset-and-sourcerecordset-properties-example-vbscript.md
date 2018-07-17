---
title: Recordset and SourceRecordset Properties Example (VBScript)
ms.prod: access
ms.assetid: 235118ce-8468-18b1-ff49-8739fde69427
ms.date: 06/08/2017
---


# Recordset and SourceRecordset Properties Example (VBScript)

  

**Applies to:** Access 2013 | Access 2016

The following example shows how to set the necessary parameters of the [RDSServer.DataFactory](http://msdn.microsoft.com/library/1de76cdd-34dc-8547-29aa-48ad6067bdea%28Office.15%29.aspx) default business rules at run time.

To test this example, cut and paste this code between the <Body> and </Body> tags in a normal HTML document and name it  **RecordsetVBS.asp**. ASP script will identify your server.



```vb

<!-- BeginRecordSetVBS --><%@ Language=VBScript %>
<html><head>
<meta name="VI60_DefaultClientScript" content=VBScript><meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<title>Recordset and SourceRecordset Properties Example (VBScript)</title><style>
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
<body><h1>Recordset and SourceRecordset Properties Example (VBScript)</h1> 
<Center><H2>RDS API Code Examples</H2>
<HR><H3>Using SourceRecordset and Recordset with RDSServer.DataFactory</H3>
<!-- RDS.DataSpace ID RDS1 --><OBJECT ID="RDS1" WIDTH=1 HEIGHT=1
CLASSID="CLSID:BD96C556-65A3-11D0-983A-00C04FC29E36"></OBJECT> 
<!-- RDS.DataControl with parameters set at Run Time --><OBJECT classid="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33"
ID=RDC WIDTH=1 HEIGHT=1></OBJECT> 
<TABLE DATASRC=#RDC><TR>
<TD> <INPUT DATAFLD="FirstName" SIZE=15> </TD><TD> <INPUT DATAFLD="LastName" SIZE=15></TD>
</TR></TABLE>
<HR><Input Size=70 Name="txtServer" Value="http://<%=Request.ServerVariables("SERVER_NAME")%>"><BR>
<Input Size=70 Name="txtConnect" Value="Provider='sqloledb';Integrated Security='SSPI';Initial Catalog='Northwind'"><BR><Input Size=70 Name="txtSQL" Value="SELECT FirstName, LastName FROM Employees">
<HR><INPUT TYPE=BUTTON NAME="Run" VALUE="Run"><BR> 
</Center><Script Language="VBScript"> 
Dim rdsDFDim strServer
strServer = "http://<%=Request.ServerVariables("SERVER_NAME")%>" 
Sub Run_OnClick() 
Dim rs' Create RDSServer.DataFactory Object
Set rdsDF = RDS1.CreateObject("RDSServer.DataFactory", strServer)' Get Recordset
Set rs = rdsDF.Query(txtConnect.Value,txtSQL.Value) 
' Set parameters of RDS.DataControl at run timeRDC.Server = txtServer.Value
RDC.SQL = txtSQL.ValueRDC.Connect = txtConnect.Value
RDC.Refresh 
End Sub 
</Script> 
</body></html>
<!-- EndRecordsetVBS -->

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

