---
title: Connect Property Example (VBScript)
ms.prod: access
ms.assetid: bd88c63f-89d9-c73b-3ee0-288ff078b938
ms.date: 06/08/2017
---


# Connect Property Example (VBScript)

  

**Applies to:** Access 2013 | Access 2016

This code shows how to set the [Connect](http://msdn.microsoft.com/library/11aa3284-18e9-6d2d-761b-c25090370b77%28Office.15%29.aspx) property at design time:




```
 
<OBJECT CLASSID="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33" ID="ADC1"> 
. 
   <PARAM NAME="SQL" VALUE="Select * from Sales"> 
   <PARAM NAME="CONNECT" VALUE="Provider=SQLOLEDB;Integrated Security=SSPI;Initial Catalog=Pubs"> 
   <PARAM NAME="Server" VALUE="http://MyWebServer"> 
. 
</OBJECT> 

```

The following example shows how to set the  **Connect** property at run time in VBScript code.
To test this example, copy and paste this code between the <Body> and </Body> tags in a normal HTML document and name it  **ConnectVBS.asp**. ASP script will identify your server.



```

<!-- BeginConnectVBS --><%@ Language=VBScript %>
<HTML><HEAD>
<title>ADO Connect Property</title><%' local style sheet used for display%>
<STYLE><!--
BODY {font-family: 'Verdana','Arial','Helvetica',sans-serif;
BACKGROUND-COLOR:white;COLOR:black;
}.tbody {
text-align: center;background-color: #f7efde;
font-family: 'Verdana','Arial','Helvetica',sans-serif;font-size: x-small;
}-->
</STYLE></HEAD>
<BODY><h1>ADO Connect Property (RDS)</h1>
<HR><H3>Set Connect Property at Run Time</H3>
<% ' RDS.DataControl with no parameters set at design time %><OBJECT classid="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33" ID=RDS HEIGHT=1 WIDTH=1></OBJECT>
<% ' Bind table to control for data display %><TABLE DATASRC=#RDS>
<TBODY><TR class="tbody">
<TD><SPAN DATAFLD="FirstName"></SPAN></TD><TD><SPAN DATAFLD="LastName"></SPAN></TD>
</TR></TBODY>
</TABLE><FORM name="frmInput">
SERVER: <INPUT Name="txtServer" Size="103" Value="http://<%=Request.ServerVariables("SERVER_NAME")%>"><BR>DATA SOURCE: <INPUT Name="txtDataSource" Size="93" Value="<%=Request.ServerVariables("SERVER_NAME")%>"><BR>
CONNECT: <INPUT Name="txtConnect" Size="100"><BR>SQL: <INPUT Name="txtSQL" Size="110" Value="Select FirstName, LastName from Employees">
<BR><INPUT TYPE=BUTTON NAME="Run" VALUE="Run">
<h4>To make data grid appear, click 'Run' to see the connect string in text box above.
</h4></FORM>
<Script Language="VBScript">' Set parameters of RDS.DataControl at Run Time
Sub Run_OnClickDim Cnxn
' build connection stringCnxn = "Provider='sqloledb';"
Cnxn = Cnxn &; "Data Source="Cnxn = Cnxn &; document.frmInput.txtDataSource.value &; ";"
Cnxn = Cnxn &; "Initial Catalog='Northwind';"Cnxn = Cnxn &; "Integrated Security='SSPI';"
' assign the valuedocument.frmInput.txtConnect.value = Cnxn
MsgBox "Here we go!"' set RDS properties
RDS.Server = document.frmInput.txtServer.valueRDS.SQL = document.frmInput.txtSQL.value
RDS.Connect = document.frmInput.txtConnect.valueRDS.Refresh
End Sub</Script>
</BODY></HTML>
<!-- EndConnectVBS -->
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

