---
title: Refresh Method Example (VBScript)
ms.prod: access
ms.assetid: b1e78418-9770-b0b4-1f24-f8ef866b7b42
ms.date: 06/08/2017
---


# Refresh Method Example (VBScript)

  

**Applies to:** Access 2013 | Access 2016



The following example shows how to set the necessary parameters of [RDS.DataControl](http://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) at run time. The manner in which a[Recordset](http://msdn.microsoft.com/library/0f963bf8-f066-dc8a-b754-f427de712df1%28Office.15%29.aspx) is retrieved using the[Refresh](http://msdn.microsoft.com/library/f1c8829f-9c7d-12b6-7470-727ff38d663e%28Office.15%29.aspx) method is determined by the settings of the[ExecuteOptions](http://msdn.microsoft.com/library/fb244cbd-9a03-9128-1373-694c9061c9da%28Office.15%29.aspx) and[FetchOptions](http://msdn.microsoft.com/library/0d86c5e4-9abc-5c0e-dc04-4183f4c278cc%28Office.15%29.aspx) properties. To test this example, cut and paste the following code into a normal ASP document and name it **RefreshVBS.asp**. Use **Find** to locate the file Adovbs.inc and place it in the directory you plan to use. ASP script will identify your server.

```vb

<!-- BeginRefreshVBS --><%@ Language=VBScript %>
<!--use the following META tag instead of adovbs.inc--><!--METADATA TYPE="typelib" uuid="00000205-0000-0010-8000-00AA006D2EA4" -->
<html><head>
<meta name="VI60_DefaultClientScript" content=VBScript><meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<title>Refresh Method Example (VBScript)</title><style>
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
<body><h1>Refresh Method Example (VBScript)</h1> 
<H2>RDS API Code Examples </H2><HR>
<TABLE DATASRC=#RDC><TR>
<TD> <INPUT DATAFLD="FirstName" SIZE=15> </TD><TD> <INPUT DATAFLD="LastName" SIZE=15> </TD>
<TD> <INPUT DATAFLD="Title" SIZE=15> </TD><TD> <INPUT DATAFLD="HireDate" SIZE=15> </TD>
</TR></TABLE> 
<!-- RDS.DataControl with no parameters set at design time --><OBJECT classid="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33"
ID=RDC HEIGHT=1 WIDTH=1></OBJECT>
<HR>Server: <Input Size=70 Name="txtServer" Value="http://<%=Request.ServerVariables("SERVER_NAME")%>"><BR>
Connect: <Input Size=70 Name="txtConnect" Value="Provider='sqloledb';Integrated Security='SSPI';Initial Catalog='Northwind'"><BR>SQL: <Input Size=70 Name="txtSQL" Value="Select * from Employees">
<HR><TABLE BORDER=1 WIDTH="60%">
<TR><TD COLSPAN=3 BGCOLOR=silver>
Choose if you want the Recordset brought back Synchronously on thecurrent calling thread or Asynchronously on another thread.
</TD></TR>
<TR><TD>Synchronously: <BR>
<Input Type="Radio" Name="optExecuteOptions" Checked OnClick="SetExO('adcExecSync')"></TD>
<TD>Asynchronously: <BR><Input Type="Radio" Name="optExecuteOptions" OnClick="SetExO('adcExecAsync')">
</TD><TD>&;nbsp;</TD>
</TR><TR>
<TD COLSPAN=3 BGCOLOR=silver>Fetch Up Front, Background Fetch with Blocking or Background Fetch
without Blocking</TD>
<TR><TD>Up Front:<BR>
<Input Type="Radio" Name="optFetchOptions" OnClick="SetFO('adcFetchUpFront')"></TD>
<TD>Background w/ Blocking:<BR><Input Type="Radio" Name="optFetchOptions" Checked OnClick="SetFO('adcFetchBackground')">
</TD><TD>Background w/o Blocking:<BR>
<Input Type="Radio" Name="optFetchOptions" OnClick="SetFO('adcFetchAsync')"></TD>
</TR></TABLE> 
<INPUT TYPE=BUTTON NAME="Run" VALUE="Run"><BR> 
<Script Language="VBScript"><!--
Dim EO 'ExecuteOptionsDim FO 'FetchOptions
EO = "adcExecSync" 'Default valueFO = "adcFetchBackground" 'Default value 
Sub SetExO(NewEO)EO = NewEO
End Sub 
Sub SetFO(NewFO)FO = NewFO
End Sub 
' Set parameters of RDS.DataControl at Run TimeSub Run_OnClick
RDC.Server = txtServer.ValueRDC.SQL = txtSQL.Value
RDC.Connect = txtConnect.ValueIf EO = "adcExecSync" Then 'Determine which ExecuteOption chosen
RDC.ExecuteOptions = adcExecSyncMsgBox "Recordset brought in on current calling thread Syncronously"
ElseRDC.ExecuteOptions = adcExecAsync
MsgBox "Recordset brought in on another thread Asyncronously"End If 
If FO = "adcFetchBackground" Then 'Determine 'which FetchOption chosenRDC.FetchOptions = adcFetchBackground
MsgBox "Control goes back to user after first batch of records returned"ElseIf FO = " adcFetchUpFront" Then
RDC.FetchOptions = adcFetchUpFrontMsgBox "All records returned before control goes back to user"
ElseRDC.FetchOptions = adcFetchAsync
MsgBox "Control goes back to user immediately"End If 
RDC.RefreshEnd Sub
--></Script> 
</body></html>
<!-- EndRefreshVBS -->

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

