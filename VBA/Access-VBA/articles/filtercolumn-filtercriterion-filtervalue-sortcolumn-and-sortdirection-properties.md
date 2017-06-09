---
title: FilterColumn, FilterCriterion, FilterValue, SortColumn, and SortDirection Properties and Reset Method Example (VBScript)
ms.prod: access
ms.assetid: bc22a6c4-b9d4-ad38-d802-4790ff3262a6
ms.date: 06/08/2017
---


# FilterColumn, FilterCriterion, FilterValue, SortColumn, and SortDirection Properties and Reset Method Example (VBScript)

  

**Applies to:** Access 2013 | Access 2016

The following code shows how to set the  **RDS.DataControl** **Server** parameter at design time and bind it to a data-aware HTML table using a data source. Cut and paste the following code to Notepad or another text editor and save it as **FilterColumnVBS.asp**.




```vb

<!-- BeginFilterColumnVBS --><%@ Language=VBScript %>
<HTML><HEAD>
<META name="VI60_DefaultClientScript" Content="VBScript"> 
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0"><TITLE>FilterColumn, FilterCriterion, FilterValue, SortColumn, and SortDirection
Properties and Reset Method Example (VBScript)</TITLE></HEAD>
<BODY><h1>FilterColumn, FilterCriterion, FilterValue, SortColumn, and SortDirection
Properties and Reset Method Example (VBScript)</h1> 
<OBJECT classid="clsid:BD96C556-65A3-11D0-983A-00C04FC29E33" ID=RDS HEIGHT=1 WIDTH=1><PARAM NAME="SQL" VALUE="Select FirstName, LastName, Title, ReportsTo, Extension from Employees">
<PARAM NAME="Connect" VALUE="Provider='sqloledb';Data Source=<%=Request.ServerVariables("SERVER_NAME")%>;Integrated Security='SSPI';Initial Catalog='Northwind'"><PARAM NAME="Server" VALUE="http://<%=Request.ServerVariables("SERVER_NAME")%>">
</OBJECT> 
Sort Column: <SELECT NAME="cboSortColumn"><OPTION VALUE=""></OPTION>
<OPTION VALUE=ID>ID</OPTION><OPTION VALUE=FirstName>FirstName</OPTION>
<OPTION VALUE=LastName>LastName</OPTION><OPTION VALUE=Title>Title</OPTION>
<OPTION VALUE=Title>ReportsTo</OPTION><OPTION VALUE=Phone>Extension</OPTION>
</SELECT><br>
Sort Direction: <SELECT NAME="cboSortDir"><OPTION VALUE=""></OPTION>
<OPTION VALUE=TRUE>Ascending</OPTION><OPTION VALUE=FALSE>Descending</OPTION>
</SELECT><HR WIDTH="25%">
Filter Column: <SELECT NAME="cboFilterColumn"><OPTION VALUE=""></OPTION>
<OPTION VALUE=FirstName>FirstName</OPTION><OPTION VALUE=LastName>LastName</OPTION>
<OPTION VALUE=Title>Title</OPTION><OPTION VALUE=Room>ReportsTo</OPTION>
<OPTION VALUE=Phone>Extension</OPTION></SELECT>
<br>Filter Criterion: <SELECT NAME="cboCriterion">
<OPTION VALUE=""></OPTION><OPTION VALUE="=">=</OPTION>
<OPTION VALUE="&;gt;">&;gt;</OPTION><OPTION VALUE="&;lt;">&;lt;</OPTION>
<OPTION VALUE="&;gt;=">&;gt;=</OPTION><OPTION VALUE="&;lt;=">&;lt;=</OPTION>
<OPTION VALUE="&;lt;&;gt;">&;lt;&;gt;</OPTION></SELECT>
<br>Filter Value: <INPUT NAME="txtFilterValue">
<HR WIDTH="25%"><INPUT TYPE=BUTTON NAME=Clear VALUE="CLEAR ALL"> &;nbsp;
<INPUT TYPE=BUTTON NAME=SortFilter VALUE="APPLY"> 
<HR><TABLE DATASRC=#RDS ID="DataTable">
<THEAD><TR>
<TH>FirstName</TH><TH>LastName</TH>
<TH>Title</TH><TH>Reports To</TH>
<TH>Extension</TH></TR>
</THEAD><TBODY>
<TR><TD><SPAN DATAFLD="FirstName"></SPAN></TD>
<TD><SPAN DATAFLD="LastName"></SPAN></TD><TD><SPAN DATAFLD="Title"></SPAN></TD>
<TD><SPAN DATAFLD="ReportsTo"></SPAN></TD><TD><SPAN DATAFLD="Extension"></SPAN></TD>
</TR></TBODY>
</TABLE> 
<Script Language="VBScript"><!--
Const adFilterNone = 0 
Sub SortFilter_OnClickDim vCriterion
Dim vSortDirDim vSortCol
Dim vFilterCol 
' The value of SortColumn will be the' value of what the user picks in the
' cboSortColumn box.vSortCol = cboSortColumn.options(cboSortColumn.selectedIndex).value 
If(vSortCol <> "") thenRDS.SortColumn = vSortCol
End If 
' The value of SortDirection will be the' value of what the user specifies in the
' cboSortdirection box. 
If (vSortCol <> "") thenvSortDir = cboSortDir.options(cboSortDir.selectedIndex).value
If (vSortDir = "") thenMsgBox "You must select a direction for the sort."
Exit SubElse
If vSortDir = "Ascending" Then vSortDir = "TRUE"If vSortDir = "Descending" Then vSortDir = "FALSE"
RDS.SortDirection = vSortDirEnd If
End If 
' The value of FilterColumn will be the' value of what the user specifies in the
' cboFilterColumn box.vFilterCol = cboFilterColumn.options(cboFilterColumn.selectedIndex).value 
If(vFilterCol <> "") thenRDS.FilterColumn = vFilterCol
End If 
' The value of FilterCriterion will be the' text value of what the user specifies in the
' cboCriterion box.vCriterion = cboCriterion.options(cboCriterion.selectedIndex).value
If (vCriterion <> "") ThenRDS.FilterCriterion = vCriterion
End If 
' txtFilterValue is a rich text box' control. The value of FilterValue will be the
' text value of what the user specifies in the' txtFilterValue box.
If (txtFilterValue.value <> "") ThenRDS.FilterValue = txtFilterValue.value
End If 
' Execute the sort and filter on a client-side' Recordset based on the specified sort and filter
' properties. Calling Reset refreshes the result set' that is displayed in the data-bound controls to
' display the filtered, sorted recordset.RDS.Reset
End Sub 
Sub Clear_onClick()'clear the HTML input controls
cboSortColumn.selectedIndex = 0cboSortDir.selectedIndex = 0
cboFilterColumn.selectedIndex = 0cboCriterion.selectedIndex = 0
txtFilterValue.value = "" 
'clear the filterRDS.FilterCriterion = ""
RDS.Reset(FALSE)End Sub
--></Script> 
</BODY></HTML>
<!-- EndFilterColumnVBS -->
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

