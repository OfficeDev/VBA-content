---
title: CreateRecordset Method Example (VBScript)
ms.prod: access
ms.assetid: 548e5c0a-74cc-0abb-f660-1be483410548
ms.date: 06/08/2017
---


# CreateRecordset Method Example (VBScript)

  

**Applies to:** Access 2013 | Access 2016

This code example creates a [Recordset](http://msdn.microsoft.com/library/0f963bf8-f066-dc8a-b754-f427de712df1%28Office.15%29.aspx) on the server side. It has two columns with four rows each. Cut and paste the following code to Notepad or another text editor and save it as **CreateRecordsetVBS.asp**.




```vb

<!-- BeginCreateRecordsetVBS --><%@ Language=VBScript %>
<html><head>
<meta name="VI60_DefaultClientScript" content=VBScript><meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<title>CreateRecordset Method Example (VBScript)</title><style>
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
<body> 
<OBJECT classid=clsid:BD96C556-65A3-11D0-983A-00C04FC29E33 height=1 id=DC1 width=1></OBJECT>
<h1>CreateRecordset Method Example (VBScript)</h1><script language = "vbscript">
' use the RDS.DataControl to create an empty recordset;' takes an array of variants where every element is itself another
' array of variants, one for every column required in the recordset 
' the elements of the inner array are the column's' name, type, size, and nullability
Sub GetRS()Dim Record(2)
Dim Field1(3)Dim Field2(3)
Dim Field3(3) 
' for each field, specify the name type, size, and nullability 
Field1(0) = "Name" ' Column name.Field1(1) = CInt(129) ' Column type.
Field1(2) = CInt(40) ' Column size.Field1(3) = False ' Nullable? 
Field2(0) = "Age"Field2 (1) = CInt(3)
Field2 (2) = CInt(-1)Field2 (3) = True 
Field3 (0) = "DateOfBirth"Field3 (1) = CInt(7)
Field3 (2) = CInt(-1)Field3 (3) = True 
' put all fields into an array of arraysRecord(0) = Field1
Record(1) = Field2Record(2) = Field3 
Dim NewRsSet NewRS = DC1.CreateRecordset(Record) 
Dim fields(2)fields(0) = Field1(0)
fields(1) = Field2(0)fields(2) = Field3(0) 
' Populate the new recordset with data values.Dim fieldVals(2) 
' Use AddNew to add the records.fieldVals(0) = "Joe"
fieldVals(1) = 5fieldVals(2) = CDate(#1/5/96#)
NewRS.AddNew fields, fieldVals 
fieldVals(0) = "Mary"fieldVals(1) = 6
fieldVals(2) = CDate(#6/5/94#)NewRS.AddNew fields, fieldVals 
fieldVals(0) = "Alex"fieldVals(1) = 13
fieldVals(2) = CDate(#1/6/88#)NewRS.AddNew fields, fieldVals 
fieldVals(0) = "Susan"fieldVals(1) = 13
fieldVals(2) = CDate(#8/6/87#)NewRS.AddNew fields, fieldVals 
NewRS.MoveFirst 
' Set the newly created and populated Recordset to' the SourceRecordset property of the
' RDS.DataControl to bind to visual controls 
Set DC1.SourceRecordset = NewRSEnd Sub
</script><table datasrc="#DC1" align="center">
<thead><tr id="ColHeaders" class="thead2">
<th>Name</th><th>Age</th>
<th>D.O.B.</th></tr>
</thead><tbody class="tbody">
<tr><td><input datafld="Name" size=15 id=text1 name=text1> </td>
<td><input datafld="Age" size=25 id=text2 name=text2> </td><td><input datafld="DateOfBirth" size=25 id=text3 name=text3> </td>
</tr></tbody>
</table><input type = "button" onclick = "GetRS()" value="Go!">
</body></html>
<!-- EndCreateRecordsetVBS -->
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

