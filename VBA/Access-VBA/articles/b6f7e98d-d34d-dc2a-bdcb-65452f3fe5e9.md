
# AddNew Method Example (JScript)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

This example uses the [AddNew](bae09be0-5707-4f38-9c74-0acd0f29dbac.md) method to create a new record with the specified name. Cut and paste the following code to Notepad or another text editor, and save it as **AddNewJS.asp**.




```js
 
<!-- BeginAddNewJS --> 
<%@LANGUAGE="JScript" %> 
<!-- Include file for JScript ADO Constants --> 
<%// use this meta tag instead of adojavas.inc%> 
<!--METADATA TYPE="typelib" uuid="00000205-0000-0010-8000-00AA006D2EA4" --> 
 
<html> 
 
<head> 
 <title>Add New Method Example (JScript)</title> 
<style> 
<!-- 
body { 
 font-family: 'Verdana','Arial','Helvetica',sans-serif; 
 BACKGROUND-COLOR:white; 
 COLOR:black; 
 } 
--> 
</style> 
</head> 
 
<body> 
<h1>AddNew Method Example (JScript)</h1> 
 
<% 
 if (Request.Form("Addit") == "AddNew") 
 { 
 // connection and recordset variables 
 var Cnxn = Server.CreateObject("ADODB.Connection") 
 var strCnxn = "Provider='sqloledb';Data Source=" + Request.ServerVariables("SERVER_NAME") + ";" + 
 "Initial Catalog='Northwind';Integrated Security='SSPI';"; 
 var rsEmployee = Server.CreateObject("ADODB.Recordset"); 
 //record variables 
 var FName = String(Request.Form("FirstName")); 
 var LName = String(Request.Form("LastName")); 
 
 try 
 { 
 // open connection 
 Cnxn.Open(strCnxn) 
 
 // open Employee recordset using client-side cursor 
 rsEmployee.CursorLocation = adUseClient; 
 rsEmployee.Open("Employees", strCnxn, adOpenKeyset, adLockOptimistic, adCmdTable); 
 
 rsEmployee.AddNew(); 
 rsEmployee("FirstName") = FName; 
 rsEmployee("LastName") = LName; 
 rsEmployee.Update; 
 
 // of course, you would normally do error handling here 
 Response.Write("New record added.") 
 } 
 catch (e) 
 { 
 Response.Write(e.message); 
 } 
 finally 
 { 
 // clean up 
 if (rsEmployee.State == adStateOpen) 
 rsEmployee.Close; 
 if (Cnxn.State == adStateOpen) 
 Cnxn.Close; 
 rsEmployee = null; 
 Cnxn = null; 
 } 
 } 
%> 
 
<form method="post" action="AddNewJS.asp" id=form1 name=form1> 
<table> 
<tr> 
 <td colspan="2"> 
 <h4>Please enter the record to add:</h4> 
 </td> 
</tr> 
<tr> 
 <td> 
 First Name: 
 </td> 
 <td> 
 <input name="FirstName" maxLength=20> 
 </td> 
</tr> 
<tr> 
 <td> 
 Last Name: 
 </td> 
 <td> 
 <input name="LastName" size="30" maxLength=30> 
 </td> 
</tr> 
<tr> 
 <td align="right"> 
 <input type="submit" value="Submit" name="Submit"> 
 </td> 
 <TD align="left"> 
 <INPUT type="reset" value="Reset" name="Reset"> 
 </TD> 
</tr> 
</table> 
<INPUT type="hidden" value="AddNew" name="Addit"> 
</form> 
</body> 
</HTML> 
<!-- EndAddNewJS --> 
 

```

