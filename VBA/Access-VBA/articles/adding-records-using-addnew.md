---
title: Adding Records Using AddNew
ms.prod: access
ms.assetid: b37978c6-cb5c-d54d-d6d8-b088b6218c5b
ms.date: 06/08/2017
---


# Adding Records Using AddNew


**Applies to:** Access 2013 | Access 2016

This is the basic syntax of the  **AddNew** method:




```
recordset.AddNewFieldList,Values
```

The  _FieldList_ and _Values_ arguments are optional. _FieldList_ is either a single name or an array of names or ordinal positions of the fields in the new record.
The  _Values_ argument is either a single value or an array of values for the fields in the new record.
Typically, when you intend to add a single record, you will call the  **AddNew** method without any arguments. Specifically, you will call **AddNew,** set the **Value** of each field in the new record, and then call **Update** and/or **UpdateBatch**. You can ensure that your **Recordset** supports adding new records by using the **Supports** property with the **adAddNew** enumerated constant.
The following code uses this technique to add a new Shipper to the sample  **Recordset**. The ShipperID field value is supplied automatically by SQL Server, so the code does not attempt to supply a field value for the new records.



```vb
 
'BeginAddNew1.1 
 If objRs1.Supports(adAddNew) Then 
 With objRs1 
 .AddNew 
 .Fields("CompanyName") = "Sample Shipper" 
 .Fields("Phone") = "(931) 555-6334" 
 .Update 
 End With 
 End If 
'EndAddNew1.1 

```

Because this code uses a disconnected  **Recordset** with a client-side cursor in batch mode, you must reconnect the **Recordset** to the data source with a new **Connection** object before you can call the **UpdateBatch** method to post changes to the database. This is easily done by using the new function GetNewConnection.



```
 
'BeginAddNew1.2 
 'Re-establish a Connection and update 
 Set objRs1.ActiveConnection = GetNewConnection 
 objRs1.UpdateBatch 
'EndAddNew1.2 

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

