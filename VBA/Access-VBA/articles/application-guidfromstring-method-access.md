---
title: Application.GUIDFromString Method (Access)
keywords: vbaac10.chm12558
f1_keywords:
- vbaac10.chm12558
ms.prod: access
api_name:
- Access.Application.GUIDFromString
ms.assetid: 943da2f6-a578-f05d-5778-990b6892fc64
ms.date: 06/08/2017
---


# Application.GUIDFromString Method (Access)

The  **GUIDFromString** function converts a string to a GUID, which is an array of type **Byte**.


## Syntax

 _expression_. **GUIDFromString**( ** _String_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _String_|Required|**Variant**|A string expression which evaluates to a GUID in string form.|

### Return Value

Variant


## Remarks

The Microsoft Access database engine stores GUIDs as arrays of type  **Byte**. However, Microsoft Access can't return **Byte** data from a control on a form or report. In order to return the value of a GUID from a control, you must convert it to a string. To convert a GUID to a string, use the **StringFromGUID** function. To convert a string to a GUID, use the **GUIDFromString** function.


## Example

The following example uses the  **GUIDFromString** function to convert a string to a GUID. The string is a GUID stored in string form in a replicated Employees table. The field, s_GUID, is a hidden field added to every replicated table in a replicated database.


```vb
Sub CheckGUIDType() 
 
 Dim dbsConn As ADODB.Connection 
 Dim rstEmployees As ADODB.Recordset 
 
 ' Make a connection to the current database. 
 Set dbsConn = Application.CurrentProject.Connection 
 Set rstEmployees = New ADODB.Recordset 
 rstEmployees.Open "Employees", dbsConn, , , adCmdTable 
 
 ' Print the GUID to the immediate window. 
 Debug.Print rst!s_GUID 
 Debug.Print TypeName(rst!s_GUID) 
 Debug.Print TypeName(GuidFromString(rst!s_GUID)) 
 
 Set rstEmployees = Nothing 
 Set dbsConn = Nothing 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

