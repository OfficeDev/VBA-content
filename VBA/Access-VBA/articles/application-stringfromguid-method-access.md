---
title: Application.StringFromGUID Method (Access)
keywords: vbaac10.chm12557
f1_keywords:
- vbaac10.chm12557
ms.prod: access
api_name:
- Access.Application.StringFromGUID
ms.assetid: 527c9459-a62a-9f01-dcda-3c21987b2662
ms.date: 06/08/2017
---


# Application.StringFromGUID Method (Access)

The  **StringFromGUID** function converts a GUID, which is an array of type **Byte**, to a string.


## Syntax

 _expression_. **StringFromGUID**( ** _Guid_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Guid_|Required|**Variant**|An array of  **Byte** data used to uniquely identify an application, component, or item of data to the operating system.|

### Return Value

Variant


## Remarks

The Microsoft Access database engine stores GUIDs as arrays of type  **Byte**. However, Microsoft Access can't return **Byte** data from a control on a form or report. In order to return the value of a GUID from a control, you must convert it to a string. To convert a GUID to a string, use the **StringFromGUID** function. To convert a string back to a GUID, use the **GUIDFromString** function.

For example, you may need to refer to a field that contains a GUID when using database replication. To return the value of a control on a form bound to a field that contains a GUID, use the  **StringFromGUID** function to convert the GUID to a string.


## Example

The following example returns the value of the s_GUID control on an Employees form in string form and assigns it to a string variable. The s_GUID control is bound to the s_GUID field, one of the system fields added to each replicated table in a replicated database.


```vb
Public Sub StringValueOfGUID() 
 
 Dim ctl As Control 
 Dim strGUID As String 
 
 ' Get the GUID. 
 Set ctl = Forms!Employees!s_GUID 
 Debug.Print TypeName(ctl.Value) 
 
 ' Convert the GUID to a string. 
 strGUID = StringFromGUID(ctl.Value) 
 Debug.Print TypeName(strGUID) 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

