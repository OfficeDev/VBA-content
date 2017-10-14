---
title: AllowEditRanges.Add Method (Excel)
keywords: vbaxl10.chm724075
f1_keywords:
- vbaxl10.chm724075
ms.prod: excel
api_name:
- Excel.AllowEditRanges.Add
ms.assetid: f88d900d-4974-4d8d-6279-0be6376fc232
ms.date: 06/08/2017
---


# AllowEditRanges.Add Method (Excel)

Adds a range that can be edited on a protected worksheet. Returns a  **[AllowEditRange](alloweditrange-object-excel.md)** object.


## Syntax

 _expression_ . **Add**( **_Title_** , **_Range_** , **_Password_** )

 _expression_ A variable that represents an **AllowEditRanges** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Title_|Required| **String**|The title of range.|
| _Range_|Required| **Range**| **Range** object. The range allowed to be edited.|
| _Password_|Optional| **Variant**|The password for the range.|

### Return Value

An  **AllowEditRange** object that represents the range.


## Example

This example allows edits to range "A1:A4" on the active worksheet, notifies the user, then changes the password for this specified range and notifies the user of this change.


```vb
Sub UseChangePassword() 
 
 Dim wksOne As Worksheet 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Protect the worksheet. 
 wksOne.Protect 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:="secret" 
 
 MsgBox "Cells A1 to A4 can be edited on the protected worksheet." 
 
 ' Change the password. 
 wksOne.Protection.AllowEditRanges(1).ChangePassword _ 
 Password:="moresecret" 
 
 MsgBox "The password for these cells has been changed." 
 
End Sub
```


## See also


#### Concepts


[AllowEditRanges Collection](alloweditranges-object-excel.md)

