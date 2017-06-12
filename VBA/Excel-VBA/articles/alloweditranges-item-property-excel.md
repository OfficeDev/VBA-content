---
title: AllowEditRanges.Item Property (Excel)
keywords: vbaxl10.chm724074
f1_keywords:
- vbaxl10.chm724074
ms.prod: excel
api_name:
- Excel.AllowEditRanges.Item
ms.assetid: c6ac67af-258d-c2bf-3169-f42a5b037f2e
ms.date: 06/08/2017
---


# AllowEditRanges.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents an **AllowEditRanges** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

This example allows edits to range ("A1:A4") on the active worksheet, notifies the user, then changes the password for this specified range and notifies the user of this change.


```vb
Sub UseChangePassword() 
 
 Dim wksOne As Worksheet 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:="secret" 
 
 MsgBox "Cells A1 to A4 can be edited on the protected worksheet." 
 
 ' Change the password. 
 wksOne.Protection.AllowEditRanges.Item(1).ChangePassword _ 
 Password:="moresecret" 
 
 MsgBox "The password for these cells has been changed." 
 
End Sub
```


## See also


#### Concepts


[AllowEditRanges Collection](alloweditranges-object-excel.md)

