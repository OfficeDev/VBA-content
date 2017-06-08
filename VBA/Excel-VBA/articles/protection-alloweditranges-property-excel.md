---
title: Protection.AllowEditRanges Property (Excel)
keywords: vbaxl10.chm719084
f1_keywords:
- vbaxl10.chm719084
ms.prod: excel
api_name:
- Excel.Protection.AllowEditRanges
ms.assetid: 829ec57c-2fe1-27b0-5987-83bd4dd50eed
ms.date: 06/08/2017
---


# Protection.AllowEditRanges Property (Excel)

Returns an  **[AllowEditRanges](alloweditranges-object-excel.md)** object.


## Syntax

 _expression_ . **AllowEditRanges**

 _expression_ A variable that represents a **Protection** object.


## Example

In this example, Microsoft Excel allows edits to range A1:A4 on the active worksheet and notifies the user of the title and address of the specified range.


```vb
Sub UseAllowEditRanges() 
 
 Dim wksOne As Worksheet 
 Dim strPwd1 As String 
 
 Set wksOne = Application.ActiveSheet 
 
 strPwd1 = InputBox("Enter Password") 
 
 ' Unprotect worksheet. 
 wksOne.Unprotect 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:=strPwd1 
 
 ' Notify the user 
 ' the title and address of the range. 
 With wksOne.Protection.AllowEditRanges.Item(1) 
 MsgBox "Title of range: " &; .Title 
 MsgBox "Address of range: " &; .Range.Address 
 End With 
 
End Sub
```


## See also


#### Concepts


[Protection Object](protection-object-excel.md)

