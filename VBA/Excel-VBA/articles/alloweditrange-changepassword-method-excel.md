---
title: AllowEditRange.ChangePassword Method (Excel)
keywords: vbaxl10.chm725075
f1_keywords:
- vbaxl10.chm725075
ms.prod: excel
api_name:
- Excel.AllowEditRange.ChangePassword
ms.assetid: 1cc52121-f626-eaaa-9ea0-879634e34af7
ms.date: 06/08/2017
---


# AllowEditRange.ChangePassword Method (Excel)

Changes the password for a range that can be edited on a protected worksheet.


## Syntax

 _expression_ . **ChangePassword**( **_Password_** )

 _expression_ A variable that represents an **AllowEditRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Password_|Required| **String**|The new password.|

## Example

In this example, Microsoft Excel allows edits to range "A1:A4" on the active worksheet, notifies the user, changes the password for this specified range, and notifies the user of the change. The worksheet must be unprotected before running this code.


```vb
Sub UseChangePassword() 
 
 Dim wksOne As Worksheet 
 Dim strPassword As String 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 
 strPassword = InputBox("Please enter the password for the range") 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:=strPassword 
 
 strPassword = InputBox("Please enter the new password for the range") 
 
 ' Change the password. 
 wksOne.Protection.AllowEditRanges("Classified").ChangePassword _ 
 Password:="strPassword" 
 
 MsgBox "The password for these cells has been changed." 
 
End Sub
```


## See also


#### Concepts


[AllowEditRange Object](alloweditrange-object-excel.md)

