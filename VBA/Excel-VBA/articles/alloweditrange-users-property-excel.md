---
title: AllowEditRange.Users Property (Excel)
keywords: vbaxl10.chm725078
f1_keywords:
- vbaxl10.chm725078
ms.prod: excel
api_name:
- Excel.AllowEditRange.Users
ms.assetid: 71f3c7ed-2fba-d97b-e443-674836e6bddb
ms.date: 06/08/2017
---


# AllowEditRange.Users Property (Excel)

Returns a  **[UserAccessList](useraccesslist-object-excel.md)** object for the protected range on a worksheet.


## Syntax

 _expression_ . **Users**

 _expression_ A variable that represents an **AllowEditRange** object.


## Example

In this example, Microsoft Excel displays the name of the first user allowed access to the first protected range on the active worksheet. This example assumes that a range has been chosen to be protected and that a particular user has been given access to this range.


```vb
Sub DisplayUserName() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Display name of user with access to protected range. 
 MsgBox wksSheet.Protection.AllowEditRanges(1).Users(1).Name 
 
End Sub
```


## See also


#### Concepts


[AllowEditRange Object](alloweditrange-object-excel.md)

