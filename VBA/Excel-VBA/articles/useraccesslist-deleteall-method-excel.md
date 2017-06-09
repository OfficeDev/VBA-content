---
title: UserAccessList.DeleteAll Method (Excel)
keywords: vbaxl10.chm726076
f1_keywords:
- vbaxl10.chm726076
ms.prod: excel
api_name:
- Excel.UserAccessList.DeleteAll
ms.assetid: c162c9cf-8257-e97a-ebe8-ab1d700924ca
ms.date: 06/08/2017
---


# UserAccessList.DeleteAll Method (Excel)

Removes all users who have access to a protected range on a worksheet.


## Syntax

 _expression_ . **DeleteAll**

 _expression_ A variable that represents a **UserAccessList** object.


## Example

In this example, Microsoft Excel removes all users that have access to the first protected range on the active worksheet. This example assumes the worksheet is not protected.


```vb
Sub UseDeleteAll() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Remove all users with access to the first protected range. 
 wksSheet.Protection.AllowEditRanges(1).Users.DeleteAll 
 
End Sub
```


## See also


#### Concepts


[UserAccessList Object](useraccesslist-object-excel.md)

