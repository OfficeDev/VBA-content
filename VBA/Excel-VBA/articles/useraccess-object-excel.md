---
title: UserAccess Object (Excel)
keywords: vbaxl10.chm727072
f1_keywords:
- vbaxl10.chm727072
ms.prod: excel
api_name:
- Excel.UserAccess
ms.assetid: 44df1865-a5f9-e1b7-b724-41d375e9ea44
ms.date: 06/08/2017
---


# UserAccess Object (Excel)

Represents the user access for a protected range.


## Example

Use the  **[Add](useraccesslist-add-method-excel.md)** method or the[Item](useraccesslist-item-property-excel.md) property of the[UserAccessList](useraccesslist-object-excel.md) collection to return a **UserAccess** object.



Once a  **UserAccess** object is returned, you can determine if access is allowed for a particular range in an worksheet, using the **[AllowEdit](useraccess-allowedit-property-excel.md)** property. The following example adds a range that can be edited on a protected worksheet and notifies the user the title of that range.




```vb
Sub UseAllowEditRanges() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Add a range that can be edited on the protected worksheet. 
 wksSheet.Protection.AllowEditRanges.Add "Test", Range("A1") 
 
 ' Notify the user the title of the range that can be edited. 
 MsgBox wksSheet.Protection.AllowEditRanges(1).Title 
 
End Sub
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

