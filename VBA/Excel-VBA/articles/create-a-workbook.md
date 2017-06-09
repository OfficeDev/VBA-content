---
title: Create a Workbook
keywords: vbaxl10.chm5200215
f1_keywords:
- vbaxl10.chm5200215
ms.prod: excel
ms.assetid: b505b4bc-a3c3-3362-28cb-c119c2af5a3d
ms.date: 06/08/2017
---


# Create a Workbook

To create a workbook in Visual Basic, use the  **[Add](workbooks-add-method-excel.md)** method. The following procedure creates a workbook. Microsoft Excel automatically names the workbook Book _N_, where  _N_ is the next available number. The new workbook becomes the active workbook.


```vb
Sub AddOne() 
 Workbooks.Add 
End Sub
```


A better way to create a workbook is to assign it to an object variable. In the following example, the  **[Workbook](workbook-object-excel.md)** object returned by the  **Add** method is assigned to an object variable, `newBook`. Next, several properties of  `newBook` are set. You can easily control the new workbook by using the object variable.




```vb
Sub AddNew() 
Set NewBook = Workbooks.Add 
 With NewBook 
 .Title = "All Sales" 
 .Subject = "Sales" 
 .SaveAs Filename:="Allsales.xls" 
 End With 
End Sub
```


