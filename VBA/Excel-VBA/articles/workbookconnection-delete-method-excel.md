---
title: WorkbookConnection.Delete Method (Excel)
keywords: vbaxl10.chm774080
f1_keywords:
- vbaxl10.chm774080
ms.prod: excel
api_name:
- Excel.WorkbookConnection.Delete
ms.assetid: d1312b91-04d7-2695-0c20-c18a31776fb0
ms.date: 06/08/2017
---


# WorkbookConnection.Delete Method (Excel)

Deletes a workbook connection.


## Syntax

 _expression_ . **Delete**

 _expression_ A variable that represents a **WorkbookConnection** object.


## Remarks

Use this method to delete an external data connection. This method does not apply to links to other workbooks. 

Deleting a connection will not delete or remove any objects that were using that connection. Deleting a connection will not cause any of the connection files to be deleted from the file system. If you edit any of those objects to use another connection, everything will start working again.

Objects that use a deleted connection behave as if the connection could not be established. 


## See also


#### Concepts


[WorkbookConnection Object](workbookconnection-object-excel.md)

