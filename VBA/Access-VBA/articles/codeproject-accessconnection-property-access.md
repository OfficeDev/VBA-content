---
title: CodeProject.AccessConnection Property (Access)
keywords: vbaac10.chm12726
f1_keywords:
- vbaac10.chm12726
ms.prod: access
api_name:
- Access.CodeProject.AccessConnection
ms.assetid: 04b389d0-b87f-9eb9-f067-6b5e0d68e3f8
ms.date: 06/08/2017
---


# CodeProject.AccessConnection Property (Access)

You can use the  **AccessConnection** property to return a reference to the current Microsoft ActiveX Data Objects (ADO) **Connection** object and its related properties. Read-only **Connection**.


## Syntax

 _expression_. **AccessConnection**

 _expression_ A variable that represents a **CodeProject** object.


## Remarks

You should use the AccessConnection property if you intend to create ADO recordsets that will be bound to Access forms. The form will not be updateable unless it is created by using the OLE DB Provider for Microsoft Access, even if the recordset is updateable in ADO. 


## See also


#### Concepts


[CodeProject Object](codeproject-object-access.md)

