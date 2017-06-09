---
title: CurrentProject.AccessConnection Property (Access)
keywords: vbaac10.chm12726
f1_keywords:
- vbaac10.chm12726
ms.prod: access
api_name:
- Access.CurrentProject.AccessConnection
ms.assetid: c2bf2846-c5ab-34a2-4b24-33c9cc9820c4
ms.date: 06/08/2017
---


# CurrentProject.AccessConnection Property (Access)

You can use the  **AccessConnection** property to return a reference to the current Microsoft ActiveX Data Objects (ADO) **Connection** object and its related properties. Read-only **Connection**.


## Syntax

 _expression_. **AccessConnection**

 _expression_ A variable that represents a **CurrentProject** object.


## Remarks

You should use the AccessConnection property if you intend to create ADO recordsets that will be bound to Access forms. The form will not be updateable unless it is created by using the OLE DB Provider for Microsoft Access, even if the recordset is updateable in ADO. 


## See also


#### Concepts


[CurrentProject Object](currentproject-object-access.md)

