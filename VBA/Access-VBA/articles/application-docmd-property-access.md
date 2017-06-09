---
title: Application.DoCmd Property (Access)
keywords: vbaac10.chm12511
f1_keywords:
- vbaac10.chm12511
ms.prod: access
api_name:
- Access.Application.DoCmd
ms.assetid: 171fb56a-b39f-4439-e841-ae4bbbd71719
ms.date: 06/08/2017
---


# Application.DoCmd Property (Access)

You can use the  **DoCmd** property to access the read-only **[DoCmd](docmd-object-access.md)** object and its related methods. Read-only **DoCmd**.


## Syntax

 _expression_. **DoCmd**

 _expression_ A variable that represents an **Application** object.


## Example

The following example opens a form in Form view and moves to a new record.


```vb
Sub ShowNewRecord() 
 DoCmd.OpenForm "Employees", acNormal 
 DoCmd.GoToRecord , , acNewRec 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

