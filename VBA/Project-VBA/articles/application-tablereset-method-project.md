---
title: Application.TableReset Method (Project)
keywords: vbapj.chm404
f1_keywords:
- vbapj.chm404
ms.prod: project-server
api_name:
- Project.Application.TableReset
ms.assetid: 1db786fb-b79d-0404-fe39-4118e10f3cb4
ms.date: 06/08/2017
---


# Application.TableReset Method (Project)

Resets the active table to the default table definition.


## Syntax

 _expression_. **TableReset**

 _expression_ An expression that returns an **Application** object.


### Return Value

 **Boolean**


## Remarks

If the user has added or hidden columns, the  **TableReset** method changes the table in the active view back to the default definition. The **TableReset** method has the same effect as the **Reset to Default** command in the **Tables** drop-down menu on the **VIEW** ribbon.


 **Note**  When a column is added or hidden, the modified table shows in the  **Table Definition** dialog box when you edit the table.


