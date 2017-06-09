---
title: Application.Undo Method (Project)
keywords: vbapj.chm132718
f1_keywords:
- vbapj.chm132718
ms.prod: project-server
api_name:
- Project.Application.Undo
ms.assetid: 50e1b5ba-fe4b-d53d-5712-8e2023eb2755
ms.date: 06/08/2017
---


# Application.Undo Method (Project)

Executes an undo action on items in the  **Undo** list.


## Syntax

 _expression_. **Undo**( ** _HowManyUndos_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _HowManyUndos_|Optional|**Long**|Specifies the number of items from the list to undo. The default is 1.|

### Return Value

 **Boolean**


## Remarks

Many actions you perform in Project, such as adding a task, add items to the  **Undo** list. To redo one or more actions after using the **Undo** method, you can use the **[Redo](application-redo-method-project.md)** method or click **Redo** in the Quick Access Toolbar.


