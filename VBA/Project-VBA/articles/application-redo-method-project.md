---
title: Application.Redo Method (Project)
keywords: vbapj.chm132540
f1_keywords:
- vbapj.chm132540
ms.prod: project-server
api_name:
- Project.Application.Redo
ms.assetid: 25a43bd7-4bfd-2be6-172d-8e5bef781f00
ms.date: 06/08/2017
---


# Application.Redo Method (Project)

Executes a redo action on items in the  **Redo** list.


## Syntax

 _expression_. **Redo**( ** _HowManyRedos_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _HowManyRedos_|Optional|**Long**|Specifies the number of items from the list to redo. The default is 1.|

### Return Value

 **Boolean**


## Remarks

You can add items to the  **Redo** list by using the **[Undo](application-undo-method-project.md)** method or clicking **Undo** in the Quick Access Toolbar.


