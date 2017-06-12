---
title: Toolbars.Item Property (Visio)
keywords: vis_sdr.chm13813765
f1_keywords:
- vis_sdr.chm13813765
ms.prod: visio
api_name:
- Visio.Toolbars.Item
ms.assetid: 0f56cab6-edcd-a153-f8a7-e6c3292cdfbb
ms.date: 06/08/2017
---


# Toolbars.Item Property (Visio)

Returns an object from a collection. The  **Item** property is the default property for all collections. Read-only.


## Syntax

 _expression_ . **Item**( **_lIndex_** )

 _expression_ A variable that represents a **Toolbars** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _lIndex_|Required| **Long**|Contains the index of the object to retrieve.|

### Return Value

Toolbar


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```
objRet = object(index )
```


