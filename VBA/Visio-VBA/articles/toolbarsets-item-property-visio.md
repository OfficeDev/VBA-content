---
title: ToolbarSets.Item Property (Visio)
keywords: vis_sdr.chm14013765
f1_keywords:
- vis_sdr.chm14013765
ms.prod: visio
api_name:
- Visio.ToolbarSets.Item
ms.assetid: d2880910-4e77-1c23-9785-727fa55afc98
ms.date: 06/08/2017
---


# ToolbarSets.Item Property (Visio)

Returns an object from a collection. The  **Item** property is the default property for all collections. Read-only.


## Syntax

 _expression_ . **Item**( **_lIndex_** )

 _expression_ A variable that represents a **ToolbarSets** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _lIndex_|Required| **Long**|Contains the index of the object to retrieve.|

### Return Value

ToolbarSet


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```
objRet = object(index )
```


