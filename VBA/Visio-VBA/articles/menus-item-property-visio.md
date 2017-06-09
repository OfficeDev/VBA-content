---
title: Menus.Item Property (Visio)
keywords: vis_sdr.chm13213765
f1_keywords:
- vis_sdr.chm13213765
ms.prod: visio
api_name:
- Visio.Menus.Item
ms.assetid: 6b09568f-4ae0-1818-b484-456749fe3676
ms.date: 06/08/2017
---


# Menus.Item Property (Visio)

Returns a  **Menu** object from the **Menus** collection. Read-only.


## Syntax

 _expression_ . **Item**( **_lIndex_** )

 _expression_ A variable that represents a **Menus** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _lIndex_|Required| **Long**|The index of the object to retrieve.|

### Return Value

Menu


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```
objRet = object(index )
```


