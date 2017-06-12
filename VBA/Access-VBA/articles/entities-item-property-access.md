---
title: Entities.Item Property (Access)
keywords: vbaac10.chm14562
f1_keywords:
- vbaac10.chm14562
ms.prod: access
api_name:
- Access.Entities.Item
ms.assetid: 6e8e9b66-35c9-d436-6391-df424ad0f66f
ms.date: 06/08/2017
---


# Entities.Item Property (Access)

The  **Item** property returns a specific member of a collection either by position or by index. Read-only **Object**.


## Syntax

 _expression_. **Item**( ** _Index_** )

 _expression_ A variable that represents an **Entities** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**||

## Remarks

If the value provided for the  _index_ argument doesn't match any existing member of the collection, an error occurs.

The  **Item** property is the default member of a collection, so you don't have to specify it explicitly. For example, the following two lines of code are equivalent:




```vb
Debug.Print Modules(0)
```




```vb
Debug.Print Modules.Item(0)
```


## See also


#### Concepts


[Entities Collection](entities-object-access.md)

