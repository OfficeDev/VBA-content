---
title: Operations.Item Property (Access)
keywords: vbaac10.chm14571
f1_keywords:
- vbaac10.chm14571
ms.prod: access
api_name:
- Access.Operations.Item
ms.assetid: 292f3492-ca44-21e3-245a-aaf0f9167e4d
ms.date: 06/08/2017
---


# Operations.Item Property (Access)

The  **Item** property returns a specific member of a collection either by position or by index. Read-only **Object**.


## Syntax

 _expression_. **Item**( ** _Index_** )

 _expression_ A variable that represents an **Operations** object.


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


[Operations Collection](operations-object-access.md)

