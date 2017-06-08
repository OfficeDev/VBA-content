---
title: WSParameters.Item Property (Access)
keywords: vbaac10.chm14579
f1_keywords:
- vbaac10.chm14579
ms.prod: access
api_name:
- Access.WSParameters.Item
ms.assetid: fe40b7f4-58e6-c632-0303-0925ab3a56c2
ms.date: 06/08/2017
---


# WSParameters.Item Property (Access)

The  **Item** property returns a specific member of a collection either by position or by index. Read-only **Object**.


## Syntax

 _expression_. **Item**( ** _Index_** )

 _expression_ A variable that represents a **WSParameters** object.


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


[WSParameters Collection](wsparameters-object-access.md)

