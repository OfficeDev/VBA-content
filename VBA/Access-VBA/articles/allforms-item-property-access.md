---
title: AllForms.Item Property (Access)
keywords: vbaac10.chm12680
f1_keywords:
- vbaac10.chm12680
ms.prod: access
api_name:
- Access.AllForms.Item
ms.assetid: 07536c98-57e1-8660-b35e-0e79e4e797cb
ms.date: 06/08/2017
---


# AllForms.Item Property (Access)

The  **Item** property returns a specific member of a collection either by position or by index. Read-only **AccessObject**.


## Syntax

 _expression_. **Item**( ** _var_** )

 _expression_ A variable that represents an **AllForms** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _var_|Required|**Variant**|An expression that specifies the position of a member of the collection referred to by the  _expression_ argument. If a numeric expression, the _index_ argument must be a number from 0 to the value of the collection's **Count** property minus 1. If a string expression, the _index_ argument must be the name of a member of the collection|

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


[AllForms Collection](allforms-object-access.md)

