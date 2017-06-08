---
title: ODSOFilters.Item Method (Office)
keywords: vbaof11.chm241003
f1_keywords:
- vbaof11.chm241003
ms.prod: office
api_name:
- Office.ODSOFilters.Item
ms.assetid: eff21bc3-dc55-82a4-d405-2d4842c8bfa0
ms.date: 06/08/2017
---


# ODSOFilters.Item Method (Office)

Represents a  **ODSOFilter** object in the **ODSOFilters** collection.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ A variable that represents an **ODSOFilters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The number of the item.|

### Return Value

Object


## Example

The following example retrieves an  **ODSOFilter** object from the **ODSOFilters** collection.


```
oOdsoFilter = oOdsoFilters.Item(1)
```


## See also


#### Concepts


[ODSOFilters Object](odsofilters-object-office.md)
#### Other resources


[ODSOFilters Object Members](odsofilters-members-office.md)

