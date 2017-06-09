---
title: Placeholders.FindByName Method (PowerPoint)
keywords: vbapp10.chm544004
f1_keywords:
- vbapp10.chm544004
ms.prod: powerpoint
api_name:
- PowerPoint.Placeholders.FindByName
ms.assetid: 8911f52e-b544-4246-8b75-8af3650da4de
ms.date: 06/08/2017
---


# Placeholders.FindByName Method (PowerPoint)

Finds the placeholder in the  **[Placeholders](placeholders-object-powerpoint.md)** collection at the specified index location or with the specified name.


## Syntax

 _expression_. **FindByName**( **_Index_** )

 _expression_ An expression that returns a **Placeholders** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The index of the placeholder to be found.|

### Return Value

Shape


## Remarks

Because it takes a  **Variant**, the **FindByName** method makes it possible to locate a member of the **Placeholders** collection by specifying either the index number (an **Integer** or **Long** ) or the name (a **String** ) of an individual placeholder. Unlike the corresponding methods of other collections, such as **[Shapes](shapes-object-powerpoint.md)** or **[Slides](slides-object-powerpoint.md)**, the **[Item](placeholders-item-method-powerpoint.md)** method of the **Placeholders** collection takes only a **Long**.


## Example

The following example shows how to use the  **FindByName** method to select the title placeholder in slide one in the active presentation.


```vb
Public Sub FindByName_Example()

    

    PowerPoint.ActivePresentation.Slides(1).Shapes.Placeholders.FindByName("Title 1").Select



End Sub
```


## See also


#### Concepts


[Placeholders Object](placeholders-object-powerpoint.md)

