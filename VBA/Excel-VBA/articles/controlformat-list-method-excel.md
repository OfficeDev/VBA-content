---
title: ControlFormat.List Method (Excel)
keywords: vbaxl10.chm630080
f1_keywords:
- vbaxl10.chm630080
ms.prod: excel
api_name:
- Excel.ControlFormat.List
ms.assetid: 8ec9abd2-d5cf-8179-96e9-a8b583bb8bcc
ms.date: 06/08/2017
---


# ControlFormat.List Method (Excel)

Returns or sets the text entries in the specified list box or a combo box, as an array of strings, or returns or sets a single text entry. An error occurs if there are no entries in the list.


## Syntax

 _expression_ . **List**( **_Index_** )

 _expression_ A variable that represents a **ControlFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The index number of a single text entry to be set or returned. If this argument is omitted, the entire list is returned or set as an array of strings.|

### Return Value

Variant


## Remarks

Setting this property clears any range specified by the  **[ListFillRange](controlformat-listfillrange-property-excel.md)** property.


## Example

This example sets the entries in a list box on worksheet one. If  `Shapes(2)` doesn?t represent a list box, this example fails.


```vb
Worksheets(1).Shapes(2).ControlFormat.List = _ 
 Array("cogs", "widgets", "sprockets", "gizmos")
```

This example sets entry four in a list box on worksheet one. If  `Shapes(2)` doesn?t represent a list box, this example fails.




```vb
Worksheets(1).Shapes(2).ControlFormat.List(4) = "gadgets"
```


## See also


#### Concepts


[ControlFormat Object](controlformat-object-excel.md)

