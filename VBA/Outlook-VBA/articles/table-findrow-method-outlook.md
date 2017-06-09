---
title: Table.FindRow Method (Outlook)
keywords: vbaol11.chm2228
f1_keywords:
- vbaol11.chm2228
ms.prod: outlook
api_name:
- Outlook.Table.FindRow
ms.assetid: 5722cf58-d026-007a-558f-90b73bad920d
ms.date: 06/08/2017
---


# Table.FindRow Method (Outlook)

Finds the first row in the  **[Table](table-object-outlook.md)** that meets the criteria specified in _Filter_ .


## Syntax

 _expression_ . **FindRow**( **_Filter_** )

 _expression_ A variable that represents a **Table** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filter_|Required| **String**|Specifies the condition that a row in the  **Table** has to meet.|

### Return Value

A  **[Row](row-object-outlook.md)** object that represents the first row in the **Table** that meets the filter criteria. Returns **Null** ( **Nothing** in Visual Basic) if no such row can be found, or the **Table** does not contain any rows.


## Remarks

 **FindRow** always starts from the first row in the **Table** .

 **FindRow** returns **Null** ( **Nothing** in Visual Basic) if a property in _Filter_ does not exist in the specified namespace. The property is considered a named property in the MAPI property set, **PS_PUBLIC_STRINGS** . **FindRow** does not return an error in this case.

 **FindRow** returns an error if _Filter_ is a blank string or an invalid restriction. In cases where **FindRow** does not find any row, the current row will not be repositioned to where it was before the call to **FindRow** .

To use content indexing search in a  **Table** , use the **[Restrict](table-restrict-method-outlook.md)** method. **FindRow** returns an error if _Filter_ contains content indexing keywords.


## See also


#### Concepts


[Table Object](table-object-outlook.md)

