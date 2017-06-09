---
title: Row.LocalTimeToUTC Method (Outlook)
keywords: vbaol11.chm2246
f1_keywords:
- vbaol11.chm2246
ms.prod: outlook
api_name:
- Outlook.Row.LocalTimeToUTC
ms.assetid: 10e24b21-8fd5-8740-b120-a49340cb9670
ms.date: 06/08/2017
---


# Row.LocalTimeToUTC Method (Outlook)

Obtains a  **Date** value in a **[Table](table-object-outlook.md)** specified by the **[Row](row-object-outlook.md)** object at _Index_ , that has been converted from local time to Coordinated Universal Time (UTC).


## Syntax

 _expression_ . **LocalTimeToUTC**( **_Index_** )

 _expression_ A variable that represents a **Row** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|A 1-based index value that can be either a  **Long** representing the column index for the **[Columns](columns-object-outlook.md)** collection or a **String** representing the **[Name](column-name-property-outlook.md)** of the **[Column](column-object-outlook.md)** .|

### Return Value

A  **Date** value that has been converted from a representation in local time to UTC. An error is returned if _Index_ is invalid or the row value indicated by _Index_ is not a **Date** value.


## Remarks

Use the helper functions  **[Row.BinaryToString](row-binarytostring-method-outlook.md)** , **Row.LocalTimeToUTC** , and **[Row.UTCToLocalTime](row-utctolocaltime-method-outlook.md)** to facilitate type conversion of column values at a specific row.

For information on property value representation in a  **Table** , see[Factors Affecting Property Value Representation in the Table and View Classes](http://msdn.microsoft.com/library/13cf9945-a9e0-bb32-a2cb-74366a365ae1%28Office.15%29.aspx). For information on using Date-time comparisons in  **Table** filters, see[Filtering Items Using a Date-time Comparison](http://msdn.microsoft.com/library/668e0993-c3d2-835f-0645-ba79bcffe67f%28Office.15%29.aspx).


## See also


#### Concepts


[Row Object](row-object-outlook.md)

