---
title: Row.BinaryToString Method (Outlook)
keywords: vbaol11.chm2243
f1_keywords:
- vbaol11.chm2243
ms.prod: outlook
api_name:
- Outlook.Row.BinaryToString
ms.assetid: 2416a69f-f0a2-b9a6-6f55-688dcf702824
ms.date: 06/08/2017
---


# Row.BinaryToString Method (Outlook)

Obtains a  **String** representing a value that has been converted from a binary value for the parent **[Row](row-object-outlook.md)** at the column specified by _Index_ .


## Syntax

 _expression_ . **BinaryToString**( **_Index_** )

 _expression_ A variable that represents a **Row** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|A 1-based index value that can be either a  **Long** representing the column index for the **[Columns](columns-object-outlook.md)** collection or a **String** representing the **[Name](column-name-property-outlook.md)** of the **Column** .|

### Return Value

A hexadecimal  **String** value that has been converted from a **PT_BINARY** value for the parent **Row** at the column specified by _Index_ . Returns the error, "Cannot convert the column specified by Index to String" if the value specified by _Index_ is not **PT_BINARY**.


## Remarks

Use the helper functions  **Row.BinaryToString** , **[Row.LocalTimeToUTC](row-localtimetoutc-method-outlook.md)** , and **[Row.UTCToLocalTime](row-utctolocaltime-method-outlook.md)** to facilitate type conversion of column values at a specific row. For more information on property value representation in a **[Table](table-object-outlook.md)** , see[Factors Affecting Property Value Representation in the Table and View Classes](http://msdn.microsoft.com/library/13cf9945-a9e0-bb32-a2cb-74366a365ae1%28Office.15%29.aspx).


## See also


#### Concepts


[Row Object](row-object-outlook.md)

