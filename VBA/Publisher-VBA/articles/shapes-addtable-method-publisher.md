---
title: Shapes.AddTable Method (Publisher)
keywords: vbapb10.chm2162713
f1_keywords:
- vbapb10.chm2162713
ms.prod: publisher
api_name:
- Publisher.Shapes.AddTable
ms.assetid: 1aa00f40-de41-12ed-8d4f-5e9c91cbf5af
ms.date: 06/08/2017
---


# Shapes.AddTable Method (Publisher)

Adds a new  **Shape** object representing a table to the specified **Shapes** collection.


## Syntax

 _expression_. **AddTable**( **_NumRows_**,  **_NumColumns_**,  **_Left_**,  **_Top_**,  **_Width_**,  **_Height_**,  **_FixedSize_**,  **_Direction_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|NumRows|Required| **Long**|The number of rows in the new table. Values between 1 and 128 are valid; any values outside this range will generate an error.|
|NumColumns|Required| **Long**|The number of columns in the new table. Values between 1 and 128 are valid; any values outside this range will generate an error.|
|Left|Required| **Variant**|The position of the left edge of the shape representing the table.|
|Top|Required| **Variant**|The position of the top edge of the shape representing the table.|
|Width|Required| **Variant**|The width of the shape representing the table.|
|Height|Required| **Variant**|The height of the shape representing the table.|
|FixedSize|Optional| **Boolean**| **True** if Microsoft Publisher reduces the number of rows and columns of the table to fit the specified width and height. **False** if Publisher automatically increases the width and height of the table frame to accommodate the number of rows and columns in the table. Default is **False**.|
|Direction|Optional| **PbTableDirectionType**|The direction in which table columns are numbered. The default depends on the current language setting.|

### Return Value

Shape


## Remarks

For the Left, Top, Width, and Height arguments, numeric values are evaluated in points; strings can be in any units supported by Publisher (for example, "2.5 in").

The Direction parameter can be one of the  **PbTableDirectionType** constants declared in the Microsoft Publisher type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **pbTableDirectionLeftToRight**|Table columns are numbered from left to right. Default for left-to-right languages.|
| **pbTableDirectionRightToLeft**|Table columns are numbered from right to left. Default for right-to-left languages.|

## Example

This example creates a new table on the first page of the active publication.


```vb
Dim shpTable As Shape 
 
Set shpTable = ActiveDocument.Pages(1).Shapes.AddTable _ 
 (NumRows:=3, NumColumns:=4, _ 
 Left:=10, Top:=10, _ 
 Width:=288, Height:=216) 

```


