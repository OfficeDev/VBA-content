---
title: Range.ConvertToTable Method (Word)
keywords: vbawd10.chm157155826
f1_keywords:
- vbawd10.chm157155826
ms.prod: word
api_name:
- Word.Range.ConvertToTable
ms.assetid: a7d005ec-774e-151c-ff38-64df3ea36646
ms.date: 06/08/2017
---


# Range.ConvertToTable Method (Word)

Converts text within a range to a table. Returns the table as a  **Table** object.


## Syntax

 _expression_ . **ConvertToTable**( **_Separator_** , **_NumRows_** , **_NumColumns_** , **_InitialColumnWidth_** , **_Format_** , **_ApplyBorders_** , **_ApplyShading_** , **_ApplyFont_** , **_ApplyColor_** , **_ApplyHeadingRows_** , **_ApplyLastRow_** , **_ApplyFirstColumn_** , **_ApplyLastColumn_** , **_AutoFit_** , **_AutoFitBehavior_** , **_DefaultTableBehavior_** )

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Separator_|Optional| **Variant**|Specifies the character used to separate text into cells. Can be a character or one of the following  **WdTableFieldSeparator** constant. If this argument is omitted, the value of the **DefaultTableSeparator** property is used.|
| _NumRows_|Optional| **Variant**|The number of rows in the table. If this argument is omitted, Microsoft Word sets the number of rows, based on the contents of the range.|
| _NumColumns_|Optional| **Variant**|The number of columns in the table. If this argument is omitted, Word sets the number of columns, based on the contents of the range.|
| _InitialColumnWidth_|Optional| **Variant**|The initial width of each column, in points. If this argument is omitted, Word calculates and adjusts the column width so that the table stretches from margin to margin.|
| _Format_|Optional| **Variant**|Specifies one of the predefined formats listed in the  **Table AutoFormat** dialog box. Can be one of the **WdTableFormat** constants.|
| _ApplyBorders_|Optional| **Variant**| **True** to apply the border properties of the specified format.|
| _ApplyShading_|Optional| **Variant**| **True** to apply the shading properties of the specified format.|
| _ApplyFont_|Optional| **Variant**| **True** to apply the font properties of the specified format.|
| _ApplyColor_|Optional| **Variant**| **True** to apply the color properties of the specified format.|
| _ApplyHeadingRows_|Optional| **Variant**| **True** to apply the heading-row properties of the specified format.|
| _ApplyLastRow_|Optional| **Variant**| **True** to apply the last-row properties of the specified format.|
| _ApplyFirstColumn_|Optional| **Variant**| **True** to apply the first-column properties of the specified format.|
| _ApplyLastColumn_|Optional| **Variant**| **True** to apply the last-column properties of the specified format.|
| _AutoFit_|Optional| **Variant**| **True** to decrease the width of the table columns as much as possible without changing the way text wraps in the cells.|
| _AutoFitBehavior_|Optional| **Variant**|Sets the AutoFit rules for how Word sizes a table. Can be one of the following  **WdAutoFitBehavior** constant. If DefaultTableBehavior is **wdWord8TableBehavior** , this argument is ignored.|
| _DefaultTableBehavior_|Optional| **Variant**| Sets a value that specifies whether Microsoft Word automatically resizes cells in a table to fit the contents (AutoFit). Can be one of the **WdDefaultTableBehavior** constant.|

### Return Value

Table


## Example

This example converts the first three paragraphs in the active document to a table.


```vb
Set aDoc = ActiveDocument 
Set myRange = aDoc.Range(Start:=aDoc.Paragraphs(1).Range.Start, _ 
 End:=aDoc.Paragraphs(3).Range.End) 
myRange.ConvertToTable Separator:=wdSeparateByParagraphs
```


## See also


#### Concepts


[Range Object](range-object-word.md)

