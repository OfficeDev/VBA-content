---
title: Table.AutoFormat Method (Word)
keywords: vbawd10.chm156303374
f1_keywords:
- vbawd10.chm156303374
ms.prod: word
api_name:
- Word.Table.AutoFormat
ms.assetid: c76452fa-e1e8-3787-726a-b1c9967d96c2
ms.date: 06/08/2017
---


# Table.AutoFormat Method (Word)

Applies a predefined look to a table.


## Syntax

 _expression_ . **AutoFormat**( **_Format_** , **_ApplyBorders_** , **_ApplyShading_** , **_ApplyFont_** , **_ApplyColor_** , **_ApplyHeadingRows_** , **_ApplyLastRow_** , **_ApplyFirstColumn_** , **_ApplyLastColumn_** , **_AutoFit_** )

 _expression_ Required. A variable that represents a **[Table](table-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Format_|Optional| **Variant**|The format to apply. This parameter can be a  **[WdTableFormat](wdtableformat-enumeration-word.md)** constant, a **[WdTableFormatApply](wdtableformatapply-enumeration-word.md)** constant, or a **TableStyle** object.|
| _ApplyBorders_|Optional| **Variant**| **True** to apply the border properties of the specified format. The default value is **True** .|
| _ApplyShading_|Optional| **Variant**| **True** to apply the shading properties of the specified format. The default value is **True** .|
| _ApplyFont_|Optional| **Variant**| **True** to apply the font properties of the specified format. The default value is **True** .|
| _ApplyColor_|Optional| **Variant**| **True** to apply the color properties of the specified format. The default value is **True** .|
| _ApplyHeadingRows_|Optional| **Variant**| **True** to apply the heading-row properties of the specified format. The default value is **True** .|
| _ApplyLastRow_|Optional| **Variant**| **True** to apply the last-row properties of the specified format. The default value is **False** .|
| _ApplyFirstColumn_|Optional| **Variant**| **True** to apply the first-column properties of the specified format. The default value is **True** .|
| _ApplyLastColumn_|Optional| **Variant**| **True** to apply the last-column properties of the specified format. The default value is **False** .|
| _AutoFit_|Optional| **Variant**| **True** to decrease the width of the table columns as much as possible without changing the way text wraps in the cells. The default value is **True** .|

## Remarks

The arguments for this method correspond to the options in the  **Table AutoFormat** dialog box.


## Example

This example creates a 5x5 table in a new document and applies all the properties of the Colorful 2 format to the table.


```vb
Set newDoc = Documents.Add 
Set myTable = newDoc.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=5, NumColumns:=5) 
myTable.AutoFormat Format:=wdTableFormatColorful2
```

This example applies all the properties of the Classic 2 format to the table in which the insertion point is currently located. If the insertion point isn't in a table, a message box is displayed.




```vb
Selection.Collapse Direction:=wdCollapseStart 
If Selection.Information(wdWithInTable) = True Then 
 Selection.Tables(1).AutoFormat Format:=wdTableFormatClassic2 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Table Object](table-object-word.md)

