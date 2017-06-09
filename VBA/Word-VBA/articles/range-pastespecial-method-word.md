---
title: Range.PasteSpecial Method (Word)
keywords: vbawd10.chm157155504
f1_keywords:
- vbawd10.chm157155504
ms.prod: word
api_name:
- Word.Range.PasteSpecial
ms.assetid: 76d074ee-f0d8-8bdd-e7c2-d0aa7b5f6702
ms.date: 06/08/2017
---


# Range.PasteSpecial Method (Word)

Inserts the contents of the Clipboard. .


## Syntax

 _expression_ . **PasteSpecial**( **_IconIndex_** , **_Link_** , **_Placement_** , **_DisplayAsIcon_** , **_DataType_** , **_IconFileName_** , **_IconLabel_** )

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _IconIndex_|Optional| **Variant**|If DisplayAsIcon is  **True** , this argument is a number that corresponds to the icon you want to use in the program file specified by IconFilename. Icons appear in the **Change Icon** dialog box: 0 (zero) corresponds to the first icon, 1 corresponds to the second icon, and so on. If this argument is omitted, the first (default) icon is used.|
| _Link_|Optional| **Variant**| **True** to create a link to the source file of the Clipboard contents. The default value is **False** .|
| _Placement_|Optional| **Variant**|Can be either of the following  **WdOLEPlacement** constants: **wdFloatOverText** or **wdInLine** . The default value is **wdInLine** .|
| _DisplayAsIcon_|Optional| **Variant**| **True** to display the link as an icon. The default value is **False** .|
| _DataType_|Optional| **Variant**|A format for the Clipboard contents when they're inserted into the document. Can be any  **WdPasteDataType** constant.|
| _IconFileName_|Optional| **Variant**|If DisplayAsIcon is  **True** , this argument is the path and file name for the file in which the icon to be displayed is stored.|
| _IconLabel_|Optional| **Variant**|If DisplayAsIcon is  **True** , this argument is the text that appears below the icon.|

## Example

This example inserts the Clipboard contents at the insertion point as unformatted text.


```
Selection.Collapse Direction:=wdCollapseStart 
Selection.Range.PasteSpecial DataType:=wdPasteText
```

This example copies the selected text and pastes it into a new document as a hyperlink. The source document must first be saved for this example to work.




```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Copy 
 Documents.Add.Content.PasteSpecial Link:=True, _ 
 DataType:=wdPasteHyperlink 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

