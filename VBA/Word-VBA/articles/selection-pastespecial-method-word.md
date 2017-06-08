---
title: Selection.PasteSpecial Method (Word)
keywords: vbawd10.chm158662832
f1_keywords:
- vbawd10.chm158662832
ms.prod: word
api_name:
- Word.Selection.PasteSpecial
ms.assetid: 186ddf42-f8ab-e334-ccfe-245b2cc82224
ms.date: 06/08/2017
---


# Selection.PasteSpecial Method (Word)

Inserts the contents of the Clipboard.


## Syntax

 _expression_ . **PasteSpecial**( **_IconIndex_** , **_Link_** , **_Placement_** , **_DisplayAsIcon_** , **_DataType_** , **_IconFileName_** , **_IconLabel_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _IconIndex_|Optional| **Variant**|If DisplayAsIcon is  **True** , this argument is a number that corresponds to the icon you want to use in the program file specified by IconFilename. If this argument is omitted, this method uses the first (default) icon.|
| _Link_|Optional| **Variant**| **True** to create a link to the source file of the Clipboard contents. The default value is **False** .|
| _Placement_|Optional| **Variant**|Can be either of the  **[WdOLEPlacement](wdoleplacement-enumeration-word.md)** constants.|
| _DisplayAsIcon_|Optional| **Variant**| **True** to display the link as an icon. The default value is **False** .|
| _DataType_|Optional| **Variant**|A format for the Clipboard contents when they are inserted into the document.  **[WdPasteDataType](wdpastedatatype-enumeration-word.md)** .|
| _IconFileName_|Optional| **Variant**|If DisplayAsIcon is  **True** , this argument is the path and file name for the file in which the icon to be displayed is stored.|
| _IconLabel_|Optional| **Variant**|If DisplayAsIcon is  **True** , this argument is the text that appears below the icon.|

## Remarks

Unlike with the  **[Paste](selection-paste-method-word.md)** method, with **PasteSpecial** you can control the format of the pasted information and (optionally) establish a link to the source file (for example, a Microsoft Excel worksheet). If you do not want to replace the contents of the specified selection, use the **[Collapse](selection-collapse-method-word.md)** method before you use this method. When you use this method, the selection does not expand to include the contents of the Clipboard.


## Example

This example inserts the Clipboard contents at the insertion point as unformatted text.


```
Selection.Collapse Direction:=wdCollapseStart 
Selection.PasteSpecial DataType:=wdPasteText
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

