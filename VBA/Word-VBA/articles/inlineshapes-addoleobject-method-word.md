---
title: InlineShapes.AddOLEObject Method (Word)
keywords: vbawd10.chm162070552
f1_keywords:
- vbawd10.chm162070552
ms.prod: word
api_name:
- Word.InlineShapes.AddOLEObject
ms.assetid: a7954807-160b-44de-1a50-c35a831d8651
ms.date: 06/08/2017
---


# InlineShapes.AddOLEObject Method (Word)

Creates an OLE object. Returns the  **[InlineShape](inlineshape-object-word.md)** object that represents the new OLE object.


## Syntax

 _expression_ . **AddOLEObject**( **_ClassType_** , **_FileName_** , **_LinkToFile_** , **_DisplayAsIcon_** , **_IconFileName_** , **_IconIndex_** , **_IconLabel_** , **_Range_** )

 _expression_ Required. A variable that represents an **[InlineShapes](inlineshapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ClassType_|Optional| **Variant**|The name of the application used to activate the specified OLE object.|
| _FileName_|Optional| **Variant**|The file from which the object is to be created. If this argument is omitted, the current folder is used. You must specify either the ClassType or FileName argument for the object, but not both.|
| _LinkToFile_|Optional| **Variant**| **True** to link the OLE object to the file from which it was created. **False** to make the OLE object an independent copy of the file. If you specified a value for ClassType, the LinkToFile argument must be **False** . The default value is **False** .|
| _DisplayAsIcon_|Optional| **Variant**| **True** to display the OLE object as an icon. The default value is **False** .|
| _IconFileName_|Optional| **Variant**|The file that contains the icon to be displayed.|
| _IconIndex_|Optional| **Variant**|The index number of the icon within IconFileName. The order of icons in the specified file corresponds to the order in which the icons appear in the  **Change Icon** dialog box ( **Insert** menu, **Object** dialog box) when the **Display as icon** check box is selected. The first icon in the file has the index number 0 (zero). If an icon with the given index number doesn't exist in IconFileName, the icon with the index number 1 (the second icon in the file) is used. The default value is 0 (zero).|
| _IconLabel_|Optional| **Variant**|A label (caption) to be displayed beneath the icon.|
| _Range_|Optional| **Variant**|The range where the OLE object will be placed in the text. The OLE object replaces the range, unless the range is collapsed. If this argument is omitted, the object is placed automatically.|

## Example

This example adds a new Microsoft Excel worksheet to the active document at the second paragraph.


```vb
ActiveDocument.InlineShapes.AddOLEObject _ 
 ClassType:="Excel.Sheet", DisplayAsIcon:=False, _ 
 Range:=ActiveDocument.Paragraphs(2).Range
```


## See also


#### Concepts


[InlineShapes Collection Object](inlineshapes-object-word.md)

