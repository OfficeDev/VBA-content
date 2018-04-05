---
title: Shapes.PasteSpecial Method (PowerPoint)
keywords: vbapp10.chm543028
f1_keywords:
- vbapp10.chm543028
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.PasteSpecial
ms.assetid: 6a1e5b6d-da09-fae8-7165-0c9bf71d525c
ms.date: 11/24/2017
---


# Shapes.PasteSpecial Method (PowerPoint)

Pastes the contents of the Clipboard, using a special format.


## Syntax

 _expression_. **PasteSpecial**(**_DataType_**, **_DisplayAsIcon_**, **_IconFileName_**, **_IconIndex_**, **_IconLabel_**, **_Link_**)

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataType_|Optional|**PpPasteDataType**|A format for the Clipboard contents when they're inserted into the document. The default value varies, depending on the contents in the Clipboard. An error occurs if the specified data type in the DataType argument is not supported by the clipboard contents.|
| _DisplayAsIcon_|Optional|**MsoTriState**|**MsoTrue** to display the embedded object (or link) as an icon.|
| _IconFileName_|Optional|**String**|If DisplayAsIcon is set to  **msoTrue**, this argument is the path and file name for the file in which the icon to be displayed is stored. If DisplayAsIcon is set to **msoFalse**, this argument is ignored.|
| _IconIndex_|Optional|**Long**|If DisplayAsIcon is set to  **msoTrue**, this argument is a number that corresponds to the icon you want to use in the program file specified by IconFilename. For example, 0 (zero) corresponds to the first icon, and 1 corresponds to the second icon. If this argument is omitted, the first (default) icon is used. If DisplayAsIcon is set to **msoFalse**, then this argument is ignored. If IconIndex is outside the valid range, then the default icon (index 0) is used.|
| _IconLabel_|Optional|**String**|If DisplayAsIcon is set to  **msoTrue**, this argument is the text that appears below the icon. If this label is missing, Microsoft PowerPoint generates an icon label based on the Clipboard contents. If DisplayAsIcon is set to **msoFalse**, then this argument is ignored.|
| _Link_|Optional|**MsoTriState**|Determines whether to create a link to the source file of the Clipboard contents. An error occurs if the Clipboard contents do not support a link.|

### Return Value

ShapeRange


## Remarks

Adds the shape to the collection of shapes in the specified format. If the specified data type is a text data type, a new text box is created with the text. If the paste succeeds, the **PasteSpecial** method returns a **[ShapeRange](shaperange-object-powerpoint.md)** object representing the shape range that was pasted.

The _DataType_ parameter value can be one of these **PpPasteDataType** constants:

- **ppPasteBitmap**
- **ppPasteDefault**
- **ppPasteEnhancedMetafile**
- **ppPasteHTML**
- **ppPasteGIF**
- **ppPasteJPG**
- **ppPasteMetafilePicture**
- **ppPastePNG**
- **ppPasteShape**

<br/>

The _DisplayAsIcon_ parameter value can be one of these **MsoTriState** constants.

|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default. Does not display the embedded object (or link) as an icon.|
|**msoTrue**|Displays the embedded object (or link) as an icon.|


An error occurs if there is no data on the Clipboard when the **PasteSpecial** method is called.

## See also

[Shapes Object](shapes-object-powerpoint.md)

