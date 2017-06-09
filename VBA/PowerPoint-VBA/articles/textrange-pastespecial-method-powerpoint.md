---
title: TextRange.PasteSpecial Method (PowerPoint)
keywords: vbapp10.chm569040
f1_keywords:
- vbapp10.chm569040
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.PasteSpecial
ms.assetid: 97bfd298-f8e8-32f0-b05c-6a93ed651954
ms.date: 06/08/2017
---


# TextRange.PasteSpecial Method (PowerPoint)

Replaces the text range with the contents of the Clipboard in the format specified. 


## Syntax

 _expression_. **PasteSpecial**( **_DataType_**, **_DisplayAsIcon_**, **_IconFileName_**, **_IconIndex_**, **_IconLabel_**, **_Link_** )

 _expression_ A variable that represents a **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DataType_|Optional|**PpPasteDataType**|A format for the Clipboard contents when they're inserted into the document. The default value varies, depending on the contents in the Clipboard. An error occurs if the specified data type in the DataType argument is not supported by the clipboard contents.|
| _DisplayAsIcon_|Optional|**MsoTriState**|**MsoTrue** to display the embedded object (or link) as an icon.|
| _IconFileName_|Optional|**String**|If DisplayAsIcon is set to  **msoTrue**, this argument is the path and file name for the file in which the icon to be displayed is stored. If DisplayAsIcon is set to **msoFalse**, this argument is ignored.|
| _IconIndex_|Optional|**Long**|If DisplayAsIcon is set to  **msoTrue**, this argument is a number that corresponds to the icon you want to use in the program file specified by IconFilename. For example, 0 (zero) corresponds to the first icon, 1 corresponds to the second icon. If this argument is omitted, the first (default) icon is used. If DisplayAsIcon is set to **msoFalse**, then this argument is ignored. If IconIndex is outside the valid range, then the default icon (index 0) is used.|
| _IconLabel_|Optional|**String**|If DisplayAsIcon is set to  **msoTrue**, this argument is the text that appears below the icon. If this label is missing, Microsoft PowerPoint generates an icon label based on the Clipboard contents. If DisplayAsIcon is set to **msoFalse**, then this argument is ignored.|
| _Link_|Optional|**MsoTriState**|Determines whether to create a link to the source file of the Clipboard contents. An error occurs if the Clipboard contents do not support a link.|

### Return Value

TextRange


## Remarks

Valid data types for the  **TextRange** object are **ppPasteText**, **ppPasteHTML**, and **ppPasteRTF** (any other format generates an error). If the paste succeeds, this method returns a **TextRange** object representing the text range that was pasted.

The  _DataType_ parameter value can be one of these **PpPasteDataType** constants.


||
|:-----|
|**ppPasteDefault**|
|**ppPasteHTML**|
|**ppPasteRTF**|
|**ppPasteText**|
The  _DisplayAsIcon_ parameter value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default. Does not display the embedded object (or link) as an icon.|
|**msoTrue**|Displays the embedded object (or link) as an icon.|
An error occurs if there is no data on the Clipboard when the  **PasteSpecial** method is called.


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

