---
title: Application.ClipboardFormats Property (Excel)
keywords: vbaxl10.chm133092
f1_keywords:
- vbaxl10.chm133092
ms.prod: excel
api_name:
- Excel.Application.ClipboardFormats
ms.assetid: 9b0de0b9-6acf-a73c-6d29-a405d0784170
ms.date: 06/08/2017
---


# Application.ClipboardFormats Property (Excel)

Returns the formats that are currently on the Clipboard, as an array of numeric values. To determine whether a particular format is on the Clipboard, compare each element in the array with the appropriate constant listed in the Remarks section. Read-only  **Variant** .


## Syntax

 _expression_ . **ClipboardFormats**( **_Index_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The array element to be returned. If this argument is omitted, the property returns the entire array of formats that are currently on the Clipboard. For more information, see the Remarks section.|

## Remarks

This property returns an array of numeric values. To determine whether a particular format is on the Clipboard compare each element of the array with one of the  **[XlClipboardFormat](xlclipboardformat-enumeration-excel.md)** constants.


## Example

This example displays a message box if the Clipboard contains a rich-text format (RTF) object. You can create an RTF object by copying text from a Word document.


```vb
aFmts = Application.ClipboardFormats 
For Each fmt In aFmts 
 If fmt = xlClipboardFormatRTF Then 
 MsgBox "Clipboard contains rich text" 
 End If 
Next
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

