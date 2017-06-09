---
title: TextRange2.PasteSpecial Method (Office)
ms.prod: office
api_name:
- Office.TextRange2.PasteSpecial
ms.assetid: 79f88454-2f95-ea10-6ec4-5fb78ca8036d
ms.date: 06/08/2017
---


# TextRange2.PasteSpecial Method (Office)

Replaces the text range with the contents of the Clipboard in the format specified. If the paste succeeds, this method returns a  **TextRange2** object including the text range that was pasted.


## Syntax

 _expression_. **PasteSpecial**( **_Format_** )

 _expression_ An expression that returns a **TextRange2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Format_|Required|**MsoClipboardFormat**|Determines the format for the Clipboard contents when they're inserted into the document.|

### Return Value

TextRange2


## See also


#### Concepts


[TextRange2 Object](textrange2-object-office.md)
#### Other resources


[TextRange2 Object Members](textrange2-members-office.md)

