---
title: ContentControl.Copy Method (Word)
keywords: vbawd10.chm266534918
f1_keywords:
- vbawd10.chm266534918
ms.prod: word
api_name:
- Word.ContentControl.Copy
ms.assetid: ce3ba4ce-aef7-cb7d-ec7b-a160155a939d
ms.date: 06/08/2017
---


# ContentControl.Copy Method (Word)

Copies the content control from the active document to the Clipboard.


## Syntax

 _expression_ . **Copy**

 _expression_ An expression that returns a **ContentControl** object.


## Remarks

When you use the  **Copy** method, the original content control remains in the active document, but a copy of the control, including all text and property settings, is moved to the Clipboard. You can then paste the content control into other sections of the active document. Use the **Paste** method of the **[Selection](selection-object-word.md)** object or the **Paste** method of the **[Range](range-object-word.md)** object to insert the copied content control, or use the **Paste** function from within Microsoft Word.


## See also


#### Concepts


[ContentControl Object](contentcontrol-object-word.md)

