---
title: ContentControl.Delete Method (Word)
keywords: vbawd10.chm266534920
f1_keywords:
- vbawd10.chm266534920
ms.prod: word
api_name:
- Word.ContentControl.Delete
ms.assetid: 46fe3237-5d22-008e-3c2f-56a98f060723
ms.date: 06/08/2017
---


# ContentControl.Delete Method (Word)

Deletes the specified content control and the contents of the content control.


## Syntax

 _expression_ . **Delete**( **_DeleteContents_** )

 _expression_ An expression that returns a **ContentControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DeleteContents_|Optional| **Boolean**|Specifies whether to delete the contents of the content control.  **True** removes both the content control and its contents. **False** removes the control but leaves the contents of the content control in the active document. The default value is **False** .|

## Example

The following example removes all content controls and their contents from the active document.


```vb
Dim objCC As ContentControl 
 
Do While ActiveDocument.ContentControls.Count > 0 
 For Each objCC In ActiveDocument.ContentControls 
 objCC.Delete True 
 Next 
Loop
```


## See also


#### Concepts


[ContentControl Object](contentcontrol-object-word.md)

