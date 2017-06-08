---
title: ContentControl.Checked Property (Word)
keywords: vbawd10.chm266534940
f1_keywords:
- vbawd10.chm266534940
ms.prod: word
api_name:
- Word.ContentControl.Checked
ms.assetid: 43315939-8ecb-788f-ddd5-3256cca5c9b6
ms.date: 06/08/2017
---


# ContentControl.Checked Property (Word)

Returns or sets a  **Boolean** that represents the current state (checked/unchecked) for a check box. Read/Write.


## Syntax

 _expression_ . **Checked**

 _expression_ An expression that returns a **[ContentControl](contentcontrol-object-word.md)** object.


## Remarks

Use the  **Checked** property to get/set the current state for a check box content control. If the control is not a check box, attempts to access the property will fail with the run-time error "This property is only available for check box content controls."


## Example

The following code example sets the specified check box content control  **Checked** property.


```vb
Dim objCC As ContentControl 
 
Set objCC = ActiveDocument.ContentControls.Add (wdContentControlCheckbox) 
objCC.Title = "Send Reminder" 
objCC.Checked = true 

```


## See also


#### Concepts


[ContentControl Object](contentcontrol-object-word.md)

