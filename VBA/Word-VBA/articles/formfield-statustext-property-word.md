---
title: FormField.StatusText Property (Word)
keywords: vbawd10.chm153616392
f1_keywords:
- vbawd10.chm153616392
ms.prod: word
api_name:
- Word.FormField.StatusText
ms.assetid: e374b94a-6faa-a2ea-9085-d9b987376fa8
ms.date: 06/08/2017
---


# FormField.StatusText Property (Word)

Returns or sets the text that is displayed in the status bar when a form field has the focus. Read/write  **String** .


## Syntax

 _expression_ . **StatusText**

 _expression_ A variable that represents a **[FormField](formfield-object-word.md)** object.


## Remarks

If the  **[OwnStatus](formfield-ownstatus-property-word.md)** property is set to **True** , the **StatusText** property specifies the status bar text. If the **OwnStatus** property is set to **False** , the **StatusText** property specifies the name of an AutoText entry that contains status bar text for the form field.


## Example

This example sets the status bar help text for the form field named "Age."


```vb
With ActiveDocument.FormFields("Age") 
 .OwnStatus = True 
 .StatusText = "Type your current age." 
End With
```


## See also


#### Concepts


[FormField Object](formfield-object-word.md)

