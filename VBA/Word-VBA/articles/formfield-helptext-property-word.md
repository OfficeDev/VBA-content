---
title: FormField.HelpText Property (Word)
keywords: vbawd10.chm153616391
f1_keywords:
- vbawd10.chm153616391
ms.prod: word
api_name:
- Word.FormField.HelpText
ms.assetid: 98069a1f-03eb-933b-9f7a-7d20cb83ce8c
ms.date: 06/08/2017
---


# FormField.HelpText Property (Word)

Returns or sets the text that's displayed in a message box when the form field has the focus and the user presses F1. Read/write  **String** .


## Syntax

 _expression_ . **HelpText**

 _expression_ A variable that represents a **[FormField](formfield-object-word.md)** object.


## Remarks

If the  **[OwnHelp](formfield-ownhelp-property-word.md)** property is set to **True** , **HelpText** specifies the text string value. If **OwnHelp** is set to **False** , **HelpText** specifies the name of an AutoText entry that contains help text for the form field.


## Example

This example sets the help text for the form field named "Name."


```vb
With ActiveDocument.FormFields("Name") 
 .OwnHelp = True 
 .HelpText = "Type your full legal name." 
End With
```


## See also


#### Concepts


[FormField Object](formfield-object-word.md)

