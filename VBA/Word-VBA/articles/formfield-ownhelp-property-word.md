---
title: FormField.OwnHelp Property (Word)
keywords: vbawd10.chm153616389
f1_keywords:
- vbawd10.chm153616389
ms.prod: word
api_name:
- Word.FormField.OwnHelp
ms.assetid: a066ffc1-89d3-12d4-0bf1-bf338679d2d4
ms.date: 06/08/2017
---


# FormField.OwnHelp Property (Word)

Specifies the source of the text that's displayed in a message box when a form field has the focus and the user presses F1. Read/write  **Boolean** .


## Syntax

 _expression_ . **OwnHelp**

 _expression_ An expression that returns a **[FormField](formfield-object-word.md)** object.


## Remarks

If  **True** , the text specified by the **[HelpText](formfield-helptext-property-word.md)** property is displayed. If **False** , the text in the AutoText entry specified by the **HelpText** property is displayed.


## Example

This example sets the help text for the first form field in the current section to the contents of the AutoText entry named "Sample."


```vb
With Selection.Sections(1).Range.FormFields(1) 
 .OwnHelp = False 
 .HelpText = "Sample" 
End With
```


## See also


#### Concepts


[FormField Object](formfield-object-word.md)

