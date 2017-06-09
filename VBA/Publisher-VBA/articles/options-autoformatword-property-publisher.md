---
title: Options.AutoFormatWord Property (Publisher)
keywords: vbapb10.chm1048579
f1_keywords:
- vbapb10.chm1048579
ms.prod: publisher
api_name:
- Publisher.Options.AutoFormatWord
ms.assetid: b0466bd7-f0a1-44a8-480f-5d046e24e759
ms.date: 06/08/2017
---


# Options.AutoFormatWord Property (Publisher)

 **True** for Microsoft Publisher to automatically format the entire word at the cursor position, even when no text is selected. Read/write **Boolean**.


## Syntax

 _expression_. **AutoFormatWord**

 _expression_A variable that represents an  **Options** object.


### Return Value

Boolean


## Remarks

If only one or two characters in a word are selected, only the selected characters are affected by a formatting change, not the whole word.


## Example

This example sets global options for Microsoft Publisher, including enabling automatic formatting of the entire word.


```vb
Sub SetGlobalOptions() 
 With Options 
 .AutoFormatWord = True 
 .AutoKeyboardSwitching = True 
 .AutoSelectWord = True 
 .DragAndDropText = True 
 .UseCatalogAtStartup = False 
 .UseHelpfulMousePointers = False 
 End With 
End Sub
```


