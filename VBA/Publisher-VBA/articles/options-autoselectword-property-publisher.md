---
title: Options.AutoSelectWord Property (Publisher)
keywords: vbapb10.chm1048581
f1_keywords:
- vbapb10.chm1048581
ms.prod: publisher
api_name:
- Publisher.Options.AutoSelectWord
ms.assetid: 2b36f0d2-3260-aa3d-13b2-ae08b8d631d1
ms.date: 06/08/2017
---


# Options.AutoSelectWord Property (Publisher)

 **True** for Microsoft Publisher to automatically select the entire word when selecting text. Read/write **Boolean**.


## Syntax

 _expression_. **AutoSelectWord**

 _expression_A variable that represents an  **Options** object.


### Return Value

Boolean


## Example

This example sets Publisher global options, including enabling automatically selecting an entire word when selecting text.


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


