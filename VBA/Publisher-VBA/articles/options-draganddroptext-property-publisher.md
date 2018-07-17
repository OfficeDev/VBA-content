---
title: Options.DragAndDropText Property (Publisher)
keywords: vbapb10.chm1048584
f1_keywords:
- vbapb10.chm1048584
ms.prod: publisher
api_name:
- Publisher.Options.DragAndDropText
ms.assetid: 55fb68e8-4ddc-6866-00d8-bdd6a1e25ec3
ms.date: 06/08/2017
---


# Options.DragAndDropText Property (Publisher)

 **True** to enable dragging of text. Read/write **Boolean**.


## Syntax

 _expression_. **DragAndDropText**

 _expression_A variable that represents a  **Options** object.


### Return Value

Boolean


## Example

This example sets global options for Microsoft Publisher, including enabling dragging to reposition text.


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


