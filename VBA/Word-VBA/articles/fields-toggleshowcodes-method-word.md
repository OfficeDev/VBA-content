---
title: Fields.ToggleShowCodes Method (Word)
keywords: vbawd10.chm154140772
f1_keywords:
- vbawd10.chm154140772
ms.prod: word
api_name:
- Word.Fields.ToggleShowCodes
ms.assetid: 71f5aabf-7570-3594-d97c-de9cfcee0650
ms.date: 06/08/2017
---


# Fields.ToggleShowCodes Method (Word)

Switches the display of the fields between field codes and field results. Use the  **ShowCodes** property to control the display of an individual field.


## Syntax

 _expression_ . **ToggleShowCodes**

 _expression_ Required. A variable that represents a **[Fields](fields-object-word.md)** collection.


## Example

This example switches on or switches off the display of fields in the selection (the equivalent of pressing SHIFT+F9).


```
Selection.Fields.ToggleShowCodes
```

This example switches on or switches off the display of fields in the active document (the equivalent of pressing ALT+F9).




```vb
ActiveDocument.Fields.ToggleShowCodes
```


## See also


#### Concepts


[Fields Collection Object](fields-object-word.md)

