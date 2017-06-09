---
title: Options.DisplayGridLines Property (Word)
keywords: vbawd10.chm162988338
f1_keywords:
- vbawd10.chm162988338
ms.prod: word
api_name:
- Word.Options.DisplayGridLines
ms.assetid: b4bb7db3-bdfb-74bb-891d-cd11c31d66ba
ms.date: 06/08/2017
---


# Options.DisplayGridLines Property (Word)

 **True** if Microsoft Word displays the document grid. This property is the equivalent of the **Gridlines** command on the **View** menu. Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayGridLines**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Remarks

This property affects only the document grid. For table gridlines, use the  **[TableGridlines](view-tablegridlines-property-word.md)** property.


## Example

This example switches between displaying and hiding the document grid in the active window.


```vb
Options.DisplayGridLines = Not Options.DisplayGridLines
```


## See also


#### Concepts


[Options Object](options-object-word.md)

