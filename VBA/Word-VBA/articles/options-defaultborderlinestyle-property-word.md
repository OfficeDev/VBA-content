---
title: Options.DefaultBorderLineStyle Property (Word)
keywords: vbawd10.chm162988307
f1_keywords:
- vbawd10.chm162988307
ms.prod: word
api_name:
- Word.Options.DefaultBorderLineStyle
ms.assetid: 677ffe8a-ca89-fd4e-158e-158bd4c98f0c
ms.date: 06/08/2017
---


# Options.DefaultBorderLineStyle Property (Word)

Returns or sets the default border line style. Read/write  **WdLineStyle** .


## Syntax

 _expression_ . **DefaultBorderLineStyle**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Example

This example sets the default line style to double.


```
Options.DefaultBorderLineStyle = wdLineStyleDouble
```

This example returns the current default line style.




```vb
Dim lngTemp As Long 
 
lngTemp= Options.DefaultBorderLineStyle
```


## See also


#### Concepts


[Options Object](options-object-word.md)

