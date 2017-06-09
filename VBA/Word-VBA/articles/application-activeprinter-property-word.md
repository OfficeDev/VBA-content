---
title: Application.ActivePrinter Property (Word)
keywords: vbawd10.chm158335042
f1_keywords:
- vbawd10.chm158335042
ms.prod: word
api_name:
- Word.Application.ActivePrinter
ms.assetid: 835e350a-e069-e751-a7d7-1e9bb2483b4a
ms.date: 06/08/2017
---


# Application.ActivePrinter Property (Word)

Returns or sets the name of the active printer. Read/write  **String** .


## Syntax

 _expression_ . **ActivePrinter**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

Setting the printer using the  **ActivePrinter** property changes the default printer. For more information, see[Setting ActivePrinter Changes System Default Printer](http://go.microsoft.com/fwlink/?LinkId=61996) .


## Example

This example displays the name of the active printer.


```vb
MsgBox "The name of the active printer is " &; ActivePrinter
```

This example makes a network HP LaserJet IIISi printer the active printer.




```vb
Application.ActivePrinter = "HP LaserJet IIISi on \\printers\laser"
```

This example makes a local HP LaserJet 4 printer on LPT1 the active printer.




```vb
Application.ActivePrinter = "HP LaserJet 4 local on LPT1:"
```


## See also


#### Concepts


[Application Object](application-object-word.md)

