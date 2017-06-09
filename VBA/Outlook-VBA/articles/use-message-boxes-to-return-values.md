---
title: Use Message Boxes to Return Values
keywords: olfm10.chm3077357
f1_keywords:
- olfm10.chm3077357
ms.prod: outlook
ms.assetid: c63ad579-a2cd-ccc7-602c-7a83476d3060
ms.date: 06/08/2017
---


# Use Message Boxes to Return Values

One way to isolate errors is to use a message box to display the value of a variable or property at a particular point in the code. This code example shows the selection length returned from the  ** [TextBox.SelLength](olktextbox-sellength-property-outlook.md)** property in a message box.


```vb
MsgBox Item.GetInspector.ModifiedFormPages("Test").Textbox1.SelLength
```


