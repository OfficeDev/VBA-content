---
title: Options.AddHebDoubleQuote Property (Publisher)
keywords: vbapb10.chm1048629
f1_keywords:
- vbapb10.chm1048629
ms.prod: publisher
api_name:
- Publisher.Options.AddHebDoubleQuote
ms.assetid: 9c71b52e-0273-7ca9-1f50-5beed65c2e73
ms.date: 06/08/2017
---


# Options.AddHebDoubleQuote Property (Publisher)

 **True** for Microsoft Publisher to display double quotes for Hebrew alphabet numbering. Default is **False**. Read/write  **Boolean**.


## Syntax

 _expression_. **AddHebDoubleQuote**

 _expression_A variable that represents an  **Options** object.


### Return Value

Boolean


## Remarks

This property is accessible only if Hebrew is enabled for Microsoft Office on your computer. 

This property applies only to Hebrew alphabet numbering.

As with all the properties of the  **[Options](options-object-publisher.md)** object, the current value of the **AddHebDoubleQuote** property becomes the default setting applied to all new publications.

This property corresponds to the  **Add double quotes for Hebrew alphabet numbering** check box on the **Bullets and Numbering** dialog box.


## Example

The following example sets Publisher to display double quotes for Hebrew alphabet numbering.


```vb
Publisher.Options.AddHebDoubleQuote = True
```


