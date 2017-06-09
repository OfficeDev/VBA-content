---
title: Options.PrintXMLTag Property (Word)
keywords: vbawd10.chm162988487
f1_keywords:
- vbawd10.chm162988487
ms.prod: word
api_name:
- Word.Options.PrintXMLTag
ms.assetid: f0fd4863-d57a-f1cb-f87d-b60190b8093e
ms.date: 06/08/2017
---


# Options.PrintXMLTag Property (Word)

Returns a  **Boolean** that represents whether to print the XML tags when printing a document. Corresponds to the **XML tags** check box on the **Print** tab in the **Options** dialog box. .


## Syntax

 _expression_ . **PrintXMLTag**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

 **True** indicates that tags are printed. **False** indicates tags are not printed.


## Example

The following example specifies that when documents are printed tags will also be printed.


```vb
Options.PrintXMLTag = True
```


## See also


#### Concepts


[Options Object](options-object-word.md)

