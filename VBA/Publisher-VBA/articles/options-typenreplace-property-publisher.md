---
title: Options.TypeNReplace Property (Publisher)
keywords: vbapb10.chm1048626
f1_keywords:
- vbapb10.chm1048626
ms.prod: publisher
api_name:
- Publisher.Options.TypeNReplace
ms.assetid: 0eb378d2-3554-6a46-8b6b-4a990b4638db
ms.date: 06/08/2017
---


# Options.TypeNReplace Property (Publisher)

 **True** for Microsoft Publisher to replace unreadable Asian character clusters resulting from invalid keyboard sequences. Read/write **Boolean**.


## Syntax

 _expression_. **TypeNReplace**

 _expression_A variable that represents a  **Options** object.


### Return Value

Boolean


## Example

This example instructs Publisher to replace unreadable Asian character clusters resulting from invalid keyboard sequences.


```vb
Sub TypeReplace() 
 Options.TypeNReplace = True 
End Sub
```


