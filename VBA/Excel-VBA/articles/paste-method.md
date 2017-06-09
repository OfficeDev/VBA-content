---
title: Paste Method
keywords: vbagr10.chm5207755
f1_keywords:
- vbagr10.chm5207755
ms.prod: excel
api_name:
- Excel.Paste
ms.assetid: 4cb4fa45-b319-f3a8-e477-80b96060905b
ms.date: 06/08/2017
---


# Paste Method

Pastes the contents of the Clipboard into the specified range on the datasheet.

 _expression_. **Paste**( **_Link_**)

 _expression_ Required. An expression that returns a **Range** object.

 **Link** Optional **Variant**.  **True** to establish a link to the source of the pasted data. The default value is **False**.

## Example

This example pastes the contents of the Clipboard into cell A1 on the datasheet.


```
myChart.Application.DataSheet.Range("A1").Paste
```


