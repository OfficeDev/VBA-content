---
title: View.PageMovementType Property (Word)
keywords: vbawd10.chm161808449
f1_keywords:
- vbawd10.chm161808449
ms.prod: word
api_name:
- Word.View.PageMovementType
ms.date: 08/15/2017
---

# View.PageMovementType Property (Word)

Returns or sets the page movement type. Read/write **[WdPageMovementType](wdpagemovementtype-enumeration-word.md)**.

## Syntax

 _expression_ .**PageMovementType**

 _expression_ Required. A variable that represents a **[View](view-object-word.md)** object.

## Example

This example sets the page movement type to side-to-side.

```vb
ActiveWindow.View.PageMovementType = wdSideToSide
```

## See also

#### Concepts

[View Object](view-object-word.md)