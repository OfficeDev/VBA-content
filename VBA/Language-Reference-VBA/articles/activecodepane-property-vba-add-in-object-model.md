---
title: ActiveCodePane Property (VBA Add-In Object Model)
keywords: vbob6.chm1070961
f1_keywords:
- vbob6.chm1070961
ms.prod: office
ms.assetid: 7c9839e2-e458-1dc5-f402-b05305503824
ms.date: 06/08/2017
---


# ActiveCodePane Property (VBA Add-In Object Model)



Returns the active or last active  **CodePane** object or sets the active **CodePane** object. Read/write.
 **Remarks**
You can set the  **ActiveCodePane** property to any valid **CodePane** object, as shown in the following example:



```vb
Set MyApp.VBE. ActiveCodePane = MyApp.VBE.CodePanes(1)

```

The preceding example sets the first [code pane](vbe-glossary.md) in a[collection](vbe-glossary.md) of code panes to be the active code pane. You can also activate a code pane using the **Set** method.

