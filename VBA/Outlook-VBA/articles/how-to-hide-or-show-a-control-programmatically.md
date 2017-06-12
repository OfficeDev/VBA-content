---
title: "How to: Hide or Show a Control Programmatically"
keywords: olmain11.chm1045238
f1_keywords:
- olmain11.chm1045238
ms.prod: outlook
ms.assetid: c6cbadf7-7b10-81de-0abe-65b24c3f46d4
ms.date: 06/08/2017
---


# How to: Hide or Show a Control Programmatically

The following code example uses the  **[ModifiedFormPages](inspector-modifiedformpages-property-outlook.md)** property of the current **[Inspector](inspector-object-outlook.md)** object to set the Microsoft Forms 2.0 **Visible** property of a **[CheckBox](checkbox-object-outlook-forms-script.md)** on a page named "Test."


```vb
Item.GetInspector.ModifiedFormPages("Test").Checkbox1.Visible = False
```


