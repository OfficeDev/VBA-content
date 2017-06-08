---
title: "How to: Change the Name of a Control"
keywords: olmain11.chm1045891
f1_keywords:
- olmain11.chm1045891
ms.prod: outlook
ms.assetid: cc3adf8b-526b-9eed-aabd-34991be2a85e
ms.date: 06/08/2017
---


# How to: Change the Name of a Control

The following code example uses the  **[ModifiedFormPages](inspector-modifiedformpages-property-outlook.md)** property of the current **[Inspector](inspector-object-outlook.md)** object to set the Microsoft Forms 2.0 **Name** property of a **[CheckBox](checkbox-object-outlook-forms-script.md)** on a page named "Test" to "Selection."


```
Item.GetInspector.ModifiedFormPages("Test").Checkbox1.Name = "Selection"
```


 **Note**  Each control should have a unique name.


