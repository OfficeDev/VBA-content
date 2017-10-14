---
title: WebNavigationBarSet.ButtonStyle Property (Publisher)
keywords: vbapb10.chm8519685
f1_keywords:
- vbapb10.chm8519685
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.ButtonStyle
ms.assetid: 39251032-d51e-3895-af18-cb4b613a38f4
ms.date: 06/08/2017
---


# WebNavigationBarSet.ButtonStyle Property (Publisher)

Sets or returns a  **PbWizardNavBarButtonStyle** constant that represents the style of the navigation bar buttons: large, small, or text-only. Read/write.


## Syntax

 _expression_. **ButtonStyle**

 _expression_A variable that represents a  **WebNavigationBarSet** object.


### Return Value

PbWizardNavBarButtonStyle


## Remarks

The  **ButtonStyle** property value can be one of the **[PbWizardNavBarButtonStyle](pbwizardnavbarbuttonstyle-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

The following example sets the button style to  **pbnbButtonStyleLarge** for the first Web navigation bar set of the active document.


```vb
ActiveDocument.WebNavigationBarSets(1).ButtonStyle = pbnbButtonStyleLarge
```


