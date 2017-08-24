---
title: WebNavigationBarSet.Design Property (Publisher)
keywords: vbapb10.chm8519684
f1_keywords:
- vbapb10.chm8519684
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.Design
ms.assetid: 643d0b88-3b6d-65fd-7607-2f81c593a568
ms.date: 06/08/2017
---


# WebNavigationBarSet.Design Property (Publisher)

Sets or returns a  **PbWizardNavBarDesign** constant representing the design of the specified Web navigation bar set. Read/write.


## Syntax

 _expression_. **Design**

 _expression_A variable that represents a  **WebNavigationBarSet** object.


### Return Value

PbWizardNavBarDesign


## Remarks

The  **Design** property value can be one of the **[PbWizardNavBarDesign](pbwizardnavbardesign-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example adds a new Web navigation bar set to every page in the active document, sets the button style to large, and then sets the design property to  **pbnbDesignCapsule**.


```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets.AddSet(Name:="newNavBar") 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 .ButtonStyle = pbnbButtonStyleLarge 
 .Design = pbnbDesignCapsule 
End With
```


