---
title: SpinDown, SpinUp Events
keywords: fm20.chm2000220
f1_keywords:
- fm20.chm2000220
ms.prod: office
ms.assetid: 4e6e4395-1622-eb97-59d0-2b52a22d6528
ms.date: 06/08/2017
---


# SpinDown, SpinUp Events



SpinDown occurs when the user clicks the lower or left spin-button arrow. SpinUp occurs when the user clicks the upper or right spin-button arrow.
 **Syntax**
 **Private Sub**_object_ _**SpinDown( )**
 **Private Sub**_object_ _**SpinUp( )**
The  **SpinDown** and **SpinUp** event syntaxes have these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
 **Remarks**
The SpinDown event decreases the  **Value** property. The SpinUp event increases **Value**.

