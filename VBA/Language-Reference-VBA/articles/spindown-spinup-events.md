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
 <strong>Syntax</strong>
 
<strong>Private Sub</strong><em>object</em> <em><strong>SpinDown( )</strong>
 <strong>Private Sub</strong>_object</em> _<strong>SpinUp( )</strong>
The  
<strong>SpinDown</strong> and <strong>SpinUp</strong> event syntaxes have these parts:


| <strong>Part</strong> | <strong>Description</strong> |
|:----------------------|:-----------------------------|
| <em>object</em>       | Required. A valid object.    |

 **Remarks**
The SpinDown event decreases the  **Value** property. The SpinUp event increases **Value**.

