---
title: Shape.CharCount Property (Visio)
keywords: vis_sdr.chm11213220
f1_keywords:
- vis_sdr.chm11213220
ms.prod: visio
api_name:
- Visio.Shape.CharCount
ms.assetid: 2da9c359-d86c-bdf6-3553-01ded11d9208
ms.date: 06/08/2017
---


# Shape.CharCount Property (Visio)

Returns the number of characters in an object. Read-only.


## Syntax

 _expression_ . **CharCount**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Long


## Remarks

For a  **Shape** object, the **CharCount** property returns the number of characters in the shape's text. For a **Characters** object, the **CharCount** property returns the number of characters in the text range represented by that object.

The value returned by the  **CharCount** property includes the expanded number of characters for any fields in the object's text. For example, if the text contains a field that displays the file name of a drawing, the **CharCount** property includes the number of characters in the file name, rather than the one-character escape sequence used to represent a field in the **Text** property of a **Shape** object.


