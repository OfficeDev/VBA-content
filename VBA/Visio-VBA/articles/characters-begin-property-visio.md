---
title: Characters.Begin Property (Visio)
keywords: vis_sdr.chm10213140
f1_keywords:
- vis_sdr.chm10213140
ms.prod: visio
api_name:
- Visio.Characters.Begin
ms.assetid: 885adb4d-aca8-b275-806b-34c76a14e7a7
ms.date: 06/08/2017
---


# Characters.Begin Property (Visio)

Gets or sets the beginning index of a  **Characters** object, which represents a range of text in a shape. Read/write.


## Syntax

 _expression_ . **Begin**

 _expression_ A variable that represents a **Characters** object.


### Return Value

Long


## Remarks

The  **Begin** property determines the beginning of the text range represented by a **Characters** object. The value of the **Begin** property is an index that represents the boundary between two characters, similar to an insertion point in text. Like selected text in a drawing window, a **Characters** object represents the sequence of characters that are affected by subsequent actions, such as the **Cut** or **Copy** method. When you retrieve a **Characters** object, its current text range includes all the shape's text. You can change the text range by setting the **Characters** object's **Begin** and **End** properties. Changing the text range of a **Characters** object has no effect on the text of the corresponding shape.

The  **Begin** property can have a value from zero (0) to the value of the **CharCount** property for the corresponding shape. An index of 0 is before the first character in the shape's text. An index that is the same as the **CharCount** property is after the last character in the shape's text. If you specify a value less than 0, Microsoft Visio uses 0. If you specify a value that is inside the expanded characters of a field, Visio sets the value of the **Begin** property to the start of the field.

The value of the  **Begin** property must always be less than or equal to the value of the **End** property. If you attempt to set the value of the **Begin** property to a value greater than the **End** property, Visio sets both the **Begin** and **End** properties to the value specified for the **Begin** property.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVCharacters.Begin**
    

