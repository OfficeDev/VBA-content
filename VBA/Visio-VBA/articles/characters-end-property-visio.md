---
title: Characters.End Property (Visio)
keywords: vis_sdr.chm10213460
f1_keywords:
- vis_sdr.chm10213460
ms.prod: visio
api_name:
- Visio.Characters.End
ms.assetid: 61b8fdb4-e00e-b7a5-2f0b-42d46684c626
ms.date: 06/08/2017
---


# Characters.End Property (Visio)

Returns or sets the ending index of the indicated  **Characters** object representing a range of text in a shape. Read/write.


## Syntax

 _expression_ . **End**

 _expression_ A variable that represents a **Characters** object.


### Return Value

Long


## Remarks

The  **End** property determines the end of the text range represented by a **Characters** object. The value of the **End** property is an index that represents the boundary between two characters, similar to an insertion point in text. Like selected text in a drawing window, a **Characters** object represents the sequence of characters that are affected by subsequent actions, such as the **Cut** or **Copy** method. When you retrieve a **Characters** object, its current text range includes all the shape's text. You can change the text range by setting the **Characters** object's **Begin** and **End** properties. Changing the text range of a **Characters** object has no effect on the text of the corresponding shape.

The  **End** property can have a value from zero (0) to the value of the **CharCount** property for the corresponding shape. An index of 0 is positioned before the first character in the shape's text. An index that is the same as the **CharCount** property is positioned after the last character in the shape's text. If you specify a value less than 0, Microsoft Visio uses 0. If you specify a value that is inside the expanded characters of a field, Visio sets the value of the **End** property to the end of the field.

The value of the  **End** property must always be greater than or equal to the value of the **Begin** property. If you attempt to set the value of the **End** property to a value lower than the **Begin** property, Visio sets both the **End** and **Begin** properties to the value specified for the **End** property.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVCharacters.End**
    

