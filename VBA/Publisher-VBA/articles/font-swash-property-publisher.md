---
title: Font.Swash Property (Publisher)
keywords: vbapb10.chm5374005
f1_keywords:
- vbapb10.chm5374005
ms.prod: publisher
api_name:
- Publisher.Font.Swash
ms.assetid: 71537393-167a-f9e3-e3b3-ae743fdbb0ff
ms.date: 06/08/2017
---


# Font.Swash Property (Publisher)

Returns or sets an  **MsoTriState** constant that represents the state of the **Swash** property on the characters in a text range. The **Swash** property enables embellishments to the characters, often in the form of bigger and more flamboyant serifs. Read/write.


## Syntax

 _expression_. **Swash**

 _expression_A variable that represents a  **[Font](font-object-publisher.md)** object.


## Return Value

 **MsoTriState**


## Remarks


 **Note**  The  **Swash** property has an effect only for OpenType fonts that contain swashes.

The  **Swash** property value can be one of the following **MsoTriState** constants declared in the Microsoft Office type library.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|None of the characters in the range are formatted as swash.|
| **msoTriStateMixed**|Return value indicating that the range contains some text formatted as swash and some text not formatted as swash.|
| **msoTriStateToggle**|Set value that switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**|All characters in the range are formatted as swash.|

