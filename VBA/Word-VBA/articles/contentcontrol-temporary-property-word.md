---
title: ContentControl.Temporary Property (Word)
keywords: vbawd10.chm266534929
f1_keywords:
- vbawd10.chm266534929
ms.prod: word
api_name:
- Word.ContentControl.Temporary
ms.assetid: 66c1e5d6-9eb9-7d2e-dece-6b5c02373cb8
ms.date: 06/08/2017
---


# ContentControl.Temporary Property (Word)

Returns or sets a  **Boolean** that represents whether to remove a content control from the active document when the user edits the contents of the control. Read/write.


## Syntax

 _expression_ . **Temporary**

 _expression_ An expression that returns a **ContentControl** object.


## Remarks

The default value is  **False** . This property corresponds to the **Remove content control when contents are edited** check box in the **Content Control Properties** dialog box.


 **Note**  You cannot set this property if the  **[LockContentControl](contentcontrol-lockcontentcontrol-property-word.md)** property is set to **True** .


## See also


#### Concepts


[ContentControl Object](contentcontrol-object-word.md)

