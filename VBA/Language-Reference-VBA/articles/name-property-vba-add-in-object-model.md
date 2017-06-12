---
title: Name Property (VBA Add-In Object Model)
keywords: vbob6.chm1071231
f1_keywords:
- vbob6.chm1071231
ms.prod: office
ms.assetid: c393694c-f28f-acda-968c-f93defaad3c1
ms.date: 06/08/2017
---


# Name Property (VBA Add-In Object Model)



Returns or sets a [String](vbe-glossary.md) containing the name used in code to identify an object. For the **VBProject** object and the **VBComponent** object, read/write; for the **Property** object and the **Reference** object, read-only.
 **Remarks**
The following table describes how the  **Name** property setting applies to different objects.


|**Object**|**Result of Using Name Property Setting**|
|:-----|:-----|
|**VBProject**|Returns or sets the name of the active [project](vbe-glossary.md).|
|**VBComponent**|Returns or sets the name of the component. An error occurs if you try to set the  **Name** property to a name already being used or an invalid name.|
|**Property**|Returns the name of the property as it appears in the  **Property Browser**. This is the value used to index the **Properties**[collection](vbe-glossary.md). The name can't be set.|
|**Reference**|Returns the name of the reference in code. The name can't be set.|
The default name for new objects is the type of object plus a unique integer. For example, the first new Form object is Form1, a new Form object is Form1, and the third TextBox control you create on a form is TextBox3.
An object's  **Name** property must start with a letter and can be a maximum of 40 characters. It can include numbers and underline (_) characters but can't include punctuation or spaces.[Forms](vbe-glossary.md) and[modules](vbe-glossary.md) can't have the same name as another public object such as **Clipboard**, **Screen**, or **App**. Although the **Name** property setting can be a[keyword](vbe-glossary.md), property name, or the name of another object, this can create conflicts in your code.

