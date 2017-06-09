---
title: Description Property (Visual Basic for Applications)
keywords: vblr6.chm1014191
f1_keywords:
- vblr6.chm1014191
ms.prod: office
ms.assetid: cab35a69-b45a-2d96-f495-2fae208fca6a
ms.date: 06/08/2017
---


# Description Property (Visual Basic for Applications)



Returns or sets a [string expression](vbe-glossary.md) containing a descriptive string associated with an object. Read/write.
For the  **Err** object, returns or sets a descriptive string associated with an error.
 **Remarks**
The  **Description** property setting consists of a short description of the error. Use this[property](vbe-glossary.md) to alert the user to an error that you either can't or don't want to handle. When generating a user-defined error, assign a short description of your error to the **Description** property. If **Description** isn't filled in, and the value of **Number** corresponds to a Visual Basic[run-time error](vbe-glossary.md), the string returned by the  **Error** function is placed in **Description** when the error is generated.

## Example

This example assigns a user-defined message to the  **Description** property of the **Err** object.


```
Err. Description = "It was not possible to access an object necessary " _
&; "for this operation."

```


