---
title: Chr Function
keywords: vblr6.chm1020927
f1_keywords:
- vblr6.chm1020927
ms.prod: office
ms.assetid: a9dc96ec-4719-8d24-144b-61d45fa58fe5
ms.date: 06/08/2017
---


# Chr Function



Returns a [String](vbe-glossary.md) containing the character associated with the specified[character code](vbe-glossary.md).
 **Syntax**
 **Chr(**_charcode_**)**
The required  _charcode_[argument](vbe-glossary.md) is a[Long](vbe-glossary.md) that identifies a character.
 **Remarks**
Numbers from 0 - 31 are the same as standard, nonprintable [ASCII](vbe-glossary.md) codes. For example, **Chr(** 10 **)** returns a linefeed character. The normal range for _charcode_ is 0 - 255. However, on[DBCS](vbe-glossary.md) systems, the actual range for _charcode_ is -32768 to 65535.

 **Note**  The  **ChrB** function is used with byte data contained in a **String**. Instead of returning a character, which may be one or two bytes, **ChrB** always returns a single byte. The **ChrW** function returns a **String** containing the[Unicode](vbe-glossary.md) character except on platforms where Unicode is not supported, in which case, the behavior is identical to the **Chr** function.


 **Note**  Visual Basic for the Macintosh does not support Unicode strings. Therefore,  **ChrW** ( _n_ ) cannot return all Unicode characters for n values in the range of 128 - 65,535, as it does in the Windows environment. Instead, **ChrW** ( _n_ ) attempts a "best guess" for Unicode values n greater than 127. Therefore, you should not use **ChrW** in the Macintosh environment.


## Example

This example uses the  **Chr** function to return the character associated with the specified character code.


```vb
Dim MyChar
MyChar = Chr(65)    ' Returns A.
MyChar = Chr(97)    ' Returns a.
MyChar = Chr(62)    ' Returns >.
MyChar = Chr(37)    ' Returns %.


```


