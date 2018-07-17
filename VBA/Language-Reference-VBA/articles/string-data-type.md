---
title: String Data Type
keywords: vblr6.chm1009036
f1_keywords:
- vblr6.chm1009036
ms.prod: office
ms.assetid: 1c16e27a-ea31-cdbd-adbf-c9a7c81cc81c
ms.date: 06/08/2017
---


# String Data Type

There are two kinds of strings: variable-length and fixed-length strings.

- A variable-length string can contain up to approximately 2 billion (2^31) characters.
    
- A fixed-length string can contain 1 to approximately 64K (2^16) characters.
    
     **Note**  A [Public](vbe-glossary.md) fixed-length string can't be used in a [class module](vbe-glossary.md).

The codes for [String](vbe-glossary.md) characters range from 0-255. The first 128 characters (0-127) of the character set correspond to the letters and symbols on a standard U.S. keyboard. These first 128 characters are the same as those defined by the [ASCII](vbe-glossary.md) character set. The second 128 characters (128-255) represent special characters, such as letters in international alphabets, accents, currency symbols, and fractions.

The [type-declaration character](vbe-glossary.md) for **String** is the dollar sign ( **$** ).

A double-quotation-mark can be embedded within a [string literal](vbe-glossary.md#string-literal) in one of two ways:

- Use two double-quotation-marks:

        Dim s As String
        s = "This string literal has an embedded "" in it."

- Use the Chr function; character code 34 is a double-quotation-mark:

        Dim s As String
        s = "This string literal has an embedded " & Chr(34) & " in it."
