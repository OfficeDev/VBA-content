---
title: Returning Strings from Functions
keywords: vbcn6.chm1009791
f1_keywords:
- vbcn6.chm1009791
ms.prod: office
ms.assetid: 7d344b4f-e262-7f3c-71e0-7e4a884db54e
ms.date: 06/08/2017
---


# Returning Strings from Functions

Some functions have two versions: one that returns a [Variant data type](vbe-glossary.md) and one that returns a [String data type](vbe-glossary.md). The  **Variant** versions are more convenient because variants handle conversions between different types of data automatically. They also allow [Null](vbe-glossary.md) to be propagated through an [expression](vbe-glossary.md). The  **String** versions are more efficient because they use less memory.

Consider using the  **String** version when:




- Your program is very large and uses many [variables](vbe-glossary.md).
    
- You write data directly to random-access files.
    

The following functions return values in a  **String** variable when you append a dollar sign ( **$** ) to the function name. These functions have the same usage and syntax as their **Variant** equivalents without the dollar sign.

|**Function**|||
|:-----|:-----|:-----|
|[Chr$](vbe-glossary.md)|[ChrB$](vbe-glossary.md)|*[Command$](vbe-glossary.md)|
|[CurDir$](vbe-glossary.md)|[Date$](vbe-glossary.md)|[Dir$](vbe-glossary.md)|
|[Error$](vbe-glossary.md)|[Format$](vbe-glossary.md)|[Hex$](vbe-glossary.md)|
|[Input$](vbe-glossary.md)|[InputB$](vbe-glossary.md)|[LCase$](vbe-glossary.md)|
|[Left$](vbe-glossary.md)|[LeftB$](vbe-glossary.md)|[LTrim$](vbe-glossary.md)|
|[Mid$](vbe-glossary.md)|[MidB$](vbe-glossary.md)|[Oct$](vbe-glossary.md)|
|[Right$](vbe-glossary.md)|[RightB$](vbe-glossary.md)|[RTrim$](vbe-glossary.md)|
|[Space$](vbe-glossary.md)|[Str$](vbe-glossary.md)|[String$](vbe-glossary.md)|
|[Time$](vbe-glossary.md)|[Trim$](vbe-glossary.md)|[UCase$](vbe-glossary.md)|


* May not be available in all applications.

