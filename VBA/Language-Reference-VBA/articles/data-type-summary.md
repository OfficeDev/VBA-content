---
title: Data Type Summary
keywords: vblr6.chm1008885
f1_keywords:
- vblr6.chm1008885
ms.prod: office
ms.assetid: 24723bdf-8454-f661-7914-d731e74d2e7b
ms.date: 06/08/2017
---


# Data Type Summary



The following table shows the supported [data types](vbe-glossary.md), including storage sizes and ranges.


|**Data type**|**Storage size**|**Range**|
|:-----|:-----|:-----|
|**Byte**|1 byte|0 to 255|
|**Boolean**|2 bytes|**True** or **False**|
|**Integer**|2 bytes|-32,768 to 32,767|
|**Long**(long integer)|4 bytes|-2,147,483,648 to 2,147,483,647|
|**LongLong**(LongLong integer)|8 bytes|-9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 (Valid on 64-bit platforms only.)|
|**LongPtr**(Long integer on 32-bit systems, LongLong integer on 64-bit systems)|4 bytes on 32-bit systems, 8 bytes on 64-bit systems|-2,147,483,648 to 2,147,483,647 on 32-bit systems, -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 on 64-bit systems|
|**Single**(single-precision floating-point)|4 bytes|-3.402823E38 to -1.401298E-45 for negative values; 1.401298E-45 to 3.402823E38 for positive values|
|**Double**(double-precision floating-point)|8 bytes|-1.79769313486231E308 to-4.94065645841247E-324 for negative values; 4.94065645841247E-324 to 1.79769313486232E308 for positive values|
|**Currency**(scaled integer)|8 bytes|-922,337,203,685,477.5808 to 922,337,203,685,477.5807|
|**Decimal**|14 bytes|+/-79,228,162,514,264,337,593,543,950,335 with no decimal point;+/-7.9228162514264337593543950335 with 28 places to the right of the decimal; smallest non-zero number is+/-0.0000000000000000000000000001|
|**Date**|8 bytes|January 1, 100 to December 31, 9999|
|**Object**|4 bytes|Any **Object** reference|
|**String**(variable-length)|10 bytes + string length|0 to approximately 2 billion|
|**String**(fixed-length)|Length of string|1 to approximately 65,400|
|**Variant**(with numbers)|16 bytes|Any numeric value up to the range of a **Double**|
|**Variant**(with characters)|22 bytes + string length (24 bytes on 64-bit systems)|Same range as for variable-length **String**|
|User-defined(using **Type** )|Number required by elements|The range of each element is the same as the range of its data type.|

 **Note** [Arrays](vbe-glossary.md) of any data type require 20 bytes of memory plus 4 bytes for each array dimension plus the number of bytes occupied by the data itself. The memory occupied by the data can be calculated by multiplying the number of data elements by the size of each element. For example, the data in a single-dimension array consisting of 4 **Integer** data elements of 2 bytes each occupies 8 bytes. The 8 bytes required for the data plus the 24 bytes of overhead brings the total memory requirement for the array to 32 bytes. On 64-bit platforms, SAFEARRAY's take up 24-bits (plus 4 bytes per Dim statement). The pvData member is an 8-byte pointer and it must be aligned on 8 byte boundaries.


 **Note** [LongPtr](longptr-data-type.md) is not a true data type because it transforms to a [Long](long-data-type.md) in 32-bit environments, or a [LongLong](longlong-data-type.md) in 64-bit environments. **LongPtr** should be used to represent pointer and handle values in [Declare Statements](declare-statement.md) and enables writing portable code that can run in both 32-bit and 64-bit environments.

A **Variant** containing an array requires 12 bytes more than the array alone.

 **Note** Use the **StrConv** function to convert one type of string data to another.


