---
title: Decimal Data Type
keywords: vblr6.chm1099868
f1_keywords:
- vblr6.chm1099868
ms.prod: office
ms.assetid: 5f70e06b-61da-e0be-9f96-7dd84f377c74
ms.date: 06/08/2017
---


# Decimal Data Type

[Decimal variables](vbe-glossary.md) are stored as 96-bit (12-byte) signed integers scaled by a variable power of 10. The power of 10 scaling factor specifies the number of digits to the right of the decimal point, and ranges from 0 to 28. With a scale of 0 (no decimal places), the largest possible value is +/-79,228,162,514,264,337,593,543,950,335. With a 28 decimal places, the largest value is +/-7.9228162514264337593543950335 and the smallest, non-zero value is +/-0.0000000000000000000000000001.


 **Note**  At this time the  **Decimal** data type can only be used within a[Variant](vbe-glossary.md), that is, you cannot declare a variable to be of type  **Decimal**. You can, however, create a **Variant** whose subtype is **Decimal** using the **CDec** function.


