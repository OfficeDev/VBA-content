---
title: Making Faster For...Next Loops
keywords: vbcn6.chm1009794
f1_keywords:
- vbcn6.chm1009794
ms.prod: office
ms.assetid: 4a483362-fd6b-f0a7-5cb0-b85a2f794937
ms.date: 06/08/2017
---


# Making Faster For...Next Loops

Integers use less memory than the [Variant data type](vbe-glossary.md) and are slightly faster to update. However, this difference is only noticeable if you perform many thousands of operations. For example:


```vb
Dim CountFaster As Integer    ' First case, use Integer. 
For CountFaster = 0 to 32766     
Next CountFaster 
 
Dim CountSlower As Variant    ' Second case, use Variant. 
For CountSlower = 0 to 32766 
Next CountSlower 

```


The first case above takes slightly less time to run than the second case. However, if  `CountFaster` exceeds 32,767, an error occurs. To fix this, you can change `CountFaster` to the [Long data type](vbe-glossary.md), which accepts a wider range of integers. In general, the smaller the [data type](vbe-glossary.md), the less time it takes to update. Variants are slightly slower than their equivalent data type.


