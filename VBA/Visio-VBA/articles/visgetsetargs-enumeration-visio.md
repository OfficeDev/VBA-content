---
title: VisGetSetArgs Enumeration (Visio)
keywords: vis_sdr.chm70055
f1_keywords:
- vis_sdr.chm70055
ms.prod: visio
ms.assetid: e6e35119-5c80-21af-5be3-47f17d616069
ms.date: 06/08/2017
---


# VisGetSetArgs Enumeration (Visio)

Flags to be passed to the  **GetResults** , **SetFormulas** , and **SetResults** methods.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visGetFloats**|0|Results returned as doubles (VT_R8).|
| **visGetFormulasU**|5|Formulas returned in universal syntax (VT_BSTR).|
| **visGetFormulas**|4|Formulas returned as strings (VT_BSTR).|
| **visGetRoundedInts**|2|Results returned as rounded long integers (VT_I4).|
| **visGetStrings**|3|Results returned as strings (VT_BSTR).|
| **visGetTruncatedInts**|1|Results returned as truncated long integers (VT_I4).|
| **visSetBlastGuards**|2|Override present cell values even if they're guarded.|
| **visSetFormulas**|1|Treat strings in results as formulas.|
| **visSetTestCircular**|4|Test for establishment of circular cell references.|
| **visSetUniversalSyntax**|8|Formulas are in universal syntax.|

