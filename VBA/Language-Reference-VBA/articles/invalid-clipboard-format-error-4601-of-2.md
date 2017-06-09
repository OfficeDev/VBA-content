---
title: Invalid Clipboard format (Error 460) [1 of 2]
keywords: vblr6.chm1117805
f1_keywords:
- vblr6.chm1117805
ms.prod: office
ms.assetid: 3ee6ca6e-b471-c044-3db4-b2c4b02749c9
ms.date: 06/08/2017
---


# Invalid Clipboard format (Error 460) [1 of 2]

The specified Clipboard format is incompatible with the method being executed. This error has the following causes and solutions:



- You tried to use the  **GetText** method or **SetText** method with a Clipboard format other than **vbCFText** or **vbCFLink**. Remove the invalid format and specify one of the two valid formats.
    
- You tried to use the  **GetData** or **SetData** method with a Clipboard format other than **vbCFBitmap**, **vbCFDIB**, or **vbCFMetafile**. Remove the invalid format and specify one of the three valid graphics formats.
    


