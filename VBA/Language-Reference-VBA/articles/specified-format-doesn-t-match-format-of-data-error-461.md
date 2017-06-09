---
title: Specified format doesn't match format of data (Error 461)
keywords: vblr6.chm461
f1_keywords:
- vblr6.chm461
ms.prod: office
ms.assetid: 3e18e1f4-f607-f4e7-d3ed-44f76ab345f2
ms.date: 06/08/2017
---


# Specified format doesn't match format of data (Error 461)

The specified Clipboard format is incompatible with the method being executed. This error has the following causes and solutions:



- You tried to use the  **GetText** method or **SetText** method with a Clipboard format other than **vbCFText** or **vbCFLink**. Before using these methods, use the **GetFormat** method to test whether the current contents of the Clipboard matches the specified format.
    
- You tried to use the  **GetData** method or **SetData** method with a Clipboard format other than **vbCFBitmap**, **vbCFDIB**, or **vbCFMetafile**. Before using these methods, use the **GetFormat** method to test whether the current contents of the Clipboard matches the specified graphics format.
    


