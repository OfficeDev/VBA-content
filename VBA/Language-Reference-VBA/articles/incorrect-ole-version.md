---
title: Incorrect OLE version
keywords: vblr6.chm1011192
f1_keywords:
- vblr6.chm1011192
ms.prod: office
ms.assetid: 577f33f5-f44e-08c1-1cb8-b64277068d01
ms.date: 06/08/2017
---


# Incorrect OLE version

Your versions of the OLE [dynamic-link libraries (DLL)](vbe-glossary.md) (Windows) or code resource (Macintosh) don't match those expected by the[host application](vbe-glossary.md). In Microsoft Windows, the application searches for the DLLs first in the current directory, then along your path settings, and then in the WINDOWS\SYSTEM directory. This error has the following cause and solution:



- Earlier OLE DLLs were encountered in the search before the DLLs expected by the host application. You should not try to use both versions of the DLLs.
    

Note that on the Macintosh, OLE files are normally only found in the Extensions folder so it is unlikely that this error will occur.
For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

