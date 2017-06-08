---
title: Can't open Clipboard (Error 521)
keywords: vblr6.chm1117803
f1_keywords:
- vblr6.chm1117803
ms.prod: office
ms.assetid: 4a98a637-81ab-c71f-e3c4-8504c441c5a5
ms.date: 06/08/2017
---


# Can't open Clipboard (Error 521)

The Clipboard has already been opened by another application. This error has the following cause and solution:



- Another application is using the Clipboard and won't release it to your application. Set an error trap for this situation in your code and provide a message box with  **Retry** and **Cancel** buttons to allow the user to try again after a short pause.
    


