---
title: Can't empty Clipboard (Error 520)
keywords: vblr6.chm520
f1_keywords:
- vblr6.chm520
ms.prod: office
ms.assetid: d1b47bdf-e48b-471c-05c4-0491c1240c0e
ms.date: 06/08/2017
---


# Can't empty Clipboard (Error 520)

The Clipboard was opened but could not be emptied. This error has the following cause and solution:



- Another application is using the Clipboard and won't release it to your application. Set an error trap for this situation in your code and provide a message box with  **Retry** and **Cancel** buttons to allow the user to try again after a short pause.
    


