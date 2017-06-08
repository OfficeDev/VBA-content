---
title: Connection to type library or object library for remote process has been lost (Error 442)
keywords: vblr6.chm1019366
f1_keywords:
- vblr6.chm1019366
ms.prod: office
ms.assetid: a362a898-f34e-8d75-c2dd-40cac5b95b3b
ms.date: 06/08/2017
---


# Connection to type library or object library for remote process has been lost (Error 442)

During remote access (that is, when accessing an object that is part of another process or is running on another machine), the connection to the library containing object information was broken. This error has the following cause and solution:



- You lost your connection to the remote process's [object library](vbe-glossary.md) or[type library](vbe-glossary.md).
    
    To reconnect, follow these steps:
    
    
    
      1. Restart the  **Application** object.
    
  2. In the error dialog box through which you entered this Help topic, click  **OK** to display the **References** dialog box.
    
  3. The lost reference appears with the word MISSING to its left.
    
  4. Remove the lost reference.
    
  5. In the  **References** dialog box, click the check box for the object you started in step 1.
    

    
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

