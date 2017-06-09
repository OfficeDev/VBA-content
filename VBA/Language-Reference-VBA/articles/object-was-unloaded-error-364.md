---
title: Object was unloaded (Error 364)
keywords: vblr6.chm1117820
f1_keywords:
- vblr6.chm1117820
ms.prod: office
ms.assetid: 155b96e2-0bb6-dea0-b25a-26abe50ab198
ms.date: 06/08/2017
---


# Object was unloaded (Error 364)

A form was unloaded from its own  **_Load** procedure. This error has the following cause and solution:



- A form with an  **Unload** statement in its **_Load** procedure was implicitly loaded. For example, the following will implicitly load `YourForm` if it isn't already loaded:
    
```vb
MyForm.BackColor = YourForm.BackColor. 
```


    Remove the  **Unload** statement from the **Form_Load** procedure.
    


