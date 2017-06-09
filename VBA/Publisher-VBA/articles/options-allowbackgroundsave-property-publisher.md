---
title: Options.AllowBackgroundSave Property (Publisher)
keywords: vbapb10.chm1048577
f1_keywords:
- vbapb10.chm1048577
ms.prod: publisher
api_name:
- Publisher.Options.AllowBackgroundSave
ms.assetid: 5bddfb2d-7fb7-99db-43ea-c6ee53e1d0b3
ms.date: 06/08/2017
---


# Options.AllowBackgroundSave Property (Publisher)

 **True** (default) for Microsoft Publisher to save publications in the background, allowing users to perform other actions at the same time. Read/write **Boolean**.


## Syntax

 _expression_. **AllowBackgroundSave**

 _expression_A variable that represents an  **Options** object.


### Return Value

Boolean


## Remarks

This setting is saved for each individual user and persists from one session to another.


## Example

This example turns off background save, so publications do not save in the background.


```vb
Sub DoNotSaveInBackground() 
 Options.AllowBackgroundSave = False 
End Sub
```


