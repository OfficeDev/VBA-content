---
title: Duplicate definition
keywords: vblr6.chm1057031
f1_keywords:
- vblr6.chm1057031
ms.prod: office
ms.assetid: 8e9f8532-28fa-8244-939a-40eeee372312
ms.date: 06/08/2017
---


# Duplicate definition

You can only define a [conditional compiler constant](vbe-glossary.md) to have one value. This error has the following cause and solution:



- You specified two different values for the same conditional compiler constant, for example:
    
  ```
  #Const Mac = 0 
#Const Mac = 1 

  ```


    Remove one of the definitions.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

