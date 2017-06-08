---
title: MacScript Function
keywords: vblr6.chm1010848
f1_keywords:
- vblr6.chm1010848
ms.prod: office
ms.assetid: d845de85-a0d8-e10e-1174-8571e42bb8d2
ms.date: 06/08/2017
---

# MacScript Function


**Note** This function has been deprecated, therefore it is no longer supported. For information, see this [Stack Overflow article](http://stackoverflow.com/a/30949324/209942).

Executes an AppleScript script and returns a value returned by the script, if any.
 **Syntax**
 **MacScript**_script_
The  _script_ argument is a [String expression](vbe-glossary.md). The  **String** expression either can be a series of AppleScript commands or can specify the name of an AppleScript script or a script file.
 **Remarks**
Multiline scripts can be created by embedding carriage-return characters ( **Chr(** 13 **)** ).


```
ThePath$ = Macscript("ChooseFile")


```


