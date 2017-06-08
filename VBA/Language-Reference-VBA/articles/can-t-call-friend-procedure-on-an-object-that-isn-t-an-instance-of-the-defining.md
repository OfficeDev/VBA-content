---
title: Can't call Friend procedure on an object that isn't an instance of the defining class (Error 97)
keywords: vblr6.chm1000097
f1_keywords:
- vblr6.chm1000097
ms.prod: office
ms.assetid: a926e392-f2b3-5f65-41a7-211eeb31e92e
ms.date: 06/08/2017
---


# Can't call Friend procedure on an object that isn't an instance of the defining class (Error 97)

A  **Friend** procedure is callable from a[module](vbe-glossary.md) that is outside the[class](vbe-glossary.md), but part of the [project](vbe-glossary.md) within which the class is defined. This error has the following causes and solutions:



- You tried to call the  **Friend** procedure of a class. Although your reference variable is of the proper type, the variable points to an instance that isn't an instance of the class. For example, this can occur if there are two classes, class _ics_ and class _y_ (that implements class _y_ ), but you mistakenly assign the instance of _classy_ to the instance of class _ics_.
    
- You tried to access a  **Friend** property or method either cross-process or cross-thread. Friend procedures are not part of a class's public interface, so they cannot be marshaled cross-process or cross-thread.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

