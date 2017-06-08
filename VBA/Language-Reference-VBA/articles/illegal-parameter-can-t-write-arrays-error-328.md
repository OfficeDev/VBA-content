---
title: Illegal parameter. Can't write arrays (Error 328)
keywords: vblr6.chm1117816
f1_keywords:
- vblr6.chm1117816
ms.prod: office
ms.assetid: 2c6082d4-a747-a50c-7d09-d26e0be98e9d
ms.date: 06/08/2017
---


# Illegal parameter. Can't write arrays (Error 328)

An illegal [parameter](vbe-glossary.md) was passed to the method. This error has the following cause and solution:



- In the  **WriteProperties** event of your User Control, you tried to do a **PropBag.WriteProperty** X, where X is an[array](vbe-glossary.md). This isn't supported.
    
    You must write out each element of the array individually.
    


