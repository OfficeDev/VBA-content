---
title: "Warning: custom language settings not portable"
keywords: vblr6.chm1011120
f1_keywords:
- vblr6.chm1011120
ms.prod: office
ms.assetid: a23a1fd7-2995-cab0-0be2-74cd84a3a98a
ms.date: 06/08/2017
---


# Warning: custom language settings not portable

Not all language settings are portable. This warning has the following cause and solution:



- You used a custom language setting in your code. When you choose a custom language/country setting for your code, the language conventions used in your code match those set in the  **Control Panel** of your system. You can use custom code[locale](vbe-glossary.md) settings, but your code may not work in other locales or on other systems with different settings. The[host application](vbe-glossary.md) parses some strings based on the **Control Panel** settings of the machine on which it is running. If the **Control Panel** settings on the target machine aren't the same as those on the machine on which the code was written, strings parsed by a host application don't work, for example, code that depends on a locale-specific decimal separator. Therefore, you should not use a custom language setting unless you don't intend to send your code to other users. If you plan to send your code to other users, select a predefined locale.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

