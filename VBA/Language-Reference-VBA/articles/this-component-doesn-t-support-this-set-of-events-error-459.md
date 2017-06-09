---
title: This component doesn't support this set of events (Error 459)
keywords: vblr6.chm1000459
f1_keywords:
- vblr6.chm1000459
ms.prod: office
ms.assetid: f74abecc-6461-d5ef-0018-a7bbf05eeb4b
ms.date: 06/08/2017
---


# This component doesn't support this set of events (Error 459)

Not every component supports client sinking of events. This error has the following cause and solution:



- You tried to use a  **WithEvents** variable with a component that can't work as an event source for the specified set of events. For example, you may be sinking events of an object, then create another object that **Implements** the first object. Although you might think you could sink the events from the implemented object, that isn't automatically the case. **Implements** only implements an interface for methods and properties. You can't sink events for a component that doesn't source events.
    
-  **WithEvents** isn't supported for Private UserControls, because the type-info needed to raise the ObjectEvent isn't available at runtime.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

