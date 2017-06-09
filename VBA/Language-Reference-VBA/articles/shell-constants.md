---
title: Shell Constants
ms.prod: office
ms.assetid: 76b5cc9e-e896-f658-7d23-ca850305a16b
ms.date: 06/08/2017
---


# Shell Constants

The following [constants](vbe-glossary.md) can be used anywhere in your code in place of the actual values:



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbHide**|0|Window is hidden and focus is passed to the hidden window.|
|**vbNormalFocus**|1|Window has focus and is restored to its original size and position.|
|**vbMinimizedFocus**|2|Window is displayed as an icon with focus.|
|**vbMaximizedFocus**|3|Window is maximized with focus.|
|**vbNormalNoFocus**|4|Window is restored to its most recent size and position. The currently active window remains active.|
|**vbMinimizedNoFocus**|6|Window is displayed as an icon. The currently active window remains active.|

On the Macintosh,  **vbNormalFocus**, **vbMinimizedFocus**, and **vbMaximizedFocus** all place the application in the foreground; **vbHide**, **vbNoFocus**, **vbMinimizedFocus** all place the application in the background.


