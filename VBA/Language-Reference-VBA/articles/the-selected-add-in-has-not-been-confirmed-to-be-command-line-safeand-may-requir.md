---
title: The selected Add-In has not been confirmed to be 'Command Line Safe,' and may require some user intervention (possible UI)
keywords: vblr6.chm60147
f1_keywords:
- vblr6.chm60147
ms.prod: office
ms.assetid: cd374ef1-3759-f3a7-6b11-865da620791d
ms.date: 06/08/2017
---


# The selected Add-In has not been confirmed to be 'Command Line Safe,' and may require some user intervention (possible UI)

This error has the following causes and solutions:



- "Command-line safe" means that the add-in is registered in a way to indicate that it contains no user interfaces that require user input when Visual Basic is invoked through a command-line. A user interface can interfere with the operation of unattended processes (such as build scripts). If you don't indicate that an add-in is command-line safe (even if it  _is_ command-line safe), when a user selects your add-in and then Command Line in the Load Behavior box, they'll receive the warning message. This isn't a serious problem, but merely a warning to the user that the selected add-in might possibly contain UI elements that can pop up unexpectedly and halt their automated scripts by pausing for user input. To specify that an add-in is command-line safe, author it that way (you can use the Add-In designer for this), or manually change the value in the registry key.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

