---
title: Run Project and Stop Project Commands (Run Menu)
ms.prod: office
ms.assetid: 73c2cd97-496e-441a-ed1c-0617e82621bd
ms.date: 06/08/2017
---


# Run Project and Stop Project Commands (Run Menu)

 **Run Project**

Puts the project into a mode in which it can be used by other applications. This is used to debug and test the stand-alone project before building a [dynamic-link library (DLL)](vbe-glossary.md) (DLL) from it. The current project is registered, replacing any existing registration information for the project (the registry information for an existing DLL version of the project, for example).

 **Stop Project**

Unregisters the project, and restores any previous registry information. This makes the in-memory project no longer able to be called from other applications.

 **Note**  The  **Run Project** and **Stop Project** commands are available only to the current stand-alone project. They are not available to[host application](vbe-glossary.md) document projects.


 **Note**  This feature is not available in all versions of the Visual Basic Editor.


