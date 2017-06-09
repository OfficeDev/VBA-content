---
title: Object library for Visual Basic for Applications not found
ms.prod: office
ms.assetid: 1776826f-0842-4a7c-8e05-6dd4d777eca3
ms.date: 06/08/2017
---


# Object library for Visual Basic for Applications not found
The Visual Basic for Applications [object library](vbe-glossary.md) is no longer a standalone file; it is integrated into the[dynamic-link library (DLL)](vbe-glossary.md).
Under unusual circumstances a previous version of the object library (vaxxx.olb or vaxxxx.olb) corresponding to the language of the [project](vbe-glossary.md) might be needed, but not found. This error has the following causes and solutions:


- The object library is missing completely, isn't in the expected directory, or is an incorrect version. Search your disk to make sure the object library is in the correct directory, as specified in the [host-application](vbe-glossary.md) documentation.
    

If the missing library is a language version that is installed by the host application, it may be easiest to simply rerun the setup program. If a project requires a different language object library than the one that accompanies your host application (for example, if someone sends you a project written on a machine set up for a different language), make sure the correct language version of the Visual Basic object library is included with the project and it is installed in the expected location.
Applications may support different language versions of their object libraries. To find out which language version is required, display the  **References** dialog box, and see which language is indicated at the bottom of the dialog box.
Object libraries exist in different versions for each platform. Therefore, when projects are moved across platforms, for example, from Macintosh to Microsoft Windows, the correct language version of the referenced library for that platform must be available in the location specified in your host application documentation. Note that some language codes are two characters while others are three characters.
The Visual Basic object library file name is constructed as follows:


- Windows: Application Code + Language Code + [Version].OLB. For example: The French Visual Basic for Applications object library for version 2 was vafr2.olb.
    
- Macintosh: Application Name Language Code [Version] OLB. For example: The French Visual Basic for Applications object library for version 2 was VA FR 2 OLB.
    

If you can't find a missing project or object library on your system, contact the [referencing project's](vbe-glossary.md) author. If the missing library is a Microsoft application object library, you can obtain it as follows:


- If you have access to Microsoft electronic technical support services, refer to the technical support section of this Help file. Under electronic services, you will find instructions on how to use the appropriate service option.
    
- If you don't have access to Microsoft electronic technical support services, Microsoft object libraries are available upon request as an application note from Microsoft. Information on how to contact your local Microsoft product support organization is also available in the technical support section of this Help file.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

