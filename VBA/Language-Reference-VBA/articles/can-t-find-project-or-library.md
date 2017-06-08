---
title: Can't find project or library
keywords: vblr6.chm1011094
f1_keywords:
- vblr6.chm1011094
ms.prod: office
ms.assetid: 078ae060-a90b-e992-2cfb-34ee6b003098
ms.date: 06/08/2017
---


# Can't find project or library

You can't run your code until all missing references are resolved. This error has the following causes and solutions:



- A [referenced project](vbe-glossary.md) could not be found, or a referenced[object library](vbe-glossary.md) corresponding to the language of the[project](vbe-glossary.md) could not be found.
    
    Unresolved references are prefixed with MISSING in the  **References** dialog box. Select the missing reference to display the path and language of the missing project or library. Follow these steps to resolve the reference or references:
    

 **To resolve the references**


1. Display the  **References** dialog box.
    
2. Select the missing reference.
    
3. Start the [Object Browser](vbe-glossary.md).
    
4. Use the  **Browse** dialog box to find the missing reference.
    
5. Click  **OK**.
    
6. Repeat the preceding steps until all missing references are resolved.
    

Once you find a missing item, the MISSING prefix is removed to indicate that the link is reestablished. If the file name of a referenced project has changed, a new reference is added, and the old reference must be removed.
To remove a reference that is no longer required, simply clear the check box next to the unnecessary reference. Note that the references to the Visual Basic object library and [host-application](vbe-glossary.md) object library can't be removed.
Applications may support different language versions of their object libraries. To find out which language version is required, click the reference and check the language indicated at the bottom of the dialog box.
Object libraries may be standalone files with the extension .OLB or they can be integrated into a [dynamic-link library (DLL)](vbe-glossary.md) They can exist in different versions for each platform. Therefore, when projects are moved across platforms, for example, from Macintosh to Microsoft Windows, the correct language version of the referenced library for that platform must be available in the location specified in your host application documentation.
Object library file names are generally constructed as follows:


- Windows (version 3.1 and earlier): Application Code + Language Code + [Version].OLB. For example: The object library for French Visual Basic for Applications, Version 2 was vafr2.olb. The French Microsoft Excel 5.0 object library was xlfr50.olb.
    
- Macintosh: Application Name Language Code [Version] OLB. For example: The object library for French Visual Basic for Applications, Version 2 was VA FR 2 OLB. The French Microsoft Excel 5.0 object library was MS Excel FR 50 OLB.
    

If you can't find a missing project or library on your system, contact the [referencing project](vbe-glossary.md)'s author. If the missing library is a Microsoft application object library, you can obtain it as follows:


- If you have access to Microsoft electronic technical support services, refer to the technical support section of this Help file. Under electronic services, you will find instructions on how to use the appropriate service option.
    
- If you don't have access to Microsoft electronic technical support services, Microsoft object libraries are available upon request as an application note from Microsoft. Information on how to contact your local Microsoft product support organization is also available in the technical support section of this Help file.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

