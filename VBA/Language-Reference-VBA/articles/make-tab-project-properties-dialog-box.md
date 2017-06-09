---
title: Make Tab (Project Properties Dialog Box)
keywords: vbui6.chm181012
f1_keywords:
- vbui6.chm181012
ms.prod: office
ms.assetid: f24e976a-8c0d-a7b2-562d-a099a27786de
ms.date: 06/08/2017
---


# Make Tab (Project Properties Dialog Box)


![Make tab](images/vamaketabsdkversion_ZA01201791.gif)




 **Note**  This feature is not available in all versions of the Visual Basic Editor.


Sets the attributes for the [executable file](vbe-glossary.md) you make. Displays the name of the current project in the title so you can determine which project will be affected by any changes you make. The current project is the item currently selected in the Project Explorer.


## Tab Options

 **Version Number**

Creates the version number for the project.




- Major — Major release number of the project; 0 - 9999.
    
- Minor — Minor release number of the project; 0 - 9999.
    
- Revision — Revision version number of the project; 0-9999.
    
- Auto Increment — If selected, automatically increases the Revision number by one each time you run the  **Make** **Project** command for this project.
    


 **Version Information**

Lets you provide specific information about the current version of your project.




- Type — Information you can use to set a value. You can enter information for your company name, file description, legal copyright, legal trademarks, product name and comments.
    
- Value — The value for the type of information selected in the Type box.
    


 **DLL Base Address**

Allows you to set the base address

 **Remove information about unused ActiveX Controls**

Allows you to specify that even if a control is unused (present in the  **Toolbox**, but not referenced in code), its information will be retained. Uncheck this when you plan to dynamically add the referenced control at run time using the **Controls** collection's **Add** method.


