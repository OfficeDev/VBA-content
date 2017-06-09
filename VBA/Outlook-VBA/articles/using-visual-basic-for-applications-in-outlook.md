---
title: Using Visual Basic for Applications in Outlook
keywords: vbaol11.chm5274249
f1_keywords:
- vbaol11.chm5274249
ms.prod: outlook
ms.assetid: 3dd39c7c-7b90-9c19-490f-258d795787e2
ms.date: 06/08/2017
---


# Using Visual Basic for Applications in Outlook

Visual Basic for Applications (VBA) makes it easy to control Microsoft Outlook within Microsoft Outlook itself. Using VBA in Outlook, you can create macros that perform complex or repetitive tasks automatically. You can also develop program code that responds to Outlook events, allowing you to automate common tasks (such as arranging windows when Outlook starts).

Visual Basic for Applications allows you to take almost full advantage of the Outlook object model, including the wide range of application-level events, without requiring you to run an external application (such as another Microsoft Office application or an application developed using Microsoft Visual Basic). And unlike form scripts developed using Microsoft Visual Basic Scripting Edition (VBScript), Outlook Visual Basic for Applications code is always available in the application; an item does not have to be open to run the code.

All Outlook Visual Basic for Applications code is contained in a project. The project is associated with a particular user, so all users who run Outlook on a computer can customize Outlook to meet their own needs. A project can contain code modules and User Form modules (note that User Form modules are not the same as Outlook forms ).

You use the Visual Basic Editor to create and remove modules, to design User Form modules, and to edit code in modules. This editor provides a powerful set of tools, including a built-in Object Browser and debugger to make developing and troubleshooting code easy. You can even use the Visual Basic Editor in Outlook to develop and test code that you can then copy to a standalone Visual Basic application or a Visual Basic for Applications application in another Microsoft Office application.

## Managing Multiple Visual Basic for Applications Projects

Outlook supports only one Visual Basic for Applications project, Project1, at a time. You cannot add and run another project in the Visual Basic Editor at the same time. Project1 is stored on your hard disk as VbaProject.OTM; each user on the same computer can have a copy of VbaProject.OTM stored for him or her. On a computer running Windows XP, VbaProject.OTM is in <drive>:\Documents and Settings\<user>\Application Data\Microsoft\Outlook.

Because you can run only one Visual Basic for Applications project at a time, before you run a different project, you should exit Outlook, rename the current VbaProject.OTM with a different file name, name the project you want to run as VbaProject.OTM, and restart Outlook to run it. If appropriate, you can also manually integrate the projects to form one VbaProject.OTM to avoid the file naming and renaming.

Outlook Visual Basic for Applications code was designed to be a personal macro development environment, and was not designed to be deployed or distributed. To move a project from one computer to another, for example, moving the project from your work computer to your home computer, you can export the forms and code modules from the work computer and import them to the home computer. You can also copy and paste the source code of the project to Project1 on the home computer using the Visual Basic Editor.

If you are developing a solution that you intend to distribute to more than a few people, you should convert your Visual Basic for Applications code into an Outlook COM Add-in. However, developing a COM Add-in typically requires considerably more programming knowledge than creating a short macro, so if your Visual Basic for Applications project is relatively simple, and there are not too many people who need to use it, you may want to send them the code with instructions on how to set it up.


