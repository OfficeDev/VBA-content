---
title: Globally Change a Default Form
keywords: olfm10.chm1048699
f1_keywords:
- olfm10.chm1048699
ms.prod: outlook
ms.assetid: 499ea2dd-e98b-a368-453d-cf3df238c324
ms.date: 06/08/2017
---


# Globally Change a Default Form

You can change the default form in Microsoft Outlook by making changes to the Microsoft Windows registry. The registry settings specify which forms are substituted for the default Outlook forms. For example, if you create a custom form called "Default," that custom form has a message class of IPM.Note.Default, instead of the standard Outlook e-mail message class of IPM.Note. You can add certain registry keys to indicate that Outlook should substitute the IPM.Note.Default form for the standard IPM.Note form.


 **Caution**  Custom forms may have certain limitations. Before you change the detault form in Outlook to a custom form, be aware of possible implications especially if the form will be deployed to many users. See Microsoft Knowledge Base article 241235 for further information.


The Forms Administrator utility does not create the registry keys in the correct location for Microsoft Office Outlook 2003 or later. However, you can use the Forms Administrator utility to create a Windows registry file and make the necessary changes. To use a Windows registry file to change the default form in Outlook 2003 or later: 


1. Download the Outlook 2000 Forms Administrator utility.
    
2. Run the Forms Administrator utility, and then change the settings as you would for Outlook 2000 or Outlook 2002.
    
3. To save the registry settings on your computer, click  **Save**. This also makes the  **Export Saved Settings** button available.
    
4. Click  **Export Saved Settings** to save a Windows registry (.reg) file.
    
5. Open the .reg file in a text editor, such as Notepad.
    
6. The registry key paths reference 9.0, which is the location for Outlook 2000 registry settings. Change all of the references of 9.0 to 11.0, 12.0, or 14.0, depending on your current version of Outlook.
    
7. Save the .reg file.
    
8. Run the .reg file on each computer where you want to change the default form, so that the keys are added to the Windows registry on that computer.
    

 **Note**  If you used the Forms Administrator utility to change the default forms in Outlook 2000 or Outlook 2002, and you then upgrade to Outlook 2003 or later, Microsoft Office or Outlook Setup migrates the registry keys to the correct location so that Outlook continues to use the substituted forms.


