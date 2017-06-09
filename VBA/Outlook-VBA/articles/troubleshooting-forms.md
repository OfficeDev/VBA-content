---
title: Troubleshooting Forms
keywords: olfm10.chm1044313
f1_keywords:
- olfm10.chm1044313
ms.prod: outlook
ms.assetid: 79c44e72-5ef8-ad43-2838-8750d73387d5
ms.date: 06/08/2017
---


# Troubleshooting Forms

My solution doesn't run on other computers.

Use the following troubleshooting tips to help troubleshoot problems if your forms or programming solution works on some computers, but not others.

 **Microsoft Visual Basic Scripting Edition (VBScript) version** If your solution involves forms that use VBScript, you may need to make sure that all computers are using the same version of VBScript that is being used on your development computer. VBScript is a shared component. Installing other software, such as a newer version of Windows Internet Explorer, may result in newer versions of VBScript being installed. For the latest information about VBScript versions, go to http://www.microsoft.com/scripting.

 **Controls** If your solution uses any nonstandard controls, check to see if the controls are properly installed on all of the user's computers. If you are using any control other than one of the Forms 2.0 controls that are installed by Microsoft Office, you should provide your users with a Setup program to ensure that all of your controls are installed correctly.
 **Permissions or user rights** Make sure that any user experiencing problems has proper permissions or rights to use any public folders or other resources that your solution uses.
 
The  **Click** event of a control doesn't fire.

The  **Click** event doesn't fire for controls bound to a field. Because the controls are bound to a field, you can use the **PropertyChange** or **CustomPropertyChange** event of the item to detect any change to the value of the field.

Is there anything that helps debugging my custom form?
 **Problems in the user interface** If your custom form is extended by an add-in, the problems you see in the user interface of the form may be caused by the add-in. In the **Other** tab of the **Options** dialog box, click **Advanced Options**. Select the check box for  **Show add-in user interface errors**. This will help you debug errors that your add-in causes in the user interface.
 **Problems in the functionality or behavior of the form** If your custom form contains a form region, check the XML that defines the form region. Form region XML can be specified inline in the Windows registry, but more commonly the XML would be defined in the corresponding form region manifest XML file. Make sure that the XML validates against the form region XML schema. For more information, see [Using the Form Region XML Manifest to Define a Form Region](using-the-form-region-xml-manifest-to-define-a-form-region.md).

