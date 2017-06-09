---
title: Automating Outlook from Other Office Applications
keywords: vbaol11.chm5273034
f1_keywords:
- vbaol11.chm5273034
ms.prod: outlook
ms.assetid: d3e44f80-df67-2d28-94dc-14d7a8c8c26c
ms.date: 06/08/2017
---


# Automating Outlook from Other Office Applications

You can use Visual Basic for Applications (VBA) in any Microsoft Office application to control Microsoft Outlook. For example, if you are developing a cross-application solution using one primary application and several secondary applications, you can write VBA code in the primary application to automate Outlook to send messages and to store and retrieve information in Outlook items. For example, in Excel you can write routines that send a workbook to an Outlook distribution list.

To control Outlook objects from outside Outlook, you must establish a reference to the Outlook object library from the project in which you are writing code. To do this, use the  **References** dialog box in the Visual Basic Editor in the primary application. You can then write code that returns a reference to the Outlook [Application](application-object-outlook.md) object. Through this reference, your code has access to all the objects, properties, methods, and constants defined in the Outlook type library.

There are several ways to return a reference to the Outlook  **Application** object.


- You can use the  [CreateObject](application-createobject-method-outlook.md) function to start a new session of Outlook and return a reference to the **Application** object that represents the new session.
    
- You can use the Visual Basic  **GetObject** function to return a reference to the **Application** object that represents a session that is already running. Note that because there can be only one instance of Outlook running at any given time, **GetObject** usually serves little purpose when used with Outlook. **CreateObject** can always be used to access the current instance of Outlook or to create a new instance if one does not exist. However, you can use error trapping with the **GetObject** method to determine whether Outlook is currently running.
    
- You can use the  **New** keyword in several types of statements to implicitly create a new instance of the Outlook **Application** object by using the **Set** statement to set an object variable to the new instance of the **Application** object. You can also use the **New** keyword with the **Dim**,  **Private**,  **Public**, or  **Static** statement to declare an object variable. The new instance of the **Application** object is then created on the first reference to the variable.
    

 **Caution**  When you create a new instance of Outlook, the new instance is not trusted and can trigger the object model guard.

For examples of using these methods of referencing the Outlook  **Application** object, see [Automating Outlook from a Visual Basic Application](automating-outlook-from-a-visual-basic-application.md).

