---
title: Using Visual Basic with Outlook
keywords: olfm10.chm3077137
f1_keywords:
- olfm10.chm3077137
ms.prod: outlook
ms.assetid: ddcdada6-7dc1-1c7d-0165-27f8b353662e
ms.date: 06/08/2017
---


# Using Visual Basic with Outlook

You can use Visual Basic to customize and extend Outlook. You can control Outlook by writing a macro in Visual Basic for Applications, a custom form in VBScript, and other languages that you can use to write an add-in, such as Visual Basic. Which you use depends on what you want your program to do.

Visual Basic is a full-featured programming language you can use to create stand-alone applications or dynamic-link libraries (DLLs) that extend other applications. Visual Basic for Applications is a subset of Visual Basic that is run within an application to extend its capabilities. VBScript is a simplified version of Visual Basic for Applications and is run within an Outlook item. In all cases, these programming languages control Outlook through its object model.

Learn about the  [Outlook object model](about-the-object-environment.md).

If you want to create a separate application that accesses data stored by Outlook and uses Outlook to send and receive messages, use Visual Basic to create the application. You can also use other programming languages, such as C++, to control Outlook through its object model. You can create a DLL that can extend Outlook as a COM add-in. One application of COM add-ins is to program form regions and create custom forms. 

You use Visual Basic for Applications in one of two ways: You can use Visual Basic for Applications in other applications (such as Microsoft Excel or Microsoft Word) to automate Outlook, or you can use Visual Basic for Applications within Outlook to control Outlook. If you expect your users to be using another application most of the time, and you want to give them the ability to send a message using Outlook or to access information stored by Outlook, write Visual Basic for Applications programs in that application that control Outlook through the Outlook object model. If, on the other hand, you want to write Visual Basic code that customizes how Outlook works (like a macro), use Visual Basic for Applications within Outlook.

While you use an add-in to extend form regions in a custom form, you can extend the functionality of form pages in custom forms by using VBScript. VBScript programs are stored within a form. Because the program code is contained within the form, it can be sent with an item to another user. Other than the consideration of whether to customize a form with a form page or a form region, another important consideration in choosing which kind of the Visual Basic programming language you will use is the type of events you want your program to respond to. 

Because VBScript code is associated with a particular item, code that responds to events in specific items (such as when a particular item is opened or a value in a field is changed) is easiest to write using VBScript. If, on the other hand, you want your program to respond to events that occur in the application, in Windows Explorer, in folders, or in all items, then you should write your program using Visual Basic or Visual Basic for Applications.

Code written for Visual Basic or Visual Basic for Applications often does not work in VBScript without modification. For example, you must replace all built-in constants written in Visual Basic for Applications with the literal numeric values of those constants in VBScript. And VBScript uses only the  **Variant** data type.
Learn about  [constants and variables in VBScript](constants-and-variables-in-vbscript.md).
In Outlook Visual Basic for Applications and VBScript, you do not need to call  **[CreateObject](application-createobject-method-outlook.md)** or **GetObject** to obtain an **[Application](application-object-outlook.md)** object. For example, the following code displays the Tasks folder:



```vb
Set olMAPI = Application.GetNameSpace("MAPI") 
olMAPI.GetDefaultFolder(13).Display
```

In Visual Basic or Visual Basic for Applications in other applications, you must either explicitly create the  **Application** object, as shown in the following code:



```vb
Set myOlApp = CreateObject("Outlook.Application") 
Set olMAPI = myOlApp.GetNameSpace("MAPI") 
olMAPI.GetDefaultFolder(olFolderTasks).Display
```

or use the  **Application** object that is passed to the **OnConnection** event of the add-in.

 **Note**  The  **Application** object returned from calling the **[CreateObject](application-createobject-method-outlook.md)** method and any of its subordinate objects, properties, and methods are not trusted. For more information on using a trusted **Application** object in a COM add-in, see [Security Behavior of the Outlook Object Model](security-behavior-of-the-outlook-object-model.md).


