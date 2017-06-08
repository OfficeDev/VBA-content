---
title: Automating Outlook from a Visual Basic Application
keywords: vbaol11.chm5273039
f1_keywords:
- vbaol11.chm5273039
ms.prod: outlook
ms.assetid: 623f91af-cd50-1ff0-9519-5a39cbcf5d18
ms.date: 06/08/2017
---


# Automating Outlook from a Visual Basic Application

Because Microsoft Outlook supports Automation, you can control Outlook from any program that is written with Microsoft Visual Basic. Automation provides a standard method for one application to access the objects, methods, properties, and events of other applications that support Automation.

The Outlook object model provides all of the functionality necessary to manipulate data that is stored in Outlook folders, and it provides the ability to control many aspects of the Outlook user interface (UI).

To start an Outlook Automation session, you can use either early or late binding. Late binding uses either the Visual Basic  **GetObject** function or the [CreateObject](application-createobject-method-outlook.md) function to initialize Outlook. For example, the following code sets an object variable to the Outlook [Application](application-object-outlook.md) object, which is the highest-level object in the Outlook object model. All Automation code must first define an Outlook **Application** object to be able to access any other Outlook objects.




```vb
Dim objOL as Object 
Set objOL = CreateObject("Outlook.Application")
```

To use early binding, you first need to set a reference to the Outlook object library. Use the Reference command on the Visual Basic for Applications (VBA) Tools menu to set a reference to  **Microsoft Outlook xx.x Object Library**, where  **xx.x** represents the version of Outlook that you are working with. You can then use the following syntax to start an Outlook session.



```vb
Dim objOL as Outlook.Application 
Set objOL = New Outlook.Application
```

Most programming solutions interact with the data stored in Outlook. Outlook stores all of its information as items in folders. Folders are contained in one or more stores. After you set an object variable to the Outlook  **Application** object, you will commonly set a [NameSpace](namespace-object-outlook.md) object to refer to MAPI, as shown in the following example.



```vb
Set objOL = New Outlook.Application 
Set objNS = objOL.GetNameSpace("MAPI") 
Set objFolder = objNS.GetDefaultFolder(olFolderContacts)
```

Once you have set an object variable to reference the folder that contains the items you wish to work with, you use appropriate code to accomplish your task, as shown in the following example.



```vb
Sub CreateNewDefaultOutlookTask() 
    Dim objOLApp As Outlook.Application 
    Dim NewTask As Outlook.TaskItem 
    ' Set the Application object 
    Set objOLApp = New Outlook.Application 
    ' You can only use CreateItem for default items 
    Set NewTask = objOLApp.CreateItem(olTaskItem) 
    ' Display the new task form so the user can fill it out 
    NewTask.Display 
End Sub
```

If you are using VBA to create macros, there are two ways you can automate Outlook. You can implement a macro that creates a new instance of the Outlook  **Application** object. The `CreateNewDefaultOutlookTask()` method above shows how to call `New Outlook.Application` to create a new **Application** object instance.

 **Caution**  This new instance of Outlook is not trusted and can trigger the object model guard. 

As an alternative to creating and automating a separate instance of Outlook, you can use VBA to implement a macro that automates the current instance of Outlook. To do so, use the  **Application** object intrinsic to the environment. This **Application** object is trusted and can avoid triggering the object model guard. For more information about the object model guard, see [What's New for Developers in Outlook 2007 (Part 1 of 2)](http://msdn.microsoft.com/library/76e3f0b7-ef2b-4e9f-8515-3002d75d7721.aspx). The following example shows the  `CreateAnotherNewDefaultOutlookTask()` method using the **Application** object from the current instance of Outlook.



```vb
Sub CreateAnotherNewDefaultOutlookTask() 
    Dim NewTask As Outlook.TaskItem 
 
    ' You can only use CreateItem for default items 
    Set NewTask = Application.CreateItem(olTaskItem) 
    ' Display the new task form so the user can fill it out 
    NewTask.Display 
End Sub
```


