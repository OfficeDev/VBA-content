---
title: Communicating with Other Applications
keywords: vbawd10.chm5210293
f1_keywords:
- vbawd10.chm5210293
ms.prod: word
ms.assetid: c54b9e38-941f-e861-ff94-3a29490ae56e
ms.date: 06/08/2017
---


# Communicating with Other Applications

In addition to working with Word data, you may want your application to exchange data with other applications, such as Excel, PowerPoint, or Access. You can communicate with other applications by using Automation (formerly OLE Automation) or dynamic data exchange (DDE).


## Automating Word from another application

Automation allows you to return, edit, and export data by referencing another application's objects, properties, and methods. Application objects that can be referenced by another application are called Automation objects.

The first step toward making Word available to another application for Automation is to make a reference to the Word  **[Application](application-object-word.md)** object. In Visual Basic, you use the Visual Basic **CreateObject** or **GetObject** function to return a reference to the Word **Application** object. For example, in a Excel procedure, you could use the following instruction.




```vb
Set wrd = CreateObject("Word.Application")
```

This instruction makes the  **Application** object in Word available for Automation. Using the objects, properties, and methods of the Word **Application** object, you can control Word. For example, the following instruction creates a new Word document.




```
wrd.Documents.Add
```

Use the  **Visible** property to make the new document visible after creating it.




```vb
wrd.Visible = True
```

The  **CreateObject** function starts a Word session that Automation will not close when the object variable that references the **Application** object expires. Setting the object reference to the Visual Basic **Nothing** keyword will not close Word. Instead, use the **[Quit](application-quit-method-word.md)** method to close the Word application. The following Excel example displays the Word startup path. The **Quit** method is used to close the new instance of Word after the startup path is displayed.




```vb
Set wrd = CreateObject("Word.Application") 
MsgBox wrd.Options.DefaultFilePath(wdStartupPath) 
wrd.Quit
```


## Automating another application from Word

To exchange data with another application using Automation from Word, you first obtain a reference to the application using the  **CreateObject** or **GetObject** function. Then, using the objects, properties, and methods of the other application, you add, change, or delete information. When you finish making your changes, close the application. The following Word example displays the Excel startup path. You can use the Visual Basic **Set** statement with the **Nothing** keyword to clear an object variable, which has the same effect as closing the application.


```vb
Set myobject = CreateObject("Excel.Application") 
MsgBox myobject.StartupPath 
Set myobject = Nothing
```


## Using dynamic data exchange (DDE)

If an application does not support Automation, DDE may be an alternative. DDE is a protocol that permits two applications to communicate by continuously and automatically exchanging data through a DDE "channel." To control a DDE conversation between two applications, you establish a channel, select a topic, request and send data, and then close the channel. The following table lists the tasks that Word performs with DDE and the methods used to control each task in Visual Basic.


 **Security Note**  





|**Task**|**Method**|
|:-----|:-----|
|Starting DDE| **[DDEInitiate](application-ddeinitiate-method-word.md)**|
|Getting text from another application| **[DDERequest](application-dderequest-method-word.md)**|
|Sending text to another application| **[DDEPoke](application-ddepoke-method-word.md)**|
|Carrying out a command in another application| **[DDEExecute](application-ddeexecute-method-word.md)**|
|Close DDE channel| **[DDETerminate](application-ddeterminate-method-word.md)**|
|Close all DDE channels| **[DDETerminateAll](application-ddeterminateall-method-word.md)**|

