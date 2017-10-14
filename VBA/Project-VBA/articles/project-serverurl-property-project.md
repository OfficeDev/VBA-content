---
title: Project.ServerURL Property (Project)
ms.prod: project-server
ms.assetid: 444dd91d-a449-db8c-3863-d85bc6e77ca1
ms.date: 06/08/2017
---


# Project.ServerURL Property (Project)
Gets the URL of the Project Web App instance with which Project Professional is connected. For a synchronized SharePoint task list, gets or sets an arbitrary value that has no effect on the project. Read/write  **String**.

## Version Information

Version added: Project


## Syntax

 _expression_. **ServerURL**

 _expression_ A variable that represents a **Project** object.


## Remarks

If Project is not connected with Project Web App, the  **ServerURL** property gets an empty string. If Project Professional is connected with Project Web App and the active project is a synchronized SharePoint task list, the **ServerURL** property still gets the URL of Project Web App, not the URL of the SharePoint task list.

For a synchronized SharePoint task list, you can set the value of  **ServerURL** to an arbitrary string, which is saved with the project. When you close Project Professional and reopen the SharePoint task list, **ServerURL** gets the arbitrary value. However, that value has no effect on the project or URL of the task list. For example, run the following code in the **Immediate** window of the VBE, and then close Project Professional.




```vb
ActiveProject.ServerURL = "http://SomeBogusServer/NOP%20No%20URL"
```

Start Project Professional again, open the SharePoint task list, and then run  `? ActiveProject.ServerURL` in the **Immediate** window. The statement returns the arbitrary string.

For an enterprise project that Project Server manages, if you try to set the value of  **ServerURL**, Project shows a run-time error 1101, "The argument value is not valid."


## Property value

 **STRING**


## See also


#### Concepts


[Project Object](project-object-project.md)
