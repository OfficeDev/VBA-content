---
title: Using events with Application and Project objects
ms.prod: project-server
ms.assetid: 64a18885-f203-c298-db11-f9e8e75bb7b6
ms.date: 06/08/2017
---


# Using events with Application and Project objects


You can write event procedures at the application level or the project level. For example, the [Activate](project-activate-event-project.md) event occurs at the project level when a project is activated and the [NewProject](application-newproject-event-project.md) event occurs at the application level when a project is created.

Event procedures in the [Project](project-object-project.md) object are available for any open project. To write event procedures for the [Application](application-object-project.md) object, you must create a new object using the **WithEvents** keyword in a class module. The following steps show how to create and test a simple application event handler:

1. In the Visual Basic Editor, on the option menu for  **VBAProject**, choose  **Insert**, and then choose  **Class Module** to create a class named **Class1**. You can rename the class module in the **Properties** pane. In the following examples, the class is named **TestClass**.
    
2. Paste the following code in the  **TestClass** module.
    
  ```
  Option Explicit 
Public WithEvents oApp As Application 
 
Private Sub oApp_NewProject(ByVal pj As Project) 
    MsgBox "You created the " &; pj.Name &; " project." 
End Sub 
 
Private Sub Class_Initialize() 
    ' Add class initialization statements here, if needed. 
End Sub 
  ```

3. Open the  **ThisProject** module and paste in the following code.
    
  ```
  Option Explicit 
Private tClass As New TestClass 
 
Sub TestNewProjectEvent() 
    Set tClass.oApp = Application 
    tClass.oApp.Projects.Add 
    Projects.Add 
End Sub
  ```

4. Run the  **TestNewProjectEvent** macro. The macro calls the **Projects.Add** method twice—once through the **TestClass** object and once directly through the **Application** object. When the Project application creates the first project, the result is a **Microsoft Project** dialog box with the message **You created the Project2 project**. When you choose  **OK**, Project creates the second project and shows another dialog box with the message  **You created the Project3 project**.
    

 **Important**  For application-level events, register event handlers  _after_ you set `Application.Visible = True`.

If you instantiate Project from another application and register an application-level event before setting the  **Visible** property of the **Application** object to **True**, the properties and methods of child objects of **Application** do not work. For example, `Application.ActiveProject.Name` is not accessible.

 **Note**  Event code in your project can run unexpectedly, or can be blocked, if event code exists in the global file (Global.mpt).


- If code exists for an event in both the global files and project files, only the code in the project event runs.


    
- If code for an event does not exist in a project, but does exist in the global file, the code in the global event runs.


    
- If code for one of the three events [Application.ProjectBeforeClose](application-projectbeforeclose-event-project.md), [Application.ProjectBeforeSave](application-projectbeforesave-event-project.md), or [Project.Open](project-open-event-project.md) exists in the global file, but not in the project, it affects both the global and project files. If code exists for those events in both the global and project files, the code in the global file affects the global file, and the code in the project file affects the project.
    

