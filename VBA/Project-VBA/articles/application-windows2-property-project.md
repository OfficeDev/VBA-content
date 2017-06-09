---
title: Application.Windows2 Property (Project)
ms.prod: project-server
api_name:
- Project.Application.Windows2
ms.assetid: 038d051c-769d-3a14-c884-7b4b669d3cc8
ms.date: 06/08/2017
---


# Application.Windows2 Property (Project)

Gets a  **[Windows2](windows2-object-project.md)** collection representing the open windows in the application. Read-only **Windows2**.


## Syntax

 _expression_. **Windows2**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **Windows2** property is recommended, in place of the **Windows** property, for all new development in VBA and external applications developed with the .NET Framework.


## Example

The following example cascades all the open windows.


```vb
Sub CascadeWindows() 
 Dim I As Integer 
 
 ActiveWindow.WindowState = pjNormal ' Restore the window. 
 
 With Application.Windows2 
 For I = 1 To .Count 
 .Item(I).Activate 
 .Item(I).Top = (I - 1) * 15 
 .Item(I).Left = (I - 1) * 15 
 Next I 
 End With 
 
End Sub
```


