---
title: Application.EnableCancelKey Property (Project)
keywords: vbapj.chm131792
f1_keywords:
- vbapj.chm131792
ms.prod: project-server
api_name:
- Project.Application.EnableCancelKey
ms.assetid: 9b5f4f90-3ef3-139b-5f76-f48d3d7710a8
ms.date: 06/08/2017
---


# Application.EnableCancelKey Property (Project)

Gets or sets a value that controls how the CTRL + BREAK key combination is handled when a macro is running. Read/write  **PjEnableCancelKey**.


## Syntax

 _expression_. **EnableCancelKey**

 _expression_ A variable that represents an **Application** object.


## Remarks

The  **EnableCancelKey** property can be one of the following **[PjEnableCancelKey](pjenablecancelkey-enumeration-project.md)** constants: **pjDisabled**, **pjErrorHandler**, or **pjInterrupt**.


## Example

The following example shows how you can use the  **EnableCancelKey** property to create a custom cancellation error handler.


```vb
Sub CancelOperation() 
 Dim X As Long 
 
 On Error GoTo handleCancel 
 
 Application.EnableCancelKey = pjErrorHandler 
 MsgBox "This may take a long time; press CTRL+BREAK to cancel." 
 
 For X = 1 To 300000000 
 ' Do something here. 
 Next X 
 
handleCancel: 
 If Err = 18 Then 
 MsgBox "Operation cancelled" 
 End If 
 
End Sub
```


