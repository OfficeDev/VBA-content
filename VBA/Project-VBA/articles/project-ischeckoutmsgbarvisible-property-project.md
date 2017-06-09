---
title: Project.IsCheckoutMsgBarVisible Property (Project)
ms.prod: project-server
ms.assetid: 7d3ef8b3-36c1-d1f1-6c10-ad82573f9d08
ms.date: 06/08/2017
---


# Project.IsCheckoutMsgBarVisible Property (Project)
Gets whether the checkout message bar is visible. Read-only  **Boolean**.

## Syntax

 _expression_. **IsCheckoutMsgBarVisible**

 _expression_ A variable that represents a **Project** object.


## Remarks

The checkout message bar is the yellow information bar near the top of the Project window that shows  **READ-ONLY This project was opened in read-only mode**, and contains a  **Check Out** button. The **IsCheckoutMsgBarVisible** property is **True** if the checkout message bar is visible; otherwise, **False**.


## Example

The following example tests whether the checkout message bar is visible; if so, it hides the message bar.


```vb
Sub TestHideCheckoutMessageBar()
    If ActiveProject.IsCheckoutMsgBarVisible Then
        ActiveProject.HideCheckoutMsgBar
    End If
End Sub
```


## Property value

 **BOOL**


## See also


#### Concepts


[Project Object](project-object-project.md)
#### Other resources


[IsCheckoutOSVisible](project-ischeckoutosvisible-property-project.md)
[HideCheckoutMsgBar Method](project-hidecheckoutmsgbar-method-project.md)
[CheckoutProject Method](project-checkoutproject-method-project.md)
