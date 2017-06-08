---
title: Project.IsCheckoutOSVisible Property (Project)
ms.prod: project-server
ms.assetid: 1b240231-dfa1-2cd3-310e-11c8c58791eb
ms.date: 06/08/2017
---


# Project.IsCheckoutOSVisible Property (Project)
Gets whether the  **Check Out** button is visible in the Backstage view. Read-only **Boolean**.

## Syntax

 _expression_. **IsCheckoutOSVisible**

 _expression_ A variable that represents a **Project** object.


## Remarks

If the active project is not checked out, the Backstage view shows a  **Check Out** button. The **IsCheckoutOSVisible** property is **True** if the **Check Out** button is visible in the Backstage view; otherwise, **False**.


## Example

The following example tests whether the checkout message bar is visible; if so, it hides the message bar. However, if the project is not checked out, the backstage view still shows the Check Out button, so the example can try to check out the project. If the project is checked out by you or checked out to someone else, Project shows an error dialog box with the message, "This project is already checked out to you on a different computer or Project Web App session."


```vb
Sub TestBackstageCheckout()
    ' Hide the checkout message bar.
    If ActiveProject.IsCheckoutMsgBarVisible Then
        ActiveProject.HideCheckoutMsgBar
    End If
    
    ' If the Backstage Check Out button is visible, then the
    ' project is not checked out.
    If ActiveProject.IsCheckoutOSVisible Then
        ActiveProject.CheckoutProject
        Debug.Print "Attempted to check out: '" &; ActiveProject.Name &; "'"
    Else
        Debug.Print "'" &; ActiveProject.Name &; "' is already checked out."
    End If
End Sub
```


## Property value

 **BOOL**


## See also


#### Concepts


[Project Object](project-object-project.md)
#### Other resources


[IsCheckoutMsgBarVisible](project-ischeckoutmsgbarvisible-property-project.md)
[HideCheckoutMsgBar Method](project-hidecheckoutmsgbar-method-project.md)
[CheckoutProject Method](project-checkoutproject-method-project.md)
