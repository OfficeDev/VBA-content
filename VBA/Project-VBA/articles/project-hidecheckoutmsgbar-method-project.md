---
title: Project.HideCheckoutMsgBar Method (Project)
keywords: vbapj.chm131099
f1_keywords:
- vbapj.chm131099
ms.prod: project-server
ms.assetid: 2a62080f-1e23-dda5-346f-4b0194173190
ms.date: 06/08/2017
---


# Project.HideCheckoutMsgBar Method (Project)
Hides the project checkout message bar.

## Syntax

 _expression_. **HideCheckoutMsgBar**

 _expression_ A variable that represents a **Project** object.


### Return value

 **Nothing**


## Remarks

The checkout message bar is the yellow information bar near the top of the Project window that shows  **READ-ONLY This project was opened in read-only mode**, and contains a  **Check Out** button. If the checkout message bar is not visible, the **HideCheckoutMsgBar** method displays run-time error 1004, "An unexpected error occurred with the method."


## Example

The following example tests whether the checkout message bar is visible; if so, it hides the message bar.


```vb
Sub TestHideCheckoutMessageBar()
    If ActiveProject.IsCheckoutMsgBarVisible Then
        ActiveProject.HideCheckoutMsgBar
    End If
End Sub
```


## See also


#### Concepts


[Project Object](project-object-project.md)
#### Other resources


[IsCheckoutMsgBarVisible Property](project-ischeckoutmsgbarvisible-property-project.md)
[CheckoutProject Method](project-checkoutproject-method-project.md)
