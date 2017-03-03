---
title: Form.Repaint Method (Access)
keywords: vbaac10.chm13505
f1_keywords:
- vbaac10.chm13505
ms.prod: ACCESS
api_name:
- Access.Form.Repaint
ms.assetid: ce386055-c4b7-9aa8-7f49-de0010467970
---


# Form.Repaint Method (Access)

The  **Repaint** method completes any pending screen updates for a specified form. When performed on a form, the **Repaint** method also completes any pending recalculations of the form's controls.


## Syntax

 _expression_. **Repaint**

 _expression_ A variable that represents a **Form** object.


### Return Value

Nothing


## Remarks

Microsoft Access sometimes waits to complete pending screen updates until it finishes other tasks. With the  **Repaint** method, you can force immediate repainting of the controls on the specified form. You can use the **Repaint** method:


- When you change values in a number of fields. Unless you force a repaint, Microsoft Access might not display the changes immediately, especially if other fields, such as those in an expression in a calculated control, depend on values in the changed fields.
    
- When you want to make sure that a form displays data in all of its fields. For example, fields containing OLE objects often don't display their data immediately after you open a form.
    
This method doesn't cause a requery of the database, nor does it show new or changed records in the form's underlying record source. You can use the  **[Requery](form-requery-method-access.md)** method to requery the source of data for the form or one of its controls.


 **Note**  


## Example

The following example uses the  **Repaint** method to repaint a form when the form receives the focus:


```vb
Private Sub Form_Activate() 
    Me.Repaint 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

