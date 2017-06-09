---
title: Form.HasModule Property (Access)
keywords: vbaac10.chm13478
f1_keywords:
- vbaac10.chm13478
ms.prod: access
api_name:
- Access.Form.HasModule
ms.assetid: ba43a8c8-89f2-e744-ed99-082510dc8f3a
ms.date: 06/08/2017
---


# Form.HasModule Property (Access)

You can use the  **HasModule** property to specify or determine whether a form or report has a class module. Read/write **Boolean**.


## Syntax

 _expression_. **HasModule**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **HasModule** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|The form or report has a class module.|
|No|**False**|(Default) The form or report doesn't have a class module.|
Setting this property to  **False** can improve the performance and decrease the size of your database.

The  **HasModule** property can be set only in form or report Design view but can be read in any view.

Forms or reports that have the  **HasModule** property set to No are considered lightweight objects. Lightweight objects are smaller and typically load and display faster than objects with associated class modules. In many cases, a form or report doesn't need to use event procedures and doesn't require a class module.

If your application uses a switchboard form to navigate to other forms, instead of using command buttons with event procedures, you can use a command button with a macro or hyperlink. For example, to open the Employees form from a command button on a switchboard, you can set the control's  **HyperlinkSubAddress** property to "Form Employees".

Lightweight objects don't appear in the Object Browser and you can't use the  **New** keyword to create a new instance of the object. A lightweight form or report can be used as a subform or subreport and will appear in the **[Forms](forms-object-access.md)** or **[Reports](reports-object-access.md)** collection. Lightweight objects support the use of macros, and public procedures that exist in standard modules when called from the object's property sheet.

Microsoft Access sets the  **HasModule** property to **True** as soon as you attempt to view an object's module, even if no code is actually added to the module. For example, selecting **Code** from the **View** menu for a form in Design view causes Microsoft Access to add a class module to the **[Form](form-object-access.md)** object and set its **HasModule** property to **True**. You can add a class module to an object in the same way by setting the **HasModule** property to Yes in an object's property sheet.

If you set the  **HasModule** property to No by using an object's property sheet or set it to **False** by using Visual Basic, Microsoft Access deletes the object's class module and any code it may contain.

When you use a method of the  **[Module](module-object-access.md)** object or refer to the **Module** property for a form or report in Design view, Microsoft Access creates the associated module and sets the object's **HasModule** property to **True**. If you refer to the **Module** property of a form or report at run-time and the object has its **HasModule** property set to **False**, an error will occur.

Objects created by using the  **CreateForm** or **CreateReport** methods are lightweight by default.


## See also


#### Concepts


[Form Object](form-object-access.md)

