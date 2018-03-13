---
title: AddControl Event
keywords: fm20.chm2000010
f1_keywords:
- fm20.chm2000010
ms.prod: office
api_name:
- Office.AddControl
ms.assetid: 9febc628-1d26-9ecf-7f04-7c9431a7b9c8
ms.date: 06/08/2017
---


# AddControl Event



Occurs when a control is inserted onto a form, a  <strong>Frame</strong>, or a <strong>Page</strong> of a <strong>MultiPage</strong>.
 
<strong>Syntax</strong>
For Frame 
<strong>Private Sub</strong><em>object</em> <em><strong>AddControl( )</strong>
For MultiPage <strong>Private Sub</strong>_object</em> <em><strong>AddControl(</strong>_index</em><strong>As Long</strong>, <em>ctrl</em><strong>As Control)</strong>
The  
<strong>AddControl</strong> event syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                         |
|:----------------------|:-------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                                                            |
| <em>index</em>        | Required. The index of the  <strong>Page</strong> that will contain the new control. |
| <em>ctrl</em>         | Required. The control to be added.                                                   |

 **Remarks**
The AddControl event occurs when a control is added at [run time](vbe-glossary.md). This event is not initiated when you add a control at [design time](vbe-glossary.md), nor is it initiated when a form is initially loaded and displayed at run time.
The default action of this event is to add a control to the specified form,  **Frame**, or **MultiPage**.
The  **Add** method initiates the AddControl event.

