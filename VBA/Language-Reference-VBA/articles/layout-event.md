---
title: Layout Event
keywords: fm20.chm5224946
f1_keywords:
- fm20.chm5224946
ms.prod: office
api_name:
- Office.Layout
ms.assetid: ae4e356a-3283-e6a0-ac29-25327ff7c3df
ms.date: 06/08/2017
---


# Layout Event



Occurs when a form,  <strong>Frame</strong>, or <strong>MultiPage</strong> changes size.
 
<strong>Syntax</strong>
For MultiPage 
<strong>Private Sub</strong><em>object</em> <em><strong>Layout(</strong>_index</em><strong>As Long)</strong>
For all other controls 
<strong>Private Sub</strong><em>object</em> _<strong>Layout( )</strong>
The  
<strong>Layout</strong> event syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                        |
|:----------------------|:------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                                                           |
| <em>index</em>        | Required. The index of the page in a  <strong>MultiPage</strong> that changed size. |

 **Remarks**
The default action of the layout event is to calculate new positions of controls and to repaint the screen.
A user can initiate the Layout event by changing the size of a control.
For controls that support the  **AutoSize** property, the Layout event is initiated when **AutoSize** changes the size of the control. This occurs when the user changes the value of a property that affects the size of a control. For example, increasing the **Font** size of a **TextBox** or **Label** can significantly change the dimensions of the control and initiate a Layout event.

