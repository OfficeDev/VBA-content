---
title: RemoveControl Event
keywords: fm20.chm2000200
f1_keywords:
- fm20.chm2000200
ms.prod: office
api_name:
- Office.RemoveControl
ms.assetid: 6e6abe85-4c0c-8fc9-668c-009e6f1a3d76
ms.date: 06/08/2017
---


# RemoveControl Event



Occurs when a control is deleted from the [container](vbe-glossary.md).
 <strong>Syntax</strong>
For MultiPage 
<strong>Private Sub</strong><em>object</em> <em><strong>RemoveControl(</strong>_index</em><strong>As Long</strong>, <em>ctrl</em><strong>As Control)</strong>
For all other controls 
<strong>Private Sub</strong><em>object</em> <em><strong>RemoveControl(</strong>_ctrl</em><strong>As Control)</strong>
The  
<strong>RemoveControl</strong> event syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                         |
|:----------------------|:-----------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object name.                                                                       |
| <em>index</em>        | Required. The index of the page in a  <strong>MultiPage</strong> that contained the deleted control. |
| <em>ctrl</em>         | Required. The deleted control.                                                                       |

 **Remarks**
This event occurs when a control is deleted from the form, not when a control is unloaded due to a form being closed.

