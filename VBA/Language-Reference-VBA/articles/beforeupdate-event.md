---
title: BeforeUpdate Event
keywords: fm20.chm2000050
f1_keywords:
- fm20.chm2000050
ms.prod: office
api_name:
- Office.BeforeUpdate
ms.assetid: ccf0fa5d-a069-cba6-5725-072b141fa80b
ms.date: 06/08/2017
---


# BeforeUpdate Event



Occurs before data in a control is changed.
 <strong>Syntax</strong>
 
<strong>Private Sub</strong><em>object</em> <em><strong>BeforeUpdate( ByVal</strong>_Cancel</em><strong>As MSForms.ReturnBoolean)</strong>
The  
<strong>BeforeUpdate</strong> event syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                  |
|:----------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                                                                                                                                                                                     |
| <em>Cancel</em>       | Required. Event status.  <strong>False</strong> indicates that the control should handle the event (default). <strong>True</strong> cancels the update and indicates the application should handle the event. |

 **Remarks**
The BeforeUpdate event occurs regardless of whether the control is [bound](glossary-vba.md) (that is, when the **RowSource** property specifies a[data source](glossary-vba.md) for the control). This event occurs before the AfterUpdate and Exit events for the control (and before the Enter event for the next control that receives[focus](vbe-glossary.md)).
If you set the  _Cancel_ argument to **True**, the focus remains on the control and neither the AfterUpdate event nor the Exit event occurs.

