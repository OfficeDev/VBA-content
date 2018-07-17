---
title: MsoEnvelope.EnvelopeHide Event (Office)
keywords: vbaof11.chm246002
f1_keywords:
- vbaof11.chm246002
ms.prod: office
api_name:
- Office.MsoEnvelope.EnvelopeHide
ms.assetid: 066b0ed0-bd5d-f37e-6c04-66e0a59733d4
ms.date: 06/08/2017
---


# MsoEnvelope.EnvelopeHide Event (Office)

Occurs when the user interface (UI) that corresponds to the  **MsoEnvelope** object is hidden.


## Syntax

 _expression_. **EnvelopeHide**

 _expression_ An expression that returns a **MsoEnvelope** object.


## Remarks

The  **MsoEnvelope** object provides access to functionality that lets you send documents as e-mail messages directly from Microsoft Office applications.


## Example

The following example sets up event-handling routines for the  **MsoEnvelope** object.


```
Public WithEvents env As MsoEnvelope 
 
Private Sub Class_Initialize() 
 Set env = Application.ActiveDocument.MailEnvelope 
End Sub 
 
Private Sub env_EnvelopeShow() 
 MsgBox "The MsoEnvelope UI is showing." 
End Sub 
 
Private Sub env_EnvelopeHide() 
 MsgBox "The MsoEnvelope UI is hidden." 
End Sub 

```


## See also


#### Concepts


[MsoEnvelope Object](msoenvelope-object-office.md)
#### Other resources


[MsoEnvelope Object Members](msoenvelope-members-office.md)

