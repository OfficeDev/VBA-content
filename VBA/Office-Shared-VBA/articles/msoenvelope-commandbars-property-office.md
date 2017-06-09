---
title: MsoEnvelope.CommandBars Property (Office)
keywords: vbaof11.chm11005
f1_keywords:
- vbaof11.chm11005
ms.prod: office
api_name:
- Office.MsoEnvelope.CommandBars
ms.assetid: ac2a7180-044a-e945-98f9-1d2fa76e7cb8
ms.date: 06/08/2017
---


# MsoEnvelope.CommandBars Property (Office)

Gets a  **CommandBars** collection. Read-only.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **CommandBars**

 _expression_ A variable that represents a **MsoEnvelope** object.


## Example

The following example return the  **CommandBars** collection from the **MsoEnvelope** object in Microsoft Word.


```
Dim cbars As CommandBars 
Set cbars = Application.ActiveDocument.MailEnvelope.Commandbars 

```


## See also


#### Concepts


[MsoEnvelope Object](msoenvelope-object-office.md)
#### Other resources


[MsoEnvelope Object Members](msoenvelope-members-office.md)

