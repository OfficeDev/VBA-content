---
title: Application.AfterPrint Event (Publisher)
keywords: vbapb10.chm268435492
f1_keywords:
- vbapb10.chm268435492
ms.prod: publisher
api_name:
- Publisher.Application.AfterPrint
ms.assetid: ddd5a1a4-8130-9e75-039c-e069a37390e8
ms.date: 06/08/2017
---


# Application.AfterPrint Event (Publisher)

Fires after all variables and fields print.


## Syntax

 _expression_. **AfterPrint**( **_Doc_**)

 _expression_An expression that returns an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Doc|Required| **Document**|The current publication.|

## Remarks

Microsoft Publisher does not return UI control to the user until the event handler is executed. The event is called after all the drawing operations are completed (in other words, after the software's job finishes and the printing hardware takes over).

For more information about using events with the  **Application** object, see [Using Events with the Application Object](using-events-with-the-application-object-publisher.md).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to handle the  **AfterPrint** event. It displays a message notifying the user that the document was printed.


```vb
Private Sub pubApplication_AfterPrint(ByVal Doc As Document) 
 MsgBox "Printing of " &; Doc.Name &; "is complete." 
End Sub
```

For this event to occur, you must place the following line of code in the  **General Declarations** section of your module.




```vb
Private WithEvents pubApplication As Application
```

Then run the following initialization procedure.




```vb
Public Sub Initialize_pubApplication() 
 Set pubApplication = Publisher.Application 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

