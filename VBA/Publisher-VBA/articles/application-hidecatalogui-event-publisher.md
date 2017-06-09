---
title: Application.HideCatalogUI Event (Publisher)
keywords: vbapb10.chm268435494
f1_keywords:
- vbapb10.chm268435494
ms.prod: publisher
api_name:
- Publisher.Application.HideCatalogUI
ms.assetid: a7ac7594-18fe-355e-d270-d205c405862a
ms.date: 06/08/2017
---


# Application.HideCatalogUI Event (Publisher)

Occurs when the catalog of publication wizards is hidden in the Microsoft Publisher user interface.


## Syntax

 _expression_. **HideCatalogUI**

 _expression_An expression that returns a  **Application** object.


## Remarks

For more information about using events with the  **Application** object, see [Using Events with the Application Object](using-events-with-the-application-object-publisher.md).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to handle the  **HideCatalogUI** event. It displays a message notifying the user that the catalog UI is hidden.


```vb
Private Sub pubApplication_HideCatalogUI() 
 MsgBox "The Wizard Catalog is hidden." 
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

