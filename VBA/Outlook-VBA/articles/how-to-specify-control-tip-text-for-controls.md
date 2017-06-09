---
title: "How to: Specify Control Tip Text for Controls"
keywords: olfm10.chm3077167
f1_keywords:
- olfm10.chm3077167
ms.prod: outlook
ms.assetid: 50ea26b3-763b-beed-6f06-30afbd205f02
ms.date: 06/08/2017
---


# How to: Specify Control Tip Text for Controls

The following example defines the  **[ControlTipText](page-controltiptext-property-outlook-forms-script.md)** property for three **[CommandButton](commandbutton-object-outlook-forms-script.md)** controls and two **[Page](page-object-outlook-forms-script.md)** objects in a **[MultiPage](multipage-object-outlook-forms-script.md)**.


 **Note**  The Microsoft Forms 2.0  **CommandButton** control includes the **ControlTipText** property.


To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:


- A  **MultiPage** named MultiPage1.
    
- Three  **CommandButton** controls named CommandButton1 through CommandButton3.
    

 **Note**  For an individual  **Page** of a **MultiPage**,  **ControlTipText** becomes enabled when the **MultiPage** or a control on the current page of the **MultiPage** has the focus.




```vb
Sub Item_Open() 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").MultiPage1 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").CommandButton1 
 Set CommandButton2 = Item.GetInspector.ModifiedFormPages("P.2").CommandButton2 
 Set CommandButton3 = Item.GetInspector.ModifiedFormPages("P.2").CommandButton3 
 
 MultiPage1.Page1.ControlTipText = "Here in page 1" 
 MultiPage1.Page2.ControlTipText = "Now in page 2" 
 
 CommandButton1.ControlTipText = "And now here's" 
 CommandButton2.ControlTipText = "a tip from" 
 CommandButton3.ControlTipText = "your controls!" 
End Sub
```


