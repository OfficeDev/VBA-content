---
title: Extra Members for Outlook Controls
ms.prod: outlook
ms.assetid: df52e46d-d4e6-b669-f4bc-7758c9c9d9e2
ms.date: 06/08/2017
---


# Extra Members for Outlook Controls

Outlook controls have access to a set of members that belong to the Microsoft Forms 2.0  **Control** interface. These members include the following properties:


-  ** [ControlTipText](controltiptext-property-outlook-controls.md)**
    
-  ** [Height](height-property-outlook-controls.md)**
    
-  ** [Left](left-property-outlook-controls.md)**
    
-  ** [Name](name-property-outlook-controls.md)**
    
-  ** [TabIndex](tabindex-property-outlook-controls.md)**
    
-  ** [TabStop](tabstop-property-outlook-controls.md)**
    
-  ** [Tag](tag-property-outlook-controls.md)**
    
-  ** [Top](top-property-outlook-controls.md)**
    
-  ** [Visible](visible-property-outlook-controls.md)**
    
-  ** [Width](width-property-outlook-controls.md)**
    



And the following methods:

-  ** [Move](move-method-outlook-controls.md)**
    
-  ** [SetFocus](setfocus-method-outlook-controls.md)**
    
-  ** [ZOrder](zorder-method-outlook-controls.md)**
    

Because these members are not part of the Outlook object model, they are not displayed in the object browser and are not supported by intellisense. However, you can search for specific help topics for these members in the Outlook Developer Help.
To access these members, you can directly reference the member, as in the following example, an  **[OlkTextBox](olktextbox-object-outlook.md)** control, `TextBoxControl`, accesses the  **ControlTipText** property directly with the following line of code.



```
TextBoxControl.ControlTipText = "Enter name of product here"
```

Alternatively, you can add a reference to the Microsoft Forms 2.0 type library (fm20.dll) and bind the Outlook control dynamically at runtime, as in the following code sample.



```vb
Sub AddControlTip() 
 Dim TextBoxControl As OlkTextBox 
 Dim ictrl As MSForms.Control 
 
 Set ictrl = TextBoxControl 
 ictrl.ControlTipText = "Enter product description here" 
End Sub
```


