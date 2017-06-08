---
title: "How to: Group Option Buttons"
keywords: olfm10.chm3077190
f1_keywords:
- olfm10.chm3077190
ms.prod: outlook
ms.assetid: ecf72f77-585b-c493-bcc4-35eb4f11e62a
ms.date: 06/08/2017
---


# How to: Group Option Buttons

The following example uses the  **[GroupName](optionbutton-groupname-property-outlook-forms-script.md)** property to create two groups of **[OptionButton](optionbutton-object-outlook-forms-script.md)** controls on the same form.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains five **OptionButton** controls named OptionButton1 through OptionButton5.



```vb
Sub Item_Open() 
 set OptionButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton1") 
 set OptionButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton2") 
 set OptionButton3 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton3") 
 set OptionButton4 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton4") 
 set OptionButton5 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton5") 
 
 OptionButton1.Caption = "Widgets" 
 OptionButton2.Caption = "Widgets" 
 OptionButton3.Caption = "Widgets" 
 OptionButton1.GroupName = "Widgets" 
 OptionButton2.GroupName = "Widgets" 
 OptionButton3.GroupName = "Widgets" 
 
 OptionButton4.Caption = "Gadgets-Group2" 
 OptionButton5.Caption = "Gadgets-Group2" 
 OptionButton4.GroupName = "Gadgets-Group2" 
 OptionButton5.GroupName = "Gadgets-Group2" 
End Sub
```


