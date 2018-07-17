---
title: "How to: Add Scroll Bars to a Page and a Frame"
keywords: olfm10.chm3077235
f1_keywords:
- olfm10.chm3077235
ms.prod: outlook
ms.assetid: 2fdc2fb5-0ee8-b39e-f4a7-c898244b13ac
ms.date: 06/08/2017
---


# How to: Add Scroll Bars to a Page and a Frame

The following example uses the  **ScrollBars** and the **KeepScrollBarsVisible** properties to add scroll bars to a page of a **[MultiPage](multipage-object-outlook-forms-script.md)** and to a **[Frame](frame-object-outlook-forms-script.md)**. The user chooses an option button that, in turn, specifies a value for  **KeepScrollBarsVisible**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **MultiPage** named MultiPage1.
    
- A  **Frame** named Frame1.
    
- Four  **[OptionButton](optionbutton-object-outlook-forms-script.md)** controls named OptionButton1 through OptionButton4.
    



```vb
Sub Item_Open() 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set Frame1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Frame1") 
 Set OptionButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton1") 
 Set OptionButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton2") 
 Set OptionButton3 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton3") 
 Set OptionButton4 = Item.GetInspector.ModifiedFormPages("P.2").Controls("OptionButton4") 
 
 MultiPage1.Pages(0).ScrollBars = 3 '3=fmScrollBarsBoth 
 MultiPage1.Pages(0).KeepScrollBarsVisible = 0 '0=fmScrollBarsNone 
 
 Frame1.ScrollBars = 3 '3=fmScrollBarsBoth 
 Frame1.KeepScrollBarsVisible = 0 '0=fmScrollBarsNone 
 
 OptionButton1.Caption = "No scroll bars" 
 OptionButton1.Value = True 
 OptionButton2.Caption = "Horizontal scroll bars" 
 OptionButton3.Caption = "Vertical scroll bars" 
 OptionButton4.Caption = "Both scroll bars" 
End Sub 
 
Sub OptionButton1_Click() 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set Frame1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Frame1") 
 
 MultiPage1.Pages(0).KeepScrollBarsVisible = 0 '0=fmScrollBarsNone 
 Frame1.KeepScrollBarsVisible = 0 '0=fmScrollBarsNonefmScrollBarsNone 
End Sub 
 
Sub OptionButton2_Click() 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set Frame1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Frame1") 
 
 MultiPage1.Pages(0).KeepScrollBarsVisible = 1 '1=fmScrollBarsHorizontal 
 Frame1.KeepScrollBarsVisible = 1 '1=fmScrollBarsHorizontal 
End Sub 
 
Sub OptionButton3_Click() 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set Frame1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Frame1") 
 
 MultiPage1.Pages(0).KeepScrollBarsVisible = 2 '2=fmScrollBarsVertical 
 Frame1.KeepScrollBarsVisible = 2 '2=fmScrollBarsVertical 
End Sub 
 
Sub OptionButton4_Click() 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set Frame1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Frame1") 
 
 MultiPage1.Pages(0).KeepScrollBarsVisible = 3 '3=fmScrollBarsBoth 
 Frame1.KeepScrollBarsVisible = 3 '3=fmScrollBarsBoth 
End Sub
```


