---
title: CommandBarButton.HyperlinkType Property (Office)
keywords: vbaof11.chm6008
f1_keywords:
- vbaof11.chm6008
ms.prod: office
api_name:
- Office.CommandBarButton.HyperlinkType
ms.assetid: 5769ce22-a9e8-3eb2-919f-a3d016cf0706
ms.date: 06/08/2017
---


# CommandBarButton.HyperlinkType Property (Office)

Sets or gets a  **MsoCommandBarButtonHyperlinkType** constant that represents the type of hyperlink associated with the specified command bar button. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **HyperlinkType**

 _expression_ A variable that represents a **CommandBarButton** object.


## Example

This example checks the  **HyperlinkType** property for the specified command bar button on the command bar named "Custom.". If **HyperlinkType** is set to **msoCommandBarButtonHyperlinkNone**, the example sets the property to **msoCommandBarButtonHyperlinkOpen** and sets the URL to www.microsoft.com.


```
Set myBar = CommandBars _ 
    .Add(Name:="Custom", Position:=msoBarTop, _ 
    Temporary:=True) 
Set myButton = myBar.Controls.Add(Type:=msoControlButton) 
With myButton 
    .FaceId = 277 
    .HyperlinkType = msoCommandBarButtonHyperlinkNone 
End With 
If myButton.HyperlinkType > _ 
    msoCommandBarButtonHyperlinkOpen Then 
    myButton.HyperlinkType = _ 
        msoCommandBarButtonHyperlinkOpen 
    myButton.TooltipText = "www.microsoft.com" 
End If
```


## See also


#### Concepts


[CommandBarButton Object](commandbarbutton-object-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

