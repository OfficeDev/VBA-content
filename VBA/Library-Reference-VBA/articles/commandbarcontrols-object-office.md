---
title: CommandBarControls Object (Office)
keywords: vbaof11.chm4000
f1_keywords:
- vbaof11.chm4000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBarControls
ms.assetid: 7ccae243-2870-95c2-1e08-140a3e638fe6
---


# CommandBarControls Object (Office)

A collection of  **CommandBarControl** objects that represent the command bar controls on a command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Example

Use the  **Controls** property to return the **CommandBarControls** collection. The following example changes the caption of every control on the toolbar named "Standard" to the current value of the **Id** property for that control.


```
For Each ctl In CommandBars("Standard").Controls 
    ctl.Caption = CStr(ctl.Id) 
Next ctl
```

Use the  **Add** method to add a new command bar control to the **CommandBarControls** collection. This example adds a new, blank button to the command bar named "Custom."




```
Set myBlankBtn = CommandBars("Custom").Controls.Add
```

Use Controls(index), where  _index_ is the caption or index number of a control, to return a **CommandBarControl**, **CommandBarButton**, **CommandBarComboBox**, or **CommandBarPopup** object. The following example copies the first control from the command bar named "Standard" to the command bar named "Custom."




```
Set myCustomBar = CommandBars("Custom") 
Set myControl = CommandBars("Standard").Controls(1) 
myControl.Copy Bar:=myCustomBar, Before:=1
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/53e2b0b9-b11a-bf52-a1a3-523aae2c35d8%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/822f709a-fe54-cca4-49d1-6a79d2eb15e5%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/1c2b4afd-2b31-bcee-53b5-6d9761203be1%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/d1728427-b84d-f313-ef73-e234571f3be6%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/a2e7339c-bf1e-0c58-c28d-19cf5682291a%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/23fdc1d0-ffb4-04a2-55d6-9490dd9e795c%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[CommandBarControls Object Members](http://msdn.microsoft.com/library/b4db50d1-f693-d4a5-da6d-41c6f624bdd3%28Office.15%29.aspx)
