---
title: CommandBarComboBox Object (Office)
keywords: vbaof11.chm243000
f1_keywords:
- vbaof11.chm243000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBarComboBox
ms.assetid: fcfe6bde-dea0-f1f1-ad30-d0e28f97dd07
---


# CommandBarComboBox Object (Office)

Represents a combo box control on a command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Remarks

Use  **Controls(index)**, where _index_ is the index number of the control, to return a **CommandBarComboBox** object. Note that the **Type** property of the control must be **msoControlEdit**, **msoControlDropdown**, **msoControlComboBox**, **msoControlButtonDropdown**, **msoControlSplitDropdown**, **msoControlOCXDropdown**, **msoControlGraphicCombo**, or **msoControlGraphicDropdown**.


## Example

The following example adds two items to the second control on the command bar named  **Custom**, and then it adjusts the size of the control.


```
Set combo = CommandBars("Custom").Controls(2) 
With combo 
    .AddItem "First Item", 1 
    .AddItem "Second Item", 2 
    .DropDownLines = 3 
    .DropDownWidth = 75 
    .ListIndex = 0 
End With
```

You can also use the  **FindControl** method to return a **CommandBarComboBox** object. The following example searches all command bars for a visible **CommandBarComboBox** object whose tag is "sheet assignments."




```
Set myControl = CommandBars.FindControl _ 
(Type:=msoControlComboBox, Tag:="sheet assignments", Visible:=True)
```


## Events



|**Name**|
|:-----|
|[Change](http://msdn.microsoft.com/library/ddf1a306-c299-36d5-9851-04d6e5185db9%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddItem](http://msdn.microsoft.com/library/66109c4e-a75b-ebca-99e8-b6848316a04f%28Office.15%29.aspx)|
|[Clear](http://msdn.microsoft.com/library/f60afda8-5740-c6f6-7f3b-315dc95c45f8%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/15eb757c-bb07-cd98-ff9e-1810db4f475c%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/7b84c512-24e2-f159-100b-5234fc78fcf0%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/13ec7924-2420-c0c0-750f-4dae8b8e1503%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/8e8ccbee-da72-1167-9f34-ccf5b535fef8%28Office.15%29.aspx)|
|[RemoveItem](http://msdn.microsoft.com/library/8a40dcca-c320-c27f-ae91-97c195d4f821%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/28609b13-8036-a956-095a-1a6a748f00ad%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/3170651c-40da-5025-8b36-195b836c8fcb%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/6d9790dd-d418-6287-06f9-27214a564dd9%28Office.15%29.aspx)|
|[BeginGroup](http://msdn.microsoft.com/library/482ec5fc-91ef-746b-2ec8-360bb7780df2%28Office.15%29.aspx)|
|[BuiltIn](http://msdn.microsoft.com/library/4dc0232c-94dd-ce40-95cd-7700fdd9a427%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/71c317d3-f3b5-da32-1db8-0fb5bd4ba8f2%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/c2c814c7-a99f-909e-8edc-16d103fd6837%28Office.15%29.aspx)|
|[DescriptionText](http://msdn.microsoft.com/library/e06b5800-eecd-6863-68f7-9b88d3c4696b%28Office.15%29.aspx)|
|[DropDownLines](http://msdn.microsoft.com/library/715bbec9-1bd6-c7b0-0d1e-e57d61689d52%28Office.15%29.aspx)|
|[DropDownWidth](http://msdn.microsoft.com/library/051ac285-c7f1-a2b7-0c9a-ed2cb08cadc9%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/f88401a5-b180-63e5-e301-a60addaacab4%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/a3afc8c0-1c35-acc0-905c-0af47e84827d%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/3b34572b-af1b-a4fc-a98e-23d51315a077%28Office.15%29.aspx)|
|[HelpFile](http://msdn.microsoft.com/library/2fbe2d70-b8f7-d800-ed46-0ac88125b8f1%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/9cc143cb-4063-b397-05c9-d50a7c2efcb0%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/a844b760-d165-02aa-41ad-0bc75c55d0ed%28Office.15%29.aspx)|
|[IsPriorityDropped](http://msdn.microsoft.com/library/c556f630-5e95-6d1a-4e94-0ecf5b20875a%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/01dc5c7c-4fc6-a2fe-fa27-c24ed0802dd6%28Office.15%29.aspx)|
|[List](http://msdn.microsoft.com/library/c90fae92-daab-1b08-6e85-8caae26d0b72%28Office.15%29.aspx)|
|[ListCount](http://msdn.microsoft.com/library/3ab55501-b82e-0380-d805-e4386c399131%28Office.15%29.aspx)|
|[ListHeaderCount](http://msdn.microsoft.com/library/54625ef5-2e09-5a39-7909-e775c4e9e0c4%28Office.15%29.aspx)|
|[ListIndex](http://msdn.microsoft.com/library/3267a20a-7b33-3a89-5def-46c8b9756c04%28Office.15%29.aspx)|
|[OLEUsage](http://msdn.microsoft.com/library/3da25257-6ffe-a00e-bada-79c6245286b7%28Office.15%29.aspx)|
|[OnAction](http://msdn.microsoft.com/library/fe666bce-9c38-4203-1059-343d1346913b%28Office.15%29.aspx)|
|[Parameter](http://msdn.microsoft.com/library/b5019fba-5124-5d9c-7abe-db10df32078b%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/a4dc7231-5366-2504-f9b0-af6dd1728bfa%28Office.15%29.aspx)|
|[Priority](http://msdn.microsoft.com/library/0166df8f-316a-8414-a3af-1156fc1a1166%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/04d1270f-23b6-da23-312c-cb75c8969864%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/0bc1957b-aa17-aaa6-e416-26db0a34f342%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/91aa73ff-260c-c241-35d0-50bebbbaf190%28Office.15%29.aspx)|
|[TooltipText](http://msdn.microsoft.com/library/65bfb3ff-a36e-dfd5-4ae0-4d2ccfb69000%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/f49930ca-9dba-9d9b-b7bb-93de87cdfcf8%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/1f8d29ac-f429-7190-f5b9-76eb0aa5a0be%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/d3fa2bfe-10ea-70d7-40f9-bf757fff6e27%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/5efb8c56-f896-c5e7-d457-f8862e655d1c%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[CommandBarComboBox Object Members](http://msdn.microsoft.com/library/223c51c0-4564-d14a-a8bf-d315a6a50b32%28Office.15%29.aspx)
