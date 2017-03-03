---
title: CommandBarPopup Object (Office)
keywords: vbaof11.chm7000
f1_keywords:
- vbaof11.chm7000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBarPopup
ms.assetid: a8ae06a3-1d7b-a531-91df-756fafee5314
---


# CommandBarPopup Object (Office)

Represents a pop-up control on a command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Remarks

Every pop-up control contains a  **CommandBar** object. To return the command bar from a pop-up control, apply the **CommandBar** property to the **CommandBarPopup** object.

 Use Controls(index), where _index_ is the number of the control, to return a **CommandBarPopup** object. Note that the **Type** property of the control must be **msoControlPopup**, **msoControlGraphicPopup**, **msoControlButtonPopup**, **msoControlSplitButtonPopup**, or **msoControlSplitButtonMRUPopup**.


## Example

You can also use the  **FindControl** method to return a **CommandBarPopup** object. The following example searches all command bars for a **CommandBarPopup** object whose tag is "Graphics."


```
Set myControl = Application.CommandBars.FindControl _ 
(Type:=msoControlPopup, Tag:="Graphics")
```


## Methods



|**Name**|
|:-----|
|[Copy](http://msdn.microsoft.com/library/d50fff50-00fd-e70f-d777-9bf1850cae37%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/65ec78a1-9f8f-fbd7-3611-c788f3e8566d%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/fedebe76-86f5-9c30-6e23-a20e0024bbf4%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/8c36e21d-0693-63c7-4f27-b1f333d240d9%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/8e31b4e2-66d1-b902-f837-dc4833b1607f%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/ce132a0d-aa1f-c8b1-2697-1cfe78b99123%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/2a247386-f2f3-5901-038a-677a4906cb82%28Office.15%29.aspx)|
|[BeginGroup](http://msdn.microsoft.com/library/0ecc5c98-5db7-792c-8f33-86f7df32d912%28Office.15%29.aspx)|
|[BuiltIn](http://msdn.microsoft.com/library/7cf5322a-b970-39da-c200-fc8303d60f29%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/fc9221e6-cfb0-9f2a-290b-73a434569e65%28Office.15%29.aspx)|
|[CommandBar](http://msdn.microsoft.com/library/e78abe18-d260-8cac-d647-322b449e4bbb%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/0b116a89-f4a8-8043-0c0c-c64eb07a3941%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/14af9c69-394c-9547-ac79-6bc1bc7f01c1%28Office.15%29.aspx)|
|[DescriptionText](http://msdn.microsoft.com/library/81a6b11d-40ea-d17d-4a28-ca423a3e29ec%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/d56d2e1d-27b3-f375-95aa-9efa3aa4d734%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/38692baa-5b41-6f38-305c-33eb1aa5f5df%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/b07d39b7-9fad-51dc-b093-de88cd1ea905%28Office.15%29.aspx)|
|[HelpFile](http://msdn.microsoft.com/library/67c79cb5-cca7-d113-49de-9f636c757867%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/7bddc643-ec4f-7fa5-d5e4-a4677cf564fa%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/6f6f6d1f-a59a-cf52-d273-a732652b4f05%28Office.15%29.aspx)|
|[IsPriorityDropped](http://msdn.microsoft.com/library/2f4846a0-d435-df3c-903c-050b0e31d19d%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/d384480a-9777-acee-d943-ec4ebb6cb5e7%28Office.15%29.aspx)|
|[OLEMenuGroup](http://msdn.microsoft.com/library/32b1bc39-19bc-d0ed-59b5-2e7fa03f329e%28Office.15%29.aspx)|
|[OLEUsage](http://msdn.microsoft.com/library/75d338e0-f5ca-f4b6-2f94-e575749e6ae9%28Office.15%29.aspx)|
|[OnAction](http://msdn.microsoft.com/library/47511647-5f1f-5e40-179b-ec589a2c39be%28Office.15%29.aspx)|
|[Parameter](http://msdn.microsoft.com/library/3ad7783e-3afd-0019-1cf9-eae93992479b%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/1bb8a043-1ad2-28d2-8c48-8426ef24579e%28Office.15%29.aspx)|
|[Priority](http://msdn.microsoft.com/library/cef115fd-fdc8-d8a3-b51d-c9fbc21a810f%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/161b54b5-e7e6-123b-1d68-244d2b64230e%28Office.15%29.aspx)|
|[TooltipText](http://msdn.microsoft.com/library/4b2d39b5-3fcd-0478-51ae-098094a8a4c6%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/8949a41f-3772-be86-d794-002c680a4ade%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/1ef5e542-7fa6-1527-26d0-cf8a6c755979%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/03b74aed-4f36-c45b-a490-a7143542307e%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/a80aaeb8-c633-215b-bd28-8d25fa97dcc9%28Office.15%29.aspx)|

## See also


#### Other resources


[CommandBarPopup Object Members](http://msdn.microsoft.com/library/8ec16deb-bb74-2871-d837-f706c7a58f2b%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
