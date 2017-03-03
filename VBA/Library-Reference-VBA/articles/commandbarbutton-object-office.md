---
title: CommandBarButton Object (Office)
keywords: vbaof11.chm244000
f1_keywords:
- vbaof11.chm244000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBarButton
ms.assetid: e6d8209d-2c87-f1b5-bc3f-d4e5e5d3ab73
---


# CommandBarButton Object (Office)

Represents a button control on a command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Example

Use  **Controls(index)**, where _index_ is the index number of the control, to return a **CommandBarButton** object. Note that the **Type** property of the control must be **msoControlButton**. Assuming that the second control on the command bar named "Custom" is a button, the following example changes the style of that button.


```
Set c = CommandBars("Custom").Controls(2) 
With c 
If .Type = msoControlButton Then 
    If .Style = msoButtonIcon Then 
        .Style = msoButtonIconAndCaption 
    Else 
        .Style = msoButtonIcon 
    End If 
End If 
End With
```


 **Note**  


 **Note**  You can also use the  **FindControl** method to return a **CommandBarButton** object.


## Events



|**Name**|
|:-----|
|[Click](http://msdn.microsoft.com/library/d4f970e6-8c37-c5cc-a0b4-4efe213a2e05%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Copy](http://msdn.microsoft.com/library/a78a7922-aa51-7b9f-d7de-a227a6869140%28Office.15%29.aspx)|
|[CopyFace](http://msdn.microsoft.com/library/09f09dbd-b70f-8b7d-1af7-7e43bffe3030%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/af94a209-b651-442f-8fa3-3a6436833d15%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/1cf36559-86ba-8a9c-ef81-ef72185dd21c%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/b2d462ec-63a7-a395-8d93-bedbf1d6941d%28Office.15%29.aspx)|
|[PasteFace](http://msdn.microsoft.com/library/1c4179c4-b6b5-527f-5027-25ced8ee907d%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/0e39c960-3928-f91a-cf7e-1df5a2fd217b%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/f6719533-1958-05d4-5f9c-7b09cb33b1c8%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/c15d6f7e-c728-0e8a-9c56-c8b4cd59822a%28Office.15%29.aspx)|
|[BeginGroup](http://msdn.microsoft.com/library/62f522cd-30de-85a6-bd2d-0bd3f6ccb44f%28Office.15%29.aspx)|
|[BuiltIn](http://msdn.microsoft.com/library/0a159c65-99d1-efdf-ec5c-f4e51060dd09%28Office.15%29.aspx)|
|[BuiltInFace](http://msdn.microsoft.com/library/47c82878-17ea-b6ff-e841-c9f07342c8a3%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/1147e08a-b9f4-3ea9-3a86-d13394aa1959%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/9c54fa96-8c97-fcae-067f-e8511560a15f%28Office.15%29.aspx)|
|[DescriptionText](http://msdn.microsoft.com/library/bc22bef9-e923-40af-296b-959f3f3aeead%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/264335ca-6506-0e86-16df-44af277ade83%28Office.15%29.aspx)|
|[FaceId](http://msdn.microsoft.com/library/c2151f20-b1c7-97eb-35ac-7a12c5ee3f28%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/b374ae8b-cce2-7562-1247-32ea90dc3c68%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/2e4f33db-7143-dd8d-65b3-d0c993f2e966%28Office.15%29.aspx)|
|[HelpFile](http://msdn.microsoft.com/library/6e97a52d-f50d-600b-26eb-b22988bd5ed5%28Office.15%29.aspx)|
|[HyperlinkType](http://msdn.microsoft.com/library/5769ce22-a9e8-3eb2-919f-a3d016cf0706%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/d559a98c-b9b2-a987-c7af-278734a9545d%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/2924d346-735b-cdb3-6237-f840f017cf3e%28Office.15%29.aspx)|
|[IsPriorityDropped](http://msdn.microsoft.com/library/68398973-675f-2180-b22c-4ad5de0582f7%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/0a3a83ce-bbb5-1884-4125-0d9f1bf20d27%28Office.15%29.aspx)|
|[Mask](http://msdn.microsoft.com/library/de7179ac-6b39-2323-d84a-23abe3ed3167%28Office.15%29.aspx)|
|[OLEUsage](http://msdn.microsoft.com/library/4ff6f74d-4eed-8a30-468c-22be5dee1c7e%28Office.15%29.aspx)|
|[OnAction](http://msdn.microsoft.com/library/c0a4148c-330a-6bd9-dd14-7ade8fc833fe%28Office.15%29.aspx)|
|[Parameter](http://msdn.microsoft.com/library/582718f1-8274-9862-c9a8-86bcd1c528b7%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/1238aea6-0a4c-0af7-7fc4-6c5fd2627b78%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/b9a2d133-23a8-ac09-8b8b-08eda1210717%28Office.15%29.aspx)|
|[Priority](http://msdn.microsoft.com/library/72599580-16d2-20b3-05ad-b454afbba6ef%28Office.15%29.aspx)|
|[ShortcutText](http://msdn.microsoft.com/library/e0c76e70-16db-d3ae-9767-069579c8ea91%28Office.15%29.aspx)|
|[State](http://msdn.microsoft.com/library/919ca064-507c-1db6-6b69-b586283ab67b%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/5a9d5a5e-8893-14db-71f2-e007e1f9249f%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/c73a12a8-8b20-1e32-ad98-ae0bb3b1daed%28Office.15%29.aspx)|
|[TooltipText](http://msdn.microsoft.com/library/12126126-f8b6-e8a4-3d32-4d5604928e8a%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/4ad019ed-a344-dac5-0063-b52bdead7916%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/f317eb14-a5d6-857e-6b6b-89391937db96%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/121d4c6d-141d-882d-c77e-2ed9357c9445%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/f0e3f562-214b-4c0c-b239-611e710349e1%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[CommandBarButton Object Members](http://msdn.microsoft.com/library/69fe57fe-dabc-9379-283c-d0a51a775592%28Office.15%29.aspx)
