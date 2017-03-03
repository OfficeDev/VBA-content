---
title: CommandBarControl Object (Office)
keywords: vbaof11.chm5000
f1_keywords:
- vbaof11.chm5000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CommandBarControl
ms.assetid: b104ec00-beeb-a927-4b7b-108f4e3164f5
---


# CommandBarControl Object (Office)

Represents a command bar control. The  **CommandBarControl** object is a member of the **CommandBarControls** collection. The properties and methods of the **CommandBarControl** object are all shared by the **CommandBarButton**, **CommandBarComboBox**, and **CommandBarPopup** objects.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Remarks

When writing Visual Basic code to work with custom command bar controls, you use the  **CommandBarButton**, **CommandBarComboBox**, and **CommandBarPopup** objects. When writing code to work with built-in controls in the container application that cannot be represented by one of those three objects, you use the **CommandBarControl** object. Use **Controls** ( _index_ ), where _index_ is the index number of a control, to return a **CommandBarControl** object. (The **Type** property of the control must be **msoControlLabel**, **msoControlExpandingGrid**, **msoControlSplitExpandingGrid**, **msoControlGrid**, or **msoControlGauge** ). Variables declared as **CommandBarControl** can be assigned **CommandBarButton**, **CommandBarComboBox**, and **CommandBarPopup** values.


## Example

You can also use the  **FindControl** method to return a **CommandBarControl** object. The following example searches for a control of type **msoControlGauge**; if it finds one, it displays the index number of the control and the name of the command bar that contains it. In this example, the variable _lbl_ represents a **CommandBarControl** object.


```
Set lbl = CommandBars.FindControl(Type:= msoControlGauge) 
If lbl Is Nothing Then 
    MsgBox "A control of type msoControlGauge was not found." 
Else 
    MsgBox "Control " &amp; lbl.Index &amp; " on command bar " _ 
        &amp; lbl.Parent.Name &amp; " is type msoControlGauge" 
End If
```


## Methods



|**Name**|
|:-----|
|[Copy](http://msdn.microsoft.com/library/4314de01-8a25-0ab4-582f-7a61f62f8a18%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/eca4abea-092b-0c11-1040-7132318b1bea%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/5b95846f-99c6-93b3-2167-6bd7acf5d508%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/91858a91-49d8-7be6-95b3-491cd9f41235%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/7b2d42c4-ac1c-209e-6fe8-bd5ec91d1c57%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/e20065eb-a1a3-f750-5585-6e38a328b946%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/b89baccc-b6c5-6557-625e-896264f5944e%28Office.15%29.aspx)|
|[BeginGroup](http://msdn.microsoft.com/library/529b8c23-ec1f-b37b-a40c-9ae6016f4dc0%28Office.15%29.aspx)|
|[BuiltIn](http://msdn.microsoft.com/library/4b3904dc-3376-28e0-6c93-4acff8101e6f%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/6e625a77-60a9-eaa5-1d75-f5d8b6688180%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/5c2e361a-fb2b-40c5-b4fb-030734af37e6%28Office.15%29.aspx)|
|[DescriptionText](http://msdn.microsoft.com/library/4f7b8e0d-1f3a-f751-86a7-3378f21ecf3d%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/74105bf5-96a0-09ea-bb00-ef102705372c%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/71dace36-3237-e94a-f45f-7d9718f13a69%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/56f41107-92ad-7cb5-f522-7a338f0d8cf9%28Office.15%29.aspx)|
|[HelpFile](http://msdn.microsoft.com/library/2372698e-1c3b-de8b-b671-356fbd9cad6b%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/0931a07a-4a6b-cc84-a43b-b57ea9a22b78%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/0f4e6561-d53a-ed9d-3d24-7306dbe69bd6%28Office.15%29.aspx)|
|[IsPriorityDropped](http://msdn.microsoft.com/library/cc537dd9-3b10-cba1-d8e0-bdf3952a1e23%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/5af66df7-cfaa-bd98-612e-07be6d0d08c5%28Office.15%29.aspx)|
|[OLEUsage](http://msdn.microsoft.com/library/c3f818a9-7481-0a2f-aa34-5c7e36ea72c1%28Office.15%29.aspx)|
|[OnAction](http://msdn.microsoft.com/library/05e40fcb-ff67-049f-6386-a9ef20b48c87%28Office.15%29.aspx)|
|[Parameter](http://msdn.microsoft.com/library/6a1fd988-0c3f-3945-307f-e4e647c3642c%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/d6727c3d-7666-2339-1271-d44e4545b97c%28Office.15%29.aspx)|
|[Priority](http://msdn.microsoft.com/library/1bb78346-a815-75f8-f2f6-8ecff2b54cbd%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/d528c260-09dc-9cb2-d8ce-8476f91ebc7b%28Office.15%29.aspx)|
|[TooltipText](http://msdn.microsoft.com/library/03e51dbd-0d5a-5094-545f-4a98a6508b4d%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/72513f35-86ec-1fde-b056-6d50c06d8a4c%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/a0f20db6-a8a2-98e2-6f4e-efd9043df0c2%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/9aa5f926-af48-5685-da7f-ea960c4cdbb3%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/a6821638-9cc8-3a9f-ced0-770f50de7d8c%28Office.15%29.aspx)|

## See also


#### Other resources


[CommandBarControl Object Members](http://msdn.microsoft.com/library/1d2360e4-7511-a3a4-9959-2f7c8282bf99%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
