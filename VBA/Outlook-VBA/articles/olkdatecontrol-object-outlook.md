---
title: OlkDateControl Object (Outlook)
keywords: vbaol11.chm1000376
f1_keywords:
- vbaol11.chm1000376
ms.prod: outlook
api_name:
- Outlook.OlkDateControl
ms.assetid: bd0c6bbe-c348-c748-41fe-0cf7ecebcc1e
ms.date: 06/08/2017
---


# OlkDateControl Object (Outlook)

A control that supports the drop-down date picker used in inspectors for task and appointment items to select a date. 


## Remarks

Before you use this control for the first time in the forms designer, add the Microsoft Outlook Date Control to the control toolbox. You can only add this control to a form region in an Outlook form using the forms designer; you cannot add this control to a Visual Basic  **UserForm** object in the Visual Basic Editor.

The following is an example of the date control at runtime. This control supports Microsoft Windows themes.


![Date](images/olDate_ZA10120280.gif)



This control can bind to any built-in or custom  **DateTime** field. However, the control does not support any date format setting for the field, nor does it support the select range behavior that is available in the appointment inspector.

If the  **[Click](http://msdn.microsoft.com/library/ec2483b8-0fe1-de86-dc01-9cafbde31e44%28Office.15%29.aspx)** event is implemented but the **[DropButtonClick](http://msdn.microsoft.com/library/425118d2-afa4-4582-1f89-857e5b7ae903%28Office.15%29.aspx)** event is not implemented, then clicking the drop button will fire only the **Click** event.

For more information about Outlook controls, see [Controls in a Custom Form](http://msdn.microsoft.com/library/fcba1b34-c526-5d01-8644-cb8852bd2348%28Office.15%29.aspx). For examples of add-ins in C# and Visual Basic .NET that use Outlook controls, see code sample downloads on MSDN. 


## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/7086c185-99a2-94e1-6041-64c58869067f%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/2347764e-dbd8-e622-ad5a-27795613abf5%28Office.15%29.aspx)|
|[Change](http://msdn.microsoft.com/library/179e600a-8ce6-b1f4-176e-ac6aa68aaa8a%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/ec2483b8-0fe1-de86-dc01-9cafbde31e44%28Office.15%29.aspx)|
|[DoubleClick](http://msdn.microsoft.com/library/190ba56e-f4b2-ff11-0df9-1e98cdcef655%28Office.15%29.aspx)|
|[DropButtonClick](http://msdn.microsoft.com/library/425118d2-afa4-4582-1f89-857e5b7ae903%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/1e6c1905-d5f3-1063-1b7e-c62e54252e43%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/6a8ec569-4e08-0400-95ad-934cbe2c20e4%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/8b24fba9-5af4-9519-8391-1a57fab6e39e%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/59b22d35-001a-4e99-3b71-d7f95a73d821%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/7776832b-fdb0-cd2b-efa3-97dab74065e6%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/df29431e-c8a6-e345-e9c3-4a4195e00d41%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/a4788848-a2dd-d19e-e969-fb353eddbfc7%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/abe4afac-3afd-7f08-3128-650f847c692c%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[DropDown](http://msdn.microsoft.com/library/7668e185-ced8-6ca9-d89c-493f08d542c9%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AutoSize](http://msdn.microsoft.com/library/fdade84d-fa98-868c-4c76-34030242dc83%28Office.15%29.aspx)|
|[AutoWordSelect](http://msdn.microsoft.com/library/cd26e65e-d25f-26e3-5b6c-736beefb0742%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/9b4bf367-18c7-deea-dab6-09d2e53ad5e9%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/af73bf4f-4288-1679-4aff-26839e73c3c9%28Office.15%29.aspx)|
|[Date](http://msdn.microsoft.com/library/f1c1a454-4c1f-7ae6-2fbd-f3875beb6cea%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/ac687fc7-6e69-2020-25d3-facc24689633%28Office.15%29.aspx)|
|[EnterFieldBehavior](http://msdn.microsoft.com/library/985b7c35-cdd7-a75b-309e-a6459beeab31%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/c05993d6-9a33-648b-ec2e-d8c442c2ad6f%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/d949651c-96a0-a6a6-65f1-03e7c58bb7d0%28Office.15%29.aspx)|
|[HideSelection](http://msdn.microsoft.com/library/74bd86f9-ab29-dc4a-0058-5f33abb2e9da%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/9f34809b-70e8-503e-e345-5eaa59ccf087%28Office.15%29.aspx)|
|[MouseIcon](http://msdn.microsoft.com/library/4d2bf497-0e80-2494-4197-e746778da519%28Office.15%29.aspx)|
|[MousePointer](http://msdn.microsoft.com/library/14ca0547-b43c-df9b-105c-ddb655629d34%28Office.15%29.aspx)|
|[ShowNoneButton](http://msdn.microsoft.com/library/9a3cb14c-484c-a25a-e233-d99a14c31eb0%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/fda479bc-c613-171f-4e81-efe9c548fd81%28Office.15%29.aspx)|
|[TextAlign](http://msdn.microsoft.com/library/2050c4f9-b648-59a3-9171-dc31c49f3b51%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/df2c96d4-42d4-fd33-a55b-2162f65069b7%28Office.15%29.aspx)|

## See also


#### Other resources


[OlkDateControl Object Members](http://msdn.microsoft.com/library/6bc09aee-2f4e-5042-a653-52c0c09068c5%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
