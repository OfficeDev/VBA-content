---
title: OptionGroup Object (Access)
keywords: vbaac10.chm10894
f1_keywords:
- vbaac10.chm10894
ms.prod: access
api_name:
- Access.OptionGroup
ms.assetid: aa9e5607-7892-9ab2-dabc-822372b23811
ms.date: 06/08/2017
---


# OptionGroup Object (Access)

An option group on a form or report displays a limited set of alternatives. An option group makes selecting a value easy since you can just click the value you want. Only one option in an option group can be selected at a time.


## Remarks

An option group consists of a group frame and a set of check boxes, toggle buttons, or option buttons.

If an option group is bound to a field, only the group frame itself is bound to the field, not the check boxes, toggle buttons, or option buttons inside the frame. Instead of etting the  **ControlSource** property for each control in the option group, you set the **OptionValue** property of each check box, toggle button, or option button to a number that's meaningful for the field to which the group frame is bound. When you select an option in an option group, Microsoft Access sets the value of the field to which the option group is bound to the value of the selected option's **OptionValue** property.




 **Note**  The  **OptionValue** property is set to a number because the value of an option group can only be a number, not text. Microsoft Access stores this number in the underlying table. In the preceding example, if you want to display the name of the shipper instead of a number in the Orders table, you can create a separate table called Shippers that stores shipper names, and then make the ShipVia field in the Orders table a **Lookup** field that looks up data in the Shippers table.

An option group can also be set to an expression, or it can be unbound. You can use an unbound option group in a custom dialog box to accept user input and then carry out an action based on that input.


## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/ea848f63-7d6d-dd03-058f-80e6cb46b1dd%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/a497ff9b-d617-df5d-9989-bc420c827575%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/7a673665-88ed-9685-d7ca-9146e224f090%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/4ef52706-64dc-38b7-7800-07d3a4d7d7cc%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/ab5f5745-b8c2-7d5c-6fd6-43fd7901abd1%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/2c8000f7-256d-232a-c2ac-f027eac7bc6a%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/f3c569de-879d-aa27-77f2-22192731febf%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/72c6d4b1-9cfe-6e34-3c87-3577e874a322%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/203556bc-5242-1aec-ec6c-b11db04df569%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Move](http://msdn.microsoft.com/library/49f4a11d-ab81-7b81-cb28-904eba61048c%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/5cc8188a-a579-3cd6-335a-afb2d05c955c%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/8aaeccc5-29eb-559c-5501-4df7b325fc72%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/53c509fe-41d8-b430-b272-5c506c237680%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/ad22e7a9-4b9c-d46c-99e1-8f1d020c32d8%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddColon](http://msdn.microsoft.com/library/2c40b39b-2c57-e719-78ed-e28080f78fd8%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/30d35bfd-6128-0d68-12c8-56ad6f19c342%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/21d0325e-4552-699e-4972-1fc5ee157b21%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/00feb954-30a3-f7ba-591c-41679e4d8f4b%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/a329bf89-7bb8-71a5-d2f1-7ae5a0649089%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/8e0d3930-4520-f759-1a12-543bcbaac693%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/f93a9b31-e806-b45b-5f23-9ede92a23ba5%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/9dfc95ad-a996-d24d-b623-130d6647e430%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/4e33a712-af8f-bffa-f6c8-0502fb292813%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/0ea86e13-03ba-9f56-ef42-e8147fa70064%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/0272785b-9b7c-c54f-c544-7727deb9f4a9%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/42badef3-8e9b-d730-f355-d535352a32ec%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/96d90ffb-9cff-6678-9c2a-58e812c97a79%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/77c8779c-8ad7-5000-1184-87bf78e46f4b%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/ba231494-097a-6814-1eb8-fcece0fc21ff%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/520ed761-de5d-9e70-3cc8-79264f6c0f3f%28Office.15%29.aspx)|
|[ColumnHidden](http://msdn.microsoft.com/library/8b37f530-7078-28dc-659b-ff8e08b53071%28Office.15%29.aspx)|
|[ColumnOrder](http://msdn.microsoft.com/library/b1222140-b035-db57-db74-40b0db56aecd%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/5b3023dc-d876-e842-2b26-de8f9a7e7b80%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/e252c2b0-ab71-ed95-da04-62cec990f63e%28Office.15%29.aspx)|
|[ControlSource](http://msdn.microsoft.com/library/9f4a87a0-f31a-8b6f-c39a-51f49c96221e%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/21069dcf-9841-6548-6c5d-3793b73af1e3%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/ec0e53ef-2c44-b8d4-1711-1c13f92783c7%28Office.15%29.aspx)|
|[DefaultValue](http://msdn.microsoft.com/library/cb19cb7b-033c-9e4d-6683-5296c306f47f%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/a69b8fd5-d388-7277-d0de-5cf0ab620a33%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/7c8a10cc-6277-778c-e7c2-c8274019e3ad%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/a1a47d5b-5ba9-5071-bdc5-a5ea13d8d78a%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/f0b715ff-a1d4-4040-59e6-818705042691%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/cd6d2aec-fc7c-5dfc-1386-568bad2a26f8%28Office.15%29.aspx)|
|[HideDuplicates](http://msdn.microsoft.com/library/d9b17b9f-1eef-eda2-674b-cc7c7b1b5c5a%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/61b8b9cf-6f56-aff1-ee78-ddea0d4e5940%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/faae80ea-95ab-bcae-d923-39d264612e84%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/6652c226-ee95-b94a-dabc-942e0d9d5226%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/701c5bc6-e81a-83e2-acf6-9756e3c86946%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/2fe79f1a-fd28-32e6-3d22-c0187e1818a4%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/5044ac6f-630d-1a09-1e8e-5eae3c38c3c4%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/da310fc7-9fb7-fddf-9cb7-a6e2a7be0bc6%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/3d131a07-41cf-a21c-afad-623f01ed14ad%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/065ce5dd-8589-ee21-b850-d5ee95fb11ba%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/d26a3888-a7c3-39f4-ca3e-484e9c3826b7%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/dc294bee-49b7-af3e-745e-63dde913c52f%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/57ea9cba-cfbd-76f6-0cf9-193a5df87d66%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/f1dfb135-716f-4db3-1d4a-89c4b28b40f8%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/1edbc677-6cf5-a14c-1bd8-b12e6c5a22cf%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/48a64bc3-df50-6fd7-8784-1413a5bb88ac%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/313ac392-639a-b9c6-b0f3-64f7d34fe839%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/0f987181-1506-51ee-8f40-5a902c86d458%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/d132898c-7dba-4048-d32a-8f4257c5668c%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/882e6786-a8c3-d865-675d-a97e3143a8ab%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/d6e75b49-9b97-6018-1277-6cc6ef8558df%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/f1d04030-8aed-8591-f83e-6a890b96c1f2%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/f08f9c3a-f267-dab7-48db-3c972131b6e8%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/6d286cb3-193b-34d3-5335-c10564165af3%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/d30da689-1716-767f-0f0a-c1d0ffee6c48%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/1ceeb9cd-e9b6-129f-72b9-3d15d9622709%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/cce6547b-9e55-2216-9f00-ba9147849e21%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/a8d4d55c-f2ff-0636-fe97-f35407dd20b9%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/60a479aa-c9df-3eef-3752-4fc0bcbc07d6%28Office.15%29.aspx)|
|[ValidationRule](http://msdn.microsoft.com/library/ac66176e-35a6-6fe5-bcbe-2b201a6d8548%28Office.15%29.aspx)|
|[ValidationText](http://msdn.microsoft.com/library/d115a085-7c22-7a88-539e-ec4461ca6d5d%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/add35170-c02e-ac1d-211d-b2b46cd19c9c%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/9f8a49f1-0bce-6db8-460a-e1676af211f1%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/3af38a57-97bf-e427-acb5-ddc21678715a%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/03db8a33-19f3-94dc-4a46-5d643ab0da14%28Office.15%29.aspx)|

## See also


#### Other resources


[OptionGroup Object Members](http://msdn.microsoft.com/library/90e68eb2-20f2-510c-4332-241eeac27f14%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
