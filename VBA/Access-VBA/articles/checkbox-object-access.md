---
title: CheckBox Object (Access)
keywords: vbaac10.chm10798
f1_keywords:
- vbaac10.chm10798
ms.prod: access
api_name:
- Access.CheckBox
ms.assetid: 63e75704-af4d-7b38-7b8b-04f7f17fa1ec
ms.date: 06/08/2017
---


# CheckBox Object (Access)

This object corresponds to a check box on a form or report. This check box is a stand-alone control that displays a Yes/No value from an underlying record source.


## Remarks


|||
|:-----|:-----|
|**Control**:|**Tool**:|
|![Check box](images/t-chkbox_ZA06053977.gif)|![Check box](images/chkbox_ZA06047229.gif)|

When you select or clear a check box that's bound to a Yes/No field, Microsoft Access displays the value in the underlying table according to the field's  **Format** property (Yes/No, **True** / **False**, or On/Off).

You can also use check boxes in an option group to display values to choose from.


## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/dfcb46c7-fe13-02a5-4d1e-e3e897b738ae%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/cc0951d0-8772-8d76-5eb6-0507026587eb%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/15c55276-ef6e-bcb4-09fd-2a457df79387%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/dea6c8ff-47d5-de41-8099-a36b4c53c665%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/512122ce-f438-46d6-4990-6fff469bc68e%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/5a805d97-8d63-1635-f41a-e18aa9437d59%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/3437bdf0-cc5e-d09d-3607-9fd283613243%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/185941fa-3ae0-47ba-b3c5-b4acd82417f8%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/946df95c-da92-1977-6bb5-ecabbb5f8ee2%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/4e86b4c2-e287-db2c-4e74-f73efd7a064c%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/b93f5eb0-4afc-28af-cd03-cbbd23500f39%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/6281cd33-662e-e73f-5365-5784aca5c5df%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/f45a89b3-eab8-0757-1ac8-b2aebaa47a1f%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/704acc3b-6ff6-fb0e-9adf-bd34185443e4%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Move](http://msdn.microsoft.com/library/147a42c1-4e1d-f814-e8a6-5a0d328cf79c%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/0385fddc-7a97-1bf3-50d2-61f0978ea359%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/68d0ec9e-7a2e-1402-6a2a-38caad5d13bb%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/79309619-c2f7-d43a-5f92-ef2c4d1af208%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/e69e5d59-398d-744c-0a99-e2ca9b290c9b%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddColon](http://msdn.microsoft.com/library/16a1bdf0-3290-a465-e275-25a1097d01fc%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/eaef525d-4447-86b5-9567-311e7324b720%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/b47ba26d-bd87-ec43-381a-c93b654c08e2%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/b79869ae-f295-a747-37d7-f48a98c300c8%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/4eb1070e-9485-7ebc-70c2-48bba4b8cd88%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/21f40b75-756c-df56-11bc-824ef796797e%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/4be270c7-a093-0315-71ba-7fd8a6ec26a3%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/8d2c64cd-b5ca-0c02-1332-3887929ecee5%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/5b7fd629-a896-ab01-b965-2a2f0d7724a7%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/57e00b53-89eb-3cee-a075-9eb3c9ab60ee%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/c10ee0fd-10b8-e35e-d042-b319276b50f8%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/15c760b6-a64f-4ea7-4923-fa0783681a1d%28Office.15%29.aspx)|
|[ColumnHidden](http://msdn.microsoft.com/library/9a6c271c-566c-e18c-ae85-8d35b4487cb2%28Office.15%29.aspx)|
|[ColumnOrder](http://msdn.microsoft.com/library/b485afc6-b19a-cd0a-990e-9de14f6f9104%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/8a545cee-33fd-8105-d3c2-665ec269c18e%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/4003f288-678f-57a7-0be7-a57517f14188%28Office.15%29.aspx)|
|[ControlSource](http://msdn.microsoft.com/library/167d8da3-0489-ca23-2821-e455b8ac2d53%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/b37aaf1a-a436-636c-f529-85720810aec0%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/d9a714d8-7381-44f1-882a-57233819e024%28Office.15%29.aspx)|
|[DefaultValue](http://msdn.microsoft.com/library/3bbeaae3-3f94-0841-306d-a73e56cac461%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/9236d99e-df4d-5342-e60c-162abe7de8d6%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/48bf27fa-f08e-6fc9-ad92-6ec489b80801%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/9ab63762-34fb-06f4-3b79-97471152c939%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/6d3343aa-3505-dbb9-7e61-6b5c8d67b9f5%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/6f59985a-9b2d-e563-f0ed-dfe938e27331%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/2326ec85-b37b-cc97-d8f3-4913c936436b%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/c4a0cf6d-488c-5978-d3db-184909c79723%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/9c78a907-1801-ca30-f24e-6cfa25560a94%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/a9bd50a3-0fc1-b39d-ab04-38b06bc2bb65%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/85f36c8d-e62e-8d41-331f-ec8abd509992%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/710894e8-4271-069f-7e3e-46d39da22daa%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/04495e96-0ee3-399e-4718-d372cdb3bc4d%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/8eff7dcf-e5fc-74d2-2685-fac6f945c661%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/7a51f6bf-bf21-2233-b74e-4d2925df0b1d%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/337537e4-7754-40a9-b5e7-c672076578f9%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/c45447cc-6659-c370-398d-fd7d4888f7a2%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/205d8d08-4060-7ac3-8bb2-99d381bbef50%28Office.15%29.aspx)|
|[HideDuplicates](http://msdn.microsoft.com/library/60f024b3-113f-4509-6556-cc51ad656c85%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/6169f797-eb38-933e-96ca-d1b3259eb2e7%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/0dd5f74a-fd36-8bc2-90f8-039d1f83004b%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/2fa958e4-1580-c69e-739a-3b9e49a5713f%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/255be436-51d3-0926-a7ce-a5b595ff59ce%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/5067374b-9e37-3e13-003c-c3688812221f%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/7f016e78-850e-f55e-bc56-b574b453cede%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/aed408d0-7e94-0b2f-7746-1a456d140a91%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/a54fcf07-a233-aa9d-4014-4fd75abe5591%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/0d4eb8e2-b45a-a293-5d71-3b13743283bc%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/8a12399e-d8bc-54a2-c4ba-88e3b0dc7d58%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/0f0b6f34-d389-8376-81fd-cff5a93ca4c1%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/545d17e4-f695-33ab-8e72-4c8e048b86d4%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/255ec4d3-dff4-d63e-38a1-ad9a36e08104%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/b0e0261d-82d2-47e1-3e0b-b9582798cd9a%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/3721a21b-77dd-5f43-baea-e7e98647c17a%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/3c4f98d5-3190-e88b-50ce-df08a3c4aac0%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/54894c2c-e0ab-8679-a55a-df44af856f8a%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/fd52a8c3-7d49-9504-9afd-f6132f138690%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/eaa59b30-d037-2b3a-1e24-e5ea9a11f0f3%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/ead6dc7b-2be4-a8c4-6f4a-7b3fcfcacc48%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/a1f83ff8-b334-0314-8041-38a357b8c5a8%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/a3e08de2-f135-b7e2-6d7e-c3030674f7be%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/b2f7b85b-73c3-b47c-5a31-b9b733208901%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/236c9263-4238-ec07-d239-2481575ab8c6%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/a3d86d09-c821-72a4-f48e-2cd022c2659d%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/898c8b44-d2f6-7d4f-f3b8-5d71d893eca1%28Office.15%29.aspx)|
|[OptionValue](http://msdn.microsoft.com/library/fffd881e-190a-aa42-b54f-f8fe629f7d02%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/c95e4144-3bf5-f38c-dbb5-02e752459c0a%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/ae014699-7594-181d-7f98-e72f7cf3c071%28Office.15%29.aspx)|
|[ReadingOrder](http://msdn.microsoft.com/library/b0c40eaf-447a-0051-6ffe-2c7895cdbb58%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/669e17f4-586f-1ea3-a239-c72902970f89%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/0b78f0d4-c34f-ef4c-8cfc-800e68e9be44%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/1f9bf8b4-d0c7-ddd3-9c4f-cb9bd863463e%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/5e09067e-1648-8f95-f10a-5e125c28def5%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/5b199d3e-b79d-f611-9e66-1816f5c60f25%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/094064c7-83f3-8d3d-25f2-b5b2956331ef%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/2949f9f9-a18d-900b-cc43-05732b91eb19%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/3aa44f1b-9373-86df-fd78-ac9f5e3f8108%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/61c03e90-c5cc-c316-64dc-26293db3cf13%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/e3fc7819-8cb1-44b9-dc13-6e5c75bff62b%28Office.15%29.aspx)|
|[TripleState](http://msdn.microsoft.com/library/f2c9f398-6e1b-00cb-4033-b0fb5a83e737%28Office.15%29.aspx)|
|[ValidationRule](http://msdn.microsoft.com/library/4ebb1371-acd0-2227-49e9-ec646a0daaad%28Office.15%29.aspx)|
|[ValidationText](http://msdn.microsoft.com/library/25f8d9be-1015-4ff7-c088-569b8995e80b%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/a19b0395-eebb-42d6-58b8-affbe56a72b5%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/71b3b605-ff9f-b383-d367-0701c078a910%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/90d15ba3-525b-81cb-5768-2b4f9c3b9a70%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/a5756720-ee33-6a47-e4eb-ec54b11cd45a%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)<br/>
[CheckBox Object Members](http://msdn.microsoft.com/library/aeefeae7-4053-ec23-80ef-1da1099f54f0%28Office.15%29.aspx)
