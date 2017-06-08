---
title: OptionButton Object (Access)
keywords: vbaac10.chm10671
f1_keywords:
- vbaac10.chm10671
ms.prod: access
api_name:
- Access.OptionButton
ms.assetid: 661ada74-d044-4a5c-2bdd-2dddfc2e79ab
ms.date: 06/08/2017
---


# OptionButton Object (Access)

An option button on a form or report is a stand-alone control used to display a Yes/No value from an underlying record source.


## Remarks

When you select or clear an option button that's bound to a Yes/No field, Microsoft Access displays the value in the underlying table according to the field's  **Format** property (Yes/No, **True** / **False**, or On/Off).

You can also use option buttons in an option group to display values to choose from.

It's also possible to use an unbound option button in a custom dialog box to accept user input.


## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/dbff2785-184c-601c-f26e-1ca99ea496a8%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/9c887502-2d9c-6f21-e5ef-adc164cde095%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/99391fc2-c114-ca68-a176-a7f2757a9aaa%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/405b3c90-b00e-d7e7-6e22-161060172615%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/62d01554-4a32-cf66-84a6-945becbee9ed%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/2be3f0b3-73a1-e1e9-28ca-ee0cbe92e040%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/f0a02ae3-b90e-2193-3c59-c0f018ace680%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/9a21c03b-9806-d0ee-8c44-9edbba49b4b8%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/e2b8a352-2fd2-8bdb-0842-6f8e73868c0c%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/5685c274-19a0-2d29-f968-50412ebd1d9b%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/e840c351-9aac-7a79-31ba-bf9929d0208a%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/6115cf77-8929-bd7c-2785-880e28809553%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/4353c0b8-469a-7046-3ff7-6f2a9089dde8%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/55ee8314-8ae6-f0d7-5fcc-ae1000bef664%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Move](http://msdn.microsoft.com/library/cbc851ee-7dec-bed5-9ddf-31006a0ea6eb%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/acbd946a-bb2c-e205-9f81-54e034a26e0a%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/ca925414-9b8a-c34a-2806-e7894231803a%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/6b436216-a814-62ab-d87a-6608959365e7%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddColon](http://msdn.microsoft.com/library/c5908dac-412f-c779-56d3-3b75c790c17f%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/02ca295b-ff5c-2f6d-12f0-ea0bc176947a%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/8c2e2c14-b66b-435c-4631-d49b8a376671%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/2ff7b57a-2a8a-84ae-def5-d8a95bff05f7%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/8940a73b-9b9c-7911-60b5-10db8445ecb9%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/fd44b63b-d1bb-3663-8f14-08069424d022%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/b0bf4c1f-f3e9-ee11-4a53-d834c40a7c63%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/44aa551d-6b08-2e55-21e8-0c7af12e1cc2%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/19717679-b20e-f5ac-fac9-5349b4227e62%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/901c10bf-1d49-2fbd-4403-8d93547d534f%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/cd4a5e9d-6444-7cac-aa04-c62b42887a16%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/4813c3b0-03c2-9f43-bb1c-e28d7eff542b%28Office.15%29.aspx)|
|[ColumnHidden](http://msdn.microsoft.com/library/962a7bf7-8898-d2e5-f26a-691b8c9b5d71%28Office.15%29.aspx)|
|[ColumnOrder](http://msdn.microsoft.com/library/5d4d8302-45b4-92e8-4d8f-dc00557ded42%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/fb173bbb-8bcc-ee35-3248-2cbaa35ce5ca%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/52e9979d-2c00-dcef-0e61-5f762fbb18f8%28Office.15%29.aspx)|
|[ControlSource](http://msdn.microsoft.com/library/a2d61057-fe0b-4c00-88f9-f375074d7b3c%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/4a2ff101-e8dc-cc96-abb7-7b66c2c8e74d%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/0f411793-1381-4cef-4d80-bcdc1046ba62%28Office.15%29.aspx)|
|[DefaultValue](http://msdn.microsoft.com/library/87be103a-bfe6-ccab-7349-4c3cbbeadc30%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/dc4956e8-a34b-f4b6-d7fb-a095c74d63ef%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/ff1a1ee6-c92f-4106-b49f-25d6a17088d7%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/95896310-8723-de8f-dec9-51fded5227bb%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/23cdfbdf-6e89-8d2a-bb4a-29ee0a13af37%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/599f0476-e468-8cb7-1cf5-0f63a2dabc8f%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/831a0590-1d50-7260-5a00-c0ecf973c5db%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/075ad462-4004-0e2c-e1af-dd79de4a9a1d%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/9ae532dd-48f9-720b-91fe-ba5d67d39176%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/dcfff0b4-431f-475f-ca1f-569d0df22243%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/58a70e63-9c82-4761-8597-c134882e04e3%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/83b2b75b-7c9d-0e6c-1015-eeead18cb20b%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/b0a50d2f-25b2-22a9-4cf2-219348ec8daa%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/2edf6a74-dbe8-bf47-afa4-21496e64839e%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/24b556be-abb4-a87f-d021-c23e7d872ff8%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/1b650e6f-e6ef-4b47-5b63-c4b26fd9feba%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/d3a95041-1e8f-5a02-019e-ecdb2f795bf0%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/0966e507-59d9-5e14-f6af-6c388b9037f5%28Office.15%29.aspx)|
|[HideDuplicates](http://msdn.microsoft.com/library/c42a89b0-2fff-e56e-0621-c2d9b6e7fc4d%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/1815fce8-2afe-8e21-8702-9bb6f779f112%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/ed6d0f6f-a8d5-0a31-342b-9def542a7e78%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/e9fdcd98-275a-7e54-bee5-74d97a6de086%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/3ecb4d1f-7e32-9699-b2c3-6918d7b2eb61%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/432534bb-9c5b-6a07-0509-97c967c04cf0%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/e5fcac2e-efa7-362f-176f-90ddc53db695%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/39dc9948-a231-4a6e-3d39-6c5e23e001d2%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/dcb40002-67e4-f11c-1e75-260f96bef440%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/bb7f0e55-e08a-a231-ad6c-55ebdd65cf3b%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/5e60f737-5cc7-97e9-af4a-b8f065a5277b%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/dac84eb2-1b12-8d4b-37a0-1cdf320f6faf%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/8ab3e829-5414-de39-adcd-b67cb27fc197%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/77dbbfbd-9ddc-951c-1376-231ff8a0a768%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/ea84a877-ef29-444a-ce08-e816ee7e3dae%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/74fe1cf7-0f17-a495-6e2d-527691eae129%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/45a7b4fa-ce24-aab3-6057-ce23b1055a74%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/e454abc8-f344-f67a-f67a-ae1a8003155e%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/a857d054-b372-e10b-0246-f0e95b742902%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/04c44e84-0a60-cef5-16eb-0a9ec90015ec%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/631cb13b-cbee-e5eb-2be8-260aa08c441b%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/6adf4d90-7922-bdb4-c09e-397f1c8c8a42%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/99b7e4be-f2fc-f221-814e-b31cd3360063%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/e2579b6b-a499-ff37-8195-29cc1aad79db%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/064de273-6dd9-091c-07cf-1241f45071b6%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/4a93846f-5774-1cf1-4dfe-a93361408497%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/8c1cfbbf-99da-0844-d435-5f911ad24b17%28Office.15%29.aspx)|
|[OptionValue](http://msdn.microsoft.com/library/23e170c7-21ac-4725-b54b-ad778bba9f31%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/598f3f9c-0f25-635f-d438-3b0cd8d2f343%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/91313756-130d-5b6f-406c-4d5768f522c9%28Office.15%29.aspx)|
|[ReadingOrder](http://msdn.microsoft.com/library/52dab78d-5c67-4031-06b4-f7fa43207f4c%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/e739fdd2-18be-eb96-f8ed-a9b4b82b4885%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/0caab057-7495-e0af-6b3c-3e8c63c06f95%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/735575cf-fccd-5de8-875b-8718b60892dc%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/33dd01c0-0ee0-640d-d8f3-f7c3590aeb90%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/d3cda3a2-1b19-6b12-6d22-0cfd1b869933%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/a962d94f-9e3d-b52e-1e0b-50aa27b98e58%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/b7bd7921-2ba3-1445-1e89-ce8fa0c2ed4e%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/147ddd8e-6fe2-d59d-2f83-71c7cdfcd263%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/21612875-5a0d-3a13-28a1-1e087c5991cb%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/2689deb0-0477-6c83-550b-a08529f1f08b%28Office.15%29.aspx)|
|[TripleState](http://msdn.microsoft.com/library/f2764290-00be-38f7-f078-fc0059340455%28Office.15%29.aspx)|
|[ValidationRule](http://msdn.microsoft.com/library/1113ce22-08cf-f29d-8290-e2c86b0c4be5%28Office.15%29.aspx)|
|[ValidationText](http://msdn.microsoft.com/library/4a0a025f-7c86-cd2c-efa3-2786fc31a675%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/79f4e783-8f3d-669b-8c6e-73611cd6c6e7%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/1f821dec-12b7-bff9-4ec3-d55bf4782cf2%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/f5481b70-82a3-d2ee-d886-e952a091a9fe%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/b5cd80f3-62cb-c0f5-1ca1-adc92e97307e%28Office.15%29.aspx)|

## See also


#### Other resources


[OptionButton Object Members](http://msdn.microsoft.com/library/5173d5c5-b898-97ee-a005-7f5a4d77efa1%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
