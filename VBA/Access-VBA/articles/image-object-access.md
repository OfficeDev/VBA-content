---
title: Image Object (Access)
keywords: vbaac10.chm10436
f1_keywords:
- vbaac10.chm10436
ms.prod: access
api_name:
- Access.Image
ms.assetid: 1bcc8552-94e2-b799-6903-392205cb4341
ms.date: 06/08/2017
---


# Image Object (Access)

This object corresponds to an image control. The image control can add a picture to a form or report. For example, you could include an image control for a logo on an Invoice report.

 **Note**: The functionality for the Image object's **image.click** and **image.doubleclick** events have been deprecated. If you want an image with click/double click events, instead use a Button control and associate an image with that control as that provides better accessibility. Button controls are part of the Tab Order loop but Image controls are not. Existing applications will not be affected by this change.

## Remarks


|||
|:-----|:-----|
|**Control**:|**Tool**:|
|![Image control](images/t-imgctl_ZA06053959.gif)|![Image tool](images/imagefrm_ZA06044465.gif)|

You can use the image control or an [Unbound object frame](http://msdn.microsoft.com/library/4a0874dc-ecac-be7c-25e2-ecc79696e2eb%28Office.15%29.aspx) for unbound pictures. The advantage of using the image control is that it's faster to display. The advantage of using the unbound object frame is that you can edit the object directly from the form or report.


## Events



|**Name**|
|:-----|
|[Click](http://msdn.microsoft.com/library/1bca7597-b536-908e-c3fd-25f9dd5e1ab8%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/605ec6dc-0159-a20e-9b02-cfd9d0a23dd1%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/03da9154-2e2b-7801-ec11-06101f7cecb0%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/651525b5-0a71-0e54-d4ed-3802e672b4c2%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/29aa863b-315a-7b4b-7c9c-89fcbb44e83a%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Move](http://msdn.microsoft.com/library/feda7964-0d93-b3e2-36b1-5c68054cdff1%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/98f16a2d-ad18-c576-11e0-43d43fcf8859%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/d7cad49f-e5ee-ed4a-567c-9706725f867e%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/81e403d6-ba9a-9117-1f87-fe6bb4b76d00%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/7c308c10-ee19-f162-a9e4-2d6d6b9eafb0%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/6003c9d8-a6bd-4718-b2ea-c6e1ccb0a76a%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/899c5320-a2ef-7861-2905-fc08f5b7a1fb%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/bd3b2a60-2b9d-7b18-63d1-5bc6f059eb5a%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/9b15a086-0ff4-3ffb-4828-c22486bfc8c5%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/67654a62-b38d-fff1-8ec3-6b4fb9605988%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/4bbc6f2a-c672-f3e3-a86d-287fa020a43d%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/24bd0510-6f97-e22d-7822-f16f97591a25%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/bec20ddf-359c-d684-6561-130c830ef62f%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/7a449370-9af6-5170-d184-13ea0d01dd79%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/f1dd7a66-941b-7ff6-eb99-208e28d27767%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/6a8d8d2a-0cfe-2557-585b-ab9e42a313bf%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/c5b6a87d-8eac-5840-2bbe-cd491b035cea%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/b6313b26-4254-fafb-923b-ef9d2b9fc0f5%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/13a3cadf-8a2e-3407-5fa8-d76e3b2c9cac%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/6f84953b-a408-a741-a2a9-18eff2406abc%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/35638607-44a6-b16a-3b58-6490965e528e%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/57817dd3-62ed-5595-8196-f914f1fda037%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/bb355521-48f5-6a6f-df05-ff5be9cc1e65%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/1ed961e6-9698-322f-361c-76e42b81433e%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/1df063c7-2354-5e57-ce0e-ea4619598726%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/830eac6e-9992-057c-5905-92a17bb1d628%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/038f4c8e-a7ba-bfa4-df87-a68baaad1c0b%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/6190a13c-af22-6793-b64e-76c0fc2fed34%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/4768daef-932f-969f-fe6f-434fc14b150f%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/40b394db-e64d-f63b-a1a2-e234dc76581b%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/f2d3a6d8-c99d-37f2-1a19-eb0a003df8a6%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/22d22121-1be3-d1c9-9288-bd3294c7c583%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/046f4bb2-2cb3-b383-8ff9-2fd304e84fd4%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/33a170d3-0f09-3fc2-8a2f-cd12e93a879a%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/cfc48a43-736f-58c0-95a5-fe248de8b9d3%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/76799f89-978a-4baa-a330-525247c1131d%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/33fa46ae-531c-eeb1-f7ab-51c90ef5c6c5%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/13a80139-3b1e-f94c-d5fc-1d5c0f305a0d%28Office.15%29.aspx)|
|[HyperlinkAddress](http://msdn.microsoft.com/library/e92e7d7e-8447-9c9d-4d17-55c479d13228%28Office.15%29.aspx)|
|[HyperlinkSubAddress](http://msdn.microsoft.com/library/ba6f27ec-d28b-e495-4e63-9355cd26630b%28Office.15%29.aspx)|
|[ImageHeight](http://msdn.microsoft.com/library/91d0cc66-8b27-40f0-8112-41410429400c%28Office.15%29.aspx)|
|[ImageWidth](http://msdn.microsoft.com/library/516ebdd4-201d-db7e-de34-7f9ad0bb4955%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/f128660e-6a28-4af4-ff00-4463ff618e7f%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/c0359df3-a60b-e0d4-e494-3ea2b237aa25%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/f0a3c620-9c27-e322-276d-23a8054126e4%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/ae6e765b-a349-f16e-ce78-671ac7f6ca1b%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/76fac9f3-aead-6824-cba3-9246c397148c%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/2a409876-3c11-515a-37f5-ac676d693550%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/6dcac8a1-037d-dad2-00bf-ed73a1b21e4a%28Office.15%29.aspx)|
|[ObjectPalette](http://msdn.microsoft.com/library/394786b9-7ee1-bc79-e84e-12bb75189f12%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/ab5295d3-9e24-4604-a541-ac5bba837c0b%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/1e2b9701-1b75-5cb9-32c8-d6585575b7e8%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/ddd7ceb8-59ad-ffc4-771d-17ed0fb42ca2%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/65700447-c7bf-7b7f-ea2e-75e4c8fff70a%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/d9c4128e-0a52-698c-a605-cbf31e183e2c%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/9890dd97-0025-7329-1751-82d69799510d%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/d0c73a5e-a478-ca9c-e5c0-c9fc9bcc6269%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/f6fb685a-1934-a2f5-ce58-f2a9e46dc90b%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/7844c00d-d56d-0473-31d6-7278f9e4d10f%28Office.15%29.aspx)|
|[PictureAlignment](http://msdn.microsoft.com/library/e0ebec64-9a26-859e-b9fd-5f4a47253bba%28Office.15%29.aspx)|
|[PictureData](http://msdn.microsoft.com/library/64d4a266-2c7f-5b08-4f32-bca25dac87d8%28Office.15%29.aspx)|
|[PictureTiling](http://msdn.microsoft.com/library/9be8cde0-4632-197e-ea3a-8db5846b8920%28Office.15%29.aspx)|
|[PictureType](http://msdn.microsoft.com/library/873fdf85-bbd5-98d3-c8f0-4b1994ed0a85%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/14c0f780-0df2-8f89-320e-0e238e324f46%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/9fe9eb52-d504-6406-894f-0a90530687b9%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/994f5290-e92c-da14-2b85-194681b56d40%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/1d85ddc5-3aa7-2267-778d-e96f1e1148b0%28Office.15%29.aspx)|
|[SizeMode](http://msdn.microsoft.com/library/feaa8002-7d5c-6ce8-dd07-49f6a7330b17%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/30b9d6c8-4071-4eb0-27b8-cf4ddd7c44f7%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/627e6f93-8812-e66e-0291-d24be9185fc2%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/4ab0b0d1-802e-08f9-27a2-08500e0f8b62%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/fb248161-837d-e455-8d9e-4fb5d1a39d3b%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/dbbd345c-b384-0a4f-fd80-22920e71c4a8%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/022201dd-2847-dba5-2a0e-86e94feab535%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/9a6641b4-8e9b-2d9b-8122-6f4d6967606c%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)<br>
[Image Object Members](http://msdn.microsoft.com/library/c2ad356b-bd6b-2b45-00b0-cd484ee06cc5%28Office.15%29.aspx)
