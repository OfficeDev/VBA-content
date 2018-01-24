---
title: ToggleButton Object (Access)
keywords: vbaac10.chm11810
f1_keywords:
- vbaac10.chm11810
ms.prod: access
api_name:
- Access.ToggleButton
ms.assetid: 1c20d809-d7db-e096-4328-ebb4d79e770e
ms.date: 06/08/2017
---


# ToggleButton Object (Access)

This object corresponds to a toggle button. A toggle button on a form is a stand-alone control used to display a Yes/No value from an underlying record source.


## Remarks


|||
|:-----|:-----|
|**Control**:|**Tool**:|
|![Toggle button with the label 'Discontinued'.](images/t-togbtn_ZA06054009.gif)|![Toggle button](images/togglbtn_ZA06044638.gif)|

 **Note**  When you click a toggle button that's bound to a Yes/No field, Microsoft Access displays the value in the underlying table according to the field's  **Format** property (Yes/No, **True** / **False**, or On/Off).

Toggle buttons are most useful when used in an option group with other buttons.

You can also use a toggle button in a custom dialog box to accept user input.


## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/8e0e74e5-018f-5e0b-2c5d-d7e3db0e47f4%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/4c910eb2-6ae9-ffef-2fd9-a95222975e49%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/ba9f17a4-70ec-f4b8-fb21-01350ebf572d%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/d73ef157-6399-8a0c-6ec3-c329567f3d5a%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/b4d4f4ca-2b1f-8a9d-a6b6-eec730275af9%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/bdff5a6a-fd28-f33e-7926-360d438b1e71%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/06f9bf2b-0a69-2d90-f238-2594a7baca8b%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/70eb32a9-aea6-5d14-7dc1-1f4d4f0a8573%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/2f4d96de-5d2e-5a52-9df2-94262ad7def2%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/94359da1-d763-43f4-8d47-c6b6f3816a04%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/3dd094a9-403b-3591-9853-349b3ed761dc%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/4bcb1d42-9ef4-ff05-cf31-36459b75a668%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/ae3b48a2-962a-2990-5922-41abc9ab7f59%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/a7db8f67-202d-21a4-f74a-3826e80bb22c%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Move](http://msdn.microsoft.com/library/98cdc2e0-63b7-ed59-0fca-3d4db5f1cf4b%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/117bca69-466d-028c-b943-3a5f8517b53a%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/232880c5-cc69-b614-f918-9d0353fdb58a%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/3533c064-f559-4eb4-4cca-add03df5e693%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/95db5f79-af3d-9577-8d7e-6d2784a016f4%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddColon](http://msdn.microsoft.com/library/962c79fd-4575-1eea-982a-27a8d55416aa%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/58e20c71-189c-d2df-54a0-42b8fad6ec07%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/904f5218-95ec-dffd-4796-845ef11e6201%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/68c5518c-a7c3-bd24-9a6b-ddedf4038e7f%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/b86516be-1bf2-8a0d-ef4d-1795880ff8c4%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/d536f879-2819-9dff-56ba-aa92f3964b50%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/9d8c3d6d-e992-b1a6-b005-487270e1fe43%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/21f063d1-28c4-d357-7d92-12c38a719295%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/a2754963-4168-aa9f-6b0c-8de4332c09e6%28Office.15%29.aspx)|
|[Bevel](http://msdn.microsoft.com/library/91cfaa50-944b-23c0-2e3b-d8b8a1cb1e34%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/0ef018d1-397f-f7e8-317e-639e85de0e98%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/d490ce4a-9c25-e6cc-adc4-4a8883167175%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/339bfae9-4320-565c-c299-eb92bc28e4f0%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/bd363da1-2123-25ba-8834-b6ebbdfaa5d4%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/5d60c105-a765-5865-66b5-b236de827960%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/cbfd0285-9332-743c-a446-dfbff4dc7443%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/91248f14-4926-cee7-39e6-f1beff11bcf8%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/7ae95889-3b92-14c1-792e-eac87a2fb910%28Office.15%29.aspx)|
|[ColumnHidden](http://msdn.microsoft.com/library/9ef40b79-555d-c7c6-cf16-307d073afacb%28Office.15%29.aspx)|
|[ColumnOrder](http://msdn.microsoft.com/library/f9344297-d639-208c-db4e-4ceac2fd56ad%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/ac3f6bd8-22ae-5a3d-2646-2350a7e3be85%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/99ef9045-10c0-d059-ea6b-be70b9c12a7a%28Office.15%29.aspx)|
|[ControlSource](http://msdn.microsoft.com/library/077297e8-6911-8cef-0aa5-4c5cbebcf4a3%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/b15ebb7a-70cd-1a0c-cdfc-17cbd965e8f6%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/9371ee97-b1d3-5564-1d9d-9e6181a433b9%28Office.15%29.aspx)|
|[DefaultValue](http://msdn.microsoft.com/library/95809409-a347-33d6-4268-2b66fb1f2ac6%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/881f7a17-be3d-436f-1511-d6af5a7f4c6e%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/836c6553-07ae-0014-6a0a-ab1fa33cf550%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/80a9cfe1-87c1-b95d-f9a7-6afeca7c4755%28Office.15%29.aspx)|
|[FontBold](http://msdn.microsoft.com/library/d9c5bca6-1a89-2eb5-07dc-f855f1ea1580%28Office.15%29.aspx)|
|[FontItalic](http://msdn.microsoft.com/library/c0c2f257-832b-ebe2-a341-040adbbf1d3c%28Office.15%29.aspx)|
|[FontName](http://msdn.microsoft.com/library/7b1d51d8-5307-1446-344a-f406f2758a36%28Office.15%29.aspx)|
|[FontSize](http://msdn.microsoft.com/library/0175a789-55cb-afeb-33ad-81705983a28d%28Office.15%29.aspx)|
|[FontUnderline](http://msdn.microsoft.com/library/fef06d9f-f21f-a753-9822-f1e823ab10b4%28Office.15%29.aspx)|
|[FontWeight](http://msdn.microsoft.com/library/8b74b5cb-c5d0-82d4-a902-42dcd49ee106%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/89eac6c0-5989-40ba-276e-53f1de2d2ed8%28Office.15%29.aspx)|
|[ForeShade](http://msdn.microsoft.com/library/266e2047-8d29-69e7-bda9-c3d152cf78ba%28Office.15%29.aspx)|
|[ForeThemeColorIndex](http://msdn.microsoft.com/library/8358b6c4-960d-e414-a6c4-657700caeeb0%28Office.15%29.aspx)|
|[ForeTint](http://msdn.microsoft.com/library/b0ea7b04-962f-bdea-d3c2-8fe9f0bf83e9%28Office.15%29.aspx)|
|[Glow](http://msdn.microsoft.com/library/f279f51c-11f7-de6c-0f47-369e9b5cb3a6%28Office.15%29.aspx)|
|[Gradient](http://msdn.microsoft.com/library/ac12829e-ec4c-7f6e-93fa-918dc84bf7ce%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/8c35e5ad-5a5e-479f-4161-82637aae376c%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/33975f40-63ca-aa3f-eb8c-7af752b8c1b3%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/02a2cc7e-f8e1-d107-6f13-075ce7448082%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/1f5fb2ce-e8e2-f14c-d30d-0d28651aed06%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/c4855cec-2481-1640-9b4e-990d5d4a25a1%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/119f981d-a6a5-c4d7-613b-cef36699a172%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/437bf229-8486-3be0-e115-b81af5a88a1c%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/9ad9a972-2b67-94ae-77a2-5b1410b94639%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/a262556c-ac3d-46ef-24a1-6215e56911b1%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/23c09d6b-56a6-2ede-a83e-e542b856d4fd%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/eea6f611-1e03-fabf-53d4-c67b43f5a079%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/f707fdde-cba6-2d09-b251-358de25db75e%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/8544f955-3891-3799-5207-de7fa2a5a224%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/9f788b15-67d8-84ca-8c6f-6ef1e67f8895%28Office.15%29.aspx)|
|[HideDuplicates](http://msdn.microsoft.com/library/3bcd4798-81fa-0cfb-4dd4-1ed9150dbb3a%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/89bab994-84a3-b363-7169-a12418ef1703%28Office.15%29.aspx)|
|[HoverColor](http://msdn.microsoft.com/library/eade7060-78da-8bea-53b2-f8eb5e40be4c%28Office.15%29.aspx)|
|[HoverForeColor](http://msdn.microsoft.com/library/0280957d-7fca-0202-b9f4-15389ff3d1d9%28Office.15%29.aspx)|
|[HoverForeShade](http://msdn.microsoft.com/library/67e4c9bf-0bcc-f79f-491c-93cb32133012%28Office.15%29.aspx)|
|[HoverForeThemeColorIndex](http://msdn.microsoft.com/library/7159df87-2817-7cab-7e3c-23f0c4613796%28Office.15%29.aspx)|
|[HoverForeTint](http://msdn.microsoft.com/library/81b67e89-3ae9-941f-4830-fcdbf02afd9e%28Office.15%29.aspx)|
|[HoverShade](http://msdn.microsoft.com/library/a9e98d48-95a1-64d0-77ba-f2cd8dadc4f8%28Office.15%29.aspx)|
|[HoverThemeColorIndex](http://msdn.microsoft.com/library/40c60375-cd0b-73eb-1999-737b6d8cfc01%28Office.15%29.aspx)|
|[HoverTint](http://msdn.microsoft.com/library/fbdb27bb-8a21-729c-17d6-a0e9b43826ae%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/c168b14d-c10d-1a0a-96cb-69555c8657d0%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/1abe4640-f2ee-4aea-e86c-cb5e8946d156%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/fa8b44e8-9e42-8088-e369-a176bb320a05%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/4693ae47-a90d-6467-4e84-5ec0a78ff2e0%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/6e85e8f2-ebcb-7bf4-9c78-f83a684deebd%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/28602d7f-17c1-a54d-82d3-dfa15a88de4a%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/af440e04-2046-507d-1d66-e8287ae5bbf8%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/da08e677-2d2c-6f06-fde9-899b82349ec2%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/97747f24-6abf-f005-f4d7-b10af6f7629d%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/1fb9951a-e531-0423-38bf-f7e4c922acc6%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/70e428c6-96ab-1747-7cfa-484e9abad0e7%28Office.15%29.aspx)|
|[ObjectPalette](http://msdn.microsoft.com/library/2634e6e7-5e50-c119-0d4c-93e7f1f3dc35%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/d23f0c45-004e-74c8-6309-a76854d79a1c%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/54a5ade7-7da4-9357-588a-7b97f0a44661%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/7d7a6627-db0f-f276-36fd-776d5e4b806c%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/06605089-613c-114b-4775-587a0357e875%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/59dd0f8d-7c77-08be-8978-ea039ad851b9%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/bcc774c8-7766-942d-b37d-d4c96dd84911%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/b6a167f8-a6a3-a0b1-e04f-7bf1b595c318%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/8fe11ce6-1566-238e-c93a-1ee5835b9c2e%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/a932ab8a-3b48-8aa3-5ee4-97593b4394a4%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/f7f9f17d-0fb3-49b1-a6d8-d9498b188651%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/a9bbf8a5-4e62-fa9e-63a4-2f59cd2734f4%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/3bfbe7b8-3f8d-5f77-2afe-e8a4f3e11c8a%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/734cb3dc-0489-d336-007c-e7a93658680f%28Office.15%29.aspx)|
|[OptionValue](http://msdn.microsoft.com/library/b86ba53d-d590-efe0-9dba-89ff919871fb%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/754bba9f-e1e0-2b46-601b-986637531fe0%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/78889614-9916-1265-611a-8ae6932187fd%28Office.15%29.aspx)|
|[PictureData](http://msdn.microsoft.com/library/1d2f0d70-0176-a43c-37a5-e527b9e48e40%28Office.15%29.aspx)|
|[PictureType](http://msdn.microsoft.com/library/b9fafc70-9398-9b22-8d3f-ae0d05671aae%28Office.15%29.aspx)|
|[PressedColor](http://msdn.microsoft.com/library/b0296b52-1207-0dfa-c4b8-fd8ef5c88338%28Office.15%29.aspx)|
|[PressedForeColor](http://msdn.microsoft.com/library/0e05a577-18ec-0d8f-1407-5449153a6156%28Office.15%29.aspx)|
|[PressedForeShade](http://msdn.microsoft.com/library/9a6ddbd0-154d-6018-e8fd-8fa9bd916356%28Office.15%29.aspx)|
|[PressedForeThemeColorIndex](http://msdn.microsoft.com/library/9c2b6020-3bb5-72f5-184d-2b1453946a26%28Office.15%29.aspx)|
|[PressedForeTint](http://msdn.microsoft.com/library/c93d5f87-9b9a-fa6e-7226-709484c1e257%28Office.15%29.aspx)|
|[PressedShade](http://msdn.microsoft.com/library/72176e9c-68bf-971c-3147-fea692240d17%28Office.15%29.aspx)|
|[PressedThemeColorIndex](http://msdn.microsoft.com/library/85609290-6641-001c-7bc2-0f14443b326f%28Office.15%29.aspx)|
|[PressedTint](http://msdn.microsoft.com/library/01fa017e-05b3-7bd7-b2bf-19bf4a641802%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/41006d09-fa35-00ee-4ce1-a88ccdfca458%28Office.15%29.aspx)|
|[QuickStyle](http://msdn.microsoft.com/library/6dc5a569-8758-86cd-5b2a-693081ef95c5%28Office.15%29.aspx)|
|[QuickStyleMask](http://msdn.microsoft.com/library/7f3e65d9-44e8-289a-2123-093aed70650c%28Office.15%29.aspx)|
|[ReadingOrder](http://msdn.microsoft.com/library/172e40bd-bdd2-a4e8-3e96-d4bd8d3c40c8%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/36e38e77-104a-0cac-9c89-1bd0958ad55a%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/65d3f3af-3c21-edb6-bff2-79737231424d%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/0095ff4e-56f0-9b56-73e2-2e3066ee8b03%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/86f39f5a-ab5b-2db2-611b-53568a99ac0c%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/ba7ac65f-644c-b75c-12cc-565cd27a7162%28Office.15%29.aspx)|
|[SoftEdges](http://msdn.microsoft.com/library/23c63821-966c-4d9f-7304-5b6e31b85675%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/098391de-a83b-b8cb-e045-b6d9edac3ff5%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/1712e879-20da-8797-e94d-ee68b0d23c59%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/28712aec-2836-9ed0-c8a0-fd5aa50828d0%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/d487494e-e987-1a2f-86c3-09bdfa1ede08%28Office.15%29.aspx)|
|[ThemeFontIndex](http://msdn.microsoft.com/library/c85eef50-220f-372d-9a86-2107a8447053%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/31f8d2d5-6372-9241-9f30-3bc1d140ae3d%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/5a53f0b9-895f-afbb-b0cf-209652d3863e%28Office.15%29.aspx)|
|[TripleState](http://msdn.microsoft.com/library/e36d31b2-25e4-ab83-4a6e-def377ec6fe7%28Office.15%29.aspx)|
|[UseTheme](http://msdn.microsoft.com/library/770bea3c-4039-f6a5-a341-93d878d74085%28Office.15%29.aspx)|
|[ValidationRule](http://msdn.microsoft.com/library/2f7f967c-f98a-9d07-c2f7-7ce717d67e4a%28Office.15%29.aspx)|
|[ValidationText](http://msdn.microsoft.com/library/d42ad483-2720-2b9b-89f6-9611e345e44a%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/ab21bb39-e6ed-068e-85b6-16674a9638aa%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/4700f630-b040-e00a-4bc0-3cf6425632d2%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/84e5926b-a6a5-6590-20eb-92a3b129bfa4%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/77c69a42-4203-77ee-9d2e-b100cad9b75b%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)<br/>
[ToggleButton Object Members](http://msdn.microsoft.com/library/487101e7-c090-eb79-3671-5c9ce86cb6b0%28Office.15%29.aspx)
