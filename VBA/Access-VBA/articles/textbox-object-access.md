---
title: TextBox Object (Access)
keywords: vbaac10.chm11201
f1_keywords:
- vbaac10.chm11201
ms.prod: access
api_name:
- Access.TextBox
ms.assetid: d74fbe9a-0d40-7d28-956f-a2bfd0cfee45
ms.date: 06/08/2017
---


# TextBox Object (Access)

This object represents a text box control on a form or report. Text boxes are used to either display data from a record source, or to display the results of a calculation, or to accept input from a user.


## Example

The following code example uses a form with a text box to receive user input. The code displays a message when the user inputs data and then presses Return


```

Private Sub txtValue1_BeforeUpdate(Cancel As Integer)

MsgBox "The Text box is being updated."

End Sub

```


## Remarks

Text boxes can be either bound or unbound. You use a bound text box to display data from a particular field. You use an unbound text box to display the results of a calculation, or to accept input from a user (as in the code example above).


|||
|:-----|:-----|
|**Control**:|**Tool**:|
|![Text box control](images/t-txtbox_ZA06054010.gif)|![Text box tool](images/textbox_ZA06044637.gif)|

## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/609ef5f3-3894-85eb-4879-5db3fc7ff188%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/0d57cbce-bdbf-e19e-7f6a-11a00cb6c5f4%28Office.15%29.aspx)|
|[Change](http://msdn.microsoft.com/library/adde0a6d-d37a-a457-0dea-f2358adbb665%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/d102a526-2051-3a36-0f7a-fc234f126c47%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/ae8787e1-3425-bfbf-acf4-bbb97d42d2da%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/d6073892-7618-8e23-1fb1-795d3c76c2b6%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/970dc73b-8b8e-5811-bd4b-c23a96306bd2%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/05b5afca-4cb9-f12b-e05b-8702e35380d0%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/bc5d12a2-476b-a91d-2ad4-cdd6f46dd44c%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/00324700-f101-48a0-242f-bdabf4f2d70d%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/87db62a8-30f6-03d8-63ae-f1a1a50caea3%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/2219075d-92e5-a472-c16a-8a99dfd991c2%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/4c3a2696-5a78-5be9-7af7-205e7eb84dcd%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/ae184752-4c7f-3d79-5b3a-08407225f9d9%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/90d5d17b-8802-ec93-11ad-6be846bb1efe%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/0dfdc0b3-4a31-fd96-481c-d13db8197edd%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/ee009e53-41be-0c9a-a92d-15572f6213b6%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Move](http://msdn.microsoft.com/library/50b25305-0b91-378d-514f-d35b8d7aed6e%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/b1f8991e-7ccc-4f0b-c50f-1d51a0abda7e%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/dc5edcd0-09af-2fdb-0b94-49af0bfa705b%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/17289703-1943-2499-48c5-f34f200fd304%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/b019355a-7b78-4f03-878f-d2830c20117d%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddColon](http://msdn.microsoft.com/library/0a908d65-921b-7722-b564-cfe7a7fa8aed%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/690bc0cd-9717-7712-c022-75ba457ca0e3%28Office.15%29.aspx)|
|[AllowAutoCorrect](http://msdn.microsoft.com/library/9cafa161-c073-855f-edee-c7c9cb32be99%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/84a7ea86-f31c-775d-2383-5ac8751dd0f1%28Office.15%29.aspx)|
|[AsianLineBreak](http://msdn.microsoft.com/library/2ee42bb4-e6ae-c6b4-ef6a-71de5d35edad%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/a5e6e68c-eadc-a242-ef83-8b388f6ca41f%28Office.15%29.aspx)|
|[AutoTab](http://msdn.microsoft.com/library/27b17921-cd58-e243-e091-2686c64a7c02%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/7880c596-7a47-39b6-74ad-8036355a8e0f%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/36db2540-6d5b-ed43-a303-70b6282398cf%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/95a277c8-df48-79a5-c232-2cfe32eae8f2%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/a66a4839-3ab9-4867-b725-e613527bc94b%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/3740b360-334c-db71-9fb6-1f7aab304811%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/de841054-a98a-7108-0d7d-020175edb1ce%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/7522b663-4ce6-34a6-51db-7de503e01f04%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/554920e1-e5ae-1c48-f5d5-ab964070bec0%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/783c9424-669f-fcc7-b23d-6f5de03bad79%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/44f012fa-9021-0910-85c0-48a3b6c82141%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/3e48aa7c-ed95-aa27-f092-70d5fb2f9fb1%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/e842887f-9ec1-4405-0558-6b3b3d3d221c%28Office.15%29.aspx)|
|[BottomMargin](http://msdn.microsoft.com/library/a6ef1155-24c8-1254-614b-c912fda8dae2%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/75d2b8bb-c5c5-1d00-b175-8db80a7525c5%28Office.15%29.aspx)|
|[CanGrow](http://msdn.microsoft.com/library/5e96e693-9e1a-1f1f-5d5d-672e6232c330%28Office.15%29.aspx)|
|[CanShrink](http://msdn.microsoft.com/library/d4ac842c-18ea-a3be-a90a-5dd9d10d7b8f%28Office.15%29.aspx)|
|[ColumnHidden](http://msdn.microsoft.com/library/4014ea78-92f8-f1a8-6d73-ae7b2c5088cb%28Office.15%29.aspx)|
|[ColumnOrder](http://msdn.microsoft.com/library/b5b271bc-5b3c-9b2c-ec87-524be29597d0%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/19060aac-ccb0-3998-39c7-42f1454c339e%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/00d5dede-0583-9f0e-191a-28f91a0327b3%28Office.15%29.aspx)|
|[ControlSource](http://msdn.microsoft.com/library/be912167-402a-1bc4-6feb-c3551eb058a8%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/a63f3624-8f31-97f6-c2cb-8c34c82c825b%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/4cc842d9-2985-b65e-e259-697cedaa56fc%28Office.15%29.aspx)|
|[DecimalPlaces](http://msdn.microsoft.com/library/cd032c51-34d1-18d3-c378-7473938ec1d7%28Office.15%29.aspx)|
|[DefaultValue](http://msdn.microsoft.com/library/fab86da0-e865-478c-80c6-7681c5733059%28Office.15%29.aspx)|
|[DisplayAsHyperlink](http://msdn.microsoft.com/library/4741039e-9985-ac0a-9b74-309fcac860bf%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/6e5fa1c0-a264-cbc1-6fdf-9aef6c7f6bab%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/a13297e5-091c-7e83-78cd-fa67f5b81153%28Office.15%29.aspx)|
|[EnterKeyBehavior](http://msdn.microsoft.com/library/b7830316-a1aa-ddc1-094f-5976c5298bc1%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/a8cd7cdc-605b-473c-95b1-9d1736e0ec96%28Office.15%29.aspx)|
|[FilterLookup](http://msdn.microsoft.com/library/5c568366-94a5-8d7a-1fb4-80b4b3ab6c7f%28Office.15%29.aspx)|
|[FontBold](http://msdn.microsoft.com/library/147d151a-b51c-5be2-56ef-8a94c212cb0b%28Office.15%29.aspx)|
|[FontItalic](http://msdn.microsoft.com/library/f982c1ce-ad47-a05e-6b12-1eb51dbc0eb7%28Office.15%29.aspx)|
|[FontName](http://msdn.microsoft.com/library/4eb7cbbe-1e7d-9e29-cbff-867a83605741%28Office.15%29.aspx)|
|[FontSize](http://msdn.microsoft.com/library/73bf8d74-c616-8824-c2e0-8eed072df582%28Office.15%29.aspx)|
|[FontUnderline](http://msdn.microsoft.com/library/67bf0551-21c0-73cd-9418-dc7b3582f53c%28Office.15%29.aspx)|
|[FontWeight](http://msdn.microsoft.com/library/4dbf8092-c09c-c6ec-9476-20af2e9cf051%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/125bc04a-b747-6397-33ff-31de47004633%28Office.15%29.aspx)|
|[ForeShade](http://msdn.microsoft.com/library/b8437ede-edd1-7d86-1c2f-78d4ed1c3d0e%28Office.15%29.aspx)|
|[ForeThemeColorIndex](http://msdn.microsoft.com/library/9b49e363-fe5b-0536-c3ed-b4836acb383b%28Office.15%29.aspx)|
|[ForeTint](http://msdn.microsoft.com/library/8229f864-5ed3-309e-ba29-6a45bf9d59a8%28Office.15%29.aspx)|
|[Format](http://msdn.microsoft.com/library/c89491e2-09f8-d928-1aed-9d839545a694%28Office.15%29.aspx)|
|[FormatConditions](http://msdn.microsoft.com/library/6c643d8b-9b90-2b50-2ba0-c46bb821d38d%28Office.15%29.aspx)|
|[FuriganaControl](http://msdn.microsoft.com/library/7d70cffa-06bb-fa9d-686a-0031558aa5a3%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/849e0843-ab35-90d6-02a6-44faa316c83f%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/33daf4ec-1587-63c8-4b23-2abdf5087bbe%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/c58d8030-fc96-a53b-4cb4-5bb21237e20e%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/f1c71748-a37c-d0d0-5d8e-9899cf1efba5%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/c841157d-6e8d-8cd4-e23a-77d00d0af8e6%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/57a47306-5b85-06e0-e59f-f86e617d9c75%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/2c67d4b5-47d6-5430-cac0-bc05c3151305%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/5dbbd8a7-0942-c39d-b702-a3c0e569e3c1%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/4569d053-008b-a4ce-374f-6078f5ea9bee%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/0794df4f-88e2-5c75-13ba-88bbb8d7eb40%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/6abe0945-a6b9-72b2-e63c-1109fc7455a8%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/bb49f001-83a9-f1b8-c095-33b8b3f820b3%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/3f3d88d9-3561-a25b-5f22-a21b9cad6673%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/6829c95e-d7fc-c3c6-a8ab-0051c8e9af24%28Office.15%29.aspx)|
|[HideDuplicates](http://msdn.microsoft.com/library/a625d232-07d8-23d9-a28a-d01c70aa3a95%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/85dc54b2-7a20-4667-ade9-47202f77d524%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/a5d80cd4-d03d-41ea-9394-214537dd6c8c%28Office.15%29.aspx)|
|[IMEHold](http://msdn.microsoft.com/library/0cb93c85-07ff-a10f-5cd0-dc4045ce1079%28Office.15%29.aspx)|
|[IMEMode](http://msdn.microsoft.com/library/fa4adf03-7c20-eade-4a28-e3c3ac64ebc3%28Office.15%29.aspx)|
|[IMESentenceMode](http://msdn.microsoft.com/library/399a28d4-83a9-33d2-5f00-4f388efe048b%28Office.15%29.aspx)|
|[InputMask](http://msdn.microsoft.com/library/a705c2a4-ff2f-74d1-4a7c-1eade3b00ae8%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/6ebb497c-00d0-a854-be22-6b034deae98a%28Office.15%29.aspx)|
|[IsHyperlink](http://msdn.microsoft.com/library/68d2ca6a-7ea2-a44d-db32-1fa040475267%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/34487db4-6377-04f2-6848-a27dc5f4bab6%28Office.15%29.aspx)|
|[KeyboardLanguage](http://msdn.microsoft.com/library/a3b55e3e-16a9-87c7-6c03-bc8392e72c17%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/4714927a-9ce9-b1b0-dbeb-302aaa48a6c8%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/4d3ce486-bd79-cd6d-5886-a0a64cc7abb5%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/398b268c-271d-566a-36a6-1d703bdb0345%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/a1c841e6-221b-3ba6-4212-d76066afda48%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/b77ccc32-fbaf-e574-b0ae-293d6f999879%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/a184b336-215d-ffe0-d7ce-92f1fdc3b656%28Office.15%29.aspx)|
|[LeftMargin](http://msdn.microsoft.com/library/9c5b798b-4afe-85be-aa06-eeff98888850%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/0ceae1bc-f075-2e5f-48bf-7f749bae0630%28Office.15%29.aspx)|
|[LineSpacing](http://msdn.microsoft.com/library/3ac1c335-4b26-1a14-e4dc-bd5d56f44a2b%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/025b88db-7409-4cb6-bcc0-c72a6a3850d3%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/e97043b5-216f-2c5c-a531-45b29477cb77%28Office.15%29.aspx)|
|[NumeralShapes](http://msdn.microsoft.com/library/f0fda4bb-2522-622c-24ab-d3324a4b8dca%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/6064f8b9-31ec-da00-0346-cd259b917daa%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/d62150d2-6dc6-85c0-0452-e9e5fee199b4%28Office.15%29.aspx)|
|[OnChange](http://msdn.microsoft.com/library/ea25341f-fd30-62b1-476d-29febf4db4b4%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/54f32b3d-16df-376d-b5c0-9bbb2ff0931a%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/571a01ff-b92b-bb9b-1363-43086ef71c02%28Office.15%29.aspx)|
|[OnDirty](http://msdn.microsoft.com/library/312418b3-29cf-0d53-d92f-aaca6ec192b3%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/d8f7aa7f-5222-ec0e-7be9-4022f5e697af%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/2489acdf-4cf5-8b49-e9fe-fc78c07a87f3%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/3a180b9a-d415-b124-f884-9ce64dba8358%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/472e4b96-a6b1-6473-ed56-64af3522281f%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/458d2e2d-3003-79e4-a911-058928c25cef%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/77ebdf97-ae3f-73f4-d670-3c99d1f4f87d%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/1606cb80-bf56-3766-d939-b545c2738e17%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/2392c2eb-6c3b-047b-1e4e-2772989ba1f2%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/7201a61b-5b69-c13f-63bf-a2a5f329ecc5%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/acd5de89-de56-e7c4-1a5d-cc560c5cffb6%28Office.15%29.aspx)|
|[OnUndo](http://msdn.microsoft.com/library/fa62ba10-c8e8-f4d4-5d48-ab73c074f2ef%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/e07da876-e24c-0828-e986-d13a0cb1f78e%28Office.15%29.aspx)|
|[PostalAddress](http://msdn.microsoft.com/library/04fb29c5-909c-a0b8-a4aa-7701abc07037%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/54a6372b-77db-5557-7af1-0c608f6d46a6%28Office.15%29.aspx)|
|[ReadingOrder](http://msdn.microsoft.com/library/1b53bb00-9252-ca99-c3b7-3a97d06552c4%28Office.15%29.aspx)|
|[RightMargin](http://msdn.microsoft.com/library/13f3fe1f-d5c3-33ac-9b9b-897df8ff5ba9%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/7f9e2e21-1e36-01c1-f4e7-b3373644f9e5%28Office.15%29.aspx)|
|[RunningSum](http://msdn.microsoft.com/library/8918a58c-8c07-84dc-f43c-2486d54cd677%28Office.15%29.aspx)|
|[ScrollBarAlign](http://msdn.microsoft.com/library/5a8a77df-571a-7294-8be8-0ff2c4546131%28Office.15%29.aspx)|
|[ScrollBars](http://msdn.microsoft.com/library/de3adbf1-4398-8782-0998-d392ab860669%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/76a43ccb-a199-b640-623c-d008b7d48e1c%28Office.15%29.aspx)|
|[SelLength](http://msdn.microsoft.com/library/0fb2371d-0f60-b0c7-5c4b-7a0689867b21%28Office.15%29.aspx)|
|[SelStart](http://msdn.microsoft.com/library/51c773bb-2b70-b812-6b6a-9e062e493ebb%28Office.15%29.aspx)|
|[SelText](http://msdn.microsoft.com/library/1625b16f-8c2d-a563-6f66-a6714f5419ec%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/620de877-2164-6426-90b8-c72a6db637fd%28Office.15%29.aspx)|
|[ShowDatePicker](http://msdn.microsoft.com/library/5d65938b-ac7b-abbd-2e50-41f41c0b1558%28Office.15%29.aspx)|
|[SmartTags](http://msdn.microsoft.com/library/200175d1-78a2-3036-72ba-4a85dfc21864%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/9d34e61b-9ba9-02e0-4bd8-30da0a043a89%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/18ae7a69-2e63-7896-1bff-da3f45b62c63%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/d52e0839-e0aa-1b67-b075-115ad7b2f774%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/ecb9ede6-e7a8-ca91-9ca3-3fad9de2ef90%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/9df21640-6bea-60a9-f9d0-dac90a60af1c%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/bb510c65-6d0d-468a-c5be-f325d86c2c7f%28Office.15%29.aspx)|
|[TextAlign](http://msdn.microsoft.com/library/2b6e5ad7-02f5-4e33-47a4-87882a3113b2%28Office.15%29.aspx)|
|[TextFormat](http://msdn.microsoft.com/library/3d164782-9d9c-5462-ac40-51772d475407%28Office.15%29.aspx)|
|[ThemeFontIndex](http://msdn.microsoft.com/library/2abe2063-4658-e441-7a7d-c4d226063172%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/6a220cec-d42c-05e3-c8c0-078687813a8d%28Office.15%29.aspx)|
|[TopMargin](http://msdn.microsoft.com/library/cd56b2b2-8bb5-b3cf-bacf-13d311e5479b%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/fd6420f1-c3d9-2374-7b3c-e1fa5dd8199a%28Office.15%29.aspx)|
|[ValidationRule](http://msdn.microsoft.com/library/e481fba1-7e08-f8da-b644-5e38c2bf445e%28Office.15%29.aspx)|
|[ValidationText](http://msdn.microsoft.com/library/5d3ab2a3-9166-714f-a0c2-d56d42b19ebc%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/4cb4c33f-dd96-0309-f30b-8e445d123756%28Office.15%29.aspx)|
|[Vertical](http://msdn.microsoft.com/library/40b9f9c0-daab-5562-395e-3e785d316d91%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/b515b37f-0566-0483-d387-8bc02c7be980%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/af1b9264-53f9-bf4c-2f05-049288a1d3d5%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/0bb72524-6682-f783-e9f9-4fd34a757a40%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)<br/>
[TextBox Object Members](http://msdn.microsoft.com/library/bb55abbc-902e-fc2d-bdff-063c55426cd0%28Office.15%29.aspx)
