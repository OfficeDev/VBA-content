---
title: CommandButton Object (Access)
keywords: vbaac10.chm10554
f1_keywords:
- vbaac10.chm10554
ms.prod: access
api_name:
- Access.CommandButton
ms.assetid: 25e7c0b7-03c1-dffe-8f52-4ec59739f6b8
ms.date: 06/08/2017
---


# CommandButton Object (Access)

This object corresponds to a command button. A command button on a form can start an action or a set of actions. For example, you could create a command button that opens another form. To make a command button do something, you write a macro or event procedure and attach it to the button's  **OnClick** property.


## Remarks


|||
|:-----|:-----|
|**Control**:|**Tool**:|
|![Command button](images/t-cmdbtn_ZA06053979.gif)|![Command button](images/command_ZA06047243.gif)|

You can display text on a command button by setting its  **Caption** property, or you can display a picture by setting its **Picture** property.


 **Note**  You can create over 30 different types of command buttons with the Command Button Wizard. When you use the Command Button Wizard, Microsoft Access creates the button and the event procedure for you.


## Events



|**Name**|
|:-----|
|[Click](http://msdn.microsoft.com/library/b84b7acd-c428-8cdb-7fc3-b1963e7102a3%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/0bce5cae-67d8-3acd-2029-be72f511e250%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/cc60adbd-eb72-92c3-a562-08adbf0dcc99%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/d31c55ca-a2d9-7576-0a7f-a19307c36e87%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/b8ad669d-6353-ff62-5b06-5fda93d50327%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/d2bc24b6-62c8-dd3f-82af-600f045e2df1%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/afdc1037-c0fd-d5f2-3ccd-bc67c98aa482%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/6466c06a-d3fc-8187-82dd-7a5c332049a3%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/a8c29b13-5757-7be9-7111-81f847c8ec32%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/8daa650a-ebd8-6e87-a933-d5b1f240ded6%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/f20d4807-42a8-5c90-e18a-1208a138241c%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/92cbef4e-deee-1c5f-ec0e-10bc5e6ebd5b%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Move](http://msdn.microsoft.com/library/58c51741-fb49-4b0a-91e0-cb9486808597%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/9a8fed17-aec2-c592-c003-92bc832d5da0%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/ec0c4c1a-72cb-f766-c05b-fc1e99e5c8e9%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/a1e8f47f-30b4-c2f4-7d95-2be75f0a4aa5%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddColon](http://msdn.microsoft.com/library/cde177a8-b5a8-5063-d061-a81dfbfc2857%28Office.15%29.aspx)|
|[Alignment](http://msdn.microsoft.com/library/b0081eea-1149-d173-646a-0800aa558415%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/c71d31ac-daa0-3790-f456-185eba48db30%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/f71bdc7f-9703-eeaa-70a8-70b6ddb72f85%28Office.15%29.aspx)|
|[AutoRepeat](http://msdn.microsoft.com/library/028a5bdd-1e37-0499-202f-c9e3fdb83838%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/aa546889-e77e-35fd-0e98-be020a94cb65%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/31628a36-f0f9-92df-99ee-1540ed3831e6%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/b7c930b0-e203-fe3a-ce54-0778d65d073f%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/891e6d86-5935-1d75-1cda-ea5bcb422583%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/db441cd0-bd88-2c76-aab1-ae846974b8bd%28Office.15%29.aspx)|
|[Bevel](http://msdn.microsoft.com/library/b9bd9082-75b3-e249-a477-ce402bff1e43%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/266c8082-30c4-0182-3004-b02b5a9c4a7b%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/52816058-36f4-3b68-38dd-5a1324b87428%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/ba7b7eb5-5f1c-addd-483f-a3104a50115b%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/497b2f7a-9b17-79ed-1ad9-fc990f6b9c7d%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/244697f0-891f-792d-3ad9-61a58973ab60%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/a59dbd51-e839-145b-2971-82bdc4c21097%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/c7806653-3e00-824e-f1af-7092369af0a7%28Office.15%29.aspx)|
|[Cancel](http://msdn.microsoft.com/library/a45d52e0-7566-2d16-8f74-7168a380f6a2%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/9141b138-5bf7-5d45-f945-f9de41e43042%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/017d583d-671e-7d9b-bdae-d67a7d94b4a8%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/394aecbe-0053-d114-1804-c4ee6a9749d0%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/c41e555b-195b-7af9-f2ee-09d92980e557%28Office.15%29.aspx)|
|[CursorOnHover](http://msdn.microsoft.com/library/98bfdba4-4b42-8bbc-e1d2-d68cc21defc3%28Office.15%29.aspx)|
|[Default](http://msdn.microsoft.com/library/b643350e-9a89-a0ff-b8dd-f1c2c1392992%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/3775036c-c483-1c2d-6845-cd999104925d%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/c48d979d-3320-d8ab-1019-c5d1bf60e01d%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/89611b46-0c56-d855-9e4d-d1a301f300ae%28Office.15%29.aspx)|
|[FontBold](http://msdn.microsoft.com/library/6a736a00-6305-74cd-47b9-aa29b8a76d62%28Office.15%29.aspx)|
|[FontItalic](http://msdn.microsoft.com/library/a82d5e83-b892-a006-e68a-cda3c2c82d1d%28Office.15%29.aspx)|
|[FontName](http://msdn.microsoft.com/library/0e1099d3-92fb-a077-9148-e2f64305faee%28Office.15%29.aspx)|
|[FontSize](http://msdn.microsoft.com/library/3ceff45a-fe5d-f692-7ad3-ab20143e12fc%28Office.15%29.aspx)|
|[FontUnderline](http://msdn.microsoft.com/library/1882cbe8-3e22-9224-bb18-a5f3aa9cf737%28Office.15%29.aspx)|
|[FontWeight](http://msdn.microsoft.com/library/a7c0b157-c25c-24e5-b05d-cc8ab726ac7b%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/6d19e4b2-2375-fe37-c226-4489ebcb808e%28Office.15%29.aspx)|
|[ForeShade](http://msdn.microsoft.com/library/c8ddc31f-83a3-c836-e1f7-2ffe5ea86d4a%28Office.15%29.aspx)|
|[ForeThemeColorIndex](http://msdn.microsoft.com/library/4831634a-6988-57ec-0e47-6c16a6c832a0%28Office.15%29.aspx)|
|[ForeTint](http://msdn.microsoft.com/library/87b29d73-fdbf-0ffa-d2eb-78d182625458%28Office.15%29.aspx)|
|[Glow](http://msdn.microsoft.com/library/e6c147b4-c378-90bd-7132-f44021994ecd%28Office.15%29.aspx)|
|[Gradient](http://msdn.microsoft.com/library/6ab8ea87-7bba-6476-14e5-0d0e7e645d0e%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/ef5addc8-5e29-ef8b-f7f6-0b91c68e9bc9%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/044e8de8-e7c9-dd59-920c-529bc3e6a51a%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/77ee45fb-5dde-2925-d88b-da62a6f9ed27%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/4e98dccd-e0d6-b24c-0a7a-f8dd54907fa0%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/496c1c59-0111-8e2f-31b9-af2ee7ff3964%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/f6fb163b-ece7-08a0-b786-e32287d40e50%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/e736f508-fc12-0244-5f46-825bbbbc24c8%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/a24518ba-866e-be3e-dde7-bb3301c83353%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/693e49bf-fd74-b00f-0663-54f577179d3a%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/8c70fe5c-cf65-49af-558a-d5f28dd79f4a%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/f3d0cd61-c03c-92ba-6b5e-030d1efed9c5%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/dfa6bb67-9841-ddf0-508a-9553fbf0229e%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/40b8e9fb-8573-7bb2-9467-12ca5b593a04%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/265cf535-68b0-f627-f09c-c09b72d41aad%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/3b16ed18-a83d-df6e-5f14-6edbd25e9490%28Office.15%29.aspx)|
|[HoverColor](http://msdn.microsoft.com/library/00d4b912-fb14-2e63-ec4e-386ad4b9f0c3%28Office.15%29.aspx)|
|[HoverForeColor](http://msdn.microsoft.com/library/a1efabe5-1cde-00f2-319b-df72e0f718c8%28Office.15%29.aspx)|
|[HoverForeShade](http://msdn.microsoft.com/library/be9e6008-4cc4-94b5-869e-068c3b73443a%28Office.15%29.aspx)|
|[HoverForeThemeColorIndex](http://msdn.microsoft.com/library/7952f076-a8ac-c6d3-72f7-23e8365d8a16%28Office.15%29.aspx)|
|[HoverForeTint](http://msdn.microsoft.com/library/88922fd3-f8ce-5f07-f364-1155ac6070fe%28Office.15%29.aspx)|
|[HoverShade](http://msdn.microsoft.com/library/9a8b86d0-3849-9902-4dbf-c911c7fbe8e2%28Office.15%29.aspx)|
|[HoverThemeColorIndex](http://msdn.microsoft.com/library/7fec39e2-f79f-1260-ff6f-9e634ff18fe0%28Office.15%29.aspx)|
|[HoverTint](http://msdn.microsoft.com/library/0eac99ff-c693-d456-c319-ec1ce60ba05d%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/2f5ce470-967a-450d-f661-ac1e1f370d56%28Office.15%29.aspx)|
|[HyperlinkAddress](http://msdn.microsoft.com/library/7efa1230-955b-183c-a459-1b2598eb9163%28Office.15%29.aspx)|
|[HyperlinkSubAddress](http://msdn.microsoft.com/library/1c8af1e0-f978-0eb2-c3b5-f5ea9ab84892%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/8b8119a7-734c-8e20-8c1a-e80f02a8ad22%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/f5438725-4628-4f8e-1bf3-0027348b9285%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/a586c545-c1b1-ff31-5213-5a3cd055a2b6%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/04582d98-dbc6-4aed-e42b-f8d6638ba4ae%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/8daa4d29-ba7f-67fc-a640-d15a3886441f%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/61e0b921-ee37-af21-e84f-64f0b682e05c%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/8b511bf2-659b-f2d4-1aeb-0c238a7972a9%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/8cae225d-1919-0c6c-7980-48294fbe8c7a%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/a94afdff-4615-529e-04de-fcf3d9e63d2d%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/1e0f700c-9114-4add-4a0a-4f93266951d5%28Office.15%29.aspx)|
|[ObjectPalette](http://msdn.microsoft.com/library/e4c8ea81-b39f-e580-9a68-c809c0deaf71%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/a03e4e4c-0c02-7e6a-0654-fafc8a0f0036%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/1034aa82-58cd-f639-d936-326049ccf38c%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/465d95b4-64e3-1d1b-e388-5c96bfd2e5c9%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/6d8f659f-a8aa-4671-509c-c82ae5dead0c%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/8ff969a9-bb7c-9185-dba3-3259647fddbd%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/4d892495-791b-05b3-0bcb-3b3c3635a0bd%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/33945139-f404-ea8a-577e-2a3623f52cb3%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/de0dd03a-e3f4-c69d-0d9e-030fefc0a2de%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/dc4ad60c-4ba5-bf80-2e83-ee75da462e27%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/e3bddd85-772e-9d3c-d079-b323f10a7d5a%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/52b27f17-3df7-b0ab-23cd-7913cebaa979%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/1b24e970-1f29-af26-2d01-e6587812bf13%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/01abc8c3-031e-eb7e-1893-a4a7c6fbd24e%28Office.15%29.aspx)|
|[OnPush](http://msdn.microsoft.com/library/38fab0d1-495e-9053-5e24-932ae0d8bdce%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/288169cc-0934-43b0-a7b4-18445844519b%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/1d0d5956-719e-13eb-e6ca-319f8da78754%28Office.15%29.aspx)|
|[PictureCaptionArrangement](http://msdn.microsoft.com/library/b33ce40a-b247-9d69-c06d-17c822c80283%28Office.15%29.aspx)|
|[PictureData](http://msdn.microsoft.com/library/7208ecc7-c057-4ad0-c55e-15a7a710f0a4%28Office.15%29.aspx)|
|[PictureType](http://msdn.microsoft.com/library/a835b294-4de1-b948-e59c-a7e9c3a4f9ae%28Office.15%29.aspx)|
|[PressedColor](http://msdn.microsoft.com/library/c5f446e8-d1a2-f4c9-08c1-7a809b5ae5b8%28Office.15%29.aspx)|
|[PressedForeColor](http://msdn.microsoft.com/library/b3174fa6-d89a-906f-ef4d-19b489734dfa%28Office.15%29.aspx)|
|[PressedForeShade](http://msdn.microsoft.com/library/496e310e-b5eb-8e6a-7079-530126e71399%28Office.15%29.aspx)|
|[PressedForeThemeColorIndex](http://msdn.microsoft.com/library/32ad73cd-3960-1516-c45d-175c7d642847%28Office.15%29.aspx)|
|[PressedForeTint](http://msdn.microsoft.com/library/3c5bce3c-e140-cd4c-ef69-7aee89b89998%28Office.15%29.aspx)|
|[PressedShade](http://msdn.microsoft.com/library/8aa77c14-e9da-d4a2-015d-f1a2c2ced859%28Office.15%29.aspx)|
|[PressedThemeColorIndex](http://msdn.microsoft.com/library/12d76216-6a16-c487-02b3-721ed5e27b79%28Office.15%29.aspx)|
|[PressedTint](http://msdn.microsoft.com/library/11439c75-f951-a551-12ee-b7b2d2e8ee94%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/2d819871-1a93-c835-7c2b-c42797dceaf8%28Office.15%29.aspx)|
|[QuickStyle](http://msdn.microsoft.com/library/ac5750b0-e4cc-4330-8391-7aaef008973d%28Office.15%29.aspx)|
|[QuickStyleMask](http://msdn.microsoft.com/library/c0661897-d71c-8c3e-b18d-1100a24ed6a2%28Office.15%29.aspx)|
|[ReadingOrder](http://msdn.microsoft.com/library/5a47e95d-7421-147f-084a-74130cf524c7%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/f5a02077-2598-3b5c-58c9-fa77d5947cff%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/0ef5f32e-b724-205a-94bc-337b76f0a1b7%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/71af60bc-6f69-1408-8f3a-076a75daddcc%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/92088237-5dd8-0b40-ed2d-e5a5bfef4495%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/fea5b3e5-da70-c3b6-95f6-bc06e7b6c762%28Office.15%29.aspx)|
|[SoftEdges](http://msdn.microsoft.com/library/a970945c-a8d7-4888-8408-33bfc803d73d%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/2dc18f10-0b6f-2ae5-21c6-52c6d21ff03b%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/f8b37846-6a65-6b39-9234-5cd77049c907%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/ec624311-cad4-87b7-e697-053c939a078a%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/5099e435-8957-e54c-9c6c-bc6b063cfe66%28Office.15%29.aspx)|
|[ThemeFontIndex](http://msdn.microsoft.com/library/8cb51c03-09a1-83ba-c6bf-7e74d7444c8b%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/0c2207a6-5d99-0409-58e5-6a9ec2716e77%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/475398a6-ab75-1e39-12dc-ba7056b8caa0%28Office.15%29.aspx)|
|[Transparent](http://msdn.microsoft.com/library/655e127e-7e2e-c2c2-a979-952f95c534a6%28Office.15%29.aspx)|
|[UseTheme](http://msdn.microsoft.com/library/b28982a6-1291-377b-91af-0421b8fcb9f4%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/e0da1883-eec3-39fa-2bff-1410d79a7b2a%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/980c1f93-ae95-3481-5358-ad5362ffc9e8%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/03729218-4c70-8312-ab61-be3cf4b7a029%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)<br/>
[CommandButton Object Members](http://msdn.microsoft.com/library/9e1c10e6-0d03-78fd-ac9d-3f086ce1e0f5%28Office.15%29.aspx)
