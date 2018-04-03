---
title: NavigationControl Object (Access)
keywords: vbaac10.chm11201
f1_keywords:
- vbaac10.chm11201
ms.prod: access
api_name:
- Access.NavigationControl
ms.assetid: ab08e35c-e5e4-444c-d169-1092d282ed15
ms.date: 06/08/2017
---


# NavigationControl Object (Access)

This object represents a navigation control on a form.


## Remarks

The  **[NavigationControl](navigationcontrol-object-access.md)** contains a collection of navigation buttons, each of which is represented by a **[NavigationButton](http://msdn.microsoft.com/library/ac6ba9b4-45aa-0d92-d01d-fd8e8b9cede6%28Office.15%29.aspx)** objects. When a user clicks a navigation button, the assocated form or report is displayed in the control specified by the **[SubForm](http://msdn.microsoft.com/library/e99cec35-3186-98ec-3318-0bcfb47e97ba%28Office.15%29.aspx)** property.

Use the  **[Tabs](http://msdn.microsoft.com/library/a8b2546c-9b1f-a8ff-1a6f-8e607415ffec%28Office.15%29.aspx)** property to return the collection of navigation buttons for the navigation control.

Use  **[SelectedTab](http://msdn.microsoft.com/library/8e6da4b2-eada-51db-b198-da8213c647ac%28Office.15%29.aspx)** property to return the navigation button that is currently selected.


## Events



|**Name**|
|:-----|
|[AfterUpdate](http://msdn.microsoft.com/library/ae34fff1-4521-4ec3-707a-f1f2c49f7946%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/397c8bb2-1c8d-fa32-5015-65b58b215b38%28Office.15%29.aspx)|
|[Change](http://msdn.microsoft.com/library/43a0c20c-24dc-3be7-42fd-c000cd2dffb3%28Office.15%29.aspx)|
|[Click](http://msdn.microsoft.com/library/c49b26bd-dbab-666a-ecc0-2b3137bb10a0%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/f8f4f4d1-fbb7-e6aa-513b-fe434e50caa9%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/6125891b-c0cf-0b0e-0678-146404b2ed31%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/30741318-953e-4dde-54df-ef6fca845844%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/501b17c7-0039-7418-e31c-7c61c49691dd%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/50ebdaad-3e2c-9eff-47f0-43a402b17938%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/02b0671c-706c-960c-73d9-76301914aa65%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/e6dd9500-c6c9-ff51-fad8-2d542cf6bff6%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/35e7a26d-617c-9e51-c246-1830cd180420%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/6098212b-fd3b-0868-1112-9f52ae886e7e%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/0406fc90-fa66-b436-6761-c16915e37b5d%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/a5676866-db8b-078d-70dc-ee159c66671c%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/174c4b0d-9906-5f73-80a2-a59b3d66aae1%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/ebab443e-6abc-ed4a-5f2a-4ad00c7f9d8c%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Move](http://msdn.microsoft.com/library/bbf4e87e-8468-7cfd-7cd4-5f423a6517c8%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/613e89e1-5e02-d2da-4881-c77f3d8cb55e%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/74232b27-17f4-78fc-9c42-0aabaad56257%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/690d17ca-866d-2f8e-fc54-a5cc166b6ad1%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/d15daeaf-5c78-5833-9fed-d57d2996e60b%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/b980f9dd-1d8e-8296-8e4a-17051b5fcd4e%28Office.15%29.aspx)|
|[AutoTab](http://msdn.microsoft.com/library/3d484269-c00b-3f5e-8492-6e0ca60460b8%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/d765586f-9454-756d-b6eb-b61bdde9ea16%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/3c3de7b4-9b86-6148-69af-f4a3ccb648ff%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/1f46ccfd-78cc-0eae-3485-b91306dc6bde%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/7f9e0ebe-0b25-28ed-5b68-e5ead2c72ef0%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/cabea08c-a59c-ac0d-d40c-62f0e7b475ac%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/7fea7ca6-0363-c741-6a29-128628c1210a%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/5464f403-791a-d324-2c7a-eb6aa26acf8f%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/9135151b-2e00-ac34-9c82-a85c76b97eb5%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/9ddd1a71-e974-c70c-0240-80c695c30e35%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/8e0a943d-f863-7bd6-6491-5661b3b58556%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/a0a39f30-18c5-2073-b463-1ffcb385357c%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/fb22d41c-a310-ed95-34ea-8a4cda1bff8b%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/68c6abcf-7bb7-4795-8c6c-685ed1c25dc9%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/3952c7f5-e5d1-7a7d-3187-d4c327a33fe0%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/c0259524-8505-71a1-e482-9f142379f9e8%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/e1e91c9b-aba6-4bf1-6b54-6c64badfa7af%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/6296dabf-95a3-6751-7572-95522f7bd57c%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/d59c7baf-7614-821b-92ce-582d6f90441c%28Office.15%29.aspx)|
|[FilterLookup](http://msdn.microsoft.com/library/c368853c-6a1c-f104-2180-ebc889cf7e6d%28Office.15%29.aspx)|
|[FormatConditions](http://msdn.microsoft.com/library/20e921d6-e800-fc75-c93a-981815d694ab%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/21502538-377c-fd82-62bb-c68cabd1b2cd%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/f095b4d4-6c8b-5e17-6282-f4e97a7ef21f%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/9bd6575e-a0a5-0757-c517-a694b04130e8%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/66383fb2-d44d-c979-a025-52c4a4a369ea%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/88e8a163-84ef-8f4c-f7b2-6dd2783389d1%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/53782607-fd23-26e2-ae48-721786cd20cc%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/58676faf-b4cb-ce1b-a28c-dd93c491b025%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/fff9f85b-c978-3a87-371d-5ad0efa85a38%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/513fdb37-b479-7022-e0c7-4f8d8209ede9%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/884b81e2-4941-364f-b195-1731706bbd3d%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/1649cfc6-d968-8e51-de44-1ece83c7a5ca%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/e9d2180e-6037-a040-7b57-1be74587e49b%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/bf46c094-6eef-452b-dca9-ff6d4a3e5006%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/b56cbf60-e760-170c-9c93-edaddabf91b6%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/2e6142a7-1d9b-ec43-5ee2-0388f5d401f4%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/2d6bdb1a-808e-1712-1846-71ffa8619f0d%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/96b49172-cea7-26e3-0bdc-6e0b85a1402f%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/659d5713-a385-bead-68a0-501a724e9210%28Office.15%29.aspx)|
|[KeyboardLanguage](http://msdn.microsoft.com/library/5a4f4c8b-2d01-4613-2bb0-8c3e2c7dfda9%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/eb8ab5e3-2443-d755-6dfa-6432223e87c0%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/2fd85cf8-90c3-9b00-6d2a-9078be79f668%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/3e4f76fa-9e5c-a501-ae7f-38dfd89a836a%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/0a082747-dd3d-2ad9-b5e4-4911bd639750%28Office.15%29.aspx)|
|[LineSpacing](http://msdn.microsoft.com/library/bf1d5cef-8f0e-f759-3499-2f567097800e%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/0daea497-ec28-769b-6722-4ac60026147c%28Office.15%29.aspx)|
|[NumeralShapes](http://msdn.microsoft.com/library/207bbece-366e-bc72-876f-98c80f7bf6b5%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/57f378e9-7211-1d05-15d0-0bc1b2f2f4b3%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/ddee64e6-38cf-d033-4963-76529744ef81%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/41352c03-f034-a882-9ef1-05b06c2f51af%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/19b575b9-a727-85e0-f5c3-c4ebe3bbd987%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/62e5608d-c002-cc2b-305c-90b9ba68b527%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/8de375d7-da00-318a-2a1a-7d2fb26bd11d%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/5efcc70d-6609-d4b3-509c-063af66195c4%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/ac069657-a9de-79f2-2e7c-92e151228f2a%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/1f7496cc-7550-d9cd-c7bb-d461775d8fed%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/c8258e0e-c115-2556-a929-753c510fdc49%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/12259131-0b06-e01f-4a94-05dabaf0e53c%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/fc251872-bc0b-d3a3-1426-fdb121b24145%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/e6b36fe8-b4d3-6571-0965-f27ac611fd29%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/682d75b4-5bfd-ea22-c47a-ceb7a4d504f2%28Office.15%29.aspx)|
|[ReadingOrder](http://msdn.microsoft.com/library/ecd7522a-3a16-2a18-a3c1-0798dba1baec%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/8c2cd0df-f629-e8d1-a2df-ba0f6203ec07%28Office.15%29.aspx)|
|[ScrollBarAlign](http://msdn.microsoft.com/library/b685e196-513e-fe57-d993-d1e2f4051a4c%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/670b7950-5f94-461e-8cd1-9c6f95169e89%28Office.15%29.aspx)|
|[SelectedTab](http://msdn.microsoft.com/library/8e6da4b2-eada-51db-b198-da8213c647ac%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/89e4e907-4d28-6c9b-424c-3400d448b222%28Office.15%29.aspx)|
|[SmartTags](http://msdn.microsoft.com/library/e4c3553a-7ce3-291e-b83a-c88e20685b4d%28Office.15%29.aspx)|
|[Span](http://msdn.microsoft.com/library/a1a26d1c-5c3d-8f3f-c12c-88a0dc40aa0f%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/ab1cb63a-d51b-cbd3-bf40-d52148925556%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/8cd0c070-a8ec-e5c3-8996-a551cd344da5%28Office.15%29.aspx)|
|[SubForm](http://msdn.microsoft.com/library/e99cec35-3186-98ec-3318-0bcfb47e97ba%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/2fc2511e-5a92-7039-cfec-2556b3384fb7%28Office.15%29.aspx)|
|[Tabs](http://msdn.microsoft.com/library/a8b2546c-9b1f-a8ff-1a6f-8e607415ffec%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/40aeb05f-b94f-ee88-5e98-0f77599c7a14%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/6bec7ae8-556c-77b1-19cf-aae36dc646ec%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/76681117-639d-8e4c-4a3b-7c68e3863928%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/56cae307-f23c-d2e1-5095-fe6b696a6d98%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/9e45f505-81d3-63e9-b0c1-7182372224ad%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/0018fcea-2b3b-3e57-8055-4aaef922f999%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/91ea0e8c-63d1-3ca7-7f26-748f1651a1c6%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/373efc78-6b33-827a-5b95-9cc9fff7f9e6%28Office.15%29.aspx)|

## See also


#### Other resources


[NavigationControl Object Members](http://msdn.microsoft.com/library/c972327e-9b46-f9fb-d69d-104d1d130ee4%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
