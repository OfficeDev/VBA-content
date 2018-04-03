---
title: BoundObjectFrame Object (Access)
keywords: vbaac10.chm11026
f1_keywords:
- vbaac10.chm11026
ms.prod: access
api_name:
- Access.BoundObjectFrame
ms.assetid: b3025672-60b8-e1d6-4769-1f724c9aa1ef
ms.date: 06/08/2017
---


# BoundObjectFrame Object (Access)

A bound object frame object displays a picture, chart, or any OLE object stored in a table in a Microsoft Access database. For example, if you store pictures of your employees in a table in Microsoft Access, you can use a bound object frame to display these pictures on a form or report.


## Remarks

This object type allows you to create or edit the object from within the form or report by using the OLE server.

A bound object frame is bound to a field in an underlying table.

The field in the underlying table to which the bound object frame is bound must be of the OLE Object data type.

The object in a bound object frame is different for each record. The bound object frame can display linked or embedded objects. If you want to display objects not stored in an underlying table, use an [unbound object frame](http://msdn.microsoft.com/library/4a0874dc-ecac-be7c-25e2-ecc79696e2eb%28Office.15%29.aspx)or an [image control](http://msdn.microsoft.com/library/1f938a6e-7aea-7787-d959-e21edaa9342c%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[Move](http://msdn.microsoft.com/library/69784663-48ac-5c7f-d21d-0b0f10ba7284%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/e2936b7f-da7e-7b61-5ada-cbca28a29385%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/e4733d2b-1cce-36c1-428e-09df2b4e23e3%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/64f75ca9-e0bf-860c-7f62-b1f37f930893%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Action](http://msdn.microsoft.com/library/d75eea30-bee7-0b8e-f67c-8682cd696262%28Office.15%29.aspx)|
|[AddColon](http://msdn.microsoft.com/library/8356291c-9c96-6d6a-b05c-4993fe7cc93a%28Office.15%29.aspx)|
|[AfterUpdate](http://msdn.microsoft.com/library/3ebda4de-49c3-bfe7-8743-1c2c98caca58%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/05b5b479-fe8b-6d03-b8de-59afa7a587b9%28Office.15%29.aspx)|
|[AutoActivate](http://msdn.microsoft.com/library/162dcc86-818c-dc84-48cd-97fbfb85b77c%28Office.15%29.aspx)|
|[AutoLabel](http://msdn.microsoft.com/library/9a2b49f1-e0e6-9f4d-065a-c24fe07b23f3%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/84bad360-2e1d-0f8d-2751-c2d23fa8bb23%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/17c2e087-d4c7-f27d-a3a0-01470aa2b348%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/335ce425-d682-831a-ecfa-4c46b9bf5a28%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/e0efd6e0-9d58-85c8-0bac-1456044013cd%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/ac815c96-c30f-57e0-01e8-db12fd98a50e%28Office.15%29.aspx)|
|[BeforeUpdate](http://msdn.microsoft.com/library/01ee3c67-76c6-b651-042b-a7aa59e7443e%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/e3c43808-1254-2635-264e-2f3e79cb2c8a%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/5a8baed2-9f1d-e835-013b-b3973e79e228%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/057564f7-bb3d-3033-538e-86db4648c6b7%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/89423dbf-44de-a2e6-d31a-6ea459c2f156%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/f171677a-d8a2-f0fb-233e-636ec13e20f8%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/f06fa232-f6cd-7736-aeb9-96461d2338fc%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/beb8e3a2-5656-7ce3-7e20-1b99705139cf%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/0938d124-efd2-63c1-4282-a06fb412185a%28Office.15%29.aspx)|
|[ColumnHidden](http://msdn.microsoft.com/library/dfe097a5-18cb-5ee3-9122-cc790159c71e%28Office.15%29.aspx)|
|[ColumnOrder](http://msdn.microsoft.com/library/182e4cdf-f6e3-bf7b-5080-23b5d3cddfe3%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/3e7601ce-5aff-9f9e-feae-7ab6b9e35869%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/65113d53-fa59-ff69-c398-2ce42abd9e0b%28Office.15%29.aspx)|
|[ControlSource](http://msdn.microsoft.com/library/f4200d00-fcb8-f15b-68e5-f1e58bfe41e8%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/a6bf0845-9733-193d-e02a-b1dc90802b02%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/16ccb55a-9866-fd21-12a1-791e2c460db2%28Office.15%29.aspx)|
|[DisplayType](http://msdn.microsoft.com/library/95213bcb-9751-b43c-9722-6326d0fa8f25%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/ef64a05d-562f-2aff-09aa-b3d5609854b8%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/58b19f0d-8460-0a51-739f-9fae5de20901%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/20d82dc1-6bb4-0338-6bfb-ce801825634d%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/2cc8616d-e480-2e10-52a6-6914d58579ac%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/3fe4929b-9545-e886-f33c-9cae9f0c5f28%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/422508e7-d735-55a5-04e9-b0297536c2f5%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/1933af20-09e9-8a62-a127-cbd40b872b1c%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/498ba715-b84b-d5d9-51a1-5e085a67422b%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/8b8a6626-a0c5-e08d-f256-3d99b47aa984%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/1d527006-46f3-fc31-a579-ff2b32a104cc%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/bdb98dd5-ec7b-1e39-d39e-66e841b1090e%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/5fe7525a-20e9-a9f8-b93b-c4bcf1ebdfcb%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/0a29f26d-b2b7-67f5-ef8e-a76bd603e462%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/5118a22e-0339-ffda-96e6-7dfe54b26cf7%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/1427956f-17ec-9195-a754-ffa2f2968ed0%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/53f59551-041e-dc9e-4eee-ed0d5cad0603%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/81fd943b-58b7-eb51-7578-6b124794d359%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/b6f0b03d-8c64-ca0e-1efc-1b017aa6b615%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/8212bebc-0d9c-6476-a8f6-f1bbd3c26066%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/e750fe64-ee9a-5b42-2f5b-da8017002960%28Office.15%29.aspx)|
|[LabelAlign](http://msdn.microsoft.com/library/760ec42b-01ee-eb3f-2998-79ea7caf5578%28Office.15%29.aspx)|
|[LabelX](http://msdn.microsoft.com/library/1e2dcc6f-f192-aac2-060c-9b848ca18d10%28Office.15%29.aspx)|
|[LabelY](http://msdn.microsoft.com/library/1c68016d-9be5-b550-1b97-1840ed36f974%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/f2c64167-b3d0-098f-8a86-755efc28f548%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/35cf3634-7e5e-8b38-27b2-b13dec239366%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/48cc6653-15b3-3f2c-9cfe-d6701099a8dc%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/38f4b774-4c64-2fda-65c9-0dd05a95ac8b%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/e43c4870-12bb-ebff-5579-21134de28c36%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/d83c19fc-2d3e-a8ee-e9ce-7e6e758cdd03%28Office.15%29.aspx)|
|[Object](http://msdn.microsoft.com/library/504f695b-c518-8004-433f-627e80d15f89%28Office.15%29.aspx)|
|[ObjectPalette](http://msdn.microsoft.com/library/6f26ca1f-d851-4914-6dfa-c419b4ceac12%28Office.15%29.aspx)|
|[ObjectVerbs](http://msdn.microsoft.com/library/892ace19-928e-aa58-4a71-6f38c64727ff%28Office.15%29.aspx)|
|[ObjectVerbsCount](http://msdn.microsoft.com/library/518eff16-aef0-9e3e-2e03-af036117a152%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/7da1a1d6-bf23-5ea8-5e73-46ff92b67952%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/89d4855e-9c7e-7c3f-4063-f9f74d7245ca%28Office.15%29.aspx)|
|[OLEType](http://msdn.microsoft.com/library/9ce7cb88-e13e-4cda-bfe7-096734b796a0%28Office.15%29.aspx)|
|[OLETypeAllowed](http://msdn.microsoft.com/library/6c5ec029-043e-9828-e451-cd3507850953%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/1afb4220-a3de-076b-5619-d758b4e8483c%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/554db576-5976-6f05-0cb4-fdc6a38fd09c%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/8374c513-ede2-4ed7-2e35-55755cfd3942%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/aec13583-19c6-b5a6-2bc1-0a46e23e9459%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/67b47b88-8a45-c1e6-68b2-fe2cf2e726fe%28Office.15%29.aspx)|
|[OnKeyDown](http://msdn.microsoft.com/library/4602eec0-96ae-1592-d8b8-d4a44d7e8312%28Office.15%29.aspx)|
|[OnKeyPress](http://msdn.microsoft.com/library/fd4c6208-d311-64dd-8683-d106d33cffc0%28Office.15%29.aspx)|
|[OnKeyUp](http://msdn.microsoft.com/library/527a0034-31e1-af3f-d518-3c3b7cb62c8b%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/78ee2d7f-89d4-e9d2-a0ce-ecd6d35a98c3%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/cf1eac07-1e0f-ad7b-05c4-405867b1be71%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/8d61c653-519b-dc0a-1025-0d4bd440930a%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/333e2527-f190-e8b1-0f3b-789f4e37bff6%28Office.15%29.aspx)|
|[OnUpdated](http://msdn.microsoft.com/library/1af7adce-8d59-d8ac-cd3a-102266e55618%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/01e45f04-c4f4-0348-04ca-7eb714bf99e5%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/da3ab868-434c-f06c-04c9-9f8b4c183980%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/766c6e34-7996-f592-6fae-cb26aa2e4b40%28Office.15%29.aspx)|
|[Scaling](http://msdn.microsoft.com/library/290104f8-663b-7865-9ac9-6dc6feb5b92f%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/ad2407c1-28dc-5055-383d-8fe35d751c60%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/05f24e86-b02b-c55a-de10-0a6896ffefe0%28Office.15%29.aspx)|
|[SizeMode](http://msdn.microsoft.com/library/2c44b16f-cb04-8e45-2a67-7424342f48de%28Office.15%29.aspx)|
|[SourceDoc](http://msdn.microsoft.com/library/5b0e6b68-6528-5a35-e31d-b93d119897cc%28Office.15%29.aspx)|
|[SourceItem](http://msdn.microsoft.com/library/ab802b9b-d17c-695b-aaf5-4f84d1935615%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/29bcf6e1-880a-9e32-840f-75a54bed18ab%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/dc8ec458-8013-f6ff-5763-d083babcb4c9%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/f312def1-7abe-67e8-7970-60f09f10853a%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/9bcec2a4-c1b1-88db-e7b4-15e744c1e340%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/390cbfb5-5b05-2298-6b23-67ca7f9e2bba%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/83edbe69-58c5-ced5-9a51-d6e38f47aec5%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/cb80b5d1-a9a5-00a7-f439-3f6e7be6439b%28Office.15%29.aspx)|
|[UpdateOptions](http://msdn.microsoft.com/library/919ad3b4-1128-947a-09c0-7c7b0373698e%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/edafe10b-c207-527f-55a0-f71066fd9a85%28Office.15%29.aspx)|
|[VarOleObject](http://msdn.microsoft.com/library/3e1a6a95-d238-45ba-172d-1a1b22fb37be%28Office.15%29.aspx)|
|[Verb](http://msdn.microsoft.com/library/edbca2b1-fe7a-f0d0-1baf-fedbccb6dfb7%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/cea08737-227c-e0f6-cc8e-5e4b9129ad03%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/9fed4568-083a-8c38-4d44-b4085c2c8613%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/6ec65f4a-a02a-4434-65f6-8302cfc10b89%28Office.15%29.aspx)|

## See also


#### Other resources


[BoundObjectFrame Object Members](http://msdn.microsoft.com/library/e2bbeb0c-1b13-5953-999a-4a0b93cb3ec7%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
