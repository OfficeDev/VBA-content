---
title: ObjectFrame Object (Access)
keywords: vbaac10.chm11665
f1_keywords:
- vbaac10.chm11665
ms.prod: access
api_name:
- Access.ObjectFrame
ms.assetid: 0eb85477-58d7-249a-2bf7-f2f3960a45a9
ms.date: 06/08/2017
---


# ObjectFrame Object (Access)

This object corresponds to an unbound object frame. The unbound object frame control displays a picture, chart, or any OLE object not stored in a table.


## Remarks

For example, you can use an unbound object frame to display a chart that you created and stored in Microsoft Graph.

This control allows you to create or edit the object from within a Microsoft Access form or report by using the application in which the object was originally created.

To display objects that are stored in a Microsoft Access database, use a [Bound object frame](http://msdn.microsoft.com/library/9d087a78-278d-1b87-d1b4-22f836707efa%28Office.15%29.aspx).

The object in an unbound object frame is the same for every record.

The unbound object frame can display linked or embedded objects.

You can use the unbound object frame or an [image control](http://msdn.microsoft.com/library/1f938a6e-7aea-7787-d959-e21edaa9342c%28Office.15%29.aspx) to display unbound pictures in a form or report. The advantage of using the unbound object frame is that you can edit the object directly from the form or report. The advantage of using the image control is that it's faster to display.


## Events



|**Name**|
|:-----|
|[Click](http://msdn.microsoft.com/library/78a80855-693e-6e6a-59c9-963802dd3b5d%28Office.15%29.aspx)|
|[DblClick](http://msdn.microsoft.com/library/83a69067-7505-f126-0fa6-12f8d06d7144%28Office.15%29.aspx)|
|[Enter](http://msdn.microsoft.com/library/85b48c3c-3b0c-dbe3-71ca-7ac144477bfc%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/9abc0214-a73c-7709-aaeb-817716694dd7%28Office.15%29.aspx)|
|[GotFocus](http://msdn.microsoft.com/library/413efc78-c011-2dd6-4c5c-7b462fa9ede2%28Office.15%29.aspx)|
|[LostFocus](http://msdn.microsoft.com/library/d503815f-1511-82d6-b940-ceba6267f571%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/54fb7a4f-8428-429d-e560-3c4b64c0f683%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/aad412fd-96ba-803d-848c-a8788fa1e0ae%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/0f818329-7817-d62f-2ccd-a35232cf67dc%28Office.15%29.aspx)|
|[Updated](http://msdn.microsoft.com/library/827a14f5-4062-e904-3f53-ccb01b59b03f%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Move](http://msdn.microsoft.com/library/63b05ea6-d761-adfa-5aa6-25d16ae5ed3c%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/470a8412-a3e4-6f06-063c-84c848d73834%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/856855b5-6b61-6aea-c039-696d4662ee4c%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/f1511d4f-367e-85e4-cc5c-cdb0f8a72d8b%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Action](http://msdn.microsoft.com/library/042d3418-fe67-c4cc-60b1-dc3b373b8d4f%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/29e3d68f-4f67-793d-6976-e18e290145fe%28Office.15%29.aspx)|
|[AutoActivate](http://msdn.microsoft.com/library/e6e0dfce-1bfe-707b-d7f0-45a216d4aa55%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/c73bd932-ebfe-8b3b-5dc2-0c88a6210c94%28Office.15%29.aspx)|
|[BackShade](http://msdn.microsoft.com/library/68800e85-9dfa-958d-e87d-1241be551f90%28Office.15%29.aspx)|
|[BackStyle](http://msdn.microsoft.com/library/4d8a384b-e796-30b2-4ce1-ce172e58b431%28Office.15%29.aspx)|
|[BackThemeColorIndex](http://msdn.microsoft.com/library/9c4cbfee-2026-2caa-922d-d7345cc026f5%28Office.15%29.aspx)|
|[BackTint](http://msdn.microsoft.com/library/80c3d5f6-7240-9001-f035-0d464e8c49f2%28Office.15%29.aspx)|
|[BorderColor](http://msdn.microsoft.com/library/840f6108-e75f-9807-799a-9fc23b8a96ec%28Office.15%29.aspx)|
|[BorderShade](http://msdn.microsoft.com/library/de92000e-95bb-12df-68ef-5ada76553e97%28Office.15%29.aspx)|
|[BorderStyle](http://msdn.microsoft.com/library/243484f6-1401-cbe9-dfb9-d5c8f7e419ce%28Office.15%29.aspx)|
|[BorderThemeColorIndex](http://msdn.microsoft.com/library/8070f9c7-bee5-a702-f874-c96af9fb71d3%28Office.15%29.aspx)|
|[BorderTint](http://msdn.microsoft.com/library/f1adfec8-7106-bf3c-db7d-ea12c9a82d7d%28Office.15%29.aspx)|
|[BorderWidth](http://msdn.microsoft.com/library/2f95633a-dea1-08a6-3c0e-1fb52f453c06%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/e0e5586d-ae5c-9d22-9876-ad717a9805d1%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/38ee5131-ffcb-3db6-0f2d-1e7f59c9a5b4%28Office.15%29.aspx)|
|[ColumnCount](http://msdn.microsoft.com/library/be9b3121-e9ea-eb78-5165-0a9d5f209b32%28Office.15%29.aspx)|
|[ColumnHeads](http://msdn.microsoft.com/library/f318f924-2629-8a7a-90b0-3ab386e50a22%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/42884347-14f3-0f0f-dc7e-3d2ae8154a49%28Office.15%29.aspx)|
|[ControlTipText](http://msdn.microsoft.com/library/6b2bf5d6-fa3d-149c-1fb7-178c8bf1cd9b%28Office.15%29.aspx)|
|[ControlType](http://msdn.microsoft.com/library/3afa6ed8-db2d-6116-85ce-f1b67990fc1f%28Office.15%29.aspx)|
|[DisplayType](http://msdn.microsoft.com/library/30df2df5-ed46-f0e4-02e3-43c3aa99dbad%28Office.15%29.aspx)|
|[DisplayWhen](http://msdn.microsoft.com/library/37e03fc6-aee9-b6cf-eafb-7af111b5b9e3%28Office.15%29.aspx)|
|[Enabled](http://msdn.microsoft.com/library/1b70cc10-3132-f8e4-5a82-19396551f1a7%28Office.15%29.aspx)|
|[EventProcPrefix](http://msdn.microsoft.com/library/a38ca887-8d70-eb89-a1ac-fd7308d17c0d%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/746790f7-cac5-5631-ae25-04b95b0c405a%28Office.15%29.aspx)|
|[GridlineShade](http://msdn.microsoft.com/library/837f3c0b-5597-7abd-e580-c92f099d4448%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/692feea8-ab41-e695-c388-38b9f7f9bf26%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/09791118-77ec-c03c-00e9-d6450d1c7fe2%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/98fe7dba-d488-3a19-7640-bab09b1aca7e%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/cf902f29-bd15-9abe-cfdb-d34fc059cf0b%28Office.15%29.aspx)|
|[GridlineThemeColorIndex](http://msdn.microsoft.com/library/47440f76-07fa-8924-4a1d-10fb005e8e5b%28Office.15%29.aspx)|
|[GridlineTint](http://msdn.microsoft.com/library/f32466e0-0924-97c7-2454-7632730ffcfa%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/f5d014d2-11ad-f404-b3bc-bafbac93c8e4%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/86e7166f-ca94-83de-06fd-5182113fbbe7%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/4838b854-1679-18ff-689e-68bf6043a49a%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/72f30e89-326e-ecf3-cf48-eb0a4e56f60d%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/53d6085e-e01e-5260-0802-3958f62e378a%28Office.15%29.aspx)|
|[HelpContextId](http://msdn.microsoft.com/library/85fa8d67-b0ff-129d-b689-ceca69e8b487%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/8476d254-0e20-8c36-961a-3732b59a5b99%28Office.15%29.aspx)|
|[InSelection](http://msdn.microsoft.com/library/5ae30220-4d7a-1838-1edc-99b54689b6ab%28Office.15%29.aspx)|
|[IsVisible](http://msdn.microsoft.com/library/0fdbf0ab-518b-6c1a-5394-a6ecad4f70f5%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/18548487-558b-7c37-c17b-00496e29b2cf%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/bd666a9f-f4b6-9b33-a6e1-d6a8570133de%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/591a05e2-d014-8e0d-036b-166d8366284e%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/30af95f2-7643-d3a1-2d6f-69ea98a227f9%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/b146f062-bf23-32d6-335a-2583b6171006%28Office.15%29.aspx)|
|[LinkChildFields](http://msdn.microsoft.com/library/f82332c1-2dd0-bd3a-3f63-e84727ea7429%28Office.15%29.aspx)|
|[LinkMasterFields](http://msdn.microsoft.com/library/1e3b8cb7-a061-369a-4ff4-44d6989c3234%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/0769b9c9-ea0b-33c8-b258-e7d775bee9e6%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/d903a75a-579e-7896-a2e2-3d1688fe5145%28Office.15%29.aspx)|
|[Object](http://msdn.microsoft.com/library/db4354f5-a92d-7341-7823-a1c5f26d74b8%28Office.15%29.aspx)|
|[ObjectPalette](http://msdn.microsoft.com/library/12d507b8-ac47-3e00-434f-4a3cab7071d3%28Office.15%29.aspx)|
|[ObjectVerbs](http://msdn.microsoft.com/library/e0e2c596-7276-3626-1ce4-ec5502bec02c%28Office.15%29.aspx)|
|[ObjectVerbsCount](http://msdn.microsoft.com/library/8c7a6302-cdf0-5997-7b71-65cfb6f0a7d3%28Office.15%29.aspx)|
|[OldBorderStyle](http://msdn.microsoft.com/library/e3676f02-337b-d347-478d-9ae8fa03c343%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/ef025309-83e8-36e4-956e-62a88d8a0e21%28Office.15%29.aspx)|
|[OLEClass](http://msdn.microsoft.com/library/ed32f15c-77da-0bd6-46da-38373ea37cc1%28Office.15%29.aspx)|
|[OLEType](http://msdn.microsoft.com/library/eb9a08ba-8fc6-247d-14c3-0791a0461f0c%28Office.15%29.aspx)|
|[OLETypeAllowed](http://msdn.microsoft.com/library/ca669834-9bce-057c-dfb7-c8411b26bdd1%28Office.15%29.aspx)|
|[OnClick](http://msdn.microsoft.com/library/521e9685-317b-aafc-3ef2-bfd0d04dd3d0%28Office.15%29.aspx)|
|[OnDblClick](http://msdn.microsoft.com/library/356cbeb6-b0e2-d5a5-434a-507a760b8631%28Office.15%29.aspx)|
|[OnEnter](http://msdn.microsoft.com/library/386524c3-8208-05dd-5d0f-9899e4619eb7%28Office.15%29.aspx)|
|[OnExit](http://msdn.microsoft.com/library/47cb4cb3-1d8a-d286-a7df-832d6aa3fb55%28Office.15%29.aspx)|
|[OnGotFocus](http://msdn.microsoft.com/library/3897a919-6180-6b57-eba9-72eea8831753%28Office.15%29.aspx)|
|[OnLostFocus](http://msdn.microsoft.com/library/47f336d6-2a89-4824-55c3-c632d2fbf2f2%28Office.15%29.aspx)|
|[OnMouseDown](http://msdn.microsoft.com/library/de03fb25-bf9b-4365-3540-68505f58048c%28Office.15%29.aspx)|
|[OnMouseMove](http://msdn.microsoft.com/library/b57730b7-a8ae-d62f-0511-793633d04969%28Office.15%29.aspx)|
|[OnMouseUp](http://msdn.microsoft.com/library/0c77a908-07f5-7838-fa61-5ee0fc197aeb%28Office.15%29.aspx)|
|[OnUpdated](http://msdn.microsoft.com/library/d2239f45-959b-beb7-fe9e-c9a9a257dd4b%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/4beb8dcd-9345-5071-a86c-5ad2deb699db%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/d7a8ccf2-9df2-db72-424a-cc6fa98abe52%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/9882c250-bbe0-8abb-8c7e-00e1f8c6af4e%28Office.15%29.aspx)|
|[RowSource](http://msdn.microsoft.com/library/de2aa92d-34e8-20e7-ece7-5e1dcb8cd877%28Office.15%29.aspx)|
|[RowSourceType](http://msdn.microsoft.com/library/d450ce8b-c2e9-f51b-61af-b46a64ab7d32%28Office.15%29.aspx)|
|[Scaling](http://msdn.microsoft.com/library/ec0ccdc1-edcd-14d1-05ca-2c3b2e200440%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/be084566-3d7f-278e-5e78-b10720631cd8%28Office.15%29.aspx)|
|[ShortcutMenuBar](http://msdn.microsoft.com/library/06ed3789-f76d-0ccd-7580-74fbfc76a983%28Office.15%29.aspx)|
|[SizeMode](http://msdn.microsoft.com/library/2aaa2f95-7982-a585-1a9f-a6ed191be79e%28Office.15%29.aspx)|
|[SourceDoc](http://msdn.microsoft.com/library/23a45f7f-b4e2-fc93-6049-c9298e199202%28Office.15%29.aspx)|
|[SourceItem](http://msdn.microsoft.com/library/86cb94a8-9c13-0b07-58c2-1b78849061c9%28Office.15%29.aspx)|
|[SourceObject](http://msdn.microsoft.com/library/985c8b01-84d8-2da6-6cad-5de08d835434%28Office.15%29.aspx)|
|[SpecialEffect](http://msdn.microsoft.com/library/40117a03-0640-5b5c-363d-19f1f5b9f2d0%28Office.15%29.aspx)|
|[StatusBarText](http://msdn.microsoft.com/library/51daa6c0-8887-9843-c899-ebb99c722866%28Office.15%29.aspx)|
|[TabIndex](http://msdn.microsoft.com/library/3eae97f2-daa4-c9e9-2e4e-a17f153d5633%28Office.15%29.aspx)|
|[TabStop](http://msdn.microsoft.com/library/2930cfb8-22be-1d39-7514-fe864b2f9373%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/0b8fdb5f-dadc-fafb-cc9a-74dfe40f9b80%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/b9a6e2fc-7b71-580c-1483-69ff799c0f0f%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/cbc4038c-e9e8-7e57-5bb2-7cafe917c6b3%28Office.15%29.aspx)|
|[UpdateMethod](http://msdn.microsoft.com/library/3c29df53-33cd-d645-2c45-6ff49fe4068e%28Office.15%29.aspx)|
|[UpdateOptions](http://msdn.microsoft.com/library/29effba2-7427-62ca-c0d6-6ed5081b0e02%28Office.15%29.aspx)|
|[VarOleObject](http://msdn.microsoft.com/library/e04e769d-07fb-dacc-aa70-ddd3a064d785%28Office.15%29.aspx)|
|[Verb](http://msdn.microsoft.com/library/ed661d02-d00f-4911-7be7-3a0e973e6456%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/972f9c07-ef2e-5bf4-2562-e411e9ae05ce%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/2461bccb-44c6-82b4-93a0-9e4f8231cf53%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/7a42f8ef-6c69-1fa8-d326-95f1aab8880a%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
[ObjectFrame Object Members](http://msdn.microsoft.com/library/65229083-68ec-b870-50f4-a6c329259a39%28Office.15%29.aspx)
