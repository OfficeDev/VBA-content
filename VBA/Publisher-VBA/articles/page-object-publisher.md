---
title: Page Object (Publisher)
keywords: vbapb10.chm458751
f1_keywords:
- vbapb10.chm458751
ms.prod: publisher
api_name:
- Publisher.Page
ms.assetid: 9b2e8f29-26c3-1008-0ffd-eea2147abca4
ms.date: 06/08/2017
---


# Page Object (Publisher)

Represents a page in a publication. The  **[Pages](http://msdn.microsoft.com/library/d6b7262c-015c-dcf3-bff4-0091dd32b78f%28Office.15%29.aspx)** collection contains all the **Page** objects in a publication.


## Example

Use  **Pages** (index) to return a single **Page** object. The following example adds new text to the first shape on the first page in the active publication.


```
Sub AddPageNumberField() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 .InsertAfter " This text is added after the existing text." 
 .Font.Size = 15 
 End With 
End Sub
```

Use the  **[FindBypageID](http://msdn.microsoft.com/library/23ff5e69-33b1-e394-9d09-7199eae19fe9%28Office.15%29.aspx)** property to locate a **Page** object using the application assigned page ID. Use the **[Add](http://msdn.microsoft.com/library/3c22aa15-c1dc-94c8-62d6-a1bc9635cd89%28Office.15%29.aspx)** method to create a new page and add it to the publication. The following example adds a new page to the active publication and then looks for that page using the page ID.




```
Sub FindPage() 
 Dim lngPageID As Long 
 
 'Get page ID 
 lngPageID = ActiveDocument.Pages.Add(Count:=1, After:=1).PageID 
 
 'Use page ID to add a new shape to the page 
 ActiveDocument.Pages.FindByPageID(PageID:=lngPageID) _ 
 .Shapes.AddShape Type:=msoShape5pointStar, _ 
 Left:=200, Top:=72, Width:=50, Height:=50 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/7a7d9a67-8856-6549-7846-97b21eaf0bd2%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/9ef9d493-d2ca-8cac-3cce-6f0878acb288%28Office.15%29.aspx)|
|[ExportEmailHTML](http://msdn.microsoft.com/library/6257e9b5-26b5-73ae-7d40-50dd0a764488%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/754cfe41-0853-a2cf-59ee-85db68fb871a%28Office.15%29.aspx)|
|[SaveAsPicture](http://msdn.microsoft.com/library/9b118126-e072-9516-9863-14ea60264f01%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/c5d05664-e1ea-7936-4d3d-3d813ff4ec45%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/1bba32dc-0e7e-40ca-0f29-b67be6be518d%28Office.15%29.aspx)|
|[Footer](http://msdn.microsoft.com/library/8ab5a59b-c8d5-6217-098c-c53336ee5311%28Office.15%29.aspx)|
|[Header](http://msdn.microsoft.com/library/f10806eb-972a-d482-935c-95d5ccbbbb36%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/7ab931d7-c4aa-4687-44f8-2d03a389cd4f%28Office.15%29.aspx)|
|[IgnoreMaster](http://msdn.microsoft.com/library/53cd7b4b-4164-c6d3-766f-885a056d9b2b%28Office.15%29.aspx)|
|[IsLeading](http://msdn.microsoft.com/library/5a65f1fe-442d-f352-bea6-b732771008d8%28Office.15%29.aspx)|
|[IsTrailing](http://msdn.microsoft.com/library/e0ed15dc-d2e8-d6b7-913d-4e72b2817e88%28Office.15%29.aspx)|
|[IsTwoPageMaster](http://msdn.microsoft.com/library/dbfc3c21-0070-3f0a-c0b0-746d83c46765%28Office.15%29.aspx)|
|[IsWizardPage](http://msdn.microsoft.com/library/09c1352d-6760-ad54-aa95-211727c968b3%28Office.15%29.aspx)|
|[LayoutGuides](http://msdn.microsoft.com/library/eb9ac463-2b9f-9c68-b58f-6d93fe4993c8%28Office.15%29.aspx)|
|[Master](http://msdn.microsoft.com/library/f206b4f1-cde3-458d-f26c-a970ad3bd21b%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/cd81994d-506a-69ca-c7f6-472705b2ccd3%28Office.15%29.aspx)|
|[PageID](http://msdn.microsoft.com/library/07a87780-fb97-93ff-6f7d-1f1b72d3cb6a%28Office.15%29.aspx)|
|[PageIndex](http://msdn.microsoft.com/library/f64cc275-0474-7b97-d840-22e1e576d6f5%28Office.15%29.aspx)|
|[PageNumber](http://msdn.microsoft.com/library/670e3f46-9cad-b85e-b627-3be8c7c4e577%28Office.15%29.aspx)|
|[PageType](http://msdn.microsoft.com/library/0bb34de5-ac3e-386c-3b9f-814a476c9695%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/908daa24-3b8b-6107-d6ce-6498e6964e8e%28Office.15%29.aspx)|
|[ReaderSpread](http://msdn.microsoft.com/library/32823d2d-4bcd-a5a6-1ad1-ca1035d4fdea%28Office.15%29.aspx)|
|[RulerGuides](http://msdn.microsoft.com/library/69605642-7722-0721-cb07-d33689eda9ab%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/4e48d4cf-d7b6-9099-ddee-46a79e7eb7bf%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/94a8be36-20c2-65bc-b1e2-41f24703b264%28Office.15%29.aspx)|
|[WebPageOptions](http://msdn.microsoft.com/library/c2e3ee01-5b49-e83c-a68b-a4d526da0215%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/cb86988c-4460-4adb-19ad-e336fa9d4316%28Office.15%29.aspx)|
|[Wizard](http://msdn.microsoft.com/library/05cf1482-bde5-9ea2-4099-69a56a2dc61a%28Office.15%29.aspx)|
|[XOffsetWithinReaderSpread](http://msdn.microsoft.com/library/42ae7545-78f5-c034-33b4-f8c8f6a0b935%28Office.15%29.aspx)|
|[YOffsetWithinReaderSpread](http://msdn.microsoft.com/library/765adae3-af5d-ae37-5b1c-284cce8891ca%28Office.15%29.aspx)|

