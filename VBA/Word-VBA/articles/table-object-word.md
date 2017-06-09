---
title: Table Object (Word)
keywords: vbawd10.chm2385
f1_keywords:
- vbawd10.chm2385
ms.prod: word
api_name:
- Word.Table
ms.assetid: 996b58dd-ebc6-ee30-5bfe-c5e51a0f71d6
ms.date: 06/08/2017
---


# Table Object (Word)

Represents a single table. The  **Table** object is a member of the **[Tables](http://msdn.microsoft.com/library/068a3d0f-0b19-3927-cb0a-7fb0d0fd8e52%28Office.15%29.aspx)** collection. The **Tables** collection includes all the tables in the specified selection, range, or document.


## Remarks

Use  **Tables** (Index), where Index is the index number, to return a single **Table** object. The index number represents the position of the table in the selection, range, or document. The following example converts the first table in the active document to text.


```
ActiveDocument.Tables(1).ConvertToText Separator:=wdSeparateByTabs
```

Use the  **Add** method to add a table at the specified range. The following example adds a 3x4 table at the beginning of the active document.




```
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.Tables.Add Range:=myRange, NumRows:=3, NumColumns:=4
```


## Methods



|**Name**|
|:-----|
|[ApplyStyleDirectFormatting](http://msdn.microsoft.com/library/239807ae-6389-4492-8d17-e450c6ba91dd%28Office.15%29.aspx)|
|[AutoFitBehavior](http://msdn.microsoft.com/library/74e162a5-cde0-bdd3-2ea6-f78fb0ecca5a%28Office.15%29.aspx)|
|[AutoFormat](http://msdn.microsoft.com/library/c76452fa-e1e8-3787-726a-b1c9967d96c2%28Office.15%29.aspx)|
|[Cell](http://msdn.microsoft.com/library/7dd91771-c72b-eefb-2492-1998c0d194bb%28Office.15%29.aspx)|
|[ConvertToText](http://msdn.microsoft.com/library/750db54e-faca-f1eb-8eb8-3a5c0dbb2c25%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/157240bf-6abb-c4a6-ef39-609fd315121a%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/4150362d-ca09-deb7-34cf-b70702c55a43%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/2c68f7ad-2d57-05ea-bd8b-cb8712c21f02%28Office.15%29.aspx)|
|[SortAscending](http://msdn.microsoft.com/library/5a73ac7a-917d-7559-99c1-cb20f39b864d%28Office.15%29.aspx)|
|[SortDescending](http://msdn.microsoft.com/library/a72b25e9-06c2-8f2f-1dff-796768d43fff%28Office.15%29.aspx)|
|[Split](http://msdn.microsoft.com/library/a96c6dff-8508-2a73-2f3a-fac755e026ff%28Office.15%29.aspx)|
|[UpdateAutoFormat](http://msdn.microsoft.com/library/d33f3b59-f05c-d51e-5f43-17d56af6693f%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AllowAutoFit](http://msdn.microsoft.com/library/e8894734-68b3-60bb-7623-9497e4e99e10%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/d97d2afc-fdc0-aad8-584d-ad960e1e41bd%28Office.15%29.aspx)|
|[ApplyStyleColumnBands](http://msdn.microsoft.com/library/da3a77b6-ae71-9552-b04c-06b8812c1dcd%28Office.15%29.aspx)|
|[ApplyStyleFirstColumn](http://msdn.microsoft.com/library/9802ff74-321d-a44c-2cac-9f17b91210d2%28Office.15%29.aspx)|
|[ApplyStyleHeadingRows](http://msdn.microsoft.com/library/1c7fb6d5-9010-fded-d882-388d1e631da2%28Office.15%29.aspx)|
|[ApplyStyleLastColumn](http://msdn.microsoft.com/library/db47720e-0351-c48d-6ebe-a149f2b8c84f%28Office.15%29.aspx)|
|[ApplyStyleLastRow](http://msdn.microsoft.com/library/007ac0c4-bec8-9c48-99e2-017567415193%28Office.15%29.aspx)|
|[ApplyStyleRowBands](http://msdn.microsoft.com/library/2957cc86-2248-ac7d-f4ae-16294c518b90%28Office.15%29.aspx)|
|[AutoFormatType](http://msdn.microsoft.com/library/366dbfab-f40e-b570-d174-96f4fe07a063%28Office.15%29.aspx)|
|[Borders](http://msdn.microsoft.com/library/904bce6b-db91-32be-f65d-7200f9a63be8%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/d4e37a85-d194-8d19-c43f-09d30187e007%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/6f4c70ef-032d-7f05-1b21-c5c86af804bd%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/0f6c6ea5-ba19-8c47-edca-db3517149f82%28Office.15%29.aspx)|
|[Descr](http://msdn.microsoft.com/library/745b446c-1371-35d5-d6bd-8ad6aa4867fe%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/f14f821b-43d6-9855-e0ab-c6420ff211c5%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/ad047ad0-7a50-6905-9e60-3a2275e49a62%28Office.15%29.aspx)|
|[NestingLevel](http://msdn.microsoft.com/library/419522f9-f102-88ef-5bf8-29f4896de5ae%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/a4ca3483-3121-0169-6251-07d23faa118a%28Office.15%29.aspx)|
|[PreferredWidth](http://msdn.microsoft.com/library/15c3d169-9c61-fb70-3cc6-15f385bab8c0%28Office.15%29.aspx)|
|[PreferredWidthType](http://msdn.microsoft.com/library/92954057-5ecd-3d43-c547-e1e1a6c83904%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/6352ee1a-7047-5efe-91ec-faa90eedcd0c%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/a41681da-9a11-9b45-fcff-495208a3ab25%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/e4cc7541-15fe-97b6-0fe6-90d561a85420%28Office.15%29.aspx)|
|[Shading](http://msdn.microsoft.com/library/0c5c0ebe-d7cb-ff55-c77c-2c0c36a6c98a%28Office.15%29.aspx)|
|[Spacing](http://msdn.microsoft.com/library/56444e6f-70b6-c815-9098-e6e3ac2d6c3b%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/5b375f41-99da-314e-f8c3-d440c6153419%28Office.15%29.aspx)|
|[TableDirection](http://msdn.microsoft.com/library/3062731b-a334-927d-3871-f845cfb662ac%28Office.15%29.aspx)|
|[Tables](http://msdn.microsoft.com/library/aba332ae-49aa-4575-8f33-66ca0c647d26%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/a7b8437a-3882-1301-4235-7491156aca3a%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/005453cf-019e-c404-3114-c555cf5a1310%28Office.15%29.aspx)|
|[Uniform](http://msdn.microsoft.com/library/a156bedf-5426-be4c-b961-84a038f9bfd6%28Office.15%29.aspx)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

