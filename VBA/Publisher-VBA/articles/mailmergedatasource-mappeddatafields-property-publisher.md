---
title: "Свойство MailMergeDataSource.MappedDataFields (издатель)"
keywords: vbapb10.chm6291475
f1_keywords: vbapb10.chm6291475
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.MappedDataFields
ms.assetid: 9f2a15a7-41b0-6025-73d6-eb70a412b830
ms.date: 06/08/2017
ms.openlocfilehash: 1ed2ed7a6336c8ab070977785108c6e9c3777aba
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourcemappeddatafields-property-publisher"></a>Свойство MailMergeDataSource.MappedDataFields (издатель)

Возвращает объект **[MailMergeMappedDataFields](mailmergemappeddatafields-object-publisher.md)** , представляющий поля сопоставленные данные, доступные в Microsoft Publisher.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MappedDataFields**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

MailMergeMappedDataFields


## <a name="example"></a>Пример

В этом примере создается таблица на новой странице текущей публикации и перечисляет поля сопоставленные данные, доступные в Publisher и поля в источнике данных, к которой они сопоставлены. В этом примере предполагается, что текущей публикации является публикацией слияния почты и у соответствующих полей источника данных сопоставленные поля данных.


```vb
Sub MappedFields() 
 Dim intCount As Integer 
 Dim intRows As Integer 
 Dim docPub As Document 
 Dim pagNew As Page 
 Dim shpTable As Shape 
 Dim tblTable As Table 
 Dim rowTable As Row 
 
 On Error Resume Next 
 
 Set docPub = ThisDocument 
 Set pagNew = ThisDocument.Pages.Add(Count:=1, After:=1) 
 intRows = docPub.MailMerge.DataSource.MappedDataFields.Count + 1 
 
 'Creates new table with a heading row 
 Set shpTable = pagNew.Shapes.AddTable(NumRows:=intRows, _ 
 numColumns:=2, Left:=100, Top:=100, Width:=400, Height:=12) 
 Set tblTable = shpTable.Table 
 With tblTable.Rows(1) 
 With .Cells(1).Text 
 .Text = "Mapped Data Field" 
 .Font.Bold = msoTrue 
 End With 
 With .Cells(2).Text 
 .Text = "Data Source Field" 
 .Font.Bold = msoTrue 
 End With 
 End With 
 
 With docPub.MailMerge.DataSource 
 For intCount = 2 To intRows - 1 
 'Inserts mapped data field name and the 
 'corresponding data source field name 
 tblTable.Rows(intCount - 1).Cells(1).Text _ 
 .Text = .MappedDataFields(Index:=intCount).Name 
 tblTable.Rows(intCount - 1).Cells(2).Text _ 
 .Text = .MappedDataFields(Index:=intCount).DataFieldName 
 Next 
 End With 
End Sub
```


