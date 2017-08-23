---
title: "Свойство MailMergeMappedDataField.DataFieldName (издатель)"
keywords: vbapb10.chm6553603
f1_keywords: vbapb10.chm6553603
ms.prod: publisher
api_name: Publisher.MailMergeMappedDataField.DataFieldName
ms.assetid: c30e56c1-c4f4-a581-00d1-eb367178e0af
ms.date: 06/08/2017
ms.openlocfilehash: b3160c65042d499573664114333d33ed94e63391
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergemappeddatafielddatafieldname-property-publisher"></a>Свойство MailMergeMappedDataField.DataFieldName (издатель)

Возвращает или задает **строку** , которая представляет имя поля в источнике данных слияния почты, с которыми сопоставляется сопоставленное поле данных. Если поля данных не связан с сопоставленное поле данных, возвращается пустая строка. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DataFieldName**

 переменная _expression_A, представляет собой объект- **MailMergeMappedDataField** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере создается таблица на новой странице текущей публикации и перечисляет поля сопоставленные данные, доступные и поля в источнике данных, к которой они сопоставлены. В этом примере предполагается, что текущей публикации является публикацией слияния почты и у соответствующих полей источника данных сопоставленные поля данных.


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


