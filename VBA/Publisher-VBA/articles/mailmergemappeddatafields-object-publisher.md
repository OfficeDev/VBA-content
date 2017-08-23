---
title: "Объект MailMergeMappedDataFields (издатель)"
keywords: vbapb10.chm6553599
f1_keywords: vbapb10.chm6553599
ms.prod: publisher
api_name: Publisher.MailMergeMappedDataFields
ms.assetid: 7f33bf07-9cbb-e171-d276-d5ccb06abb95
ms.date: 06/08/2017
ms.openlocfilehash: dc99540651a24973dea2a07bb680e6433230799e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergemappeddatafields-object-publisher"></a>Объект MailMergeMappedDataFields (издатель)

Коллекция объектов **[MailMergeMappedDataField](mailmergemappeddatafield-object-publisher.md)** , который представляет поля сопоставленные данные, доступные в Microsoft Publisher.
 


## <a name="example"></a>Пример

Используйте свойство **[MappedDataFields](mailmergedatasource-mappeddatafields-property-publisher.md)** объекта **[вывода](mailmergedatasource-object-publisher.md)** для возврата коллекции **MailMergeMappedDataFields** . В этом примере создается таблица на новой странице текущей публикации и перечисляет поля сопоставленные данные, доступные в Publisher и поля в источнике данных, к которой они сопоставлены. В этом примере предполагается, что текущей публикации является публикацией слияния почты и у соответствующих полей источника данных сопоставленные поля данных.
 

 

```
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


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Элемент](mailmergemappeddatafields-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](mailmergemappeddatafields-application-property-publisher.md)|
|[Count](mailmergemappeddatafields-count-property-publisher.md)|
|[Родительский раздел](mailmergemappeddatafields-parent-property-publisher.md)|

