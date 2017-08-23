---
title: "Объект документы (издатель)"
keywords: vbapb10.chm8716287
f1_keywords: vbapb10.chm8716287
ms.prod: publisher
api_name: Publisher.Documents
ms.assetid: 855b1677-4072-1e17-c22c-6db08e0c7569
ms.date: 06/08/2017
ms.openlocfilehash: 63984128ab9ec18bb821e2a431fea76db6c16905
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documents-object-publisher"></a>Объект документы (издатель)

Представляет все открытые публикаций. Коллекция **документов** содержит все объекты **документов** , которые открыты в Microsoft Publisher.


## <a name="example"></a>Пример

Свойство **документов** используется для возврата коллекции **документов** . В следующем примере перечисляются все открытые публикации.


```
Dim objDocument As Document 
Dim strMsg As String 
For Each objDocument In Documents 
 strMsg = strMsg &amp; objDocument.Name &amp; vbCrLf 
Next objDocument 
MsgBox Prompt:=strMsg, Title:="Current Documents Open", Buttons:=vbOKOnly
```

Используйте метод **Add** для добавления нового документа в коллекцию. Новые и видимым экземпляр Publisher создается при вызове метода **Add** . Следующий пример добавляет новый документ в коллекцию **документов** .




```
Dim objDocument As Document 
Set objDocument = Documents.Add 
With objDocument 
 .LayoutGuides.Columns = 4 
 .LayoutGuides.Rows = 9 
 .ActiveView.Zoom = pbZoomWholePage 
End With
```

Свойство **элемента** (индекс), где индекс — это индекс или имя документа в **строку**, чтобы возвратить объект конкретного документа. Следующий пример отображает имя первого открытой публикации.




```
If Documents.Count >= 1 Then 
 MsgBox Documents.Item(1).Name 
End If 

```

В следующем примере проверяется имя каждого документа в коллекции **документов** . Если имя документа «sales.doc», объектной переменной objSalesDoc имеет значение этого документа в коллекции **документов** .




```
Dim objDocument As Document 
Dim objSalesDoc As Document 
For Each objDocument In Documents 
 If objDocument.Name = "sales.pub" Then 
 Set objSalesDoc = objDocument 
 End If 
Next objDocument
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](http://msdn.microsoft.com/library/1e3536c8-8fc0-8c95-3a4c-b16fe8a99098%28Office.15%29.aspx)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](http://msdn.microsoft.com/library/41a2db19-5d56-be9b-a183-707d5e9e7e25%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/fe637a72-a96c-abfb-fa17-421848db5396%28Office.15%29.aspx)|
|[Элемент](http://msdn.microsoft.com/library/61cf3002-26d4-a678-abcb-940e7c385287%28Office.15%29.aspx)|
|[Родительский раздел](http://msdn.microsoft.com/library/a0bca15f-39a0-f7f0-9b68-f6ba30414d50%28Office.15%29.aspx)|

