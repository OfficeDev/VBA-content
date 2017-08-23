---
title: "Объект разделах (издатель)"
keywords: vbapb10.chm7405567
f1_keywords: vbapb10.chm7405567
ms.prod: publisher
api_name: Publisher.Sections
ms.assetid: 429c03b8-b574-86db-c39d-551a4c753b04
ms.date: 06/08/2017
ms.openlocfilehash: 1e7701f8e8a960d4aff8ac9b2edd26d32cb924cb
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="sections-object-publisher"></a>Объект разделах (издатель)

Коллекция всех объектов **раздел** в документе.
 


## <a name="example"></a>Пример

Используйте **разделы**. Item(Index), где номер индекса, чтобы возвратить объект одного **раздела** индекса. В следующем примере задается формат и начального номера для первого раздела активных документов.
 

 

```
With ActiveDocument.Sections.Item(1) 
 .PageNumberFormat = pbPageNumberFormatArabic 
 .PageNumberStart = 1 
End With
```

С помощью **разделов** (индекс) где индекса — номер индекса, также возвращает объект **раздела** . В следующем примере задается по-прежнему производится нумерации в предыдущем разделе, для второй раздел в активный документ.
 

 



```
ActiveDocument.Sections(2).ContinueNumbersFromPreviousSection=True
```

Используйте **разделы**. Count возвращает число разделов в публикации. Следующий пример отображения числа разделов в первом открытом документе.
 

 



```
MsgBox Documents(1).Sections.Count
```

Используйте **разделы**. Add(StartPageIndex), где StartPageIndex — номер индекса страницы, чтобы reutrn новый раздел, добавлены в документ. Если страница уже содержит раздел head, будут возвращены ошибку «Отказано в разрешении.». В следующем примере добавляется новый раздел на вторую страницу активных документов.
 

 



```
Dim objSection As Section 
Set objSection = ActiveDocument.Sections.Add(StartPageIndex:=2)
```

Используйте **разделы** (индекс). Удаление, где индекс — номер индекса для удаления указанного раздела из документа. При попытке удалить в первом разделе будут возвращены ошибку «Отказано в разрешении». Следующий пример удаляет все разделы активных документов, кроме первого.
 

 

 **Примечание**  Итерации — от последнего к первому во избежание «Индекс вне диапазона.» Ошибка при доступе к удаленный раздел в **разделах** коллекции.
 




```
Dim i As Long 
For i = ActiveDocument.Sections.Count To 1 Step -1 
 If i = 1 Then Exit For 
 ActiveDocument.Sections(i).Delete 
Next i
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](sections-add-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](sections-application-property-publisher.md)|
|[Count](sections-count-property-publisher.md)|
|[Элемент](sections-item-property-publisher.md)|
|[Родительский раздел](sections-parent-property-publisher.md)|

