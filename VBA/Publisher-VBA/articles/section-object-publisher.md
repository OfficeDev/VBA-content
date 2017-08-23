---
title: "Объект раздела (издатель)"
keywords: vbapb10.chm7471103
f1_keywords: vbapb10.chm7471103
ms.prod: publisher
api_name: Publisher.Section
ms.assetid: 7e92a8de-ed66-564b-2657-cef0fc2392b8
ms.date: 06/08/2017
ms.openlocfilehash: 8dab6b4feb9ccc1cb271eb366bd750d913d56166
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="section-object-publisher"></a>Объект раздела (издатель)

Представляет раздел публикации или документа.
 


## <a name="example"></a>Пример

Используйте **разделы**. Item(Index), где номер индекса, чтобы возвратить объект одного **раздела** индекса. В следующем примере задается объект **раздела** в первый раздел в коллекции **разделах** активных документов.
 

 

```
Dim objSection As Section 
Set objSection = ActiveDocument.Sections.Item(1)
```

Используйте **разделы**. Add(StartPageIndex), где StartPageIndex — номер индекса страницы, чтобы получить новый раздел, добавлены в документ. Если страница уже содержит раздел head, будут возвращены ошибку «Отказано в разрешении.». Следующий пример добавляет объект раздела на вторую страницу активных документов.
 

 



```
Dim objSection As Section 
Set objSection = ActiveDocument.Sections.Add(StartPageIndex:=2)
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Delete](section-delete-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](section-application-property-publisher.md)|
|[ContinueNumbersFromPreviousSection](section-continuenumbersfromprevioussection-property-publisher.md)|
|[PageNumberFormat](section-pagenumberformat-property-publisher.md)|
|[PageNumberStart](section-pagenumberstart-property-publisher.md)|
|[Родительский раздел](section-parent-property-publisher.md)|
|[ShowHeaderFooterOnFirstPage](section-showheaderfooteronfirstpage-property-publisher.md)|
|[StartPageIndex](section-startpageindex-property-publisher.md)|

