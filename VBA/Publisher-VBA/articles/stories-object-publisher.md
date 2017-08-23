---
title: "Объект материалы (издатель)"
keywords: vbapb10.chm5767167
f1_keywords: vbapb10.chm5767167
ms.prod: publisher
api_name: Publisher.Stories
ms.assetid: 694a0376-fa41-3097-180b-40b8a005ddf6
ms.date: 06/08/2017
ms.openlocfilehash: c35b0d46a1669c220d8d5c28547d6b4e251a5d94
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="stories-object-publisher"></a>Объект материалы (издатель)

Представляет весь текст в публикации.
 


## <a name="example"></a>Пример

Свойство **функциональности** объекта **Document** для возврата коллекции **функциональности** . Метод **Item** коллекции **истории (en)** для доступа к отдельным объектам **сценариев** .
 

 

 

 
Коллекция **функциональности** позволяет эффективного доступа к тексту в публикации. Простой цикл по коллекции **функциональности** для проверки всех текста в текстовых рамок или таблиц без необходимости выполнять поиск каждой фигуры на всех страницах публикации.
 

 

 

 
Коллекция **функциональности** содержит один объект **статьи** для каждого несвязанные надпись, каждый цепочки связанных текстовых рамок и каждой таблицы в публикации. Текст в кадрах WordArt, объекты OLE и рисунков не включаются в коллекцию **функциональности** .
 

 

 

 
В этом примере присваивает первая статья в активной публикации объектную переменную.
 

 



```
Dim stFirst As Story 
 
stFirst = Application.ActiveDocument.Stories(1)
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Элемент](stories-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](stories-application-property-publisher.md)|
|[Count](stories-count-property-publisher.md)|
|[Родительский раздел](stories-parent-property-publisher.md)|

