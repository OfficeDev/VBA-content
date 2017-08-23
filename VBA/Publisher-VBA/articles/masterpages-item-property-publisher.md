---
title: "Свойство MasterPages.Item (издатель)"
keywords: vbapb10.chm589824
f1_keywords: vbapb10.chm589824
ms.prod: publisher
api_name: Publisher.MasterPages.Item
ms.assetid: f0a4e9b2-cd01-01c3-b1d3-a241ea08c5d3
ms.date: 06/08/2017
ms.openlocfilehash: 3878115805c60957e07a3433be95b61c5a86c61b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="masterpagesitem-property-publisher"></a>Свойство MasterPages.Item (издатель)

Возвращает указанный объект **[страницы](page-object-publisher.md)** коллекции **страниц** или **макетом** . Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Элемент_**)

 переменная _expression_A, представляет собой объект- **макетом** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Item|Обязательное свойство.| **Длинный**|Номер страницы для возврата. Для семейств сайтов **макетом** элемента может быть 1 или 2 влево и вправо главных страниц, соответственно. Для коллекции **страниц** элементов соответствует свойству **[PageIndex](page-pageindex-property-publisher.md)** объект **Page** .|

## <a name="example"></a>Пример

В этом примере отображается номер страницы, страницы индекса и идентификатор страницы первой страницы в активной публикации.


```vb
With ActiveDocument.Pages.Item(1) 
 Debug.Print "Page number = " &; .PageNumber 
 Debug.Print "Page index = " &; .PageIndex 
 Debug.Print "Page ID = " &; .PageID 
End With
```


