---
title: "Свойство Hyperlinks.Item (издатель)"
keywords: vbapb10.chm6881280
f1_keywords: vbapb10.chm6881280
ms.prod: publisher
api_name: Publisher.Hyperlinks.Item
ms.assetid: 8d288fc6-9ded-5732-b972-6fa366ef31c3
ms.date: 06/08/2017
ms.openlocfilehash: 16976a1028d943e3027a509639f43e29804a6c3e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="hyperlinksitem-property-publisher"></a>Свойство Hyperlinks.Item (издатель)

Возвращает объект отдельных из указанного семейства сайтов. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляющий объект **гиперссылки** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Длинный**|Количество для возвращаемого объекта.|

## <a name="example"></a>Пример

В этом примере отображается адрес гиперссылки в фигуры, один из активных публикации.


```vb
MsgBox "Address of first hyperlink: " _ 
 &; ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Hyperlinks.Item(1).Address
```


