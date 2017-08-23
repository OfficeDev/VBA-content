---
title: "Свойство Plates.Item (издатель)"
keywords: vbapb10.chm2818048
f1_keywords: vbapb10.chm2818048
ms.prod: publisher
api_name: Publisher.Plates.Item
ms.assetid: 7563df76-56c3-d613-7314-846fe28a995d
ms.date: 06/08/2017
ms.openlocfilehash: ee000295416e1274d021be9d973df63e61082935
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="platesitem-property-publisher"></a>Свойство Plates.Item (издатель)

Возвращает объект отдельных из указанного семейства сайтов. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляющий объект **формы** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Длинный**|Количество для возвращаемого объекта.|

## <a name="example"></a>Пример

В этом примере имя первого цвет формы в active публикации.


```vb
MsgBox "Name of first color plate: " _ 
 &; ActiveDocument.Plates.Item(1).Name
```


