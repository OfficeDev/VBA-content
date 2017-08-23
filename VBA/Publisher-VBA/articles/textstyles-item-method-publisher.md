---
title: "Метод TextStyles.Item (издатель)"
keywords: vbapb10.chm5898240
f1_keywords: vbapb10.chm5898240
ms.prod: publisher
api_name: Publisher.TextStyles.Item
ms.assetid: 14d1871f-c2cb-31af-e22d-10b3cf59b6fc
ms.date: 06/08/2017
ms.openlocfilehash: 998d5209063942ea00e7144d463c7c09b1accf72
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textstylesitem-method-publisher"></a>Метод TextStyles.Item (издатель)

Возвращает объект отдельных в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляет собой объект- **TextStyles** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Variant**|Номер или имя поля или поля элемента списка, чтобы возвратить.|

### <a name="return-value"></a>Возвращаемое значение

Стиля текста


## <a name="example"></a>Пример

Этот пример возвращает стиль «Обычный» текст из активной публикации.


```vb
Dim txtStyle As TextStyle 
 
Set txtStyle = ActiveDocument.TextStyles.Item(Index:="Normal")
```


