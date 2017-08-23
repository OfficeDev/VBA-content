---
title: "Метод Fields.Item (издатель)"
keywords: vbapb10.chm6029312
f1_keywords: vbapb10.chm6029312
ms.prod: publisher
api_name: Publisher.Fields.Item
ms.assetid: 95783e5a-2c82-235e-75a4-5ac15938718e
ms.date: 06/08/2017
ms.openlocfilehash: 6abb664194cce787b6d7c4b6b09d39e843b7a91e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fieldsitem-method-publisher"></a>Метод Fields.Item (издатель)

Возвращает объект отдельных в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляющий объект **поля** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Длинный**|Количество для возвращаемого объекта.|

### <a name="return-value"></a>Возвращаемое значение

Поле


## <a name="example"></a>Пример

В этом примере возвращается первое поле из объекта **поля** .


```vb
Dim fldTemp As Field 
 
Set fldTemp = ActiveDocument.Pages(Index:=1) _ 
 .Shapes(1).TextFrame.TextRange.Fields.Item(Index:=1)
```


