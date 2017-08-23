---
title: "Метод CellRange.Item (издатель)"
keywords: vbapb10.chm5177344
f1_keywords: vbapb10.chm5177344
ms.prod: publisher
api_name: Publisher.CellRange.Item
ms.assetid: 8f1fe143-e00c-7112-45dd-52158153cf28
ms.date: 06/08/2017
ms.openlocfilehash: 1ca31cf461de56aa78e90ac8c9ed9a4bd9f086a0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellrangeitem-method-publisher"></a>Метод CellRange.Item (издатель)

Возвращает объект отдельные **ячейки** в указанном семействе **CellRange** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Элемент** ( **_Индекс_**)

 переменная _expression_A, представляет собой объект- **CellRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Индекс|Обязательное свойство.| **Длинный**|Количество для возвращаемого объекта.|

### <a name="return-value"></a>Возвращаемое значение

Cell


## <a name="example"></a>Пример

Этот пример возвращает первую ячейку из коллекции **CellRange** .


```vb
Dim cllTemp As Cell 
 
Set cllTemp = ActiveDocument.Pages(Index:=1).Shapes(1).Table.Cells.Item(Index:=1)
```


