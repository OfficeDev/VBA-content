---
title: "Метод метода Columns.Add (издатель)"
keywords: vbapb10.chm5046276
f1_keywords: vbapb10.chm5046276
ms.prod: publisher
api_name: Publisher.Columns.Add
ms.assetid: b3dfb892-6bda-d2c4-11f7-9bd29bf257aa
ms.date: 06/08/2017
ms.openlocfilehash: 5b939c7643c3d2403cca50ae10c3246129762b71
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="columnsadd-method-publisher"></a>Метод метода Columns.Add (издатель)

Добавляет новый объект **столбца** для указанной коллекции **столбцов** и возвращает новый объект **столбца** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_BeforeColumn_**)

 переменная _expression_A, представляет собой объект- **столбцов** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|BeforeColumn|Необязательный| **Длинный**|Число столбцов, перед которым необходимо вставить новый столбец. Если этот аргумент указан, новый столбец добавляется после существующих столбцов. Если значение этого аргумента не соответствует существующего столбца в таблице, возникает ошибка.|

### <a name="return-value"></a>Возвращаемое значение

Столбец


## <a name="example"></a>Пример

В следующем примере добавляется столбец до трех столбцов в указанной таблице.


```vb
Dim colNew As Column 
 
Set colNew = ActiveDocument.Pages(1).Shapes(1) _ 
 .Table.Columns.Add(BeforeColumn:=3)
```


