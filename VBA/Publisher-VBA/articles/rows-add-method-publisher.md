---
title: "Метод Rows.Add (издатель)"
keywords: vbapb10.chm4915204
f1_keywords: vbapb10.chm4915204
ms.prod: publisher
api_name: Publisher.Rows.Add
ms.assetid: 34d72709-92f7-ddc6-5be6-e74693466e61
ms.date: 06/08/2017
ms.openlocfilehash: be3a3cef3b601a6feedc632bdc44db4cffb98e3d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="rowsadd-method-publisher"></a>Метод Rows.Add (издатель)

Добавляет новый объект **строки** в указанный набор **строк** и возвращает новый объект **строки** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_BeforeRow_**)

 переменная _expression_A, представляет собой объект- **строк** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|BeforeRow|Необязательный| **Длинный**|Номер строки, перед которым необходимо вставить новую строку. Если этот аргумент задан, новая строка добавляется после существующих строк. Если значение этого аргумента не соответствует существующей строки в таблице, возникает ошибка.|

### <a name="return-value"></a>Возвращаемое значение

Строка


## <a name="example"></a>Пример

Следующий пример добавляет строку до трех строк в указанной таблице.


```vb
Dim rowNew As Row 
 
Set rowNew = ActiveDocument.Pages(1).Shapes(1) _ 
 .Table.Rows.Add(BeforeRow:=3)
```


