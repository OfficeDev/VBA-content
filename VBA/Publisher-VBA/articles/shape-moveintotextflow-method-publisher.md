---
title: "Метод Shape.MoveIntoTextFlow (издатель)"
keywords: vbapb10.chm2228356
f1_keywords: vbapb10.chm2228356
ms.prod: publisher
api_name: Publisher.Shape.MoveIntoTextFlow
ms.assetid: d8a2af57-f974-717e-0d97-c8a3aee16f01
ms.date: 06/08/2017
ms.openlocfilehash: 07cf409c52ae20b29511de56db21d1cd838899a7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapemoveintotextflow-method-publisher"></a>Метод Shape.MoveIntoTextFlow (издатель)

Перемещает заданной фигуры в текстовый поток, определенные в ** [Объект TextRange](textrange-object-publisher.md)**. Фигура всегда будет вставленный встроенного в начале текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MoveIntoTextFlow** ( **_Диапазон_**)

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Range|Обязательное свойство.| **TextRange**|Диапазон текста, перед которым будет вставлена заданной фигуры.|

### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Метод **MoveIntoTextFlow** завершится ошибкой, если фигуры перемещаемых уже встроенные или не является допустимым встроенный тип фигуры. Недопустимый встроенных типов фигур включают:


- Встроенных фигур
    
- Изменение группы фигур
    
- Фрагменты HTML
    
- Смарт-объекты
    
- Связанные текстовые поля
    



## <a name="example"></a>Пример

В следующем примере проверяется, если второй фигуры на второй странице публикации — inline и если это не так, вставляет его встроенного в начале текста диапазона заданный текст. 


```vb
Dim theShape As Shape 
Dim theRange As TextRange 
 
Set theRange = ActiveDocument.Pages(2).Shapes(1).TextFrame.TextRange 
Set theShape = ActiveDocument.Pages(2).Shapes(2) 
 
If Not theShape.IsInline = msoTrue Then 
 theShape.MoveIntoTextFlow Range:=theRange 
End If 

```


