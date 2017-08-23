---
title: "Метод TextRange.InsertSymbol (издатель)"
keywords: vbapb10.chm5308452
f1_keywords: vbapb10.chm5308452
ms.prod: publisher
api_name: Publisher.TextRange.InsertSymbol
ms.assetid: 607d12da-5a2d-4e0e-b45e-92275ce97bab
ms.date: 06/08/2017
ms.openlocfilehash: fe91b046cb65c33b60209446398cf2e075ee1ca6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangeinsertsymbol-method-publisher"></a>Метод TextRange.InsertSymbol (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий символ вставлен вместо заданного диапазона или выделения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InsertSymbol** ( **_FontName_**, **_CharIndex_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|FontName|Обязательное свойство.| **String**|Имя шрифта, которая содержит символ.|
|CharIndex|Обязательное свойство.| **Длинный**|Символ Юникода для указанного символа.|

### <a name="return-value"></a>Возвращаемое значение

TextRange


## <a name="remarks"></a>Заметки

Если вы не хотите заменить диапазон или выделить фрагмент, используйте [Метод TextRange.Collapse (издатель)](textrange-collapse-method-publisher.md) , прежде чем использовать этот метод.


## <a name="example"></a>Пример

В этом примере вставляет двунаправленной стрелки в текущей позиции.


```vb
Sub Insert Arrow() 
    ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
            .Paragraphs(Start:=1, Length:=1).Select
    With .TextFrame.TextRange 
            .InsertPageNumber 
            .Collapse Direction:= pbCollapseStart
            .InsertSymbol FontName:="Symbol", CharIndex:=171
        End With 
End Sub
```


