---
title: "Метод TextRange.Collapse (издатель)"
keywords: vbapb10.chm5308420
f1_keywords: vbapb10.chm5308420
ms.prod: publisher
api_name: Publisher.TextRange.Collapse
ms.assetid: ae177297-bf3b-ce0f-cf3a-29093b115996
ms.date: 06/08/2017
ms.openlocfilehash: 47541b426d2b1b0bafa9ba406c15b111b5eb6c28
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangecollapse-method-publisher"></a>Метод TextRange.Collapse (издатель)

Сворачивает диапазон или выделить фрагмент положение начальной или конечной позиции. После свертывания диапазона или выделить фрагмент равны начальную и конечную точку.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Свернуть** ( **_Направление_**)

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Direction|Обязательное свойство.| **PbCollapseDirection**|Направление свернуть диапазон или выделить фрагмент.|

## <a name="remarks"></a>Заметки

Если вы используете **pbCollapseEnd** сворачивание диапазона, на который ссылается абзаца, диапазон будут расположены после конца абзаца (начало следующий абзац). Тем не менее можно переместить назад один символ диапазона с помощью метода [MoveEnd](textrange-moveend-method-publisher.md)после свертывания диапазона.

Направление параметра может иметь одно из следующих **PbCollapseDirection** константы, описанные в библиотеке типов, Microsoft Publisher.



| **pbCollapseEnd**|| **pbCollapseStart**|

## <a name="example"></a>Пример

В этом примере вставляет текст в начало второго абзаца в первую фигуру на первой странице active публикации. Предполагается, что указанные форму фрагмент текста и не другого типа фигуры.


```vb
Sub CollapseRange() 
 Dim rngText As TextRange 
 Set rngText = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange 
 
 'Collapses range to the end of the range and 
 'enters new text and a new paragraph 
 With rngText 
 .Paragraphs(Start:=1, Length:=1).Collapse Direction:=pbCollapseEnd 
 .Text = "This is a new paragraph." &; vbCrLf 
 End With 
End Sub
```

В этом примере помещает новый текст в конце в первый абзац первую фигуру на первой странице active публикации. Предполагается, что указанные форму фрагмент текста и не другого типа фигуры.




```vb
Sub CollapseSelection() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .Paragraphs(Start:=1, Length:=1).Select 
 
 'Collapses selection to end and moves cursor back 
 'one character, then enters new text 
 With Selection.TextRange 
 .Collapse Direction:=pbCollapseEnd 
 .MoveEnd Unit:=pbTextUnitCharacter, Size:=-1 
 .Text = " This is a new test." 
 End With 
End Sub
```


