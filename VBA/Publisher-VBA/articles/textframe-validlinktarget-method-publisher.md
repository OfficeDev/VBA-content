---
title: "Метод TextFrame.ValidLinkTarget (издатель)"
keywords: vbapb10.chm3866662
f1_keywords: vbapb10.chm3866662
ms.prod: publisher
api_name: Publisher.TextFrame.ValidLinkTarget
ms.assetid: ee946f58-669f-7150-0f40-2dd3b857e274
ms.date: 06/08/2017
ms.openlocfilehash: ab3f0bdf43af102735feb622bdb96defcf375054
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframevalidlinktarget-method-publisher"></a>Метод TextFrame.ValidLinkTarget (издатель)

Определяет, могут быть связаны рамки одной фигуры с рамки другую фигуру. Возвращает **значение True,** Если **_LinkTarget_** допустимое конечное **значение False** , если **_LinkTarget_** уже содержит текст или уже связанные или не поддерживает фигуры текстом.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ValidLinkTarget** ( **_LinkTarget_**)

 переменная _expression_A, представляет собой объект- **TextFrame** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|LinkTarget|Обязательное свойство.| **Фигура**|Фигура с рамки текста, для которого необходимо создать ссылку frame текста, возвращаемые выражением.|

### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере проверяется, является ли текстовые рамки для первой и второй фигур на первой странице active публикации могут быть связаны друг с другом. Если это так, пример связывает две текстовые рамки.


```vb
Dim txtFrame1 As TextFrame 
Dim txtFrame2 As TextFrame 
 
With ActiveDocument.Pages(1) 
 Set txtFrame1 = .Shapes(1).TextFrame 
 Set txtFrame2 = .Shapes(2).TextFrame 
End With 
 
If txtFrame1.ValidLinkTarget(LinkTarget:=txtFrame2.Parent) = True Then 
 txtFrame1.NextLinkedTextFrame = txtFrame2 
End If
```


