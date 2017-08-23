---
title: "Свойство ParagraphFormat.WidowControl (издатель)"
keywords: vbapb10.chm5439536
f1_keywords: vbapb10.chm5439536
ms.prod: publisher
api_name: Publisher.ParagraphFormat.WidowControl
ms.assetid: af1f1106-60e3-3987-3710-30fae7cf3940
ms.date: 06/08/2017
ms.openlocfilehash: f25749d76604b99881384abd4d3f2220b5f65978
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatwidowcontrol-property-publisher"></a>Свойство ParagraphFormat.WidowControl (издатель)

Задает или возвращает **MsoTriState** , представляющий ли первой или последней строки указанного абзаца могут появляться, сам по себе в текстовом поле. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WidowControl**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Этот параметр гарантирует, что первой или последней строки указанного абзаца не будет отображаться сам по себе в рамке. Например если последней строки в указанном абзаце является первой строки абзаца запрет управляются, второй строке перемещается на следующий кадр текста с ним.

Значение свойства **WidowControl** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|В текстовом поле первой или последней строке могут отображаться сам по себе.|
| **msoTrue**|В текстовом поле сам по себе не появится строка или фамилии.|
Значение по умолчанию для этого свойства — **msoFalse**.


## <a name="example"></a>Пример

В этом примере задается свойство **WidowControl** **msoTrue** для указанного объекта **ParagraphFormat** .


```vb
Dim objParaForm As ParagraphFormat 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Paragraphs(1).ParagraphFormat 
objParaForm.WidowControl = msoTrue 

```


