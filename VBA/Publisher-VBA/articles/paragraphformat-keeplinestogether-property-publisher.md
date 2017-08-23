---
title: "Свойство ParagraphFormat.KeepLinesTogether (издатель)"
keywords: vbapb10.chm5439537
f1_keywords: vbapb10.chm5439537
ms.prod: publisher
api_name: Publisher.ParagraphFormat.KeepLinesTogether
ms.assetid: a0f3f2f0-d986-4928-3c4f-0665711a6876
ms.date: 06/08/2017
ms.openlocfilehash: 9acb9bc7d791e09645af0dfe4c3b042979c51689
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatkeeplinestogether-property-publisher"></a>Свойство ParagraphFormat.KeepLinesTogether (издатель)

Задает или возвращает **MsoTriState** , которое указывает, остается ли все строки в указанном абзаце в одном текстовом поле. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **KeepLinesTogether**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

msoTriState


## <a name="remarks"></a>Заметки

Этот параметр гарантирует, что не текст frame или столбца разрыв между строк указанного абзаца. Если абзацы слишком велик для текстового фрейма или столбца, первая строка будет запускаться в верхней части следующей текстовой рамке или столбца.

Значение по умолчанию для этого свойства — **msoFalse**.


## <a name="example"></a>Пример

В этом примере задается свойство **KeepLinesTogether** **msoTrue** для указанного объекта **ParagraphFormat** .


```vb
Dim objParaForm As ParagraphFormat 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Paragraphs(1).ParagraphFormat 
objParaForm.KeepLinesTogether = msoTrue 

```


