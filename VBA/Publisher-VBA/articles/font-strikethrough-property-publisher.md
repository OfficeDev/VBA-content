---
title: "Свойство Font.StrikeThrough (издатель)"
keywords: vbapb10.chm5374017
f1_keywords: vbapb10.chm5374017
ms.prod: publisher
ms.assetid: fa4bca2d-b43d-4d2b-901f-858e277df520
ms.date: 06/08/2017
ms.openlocfilehash: a42efcf151a34cd457dbc79e099754044b6e9c38
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontstrikethrough-property-publisher"></a>Свойство Font.StrikeThrough (издатель)

Возвращает или задает константой **MsoTriState** , представляющее состояние свойства **зачеркивание** символы в диапазон текста. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Зачеркивание**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

 **MsoTriState**


## <a name="remarks"></a>Заметки

Значение свойства **зачеркивание** может иметь одно из следующих **MsoTriState** константы, описанные в библиотеке типов, Microsoft Office.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Ни один из символов в диапазоне форматируются как зачеркивание.|
| **msoTriStateMixed**|Возвращает значение, указывающее, что диапазон содержит текст в формате зачеркивание и текст не в формате зачеркивание.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Все символы в диапазоне форматируются как зачеркивание.|

