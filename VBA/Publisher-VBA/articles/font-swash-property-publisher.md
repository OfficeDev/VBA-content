---
title: "Свойство Font.Swash (издатель)"
keywords: vbapb10.chm5374005
f1_keywords: vbapb10.chm5374005
ms.prod: publisher
api_name: Publisher.Font.Swash
ms.assetid: 71537393-167a-f9e3-e3b3-ae743fdbb0ff
ms.date: 06/08/2017
ms.openlocfilehash: fbc73571a2406e349ad93ee366e7e9c89478c61d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontswash-property-publisher"></a>Свойство Font.Swash (издатель)

Возвращает или задает константой **MsoTriState** , представляющее состояние свойства **Swash** символов в диапазон текста. Свойство **Swash** включает Надсимвольные элементы в символы, часто в виде больше и больше затейливым засечек. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Swash**

 переменная _expression_A, представляющий объект **[Font](font-object-publisher.md)** .


## <a name="return-value"></a>Возвращаемое значение

 **MsoTriState**


## <a name="remarks"></a>Заметки


 **Примечание**  Свойство **Swash** действует только для шрифтов OpenType, которые содержат шлейфов.

Значение свойства **Swash** может иметь одно из следующих **MsoTriState** константы, описанные в библиотеке типов, Microsoft Office.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Ни один из символов в диапазоне форматируются как swash.|
| **msoTriStateMixed**|Возвращает значение, указывающее, что диапазон содержит текст в формате каллиграфическая или не в формате swash текст.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Все символы в диапазоне форматируются как swash.|

