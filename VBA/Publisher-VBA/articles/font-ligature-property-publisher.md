---
title: "Свойство Font.Ligature (издатель)"
keywords: vbapb10.chm5374007
f1_keywords: vbapb10.chm5374007
ms.prod: publisher
ms.assetid: 17847824-8761-42b7-8d0c-00345e8b5de8
ms.date: 06/08/2017
ms.openlocfilehash: b8c52417251d28f9ef5ccee122a8c401654708f8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontligature-property-publisher"></a>Свойство Font.Ligature (издатель)

Возвращает или задает значение константы **[pbLigaturePresetType](pbligaturepresettype-enumeration-publisher.md)** , представляющий состояние свойства **в то же время** на символов в диапазон текста. Свойство **в то же время** включает Надсимвольные элементы в символы, часто в виде больше и больше затейливым засечек. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **В то же время**

 переменная _expression_A, представляющий объект **[Font](font-object-publisher.md)** .


## <a name="return-value"></a>Возвращаемое значение

 **pbLigaturePresetType**


## <a name="remarks"></a>Заметки


 **Примечание**  **В то же время** свойство действует только для шрифтов OpenType, которые содержат лигатуры.

Лигатуры — это альтернативное отображение последовательности символов; несколько символов объединяются в одного знака. Например лигатуры в режиме для _Microsoft Office_word, буквы _«ffi»_ все соединены друг в одного знака, которое отображает непрерывной строки из первого _f_ через точки в _i_.


