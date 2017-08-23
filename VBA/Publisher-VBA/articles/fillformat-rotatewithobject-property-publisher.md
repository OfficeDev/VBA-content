---
title: "Свойство FillFormat.RotateWithObject (издатель)"
keywords: vbapb10.chm2359585
f1_keywords: vbapb10.chm2359585
ms.prod: publisher
ms.assetid: a1e5f826-4200-4eac-204d-f17717160f33
ms.date: 06/08/2017
ms.openlocfilehash: cbd6463b7c4fb3323ef43c4af8c950390416422e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fillformatrotatewithobject-property-publisher"></a>Свойство FillFormat.RotateWithObject (издатель)

Возвращает или задает поворот заливки вместе с указанной фигуры. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RotateWithObject**

 переменная _expression_A, представляет собой объект- **FillFormat** .


## <a name="return-value"></a>Возвращаемое значение

 **MSOTRISTATE**


## <a name="remarks"></a>Заметки

Значение, возвращаемое свойством **RotateWithObject** может иметь одно из ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** константы, перечисленные в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Не поворачивайте заливки с формой.|
| **msoTrue**|Поворот заливки с формой.|
Значение свойства **RotateWithObject** соответствует параметру поля **Поворот с фигуры** в области **заполнения** диалогового окна **Формат фигуры** в пользовательском интерфейсе Publisher.


 **Примечание**  Поле **Поворот формы** отображается только в том случае, если у вас есть либо **градиентной заливки** или **рисунок или текстуры** переключатели выбрано на панели **заливки** в диалоговом окне " **Формат фигуры** ".


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект FillFormat](fillformat-object-publisher.md)

