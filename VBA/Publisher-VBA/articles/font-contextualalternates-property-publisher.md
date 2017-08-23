---
title: "Свойство Font.ContextualAlternates (издатель)"
keywords: vbapb10.chm5374009
f1_keywords: vbapb10.chm5374009
ms.prod: publisher
api_name: Publisher.Font.ContextualAlternates
ms.assetid: 4737d43a-4ab8-0ae7-ce45-7be62f4aae6e
ms.date: 06/08/2017
ms.openlocfilehash: 42385e868aaf31076e1680ee7c1d180131db0ba9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontcontextualalternates-property-publisher"></a>Свойство Font.ContextualAlternates (издатель)

Возвращает или задает константой **MsoTriState** , представляющий состояние свойство **ContextualAlternates** на символов в диапазон текста. Свойство **ContextualAlternates** включает различные фигуры варианты для некоторых символов в зависимости от контекста знаков и проектирования выбранного шрифта. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ContextualAlternates**

 переменная _expression_A, представляющий объект **[Font](font-object-publisher.md)** .


## <a name="return-value"></a>Возвращаемое значение

 **MsoTriState**


## <a name="remarks"></a>Заметки


 **Примечание**  Свойство **ContextualAlternates** имеет значение только для шрифтов OpenType, которые содержат контекстные варианты.

Значение свойства **ContextualAlternates** может иметь одно из следующих **MsoTriState** константы, описанные в библиотеке типов, Microsoft Office.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Ни один из символов в диапазоне, отформатированный контекстные варианты.|
| **msoTriStateMixed**|Возвращает значение, указывающее, что диапазон содержит текст отформатирован контекстные варианты и текст не отформатирован контекстные варианты.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Все символы в диапазоне, отформатированный контекстные варианты.|

