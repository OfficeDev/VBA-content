---
title: "Свойство Selection.Type (издатель)"
keywords: vbapb10.chm851971
f1_keywords: vbapb10.chm851971
ms.prod: publisher
api_name: Publisher.Selection.Type
ms.assetid: 4dfcfecc-dd76-36b6-21df-34c3865b3064
ms.date: 06/08/2017
ms.openlocfilehash: ab7e27cda6fa69f6ea7a7a1c1b2440f4f87d182b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="selectiontype-property-publisher"></a>Свойство Selection.Type (издатель)

Возвращает константу **PbSelectionType** , представляющий тип выделения. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Тип**

 переменная _expression_A, представляющий объект **Selection** .


## <a name="remarks"></a>Заметки

Значение свойства **типа** может иметь одно из следующих констант **PbSelectionType** .



| **pbSelectionNone**|| **pbSelectionShape**|| **pbSelectionShapeSubSelection**|| **pbSelectionTableCells**|| **pbSelectionText**|

## <a name="example"></a>Пример

В этом примере проверяется при выделении текста и если он установлен, позволяет выделенный текст полужирным шрифтом.


```vb
Sub IfCellData() 
 Dim rowTable As Row 
 If Selection.Type = pbSelectionText Then 
 Selection.TextRange.Font.Bold = msoTrue 
 End If 
End Sub
```


