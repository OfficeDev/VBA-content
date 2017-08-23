---
title: "Свойство ShapeRange.WizardTag (издатель)"
keywords: vbapb10.chm2293860
f1_keywords: vbapb10.chm2293860
ms.prod: publisher
api_name: Publisher.ShapeRange.WizardTag
ms.assetid: 49bdeff9-fec4-2b40-1650-cd78c9bce0d4
ms.date: 06/08/2017
ms.openlocfilehash: 9bf6083228a31314e91b7153ca8345c33f6cf61a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangewizardtag-property-publisher"></a>Свойство ShapeRange.WizardTag (издатель)

Возвращает или задает значение, указывающее, функция указанного фигуры по отношению к его дизайн публикации константы **PbWizardTag**. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WizardTag**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Значение свойства **WizardTag** может иметь одно из **[PbWizardTag](pbwizardtag-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.

Комбинация свойство **[WizardTagInstance](shape-wizardtaginstance-property-publisher.md)** и свойство **WizardTag** однозначно определяет все фигуры в публикации.


## <a name="example"></a>Пример

Следующий пример отображает сведения об экземпляре тег мастера для всех фигур и мастер тега на странице один из активных публикации.


```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 With shpLoop 
 Debug.Print "Shape: " &; .Name 
 Debug.Print " Wizard tag: " &; .WizardTag 
 Debug.Print " Wizard tag instance: " _ 
 &; .WizardTagInstance 
 End With 
Next shpLoop
```


