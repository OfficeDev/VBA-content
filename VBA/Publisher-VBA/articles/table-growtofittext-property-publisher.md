---
title: "Свойство Table.GrowToFitText (издатель)"
keywords: vbapb10.chm4784132
f1_keywords: vbapb10.chm4784132
ms.prod: publisher
api_name: Publisher.Table.GrowToFitText
ms.assetid: d8822df7-a252-a5bb-be26-83df8ec5eb94
ms.date: 06/08/2017
ms.openlocfilehash: dae6529fa9e2335e2975c2cf5fffcb33dce40348
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tablegrowtofittext-property-publisher"></a>Свойство Table.GrowToFitText (издатель)

 **Значение true** для ячеек в таблице увеличить по размеру текста по вертикали. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **GrowToFitText**

 переменная _expression_A, представляет собой объект- **таблицы** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере задается каждой строки в указанной таблице 12 пунктов, а не увеличивает высоты строки текста при добавлении ячеек в строках.


```vb
Sub DontEnlargeTableCells() 
 Dim rowTable As Row 
 With ActiveDocument.Pages(1).Shapes(1).Table 
 .GrowToFitText = False 
 For Each rowTable In .Rows 
 rowTable.Height = 12 
 Next 
 End With 
End Sub
```


