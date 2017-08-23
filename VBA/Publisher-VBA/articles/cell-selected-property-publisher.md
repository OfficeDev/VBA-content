---
title: "Свойство Cell.Selected (издатель)"
keywords: vbapb10.chm5111832
f1_keywords: vbapb10.chm5111832
ms.prod: publisher
api_name: Publisher.Cell.Selected
ms.assetid: b07f40bf-a14b-9b2a-2e0d-dc907cc78748
ms.date: 06/08/2017
ms.openlocfilehash: aa82aa49375519865396cbd970b14b7a20f861af
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="cellselected-property-publisher"></a>Свойство Cell.Selected (издатель)

Возвращает **значение True** , если при выборе ячейки. Только для чтения **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выбранные**

 переменная _expression_A, представляет собой объект- **ячейки** .


## <a name="example"></a>Пример

В этом примере определяется при выборе ячейки в указанной таблице и его при вводе текста в ячейку.


```vb
Sub IsCellSelected() 
 Dim cel As Cell 
 With ActiveDocument.Pages(1).Shapes(1).Table 
 For Each cel In .Cells 
 If cel.Selected Then 
 cel.TextRange.Text = "This cell is selected." 
 End If 
 Next cel 
 End With 
End Sub
```


