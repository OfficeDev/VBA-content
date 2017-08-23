---
title: "Свойство Table.TableDirection (издатель)"
keywords: vbapb10.chm4784135
f1_keywords: vbapb10.chm4784135
ms.prod: publisher
api_name: Publisher.Table.TableDirection
ms.assetid: ffd664a8-781f-8fdc-055c-1ea7309b3b38
ms.date: 06/08/2017
ms.openlocfilehash: abaccbb9b818b15ce78614206caf0209f9fcd2f6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tabletabledirection-property-publisher"></a>Свойство Table.TableDirection (издатель)

Возвращает или задает значение константы **PbTableDirectionType** , представляет ли текста в таблице считываются слева направо или справа налево. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TableDirection**

 переменная _expression_A, представляет собой объект- **таблицы** .


### <a name="return-value"></a>Возвращаемое значение

PbTableDirectionType


## <a name="remarks"></a>Заметки

Значение свойства **TableDirection** может иметь одно из **PbTableDirectionType** константы в библиотеке типов, Microsoft Publisher.



| **pbTableDirectionLeftToRight**|| **pbTableDirectionRightToLeft**|

## <a name="example"></a>Пример

В этом примере вводит полужирный номер в каждой ячейки в указанной таблице и затем задает направление таблицы, чтобы ячейки, какой номер справа налево. В данном примере для работы указанного фигуры должен быть таблица.


```vb
Sub CountCellsByColumn() 
 Dim tblTable As Table 
 Dim rowTable As row 
 Dim celTable As Cell 
 Dim intCount As Integer 
 
 Set tblTable = ActiveDocument.Pages(1).Shapes(1).Table 
 
 'Loops through each row in the table 
 For Each rowTable In tblTable.Rows 
 
 'Loops through each cell in the row 
 For Each celTable In rowTable.Cells 
 With celTable.TextRange 
 intCount = intCount + 1 
 .Text = intCount 
 .ParagraphFormat.Alignment = _ 
 pbParagraphAlignmentCenter 
 .Font.Bold = msoTrue 
 End With 
 Next celTable 
 Next rowTable 
 tblTable.TableDirection = pbTableDirectionRightToLeft 
End Sub
```


