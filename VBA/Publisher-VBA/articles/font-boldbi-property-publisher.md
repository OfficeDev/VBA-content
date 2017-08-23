---
title: "Свойство Font.BoldBi (издатель)"
keywords: vbapb10.chm5373956
f1_keywords: vbapb10.chm5373956
ms.prod: publisher
api_name: Publisher.Font.BoldBi
ms.assetid: f3a9fa27-6c9c-4d77-0f0d-962afa211d9d
ms.date: 06/08/2017
ms.openlocfilehash: 1f28af5756a88a3f0f2591abd757371943c9eda9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontboldbi-property-publisher"></a>Свойство Font.BoldBi (издатель)

Возвращает или задает константой **MsoTriState**, указывающее, является ли шрифт полужирным; используется с текстом на языке, справа налево. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BoldBi**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **BoldBi** может иметь одно из следующих **MsoTriState** константы, описанные в библиотеке типов, Microsoft Office.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Ни один из символов в диапазоне форматируются полужирным шрифтом.|
| **msoTriStateMixed**|Возвращает значение, указывающее, что диапазон содержит текст полужирным и не форматированный текст полужирным шрифтом.|
| **msoTriStateToggle**|Задайте значение, могут переключаться между **msoTrue** и **msoFalse**.|
| **msoTrue**|Все символы в диапазоне форматируются полужирным шрифтом.|

## <a name="example"></a>Пример

В этом примере проверяется текст в первой сценариев и отображается одно из двух возможных сообщений в зависимости от того, является ли текста справа налево отформатированный и является ли шрифт полужирным шрифтом. В этом примере для выполнения должным образом необходимо быть по крайней мере один сценариев с текстом в активной публикации.


```vb
Sub BoldRtoL() 
 
 Dim stf As Font 
 
 Set stf = Application.ActiveDocument.Stories(1).TextRange.Font 
 
 With stf 
 If .BoldBi = msoTrue Then 
 MsgBox "This story is right-to-left and is bold." 
 Else 
 MsgBox "This story is either not right-to-left" &; _ 
 " or it is not bold." 
 End If 
 End With 
 
End Sub
```


