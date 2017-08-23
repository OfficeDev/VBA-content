---
title: "Свойство Plate.InkName (издатель)"
keywords: vbapb10.chm2883603
f1_keywords: vbapb10.chm2883603
ms.prod: publisher
api_name: Publisher.Plate.InkName
ms.assetid: 248c1529-2706-5458-a13f-def479d16132
ms.date: 06/08/2017
ms.openlocfilehash: 1dcb76784a04a0f4ef6ef286f1c47b28680e9cc8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="plateinkname-property-publisher"></a>Свойство Plate.InkName (издатель)

Возвращает константу **PbInkName** , представляющий имя рукописного ввода для печати с помощью этой формы. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InkName**

 переменная _expression_A, представляющий объект **формы** .


## <a name="remarks"></a>Заметки

Значение свойства **InkName** может иметь одно из ** [PbInkName](http://msdn.microsoft.com/library/69e335b8-40b8-c984-84b6-64073a8ed7ab%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Publisher.

Используйте метод **FindPlateByInkName** **[PrintablePlates](printableplates-object-publisher.md)** коллекции для возврата определенного формы с учетом его рукописного ввода имени.


## <a name="example"></a>Пример

Следующий пример возвращает список подготовленных к печати формы в настоящее время в коллекции для активной публикации. В примере предполагается, что цветоделение были указаны в качестве режима печати active публикации.


```vb
Sub ListPrintablePlates() 
 Dim pplTemp As PrintablePlates 
 Dim pplLoop As PrintablePlate 
 
 
 Set pplTemp = ActiveDocument.AdvancedPrintOptions.PrintablePlates 
 Debug.Print "There are " &; pplTemp.Count &; " printable plates in this publication." 
 
 For Each pplLoop In pplTemp 
 With pplLoop 
 Debug.Print "Printable Plate Name: " &; .Name 
 Debug.Print "Index: " &; .Index 
 Debug.Print "Ink Name: " &; .InkName 
 Debug.Print "Plate Angle: " &; .Angle 
 Debug.Print "Plate Frequency: " &; .Frequency 
 Debug.Print "Print Plate?: " &; .PrintPlate 
 End With 
 Next pplLoop 
End Sub
```


