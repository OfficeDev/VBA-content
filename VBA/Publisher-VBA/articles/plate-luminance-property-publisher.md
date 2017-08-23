---
title: "Свойство Plate.Luminance (издатель)"
keywords: vbapb10.chm2883590
f1_keywords: vbapb10.chm2883590
ms.prod: publisher
api_name: Publisher.Plate.Luminance
ms.assetid: 8d84fe74-8421-4ec2-bf6e-a156a0c0018b
ms.date: 06/08/2017
ms.openlocfilehash: 2c4b39277fc0078ff7979bd1347ccb78e5e5c353
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="plateluminance-property-publisher"></a>Свойство Plate.Luminance (издатель)

Возвращает или задает **Long** , указывающее, вычисляемых освещенности значение для указанной формы; используется для перехват кусочков цвет. Допустимые значения: от 0 до 100. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Яркости**

 переменная _expression_A, представляющий объект **формы** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Это свойство является допустимым только для публикации с ** [ColorMode](http://msdn.microsoft.com/library/58befa97-9d9b-9294-18b2-ae10dc87f51c%28Office.15%29.aspx)** значение свойства **pbColorModeSpot** или для формы смесевых цветов в публикации со значением свойства **ColorMode** **pbColorModeSpotAndProcess**.


## <a name="example"></a>Пример

В следующем примере циклически просматривает все формы кусочков цветов в публикации и их значения яркости отчетов.


```vb
Dim plaTemp As Plates 
Dim plaLoop As Plate 
 
Set plaTemp = ActiveDocument.Plates 
 
If ActiveDocument.ColorMode <> pbColorModeSpot And _ 
 ActiveDocument.ColorMode <> pbColorModeSpotAndProcess Then 
 Debug.Print "No spot colors in this publication." 
Else 
 For Each plaLoop In plaTemp 
 With plaLoop 
 Debug.Print "Plate " &; .Name _ 
 &; " has a luminance of " &; .Luminance 
 End With 
 Next plaLoop 
End If
```


