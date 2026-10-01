вар3
Sub Variant3()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Данные").Delete
    ThisWorkbook.Sheets("Саати").Delete
    ThisWorkbook.Sheets("Предпочтения").Delete
    ThisWorkbook.Sheets("Ранг").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Данные"
    ws.Range("A1").Value = "Альтернатива"
    ws.Range("B1:E1").Value = Array("A1", "A2", "A3", "A4")
    ws.Range("A2").Value = "Эксперт 1": ws.Range("A3").Value = "Эксперт 2": ws.Range("A4").Value = "Эксперт 3"
    ws.Range("B2:E2").Value = Array(2, 1, 3, 4)
    ws.Range("B3:E3").Value = Array(1, 3, 2, 4)
    ws.Range("B4:E4").Value = Array(3, 1, 2, 4)
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Саати"
    ws.Range("A1").Value = "МЕТОД СААТИ (эксперт 1)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "A1": ws.Range("A5").Value = "A2"
    ws.Range("A6").Value = "A3": ws.Range("A7").Value = "A4"
    ws.Range("B4:E4").Value = Array(1, 1/3, 5, 7)
    ws.Range("B5:E5").Value = Array(3, 1, 7, 9)
    ws.Range("B6:E6").Value = Array(1/5, 1/7, 1, 3)
    ws.Range("B7:E7").Value = Array(1/7, 1/9, 1/3, 1)
    ws.Range("F3").Value = "Цены"
    ws.Range("G3").Formula = "=(B4*C4*D4*E4)^(1/4)"
    ws.Range("G4").Formula = "=(B5*C5*D5*E5)^(1/4)"
    ws.Range("G5").Formula = "=(B6*C6*D6*E6)^(1/4)"
    ws.Range("G6").Formula = "=(B7*C7*D7*E7)^(1/4)"
    ws.Range("F7").Value = "Сумма": ws.Range("G7").Formula = "=SUM(G3:G6)"
    ws.Range("H3").Formula = "=G3/$G$7": ws.Range("H4").Formula = "=G4/$G$7"
    ws.Range("H5").Formula = "=G5/$G$7": ws.Range("H6").Formula = "=G6/$G$7"
    ws.Range("A8").Value = "Суммы столбцов"
    ws.Range("B8").Formula = "=SUM(B4:B7)": ws.Range("C8").Formula = "=SUM(C4:C7)"
    ws.Range("D8").Formula = "=SUM(D4:D7)": ws.Range("E8").Formula = "=SUM(E4:E7)"
    ws.Range("F10").Value = "λ =": ws.Range("G10").Formula = "=B8*H3 + C8*H4 + D8*H5 + E8*H6"
    ws.Range("F11").Value = "ИС =": ws.Range("G11").Formula = "=(G10-4)/3"
    ws.Range("F12").Value = "C1C =": ws.Range("G12").Value = 0.9
    ws.Range("F13").Value = "ОС =": ws.Range("G13").Formula = "=G11/G12"
    ws.Range("F14").Value = "Вывод:": ws.Range("G14").Formula = "=IF(G13<0.2,""Непротиворечиво"",""Требует уточнения"")"
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Предпочтения"
    ws.Range("A1").Value = "МЕТОД ПРЕДПОЧТЕНИЙ (3 эксперта)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "Э1": ws.Range("A5").Value = "Э2": ws.Range("A6").Value = "Э3"
    ws.Range("B4:E4").Value = Array(2, 1, 3, 4)
    ws.Range("B5:E5").Value = Array(1, 3, 2, 4)
    ws.Range("B6:E6").Value = Array(3, 1, 2, 4)
    ws.Range("A8").Value = "Преобразование B = 4 - X"
    ws.Range("B9").Value = "A1": ws.Range("C9").Value = "A2"
    ws.Range("D9").Value = "A3": ws.Range("E9").Value = "A4"
    ws.Range("A10").Value = "Э1": ws.Range("A11").Value = "Э2": ws.Range("A12").Value = "Э3"
    ws.Range("B10:E10").Formula = "=4-B4"
    ws.Range("B11:E11").Formula = "=4-B5"
    ws.Range("B12:E12").Formula = "=4-B6"
    ws.Range("A13").Value = "Cj:"
    ws.Range("B13").Formula = "=SUM(B10:B12)": ws.Range("C13").Formula = "=SUM(C10:C12)"
    ws.Range("D13").Formula = "=SUM(D10:D12)": ws.Range("E13").Formula = "=SUM(E10:E12)"
    ws.Range("F14").Value = "Сумма C =": ws.Range("G14").Formula = "=SUM(B13:E13)"
    ws.Range("A15").Value = "Веса Vj:"
    ws.Range("B15").Formula = "=B13/$G$14": ws.Range("C15").Formula = "=C13/$G$14"
    ws.Range("D15").Formula = "=D13/$G$14": ws.Range("E15").Formula = "=E13/$G$14"
    ws.Range("A17").Value = "Проверка согласованности"
    ws.Range("A18").Value = "Sj:"
    ws.Range("B18").Formula = "=SUM(B4:B6)": ws.Range("C18").Formula = "=SUM(C4:C6)"
    ws.Range("D18").Formula = "=SUM(D4:D6)": ws.Range("E18").Formula = "=SUM(E4:E6)"
    ws.Range("A19").Value = "A =": ws.Range("B19").Formula = "=3*(4+1)/2"
    ws.Range("A20").Value = "S =": ws.Range("B20").Formula = "=(B18-B19)^2+(C18-B19)^2+(D18-B19)^2+(E18-B19)^2"
    ws.Range("A21").Value = "W =": ws.Range("B21").Formula = "=12*B20/(3^2*4*(4^2-1))"
    ws.Range("A22").Value = "Вывод:": ws.Range("B22").Formula = "=IF(B21>0.5,""Согласовано"",""Требует уточнения"")"
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Ранг"
    ws.Range("A1").Value = "МЕТОД РАНГА (3 эксперта, 10 баллов)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "Э1": ws.Range("A5").Value = "Э2": ws.Range("A6").Value = "Э3"
    ws.Range("B4:E4").Value = Array(9, 10, 5, 2)
    ws.Range("B5:E5").Value = Array(10, 5, 8, 2)
    ws.Range("B6:E6").Value = Array(5, 10, 8, 2)
    ws.Range("A7").Value = "Cj:"
    ws.Range("B7").Formula = "=SUM(B4:B6)": ws.Range("C7").Formula = "=SUM(C4:C6)"
    ws.Range("D7").Formula = "=SUM(D4:D6)": ws.Range("E7").Formula = "=SUM(E4:E6)"
    ws.Range("A8").Value = "Сумма C:": ws.Range("B8").Formula = "=SUM(B7:E7)"
    ws.Range("A9").Value = "Веса Vj:"
    ws.Range("B9").Formula = "=B7/$B$8": ws.Range("C9").Formula = "=C7/$B$8"
    ws.Range("D9").Formula = "=D7/$B$8": ws.Range("E9").Formula = "=E7/$B$8"
    ws.Range("A11").Value = "Средние Xj:"
    ws.Range("B11").Formula = "=B7/3": ws.Range("C11").Formula = "=C7/3"
    ws.Range("D11").Formula = "=D7/3": ws.Range("E11").Formula = "=E7/3"
    ws.Range("G3").Value = "Дисперсии экспертов:"
    ws.Range("G4").Formula = "=((B4-$B$11)^2+(C4-$C$11)^2+(D4-$D$11)^2+(E4-$E$11)^2)/3"
    ws.Range("G5").Formula = "=((B5-$B$11)^2+(C5-$C$11)^2+(D5-$D$11)^2+(E5-$E$11)^2)/3"
    ws.Range("G6").Formula = "=((B6-$B$11)^2+(C6-$C$11)^2+(D6-$D$11)^2+(E6-$E$11)^2)/3"
    ws.Range("A13").Value = "Дисперсии альтернатив:"
    ws.Range("B13").Formula = "=((B4-$B$11)^2+(B5-$B$11)^2+(B6-$B$11)^2)/2"
    ws.Range("C13").Formula = "=((C4-$C$11)^2+(C5-$C$11)^2+(C6-$C$11)^2)/2"
    ws.Range("D13").Formula = "=((D4-$D$11)^2+(D5-$D$11)^2+(D6-$D$11)^2)/2"
    ws.Range("E13").Formula = "=((E4-$E$11)^2+(E5-$E$11)^2+(E6-$E$11)^2)/2"
    ThisWorkbook.Sheets("Данные").Activate
    MsgBox "Готово! Создан Вариант 3."
End Sub



вар4
Sub Variant4()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Данные").Delete
    ThisWorkbook.Sheets("Саати").Delete
    ThisWorkbook.Sheets("Предпочтения").Delete
    ThisWorkbook.Sheets("Ранг").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Данные"
    ws.Range("A1").Value = "Альтернатива"
    ws.Range("B1:E1").Value = Array("A1", "A2", "A3", "A4")
    ws.Range("A2").Value = "Эксперт 1": ws.Range("A3").Value = "Эксперт 2": ws.Range("A4").Value = "Эксперт 3"
    ws.Range("B2:E2").Value = Array(3, 2, 1, 4)
    ws.Range("B3:E3").Value = Array(1, 4, 2, 3)
    ws.Range("B4:E4").Value = Array(4, 3, 1, 2)
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Саати"
    ws.Range("A1").Value = "МЕТОД СААТИ (эксперт 1)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "A1": ws.Range("A5").Value = "A2"
    ws.Range("A6").Value = "A3": ws.Range("A7").Value = "A4"
    ws.Range("B4:E4").Value = Array(1, 1/3, 1/7, 5)
    ws.Range("B5:E5").Value = Array(3, 1, 1/5, 7)
    ws.Range("B6:E6").Value = Array(7, 5, 1, 9)
    ws.Range("B7:E7").Value = Array(1/5, 1/7, 1/9, 1)
    ws.Range("F3").Value = "Цены"
    ws.Range("G3").Formula = "=(B4*C4*D4*E4)^(1/4)"
    ws.Range("G4").Formula = "=(B5*C5*D5*E5)^(1/4)"
    ws.Range("G5").Formula = "=(B6*C6*D6*E6)^(1/4)"
    ws.Range("G6").Formula = "=(B7*C7*D7*E7)^(1/4)"
    ws.Range("F7").Value = "Сумма": ws.Range("G7").Formula = "=SUM(G3:G6)"
    ws.Range("H3").Formula = "=G3/$G$7": ws.Range("H4").Formula = "=G4/$G$7"
    ws.Range("H5").Formula = "=G5/$G$7": ws.Range("H6").Formula = "=G6/$G$7"
    ws.Range("A8").Value = "Суммы столбцов"
    ws.Range("B8").Formula = "=SUM(B4:B7)": ws.Range("C8").Formula = "=SUM(C4:C7)"
    ws.Range("D8").Formula = "=SUM(D4:D7)": ws.Range("E8").Formula = "=SUM(E4:E7)"
    ws.Range("F10").Value = "λ =": ws.Range("G10").Formula = "=B8*H3 + C8*H4 + D8*H5 + E8*H6"
    ws.Range("F11").Value = "ИС =": ws.Range("G11").Formula = "=(G10-4)/3"
    ws.Range("F12").Value = "C1C =": ws.Range("G12").Value = 0.9
    ws.Range("F13").Value = "ОС =": ws.Range("G13").Formula = "=G11/G12"
    ws.Range("F14").Value = "Вывод:": ws.Range("G14").Formula = "=IF(G13<0.2,""Непротиворечиво"",""Требует уточнения"")"
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Предпочтения"
    ws.Range("A1").Value = "МЕТОД ПРЕДПОЧТЕНИЙ (3 эксперта)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "Э1": ws.Range("A5").Value = "Э2": ws.Range("A6").Value = "Э3"
    ws.Range("B4:E4").Value = Array(3, 2, 1, 4)
    ws.Range("B5:E5").Value = Array(1, 4, 2, 3)
    ws.Range("B6:E6").Value = Array(4, 3, 1, 2)
    ws.Range("A8").Value = "Преобразование B = 4 - X"
    ws.Range("B9").Value = "A1": ws.Range("C9").Value = "A2"
    ws.Range("D9").Value = "A3": ws.Range("E9").Value = "A4"
    ws.Range("A10").Value = "Э1": ws.Range("A11").Value = "Э2": ws.Range("A12").Value = "Э3"
    ws.Range("B10:E10").Formula = "=4-B4"
    ws.Range("B11:E11").Formula = "=4-B5"
    ws.Range("B12:E12").Formula = "=4-B6"
    ws.Range("A13").Value = "Cj:"
    ws.Range("B13").Formula = "=SUM(B10:B12)": ws.Range("C13").Formula = "=SUM(C10:C12)"
    ws.Range("D13").Formula = "=SUM(D10:D12)": ws.Range("E13").Formula = "=SUM(E10:E12)"
    ws.Range("F14").Value = "Сумма C =": ws.Range("G14").Formula = "=SUM(B13:E13)"
    ws.Range("A15").Value = "Веса Vj:"
    ws.Range("B15").Formula = "=B13/$G$14": ws.Range("C15").Formula = "=C13/$G$14"
    ws.Range("D15").Formula = "=D13/$G$14": ws.Range("E15").Formula = "=E13/$G$14"
    ws.Range("A17").Value = "Проверка согласованности"
    ws.Range("A18").Value = "Sj:"
    ws.Range("B18").Formula = "=SUM(B4:B6)": ws.Range("C18").Formula = "=SUM(C4:C6)"
    ws.Range("D18").Formula = "=SUM(D4:D6)": ws.Range("E18").Formula = "=SUM(E4:E6)"
    ws.Range("A19").Value = "A =": ws.Range("B19").Formula = "=3*(4+1)/2"
    ws.Range("A20").Value = "S =": ws.Range("B20").Formula = "=(B18-B19)^2+(C18-B19)^2+(D18-B19)^2+(E18-B19)^2"
    ws.Range("A21").Value = "W =": ws.Range("B21").Formula = "=12*B20/(3^2*4*(4^2-1))"
    ws.Range("A22").Value = "Вывод:": ws.Range("B22").Formula = "=IF(B21>0.5,""Согласовано"",""Требует уточнения"")"
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Ранг"
    ws.Range("A1").Value = "МЕТОД РАНГА (3 эксперта, 10 баллов)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "Э1": ws.Range("A5").Value = "Э2": ws.Range("A6").Value = "Э3"
    ws.Range("B4:E4").Value = Array(5, 8, 10, 2)
    ws.Range("B5:E5").Value = Array(10, 2, 8, 5)
    ws.Range("B6:E6").Value = Array(2, 5, 10, 8)
    ws.Range("A7").Value = "Cj:"
    ws.Range("B7").Formula = "=SUM(B4:B6)": ws.Range("C7").Formula = "=SUM(C4:C6)"
    ws.Range("D7").Formula = "=SUM(D4:D6)": ws.Range("E7").Formula = "=SUM(E4:E6)"
    ws.Range("A8").Value = "Сумма C:": ws.Range("B8").Formula = "=SUM(B7:E7)"
    ws.Range("A9").Value = "Веса Vj:"
    ws.Range("B9").Formula = "=B7/$B$8": ws.Range("C9").Formula = "=C7/$B$8"
    ws.Range("D9").Formula = "=D7/$B$8": ws.Range("E9").Formula = "=E7/$B$8"
    ws.Range("A11").Value = "Средние Xj:"
    ws.Range("B11").Formula = "=B7/3": ws.Range("C11").Formula = "=C7/3"
    ws.Range("D11").Formula = "=D7/3": ws.Range("E11").Formula = "=E7/3"
    ws.Range("G3").Value = "Дисперсии экспертов:"
    ws.Range("G4").Formula = "=((B4-$B$11)^2+(C4-$C$11)^2+(D4-$D$11)^2+(E4-$E$11)^2)/3"
    ws.Range("G5").Formula = "=((B5-$B$11)^2+(C5-$C$11)^2+(D5-$D$11)^2+(E5-$E$11)^2)/3"
    ws.Range("G6").Formula = "=((B6-$B$11)^2+(C6-$C$11)^2+(D6-$D$11)^2+(E6-$E$11)^2)/3"
    ws.Range("A13").Value = "Дисперсии альтернатив:"
    ws.Range("B13").Formula = "=((B4-$B$11)^2+(B5-$B$11)^2+(B6-$B$11)^2)/2"
    ws.Range("C13").Formula = "=((C4-$C$11)^2+(C5-$C$11)^2+(C6-$C$11)^2)/2"
    ws.Range("D13").Formula = "=((D4-$D$11)^2+(D5-$D$11)^2+(D6-$D$11)^2)/2"
    ws.Range("E13").Formula = "=((E4-$E$11)^2+(E5-$E$11)^2+(E6-$E$11)^2)/2"
    ThisWorkbook.Sheets("Данные").Activate
    MsgBox "Готово! Создан Вариант 4."
End Sub


вар5

Sub Variant5()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Данные").Delete
    ThisWorkbook.Sheets("Саати").Delete
    ThisWorkbook.Sheets("Предпочтения").Delete
    ThisWorkbook.Sheets("Ранг").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Данные"
    ws.Range("A1").Value = "Альтернатива"
    ws.Range("B1:E1").Value = Array("A1", "A2", "A3", "A4")
    ws.Range("A2").Value = "Эксперт 1": ws.Range("A3").Value = "Эксперт 2": ws.Range("A4").Value = "Эксперт 3"
    ws.Range("B2:E2").Value = Array(2, 3, 1, 4)
    ws.Range("B3:E3").Value = Array(1, 2, 3, 4)
    ws.Range("B4:E4").Value = Array(2, 4, 1, 3)
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Саати"
    ws.Range("A1").Value = "МЕТОД СААТИ (эксперт 1)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "A1": ws.Range("A5").Value = "A2"
    ws.Range("A6").Value = "A3": ws.Range("A7").Value = "A4"
    ws.Range("B4:E4").Value = Array(1, 3, 1/3, 7)
    ws.Range("B5:E5").Value = Array(1/3, 1, 1/5, 5)
    ws.Range("B6:E6").Value = Array(3, 5, 1, 9)
    ws.Range("B7:E7").Value = Array(1/7, 1/5, 1/9, 1)
    ws.Range("F3").Value = "Цены"
    ws.Range("G3").Formula = "=(B4*C4*D4*E4)^(1/4)"
    ws.Range("G4").Formula = "=(B5*C5*D5*E5)^(1/4)"
    ws.Range("G5").Formula = "=(B6*C6*D6*E6)^(1/4)"
    ws.Range("G6").Formula = "=(B7*C7*D7*E7)^(1/4)"
    ws.Range("F7").Value = "Сумма": ws.Range("G7").Formula = "=SUM(G3:G6)"
    ws.Range("H3").Formula = "=G3/$G$7": ws.Range("H4").Formula = "=G4/$G$7"
    ws.Range("H5").Formula = "=G5/$G$7": ws.Range("H6").Formula = "=G6/$G$7"
    ws.Range("A8").Value = "Суммы столбцов"
    ws.Range("B8").Formula = "=SUM(B4:B7)": ws.Range("C8").Formula = "=SUM(C4:C7)"
    ws.Range("D8").Formula = "=SUM(D4:D7)": ws.Range("E8").Formula = "=SUM(E4:E7)"
    ws.Range("F10").Value = "λ =": ws.Range("G10").Formula = "=B8*H3 + C8*H4 + D8*H5 + E8*H6"
    ws.Range("F11").Value = "ИС =": ws.Range("G11").Formula = "=(G10-4)/3"
    ws.Range("F12").Value = "C1C =": ws.Range("G12").Value = 0.9
    ws.Range("F13").Value = "ОС =": ws.Range("G13").Formula = "=G11/G12"
    ws.Range("F14").Value = "Вывод:": ws.Range("G14").Formula = "=IF(G13<0.2,""Непротиворечиво"",""Требует уточнения"")"
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Предпочтения"
    ws.Range("A1").Value = "МЕТОД ПРЕДПОЧТЕНИЙ (3 эксперта)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "Э1": ws.Range("A5").Value = "Э2": ws.Range("A6").Value = "Э3"
    ws.Range("B4:E4").Value = Array(2, 3, 1, 4)
    ws.Range("B5:E5").Value = Array(1, 2, 3, 4)
    ws.Range("B6:E6").Value = Array(2, 4, 1, 3)
    ws.Range("A8").Value = "Преобразование B = 4 - X"
    ws.Range("B9").Value = "A1": ws.Range("C9").Value = "A2"
    ws.Range("D9").Value = "A3": ws.Range("E9").Value = "A4"
    ws.Range("A10").Value = "Э1": ws.Range("A11").Value = "Э2": ws.Range("A12").Value = "Э3"
    ws.Range("B10:E10").Formula = "=4-B4"
    ws.Range("B11:E11").Formula = "=4-B5"
    ws.Range("B12:E12").Formula = "=4-B6"
    ws.Range("A13").Value = "Cj:"
    ws.Range("B13").Formula = "=SUM(B10:B12)": ws.Range("C13").Formula = "=SUM(C10:C12)"
    ws.Range("D13").Formula = "=SUM(D10:D12)": ws.Range("E13").Formula = "=SUM(E10:E12)"
    ws.Range("F14").Value = "Сумма C =": ws.Range("G14").Formula = "=SUM(B13:E13)"
    ws.Range("A15").Value = "Веса Vj:"
    ws.Range("B15").Formula = "=B13/$G$14": ws.Range("C15").Formula = "=C13/$G$14"
    ws.Range("D15").Formula = "=D13/$G$14": ws.Range("E15").Formula = "=E13/$G$14"
    ws.Range("A17").Value = "Проверка согласованности"
    ws.Range("A18").Value = "Sj:"
    ws.Range("B18").Formula = "=SUM(B4:B6)": ws.Range("C18").Formula = "=SUM(C4:C6)"
    ws.Range("D18").Formula = "=SUM(D4:D6)": ws.Range("E18").Formula = "=SUM(E4:E6)"
    ws.Range("A19").Value = "A =": ws.Range("B19").Formula = "=3*(4+1)/2"
    ws.Range("A20").Value = "S =": ws.Range("B20").Formula = "=(B18-B19)^2+(C18-B19)^2+(D18-B19)^2+(E18-B19)^2"
    ws.Range("A21").Value = "W =": ws.Range("B21").Formula = "=12*B20/(3^2*4*(4^2-1))"
    ws.Range("A22").Value = "Вывод:": ws.Range("B22").Formula = "=IF(B21>0.5,""Согласовано"",""Требует уточнения"")"
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Ранг"
    ws.Range("A1").Value = "МЕТОД РАНГА (3 эксперта, 10 баллов)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "Э1": ws.Range("A5").Value = "Э2": ws.Range("A6").Value = "Э3"
    ws.Range("B4:E4").Value = Array(8, 5, 10, 2)
    ws.Range("B5:E5").Value = Array(10, 8, 5, 2)
    ws.Range("B6:E6").Value = Array(9, 2, 10, 5)
    ws.Range("A7").Value = "Cj:"
    ws.Range("B7").Formula = "=SUM(B4:B6)": ws.Range("C7").Formula = "=SUM(C4:C6)"
    ws.Range("D7").Formula = "=SUM(D4:D6)": ws.Range("E7").Formula = "=SUM(E4:E6)"
    ws.Range("A8").Value = "Сумма C:": ws.Range("B8").Formula = "=SUM(B7:E7)"
    ws.Range("A9").Value = "Веса Vj:"
    ws.Range("B9").Formula = "=B7/$B$8": ws.Range("C9").Formula = "=C7/$B$8"
    ws.Range("D9").Formula = "=D7/$B$8": ws.Range("E9").Formula = "=E7/$B$8"
    ws.Range("A11").Value = "Средние Xj:"
    ws.Range("B11").Formula = "=B7/3": ws.Range("C11").Formula = "=C7/3"
    ws.Range("D11").Formula = "=D7/3": ws.Range("E11").Formula = "=E7/3"
    ws.Range("G3").Value = "Дисперсии экспертов:"
    ws.Range("G4").Formula = "=((B4-$B$11)^2+(C4-$C$11)^2+(D4-$D$11)^2+(E4-$E$11)^2)/3"
    ws.Range("G5").Formula = "=((B5-$B$11)^2+(C5-$C$11)^2+(D5-$D$11)^2+(E5-$E$11)^2)/3"
    ws.Range("G6").Formula = "=((B6-$B$11)^2+(C6-$C$11)^2+(D6-$D$11)^2+(E6-$E$11)^2)/3"
    ws.Range("A13").Value = "Дисперсии альтернатив:"
    ws.Range("B13").Formula = "=((B4-$B$11)^2+(B5-$B$11)^2+(B6-$B$11)^2)/2"
    ws.Range("C13").Formula = "=((C4-$C$11)^2+(C5-$C$11)^2+(C6-$C$11)^2)/2"
    ws.Range("D13").Formula = "=((D4-$D$11)^2+(D5-$D$11)^2+(D6-$D$11)^2)/2"
    ws.Range("E13").Formula = "=((E4-$E$11)^2+(E5-$E$11)^2+(E6-$E$11)^2)/2"
    ThisWorkbook.Sheets("Данные").Activate
    MsgBox "Готово! Создан Вариант 5."
End Sub


вар 6

Sub Variant6()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Данные").Delete
    ThisWorkbook.Sheets("Саати").Delete
    ThisWorkbook.Sheets("Предпочтения").Delete
    ThisWorkbook.Sheets("Ранг").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Данные"
    ws.Range("A1").Value = "Альтернатива"
    ws.Range("B1:E1").Value = Array("A1", "A2", "A3", "A4")
    ws.Range("A2").Value = "Эксперт 1": ws.Range("A3").Value = "Эксперт 2": ws.Range("A4").Value = "Эксперт 3"
    ws.Range("B2:E2").Value = Array(3, 2, 1, 4)
    ws.Range("B3:E3").Value = Array(2, 1, 3, 4)
    ws.Range("B4:E4").Value = Array(2, 4, 1, 3)
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Саати"
    ws.Range("A1").Value = "МЕТОД СААТИ (эксперт 1)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "A1": ws.Range("A5").Value = "A2"
    ws.Range("A6").Value = "A3": ws.Range("A7").Value = "A4"
    ws.Range("B4:E4").Value = Array(1, 1/3, 1/5, 7)
    ws.Range("B5:E5").Value = Array(3, 1, 1/3, 9)
    ws.Range("B6:E6").Value = Array(5, 3, 1, 9)
    ws.Range("B7:E7").Value = Array(1/7, 1/9, 1/9, 1)
    ws.Range("F3").Value = "Цены"
    ws.Range("G3").Formula = "=(B4*C4*D4*E4)^(1/4)"
    ws.Range("G4").Formula = "=(B5*C5*D5*E5)^(1/4)"
    ws.Range("G5").Formula = "=(B6*C6*D6*E6)^(1/4)"
    ws.Range("G6").Formula = "=(B7*C7*D7*E7)^(1/4)"
    ws.Range("F7").Value = "Сумма": ws.Range("G7").Formula = "=SUM(G3:G6)"
    ws.Range("H3").Formula = "=G3/$G$7": ws.Range("H4").Formula = "=G4/$G$7"
    ws.Range("H5").Formula = "=G5/$G$7": ws.Range("H6").Formula = "=G6/$G$7"
    ws.Range("A8").Value = "Суммы столбцов"
    ws.Range("B8").Formula = "=SUM(B4:B7)": ws.Range("C8").Formula = "=SUM(C4:C7)"
    ws.Range("D8").Formula = "=SUM(D4:D7)": ws.Range("E8").Formula = "=SUM(E4:E7)"
    ws.Range("F10").Value = "λ =": ws.Range("G10").Formula = "=B8*H3 + C8*H4 + D8*H5 + E8*H6"
    ws.Range("F11").Value = "ИС =": ws.Range("G11").Formula = "=(G10-4)/3"
    ws.Range("F12").Value = "C1C =": ws.Range("G12").Value = 0.9
    ws.Range("F13").Value = "ОС =": ws.Range("G13").Formula = "=G11/G12"
    ws.Range("F14").Value = "Вывод:": ws.Range("G14").Formula = "=IF(G13<0.2,""Непротиворечиво"",""Требует уточнения"")"
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Предпочтения"
    ws.Range("A1").Value = "МЕТОД ПРЕДПОЧТЕНИЙ (3 эксперта)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "Э1": ws.Range("A5").Value = "Э2": ws.Range("A6").Value = "Э3"
    ws.Range("B4:E4").Value = Array(3, 2, 1, 4)
    ws.Range("B5:E5").Value = Array(2, 1, 3, 4)
    ws.Range("B6:E6").Value = Array(2, 4, 1, 3)
    ws.Range("A8").Value = "Преобразование B = 4 - X"
    ws.Range("B9").Value = "A1": ws.Range("C9").Value = "A2"
    ws.Range("D9").Value = "A3": ws.Range("E9").Value = "A4"
    ws.Range("A10").Value = "Э1": ws.Range("A11").Value = "Э2": ws.Range("A12").Value = "Э3"
    ws.Range("B10:E10").Formula = "=4-B4"
    ws.Range("B11:E11").Formula = "=4-B5"
    ws.Range("B12:E12").Formula = "=4-B6"
    ws.Range("A13").Value = "Cj:"
    ws.Range("B13").Formula = "=SUM(B10:B12)": ws.Range("C13").Formula = "=SUM(C10:C12)"
    ws.Range("D13").Formula = "=SUM(D10:D12)": ws.Range("E13").Formula = "=SUM(E10:E12)"
    ws.Range("F14").Value = "Сумма C =": ws.Range("G14").Formula = "=SUM(B13:E13)"
    ws.Range("A15").Value = "Веса Vj:"
    ws.Range("B15").Formula = "=B13/$G$14": ws.Range("C15").Formula = "=C13/$G$14"
    ws.Range("D15").Formula = "=D13/$G$14": ws.Range("E15").Formula = "=E13/$G$14"
    ws.Range("A17").Value = "Проверка согласованности"
    ws.Range("A18").Value = "Sj:"
    ws.Range("B18").Formula = "=SUM(B4:B6)": ws.Range("C18").Formula = "=SUM(C4:C6)"
    ws.Range("D18").Formula = "=SUM(D4:D6)": ws.Range("E18").Formula = "=SUM(E4:E6)"
    ws.Range("A19").Value = "A =": ws.Range("B19").Formula = "=3*(4+1)/2"
    ws.Range("A20").Value = "S =": ws.Range("B20").Formula = "=(B18-B19)^2+(C18-B19)^2+(D18-B19)^2+(E18-B19)^2"
    ws.Range("A21").Value = "W =": ws.Range("B21").Formula = "=12*B20/(3^2*4*(4^2-1))"
    ws.Range("A22").Value = "Вывод:": ws.Range("B22").Formula = "=IF(B21>0.5,""Согласовано"",""Требует уточнения"")"
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Ранг"
    ws.Range("A1").Value = "МЕТОД РАНГА (3 эксперта, 10 баллов)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "Э1": ws.Range("A5").Value = "Э2": ws.Range("A6").Value = "Э3"
    ws.Range("B4:E4").Value = Array(5, 8, 10, 2)
    ws.Range("B5:E5").Value = Array(8, 10, 5, 2)
    ws.Range("B6:E6").Value = Array(8, 2, 10, 5)
    ws.Range("A7").Value = "Cj:"
    ws.Range("B7").Formula = "=SUM(B4:B6)": ws.Range("C7").Formula = "=SUM(C4:C6)"
    ws.Range("D7").Formula = "=SUM(D4:D6)": ws.Range("E7").Formula = "=SUM(E4:E6)"
    ws.Range("A8").Value = "Сумма C:": ws.Range("B8").Formula = "=SUM(B7:E7)"
    ws.Range("A9").Value = "Веса Vj:"
    ws.Range("B9").Formula = "=B7/$B$8": ws.Range("C9").Formula = "=C7/$B$8"
    ws.Range("D9").Formula = "=D7/$B$8": ws.Range("E9").Formula = "=E7/$B$8"
    ws.Range("A11").Value = "Средние Xj:"
    ws.Range("B11").Formula = "=B7/3": ws.Range("C11").Formula = "=C7/3"
    ws.Range("D11").Formula = "=D7/3": ws.Range("E11").Formula = "=E7/3"
    ws.Range("G3").Value = "Дисперсии экспертов:"
    ws.Range("G4").Formula = "=((B4-$B$11)^2+(C4-$C$11)^2+(D4-$D$11)^2+(E4-$E$11)^2)/3"
    ws.Range("G5").Formula = "=((B5-$B$11)^2+(C5-$C$11)^2+(D5-$D$11)^2+(E5-$E$11)^2)/3"
    ws.Range("G6").Formula = "=((B6-$B$11)^2+(C6-$C$11)^2+(D6-$D$11)^2+(E6-$E$11)^2)/3"
    ws.Range("A13").Value = "Дисперсии альтернатив:"
    ws.Range("B13").Formula = "=((B4-$B$11)^2+(B5-$B$11)^2+(B6-$B$11)^2)/2"
    ws.Range("C13").Formula = "=((C4-$C$11)^2+(C5-$C$11)^2+(C6-$C$11)^2)/2"
    ws.Range("D13").Formula = "=((D4-$D$11)^2+(D5-$D$11)^2+(D6-$D$11)^2)/2"
    ws.Range("E13").Formula = "=((E4-$E$11)^2+(E5-$E$11)^2+(E6-$E$11)^2)/2"
    ThisWorkbook.Sheets("Данные").Activate
    MsgBox "Готово! Создан Вариант 6."
End Sub


вар7
Sub Variant7()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Данные").Delete
    ThisWorkbook.Sheets("Саати").Delete
    ThisWorkbook.Sheets("Предпочтения").Delete
    ThisWorkbook.Sheets("Ранг").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Данные"
    ws.Range("A1").Value = "Альтернатива"
    ws.Range("B1:E1").Value = Array("A1", "A2", "A3", "A4")
    ws.Range("A2").Value = "Эксперт 1": ws.Range("A3").Value = "Эксперт 2": ws.Range("A4").Value = "Эксперт 3"
    ws.Range("B2:E2").Value = Array(2, 4, 1, 3)
    ws.Range("B3:E3").Value = Array(1, 3, 2, 4)
    ws.Range("B4:E4").Value = Array(3, 4, 1, 2)
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Саати"
    ws.Range("A1").Value = "МЕТОД СААТИ (эксперт 1)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "A1": ws.Range("A5").Value = "A2"
    ws.Range("A6").Value = "A3": ws.Range("A7").Value = "A4"
    ws.Range("B4:E4").Value = Array(1, 7, 1/3, 3)
    ws.Range("B5:E5").Value = Array(1/7, 1, 1/9, 1/5)
    ws.Range("B6:E6").Value = Array(3, 9, 1, 5)
    ws.Range("B7:E7").Value = Array(1/3, 5, 1/5, 1)
    ws.Range("F3").Value = "Цены"
    ws.Range("G3").Formula = "=(B4*C4*D4*E4)^(1/4)"
    ws.Range("G4").Formula = "=(B5*C5*D5*E5)^(1/4)"
    ws.Range("G5").Formula = "=(B6*C6*D6*E6)^(1/4)"
    ws.Range("G6").Formula = "=(B7*C7*D7*E7)^(1/4)"
    ws.Range("F7").Value = "Сумма": ws.Range("G7").Formula = "=SUM(G3:G6)"
    ws.Range("H3").Formula = "=G3/$G$7": ws.Range("H4").Formula = "=G4/$G$7"
    ws.Range("H5").Formula = "=G5/$G$7": ws.Range("H6").Formula = "=G6/$G$7"
    ws.Range("A8").Value = "Суммы столбцов"
    ws.Range("B8").Formula = "=SUM(B4:B7)": ws.Range("C8").Formula = "=SUM(C4:C7)"
    ws.Range("D8").Formula = "=SUM(D4:D7)": ws.Range("E8").Formula = "=SUM(E4:E7)"
    ws.Range("F10").Value = "λ =": ws.Range("G10").Formula = "=B8*H3 + C8*H4 + D8*H5 + E8*H6"
    ws.Range("F11").Value = "ИС =": ws.Range("G11").Formula = "=(G10-4)/3"
    ws.Range("F12").Value = "C1C =": ws.Range("G12").Value = 0.9
    ws.Range("F13").Value = "ОС =": ws.Range("G13").Formula = "=G11/G12"
    ws.Range("F14").Value = "Вывод:": ws.Range("G14").Formula = "=IF(G13<0.2,""Непротиворечиво"",""Требует уточнения"")"
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Предпочтения"
    ws.Range("A1").Value = "МЕТОД ПРЕДПОЧТЕНИЙ (3 эксперта)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "Э1": ws.Range("A5").Value = "Э2": ws.Range("A6").Value = "Э3"
    ws.Range("B4:E4").Value = Array(2, 4, 1, 3)
    ws.Range("B5:E5").Value = Array(1, 3, 2, 4)
    ws.Range("B6:E6").Value = Array(3, 4, 1, 2)
    ws.Range("A8").Value = "Преобразование B = 4 - X"
    ws.Range("B9").Value = "A1": ws.Range("C9").Value = "A2"
    ws.Range("D9").Value = "A3": ws.Range("E9").Value = "A4"
    ws.Range("A10").Value = "Э1": ws.Range("A11").Value = "Э2": ws.Range("A12").Value = "Э3"
    ws.Range("B10:E10").Formula = "=4-B4"
    ws.Range("B11:E11").Formula = "=4-B5"
    ws.Range("B12:E12").Formula = "=4-B6"
    ws.Range("A13").Value = "Cj:"
    ws.Range("B13").Formula = "=SUM(B10:B12)": ws.Range("C13").Formula = "=SUM(C10:C12)"
    ws.Range("D13").Formula = "=SUM(D10:D12)": ws.Range("E13").Formula = "=SUM(E10:E12)"
    ws.Range("F14").Value = "Сумма C =": ws.Range("G14").Formula = "=SUM(B13:E13)"
    ws.Range("A15").Value = "Веса Vj:"
    ws.Range("B15").Formula = "=B13/$G$14": ws.Range("C15").Formula = "=C13/$G$14"
    ws.Range("D15").Formula = "=D13/$G$14": ws.Range("E15").Formula = "=E13/$G$14"
    ws.Range("A17").Value = "Проверка согласованности"
    ws.Range("A18").Value = "Sj:"
    ws.Range("B18").Formula = "=SUM(B4:B6)": ws.Range("C18").Formula = "=SUM(C4:C6)"
    ws.Range("D18").Formula = "=SUM(D4:D6)": ws.Range("E18").Formula = "=SUM(E4:E6)"
    ws.Range("A19").Value = "A =": ws.Range("B19").Formula = "=3*(4+1)/2"
    ws.Range("A20").Value = "S =": ws.Range("B20").Formula = "=(B18-B19)^2+(C18-B19)^2+(D18-B19)^2+(E18-B19)^2"
    ws.Range("A21").Value = "W =": ws.Range("B21").Formula = "=12*B20/(3^2*4*(4^2-1))"
    ws.Range("A22").Value = "Вывод:": ws.Range("B22").Formula = "=IF(B21>0.5,""Согласовано"",""Требует уточнения"")"
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Ранг"
    ws.Range("A1").Value = "МЕТОД РАНГА (3 эксперта, 10 баллов)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "Э1": ws.Range("A5").Value = "Э2": ws.Range("A6").Value = "Э3"
    ws.Range("B4:E4").Value = Array(8, 2, 10, 5)
    ws.Range("B5:E5").Value = Array(10, 5, 8, 2)
    ws.Range("B6:E6").Value = Array(5, 2, 10, 8)
    ws.Range("A7").Value = "Cj:"
    ws.Range("B7").Formula = "=SUM(B4:B6)": ws.Range("C7").Formula = "=SUM(C4:C6)"
    ws.Range("D7").Formula = "=SUM(D4:D6)": ws.Range("E7").Formula = "=SUM(E4:E6)"
    ws.Range("A8").Value = "Сумма C:": ws.Range("B8").Formula = "=SUM(B7:E7)"
    ws.Range("A9").Value = "Веса Vj:"
    ws.Range("B9").Formula = "=B7/$B$8": ws.Range("C9").Formula = "=C7/$B$8"
    ws.Range("D9").Formula = "=D7/$B$8": ws.Range("E9").Formula = "=E7/$B$8"
    ws.Range("A11").Value = "Средние Xj:"
    ws.Range("B11").Formula = "=B7/3": ws.Range("C11").Formula = "=C7/3"
    ws.Range("D11").Formula = "=D7/3": ws.Range("E11").Formula = "=E7/3"
    ws.Range("G3").Value = "Дисперсии экспертов:"
    ws.Range("G4").Formula = "=((B4-$B$11)^2+(C4-$C$11)^2+(D4-$D$11)^2+(E4-$E$11)^2)/3"
    ws.Range("G5").Formula = "=((B5-$B$11)^2+(C5-$C$11)^2+(D5-$D$11)^2+(E5-$E$11)^2)/3"
    ws.Range("G6").Formula = "=((B6-$B$11)^2+(C6-$C$11)^2+(D6-$D$11)^2+(E6-$E$11)^2)/3"
    ws.Range("A13").Value = "Дисперсии альтернатив:"
    ws.Range("B13").Formula = "=((B4-$B$11)^2+(B5-$B$11)^2+(B6-$B$11)^2)/2"
    ws.Range("C13").Formula = "=((C4-$C$11)^2+(C5-$C$11)^2+(C6-$C$11)^2)/2"
    ws.Range("D13").Formula = "=((D4-$D$11)^2+(D5-$D$11)^2+(D6-$D$11)^2)/2"
    ws.Range("E13").Formula = "=((E4-$E$11)^2+(E5-$E$11)^2+(E6-$E$11)^2)/2"
    ThisWorkbook.Sheets("Данные").Activate
    MsgBox "Готово! Создан Вариант 7."
End Sub



вар8
Sub Variant8()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Данные").Delete
    ThisWorkbook.Sheets("Саати").Delete
    ThisWorkbook.Sheets("Предпочтения").Delete
    ThisWorkbook.Sheets("Ранг").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Данные"
    ws.Range("A1").Value = "Альтернатива"
    ws.Range("B1:E1").Value = Array("A1", "A2", "A3", "A4")
    ws.Range("A2").Value = "Эксперт 1": ws.Range("A3").Value = "Эксперт 2": ws.Range("A4").Value = "Эксперт 3"
    ws.Range("B2:E2").Value = Array(3, 4, 2, 1)
    ws.Range("B3:E3").Value = Array(1, 4, 2, 3)
    ws.Range("B4:E4").Value = Array(4, 3, 1, 2)
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Саати"
    ws.Range("A1").Value = "МЕТОД СААТИ (эксперт 1)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "A1": ws.Range("A5").Value = "A2"
    ws.Range("A6").Value = "A3": ws.Range("A7").Value = "A4"
    ws.Range("B4:E4").Value = Array(1, 3, 1/3, 1/5)
    ws.Range("B5:E5").Value = Array(1/3, 1, 1/5, 1/7)
    ws.Range("B6:E6").Value = Array(3, 5, 1, 1/3)
    ws.Range("B7:E7").Value = Array(5, 7, 3, 1)
    ws.Range("F3").Value = "Цены"
    ws.Range("G3").Formula = "=(B4*C4*D4*E4)^(1/4)"
    ws.Range("G4").Formula = "=(B5*C5*D5*E5)^(1/4)"
    ws.Range("G5").Formula = "=(B6*C6*D6*E6)^(1/4)"
    ws.Range("G6").Formula = "=(B7*C7*D7*E7)^(1/4)"
    ws.Range("F7").Value = "Сумма": ws.Range("G7").Formula = "=SUM(G3:G6)"
    ws.Range("H3").Formula = "=G3/$G$7": ws.Range("H4").Formula = "=G4/$G$7"
    ws.Range("H5").Formula = "=G5/$G$7": ws.Range("H6").Formula = "=G6/$G$7"
    ws.Range("A8").Value = "Суммы столбцов"
    ws.Range("B8").Formula = "=SUM(B4:B7)": ws.Range("C8").Formula = "=SUM(C4:C7)"
    ws.Range("D8").Formula = "=SUM(D4:D7)": ws.Range("E8").Formula = "=SUM(E4:E7)"
    ws.Range("F10").Value = "λ =": ws.Range("G10").Formula = "=B8*H3 + C8*H4 + D8*H5 + E8*H6"
    ws.Range("F11").Value = "ИС =": ws.Range("G11").Formula = "=(G10-4)/3"
    ws.Range("F12").Value = "C1C =": ws.Range("G12").Value = 0.9
    ws.Range("F13").Value = "ОС =": ws.Range("G13").Formula = "=G11/G12"
    ws.Range("F14").Value = "Вывод:": ws.Range("G14").Formula = "=IF(G13<0.2,""Непротиворечиво"",""Требует уточнения"")"
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Предпочтения"
    ws.Range("A1").Value = "МЕТОД ПРЕДПОЧТЕНИЙ (3 эксперта)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "Э1": ws.Range("A5").Value = "Э2": ws.Range("A6").Value = "Э3"
    ws.Range("B4:E4").Value = Array(3, 4, 2, 1)
    ws.Range("B5:E5").Value = Array(1, 4, 2, 3)
    ws.Range("B6:E6").Value = Array(4, 3, 1, 2)
    ws.Range("A8").Value = "Преобразование B = 4 - X"
    ws.Range("B9").Value = "A1": ws.Range("C9").Value = "A2"
    ws.Range("D9").Value = "A3": ws.Range("E9").Value = "A4"
    ws.Range("A10").Value = "Э1": ws.Range("A11").Value = "Э2": ws.Range("A12").Value = "Э3"
    ws.Range("B10:E10").Formula = "=4-B4"
    ws.Range("B11:E11").Formula = "=4-B5"
    ws.Range("B12:E12").Formula = "=4-B6"
    ws.Range("A13").Value = "Cj:"
    ws.Range("B13").Formula = "=SUM(B10:B12)": ws.Range("C13").Formula = "=SUM(C10:C12)"
    ws.Range("D13").Formula = "=SUM(D10:D12)": ws.Range("E13").Formula = "=SUM(E10:E12)"
    ws.Range("F14").Value = "Сумма C =": ws.Range("G14").Formula = "=SUM(B13:E13)"
    ws.Range("A15").Value = "Веса Vj:"
    ws.Range("B15").Formula = "=B13/$G$14": ws.Range("C15").Formula = "=C13/$G$14"
    ws.Range("D15").Formula = "=D13/$G$14": ws.Range("E15").Formula = "=E13/$G$14"
    ws.Range("A17").Value = "Проверка согласованности"
    ws.Range("A18").Value = "Sj:"
    ws.Range("B18").Formula = "=SUM(B4:B6)": ws.Range("C18").Formula = "=SUM(C4:C6)"
    ws.Range("D18").Formula = "=SUM(D4:D6)": ws.Range("E18").Formula = "=SUM(E4:E6)"
    ws.Range("A19").Value = "A =": ws.Range("B19").Formula = "=3*(4+1)/2"
    ws.Range("A20").Value = "S =": ws.Range("B20").Formula = "=(B18-B19)^2+(C18-B19)^2+(D18-B19)^2+(E18-B19)^2"
    ws.Range("A21").Value = "W =": ws.Range("B21").Formula = "=12*B20/(3^2*4*(4^2-1))"
    ws.Range("A22").Value = "Вывод:": ws.Range("B22").Formula = "=IF(B21>0.5,""Согласовано"",""Требует уточнения"")"
    
    Set ws = ThisWorkbook.Sheets.Add: ws.Name = "Ранг"
    ws.Range("A1").Value = "МЕТОД РАНГА (3 эксперта, 10 баллов)"
    ws.Range("B3").Value = "A1": ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3": ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "Э1": ws.Range("A5").Value = "Э2": ws.Range("A6").Value = "Э3"
    ws.Range("B4:E4").Value = Array(5, 2, 8, 10)
    ws.Range("B5:E5").Value = Array(10, 2, 8, 5)
    ws.Range("B6:E6").Value = Array(2, 5, 10, 8)
    ws.Range("A7").Value = "Cj:"
    ws.Range("B7").Formula = "=SUM(B4:B6)": ws.Range("C7").Formula = "=SUM(C4:C6)"
    ws.Range("D7").Formula = "=SUM(D4:D6)": ws.Range("E7").Formula = "=SUM(E4:E6)"
    ws.Range("A8").Value = "Сумма C:": ws.Range("B8").Formula = "=SUM(B7:E7)"
    ws.Range("A9").Value = "Веса Vj:"
    ws.Range("B9").Formula = "=B7/$B$8": ws.Range("C9").Formula = "=C7/$B$8"
    ws.Range("D9").Formula = "=D7/$B$8": ws.Range("E9").Formula = "=E7/$B$8"
    ws.Range("A11").Value = "Средние Xj:"
    ws.Range("B11").Formula = "=B7/3": ws.Range("C11").Formula = "=C7/3"
    ws.Range("D11").Formula = "=D7/3": ws.Range("E11").Formula = "=E7/3"
    ws.Range("G3").Value = "Дисперсии экспертов:"
    ws.Range("G4").Formula = "=((B4-$B$11)^2+(C4-$C$11)^2+(D4-$D$11)^2+(E4-$E$11)^2)/3"
    ws.Range("G5").Formula = "=((B5-$B$11)^2+(C5-$C$11)^2+(D5-$D$11)^2+(E5-$E$11)^2)/3"
    ws.Range("G6").Formula = "=((B6-$B$11)^2+(C6-$C$11)^2+(D6-$D$11)^2+(E6-$E$11)^2)/3"
    ws.Range("A13").Value = "Дисперсии альтернатив:"
    ws.Range("B13").Formula = "=((B4-$B$11)^2+(B5-$B$11)^2+(B6-$B$11)^2)/2"
    ws.Range("C13").Formula = "=((C4-$C$11)^2+(C5-$C$11)^2+(C6-$C$11)^2)/2"
    ws.Range("D13").Formula = "=((D4-$D$11)^2+(D5-$D$11)^2+(D6-$D$11)^2)/2"
    ws.Range("E13").Formula = "=((E4-$E$11)^2+(E5-$E$11)^2+(E6-$E$11)^2)/2"
    ThisWorkbook.Sheets("Данные").Activate
    MsgBox "Готово! Создан Вариант 8."
End Sub






    Sub SozdatLab4()
    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Data").Delete
    ThisWorkbook.Sheets("Saaty").Delete
    ThisWorkbook.Sheets("Preference").Delete
    ThisWorkbook.Sheets("Rank").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' ===== ЛИСТ 1: DATA =====
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Data"
    ws.Range("A1").Value = "Alternative"
    ws.Range("B1:E1").Value = Array("A1", "A2", "A3", "A4")
    ws.Range("A2").Value = "Expert1"
    ws.Range("A3").Value = "Expert2"
    ws.Range("A4").Value = "Expert3"
    ws.Range("B2:E2").Value = Array(2, 1, 4, 3)
    ws.Range("B3:E3").Value = Array(1, 2, 4, 3)
    ws.Range("B4:E4").Value = Array(1, 2, 3, 4)
    
    ' ===== ЛИСТ 2: SAATY =====
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Saaty"
    ws.Range("A1").Value = "SAATY METHOD (expert 1)"
    ws.Range("B3").Value = "A1"
    ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3"
    ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "A1"
    ws.Range("A5").Value = "A2"
    ws.Range("A6").Value = "A3"
    ws.Range("A7").Value = "A4"
    ws.Range("B4:E4").Value = Array(1, 1 / 3, 7, 3)
    ws.Range("B5:E5").Value = Array(3, 1, 9, 5)
    ws.Range("B6:E6").Value = Array(1 / 7, 1 / 9, 1, 1 / 3)
    ws.Range("B7:E7").Value = Array(1 / 3, 1 / 5, 3, 1)
    
    ws.Range("F3").Value = "Prices"
    ws.Range("G3").Formula = "=(B4*C4*D4*E4)^(1/4)"
    ws.Range("G4").Formula = "=(B5*C5*D5*E5)^(1/4)"
    ws.Range("G5").Formula = "=(B6*C6*D6*E6)^(1/4)"
    ws.Range("G6").Formula = "=(B7*C7*D7*E7)^(1/4)"
    ws.Range("F7").Value = "Sum"
    ws.Range("G7").Formula = "=SUM(G3:G6)"
    
    ws.Range("F2").Value = "Weights"
    ws.Range("H3").Formula = "=G3/$G$7"
    ws.Range("H4").Formula = "=G4/$G$7"
    ws.Range("H5").Formula = "=G5/$G$7"
    ws.Range("H6").Formula = "=G6/$G$7"
    
    ws.Range("A8").Value = "Col sums"
    ws.Range("B8").Formula = "=SUM(B4:B7)"
    ws.Range("C8").Formula = "=SUM(C4:C7)"
    ws.Range("D8").Formula = "=SUM(D4:D7)"
    ws.Range("E8").Formula = "=SUM(E4:E7)"
    
    ws.Range("F10").Value = "Lambda ="
    ws.Range("G10").Formula = "=SUMPRODUCT(B8:E8,H3:H6)"
    ws.Range("F11").Value = "IS ="
    ws.Range("G11").Formula = "=(G10-4)/3"
    ws.Range("F12").Value = "C1C ="
    ws.Range("G12").Value = 0.9
    ws.Range("F13").Value = "OS ="
    ws.Range("G13").Formula = "=G11/G12"
    ws.Range("F14").Value = "Result:"
    ws.Range("G14").Formula = "=IF(G13<0.2,""OK"",""NeedFix"")"
    
    ' ===== ЛИСТ 3: PREFERENCE =====
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Preference"
    ws.Range("A1").Value = "PREFERENCE METHOD (3 experts)"
    ws.Range("B3").Value = "A1"
    ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3"
    ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "E1"
    ws.Range("A5").Value = "E2"
    ws.Range("A6").Value = "E3"
    ws.Range("B4:E4").Value = Array(2, 1, 4, 3)
    ws.Range("B5:E5").Value = Array(1, 2, 4, 3)
    ws.Range("B6:E6").Value = Array(1, 2, 3, 4)
    
    ws.Range("A8").Value = "Transform B = 4 - X"
    ws.Range("B9").Value = "A1"
    ws.Range("C9").Value = "A2"
    ws.Range("D9").Value = "A3"
    ws.Range("E9").Value = "A4"
    ws.Range("A10").Value = "E1"
    ws.Range("A11").Value = "E2"
    ws.Range("A12").Value = "E3"
    ws.Range("B10:E10").Formula = "=4-B4"
    ws.Range("B11:E11").Formula = "=4-B5"
    ws.Range("B12:E12").Formula = "=4-B6"
    
    ws.Range("A13").Value = "Cj:"
    ws.Range("B13").Formula = "=SUM(B10:B12)"
    ws.Range("C13").Formula = "=SUM(C10:C12)"
    ws.Range("D13").Formula = "=SUM(D10:D12)"
    ws.Range("E13").Formula = "=SUM(E10:E12)"
    
    ws.Range("F14").Value = "Sum C ="
    ws.Range("G14").Formula = "=SUM(B13:E13)"
    
    ws.Range("A15").Value = "Weights Vj:"
    ws.Range("B15").Formula = "=B13/$G$14"
    ws.Range("C15").Formula = "=C13/$G$14"
    ws.Range("D15").Formula = "=D13/$G$14"
    ws.Range("E15").Formula = "=E13/$G$14"
    
    ws.Range("A17").Value = "Concordance check"
    ws.Range("A18").Value = "Sj:"
    ws.Range("B18").Formula = "=SUM(B4:B6)"
    ws.Range("C18").Formula = "=SUM(C4:C6)"
    ws.Range("D18").Formula = "=SUM(D4:D6)"
    ws.Range("E18").Formula = "=SUM(E4:E6)"
    
    ws.Range("A19").Value = "A ="
    ws.Range("B19").Formula = "=3*(4+1)/2"
    ws.Range("A20").Value = "S ="
    ws.Range("B20").Formula = "=(B18-B19)^2+(C18-B19)^2+(D18-B19)^2+(E18-B19)^2"
    ws.Range("A21").Value = "W ="
    ws.Range("B21").Formula = "=12*B20/(3^2*4*(4^2-1))"
    ws.Range("A22").Value = "Result:"
    ws.Range("B22").Formula = "=IF(B21>0.5,""OK"",""NeedFix"")"
    
    ' ===== ЛИСТ 4: RANK =====
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Rank"
    ws.Range("A1").Value = "RANK METHOD (3 experts, 10 points)"
    ws.Range("B3").Value = "A1"
    ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3"
    ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "E1"
    ws.Range("A5").Value = "E2"
    ws.Range("A6").Value = "E3"
    ws.Range("B4:E4").Value = Array(9, 10, 2, 8)
    ws.Range("B5:E5").Value = Array(10, 9, 2, 8)
    ws.Range("B6:E6").Value = Array(10, 9, 8, 2)
    
    ws.Range("A7").Value = "Cj:"
    ws.Range("B7").Formula = "=SUM(B4:B6)"
    ws.Range("C7").Formula = "=SUM(C4:C6)"
    ws.Range("D7").Formula = "=SUM(D4:D6)"
    ws.Range("E7").Formula = "=SUM(E4:E6)"
    
    ws.Range("A8").Value = "Sum C:"
    ws.Range("B8").Formula = "=SUM(B7:E7)"
    
    ws.Range("A9").Value = "Weights Vj:"
    ws.Range("B9").Formula = "=B7/$B$8"
    ws.Range("C9").Formula = "=C7/$B$8"
    ws.Range("D9").Formula = "=D7/$B$8"
    ws.Range("E9").Formula = "=E7/$B$8"
    
    ws.Range("A11").Value = "Means Xj:"
    ws.Range("B11").Formula = "=B7/3"
    ws.Range("C11").Formula = "=C7/3"
    ws.Range("D11").Formula = "=D7/3"
    ws.Range("E11").Formula = "=E7/3"
    
    ws.Range("G3").Value = "Disp experts:"
    ws.Range("G4").Formula = "=((B4-$B$11)^2+(C4-$C$11)^2+(D4-$D$11)^2+(E4-$E$11)^2)/3"
    ws.Range("G5").Formula = "=((B5-$B$11)^2+(C5-$C$11)^2+(D5-$D$11)^2+(E5-$E$11)^2)/3"
    ws.Range("G6").Formula = "=((B6-$B$11)^2+(C6-$C$11)^2+(D6-$D$11)^2+(E6-$E$11)^2)/3"
    
    ws.Range("A13").Value = "Disp alternatives:"
    ws.Range("B13").Formula = "=((B4-$B$11)^2+(B5-$B$11)^2+(B6-$B$11)^2)/2"
    ws.Range("C13").Formula = "=((C4-$C$11)^2+(C5-$C$11)^2+(C6-$C$11)^2)/2"
    ws.Range("D13").Formula = "=((D4-$D$11)^2+(D5-$D$11)^2+(D6-$D$11)^2)/2"
    ws.Range("E13").Formula = "=((E4-$E$11)^2+(E5-$E$11)^2+(E6-$E$11)^2)/2"
    
    ThisWorkbook.Sheets("Data").Activate
    MsgBox "Done! 4 sheets created."
End Sub




    Sub SozdatLab4_Variant8()
    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Data").Delete
    ThisWorkbook.Sheets("Saaty").Delete
    ThisWorkbook.Sheets("Preference").Delete
    ThisWorkbook.Sheets("Rank").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' ===== ЛИСТ 1: DATA =====
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Data"
    ws.Range("A1").Value = "Alternative"
    ws.Range("B1:E1").Value = Array("A1", "A2", "A3", "A4")
    ws.Range("A2").Value = "Expert1"
    ws.Range("A3").Value = "Expert2"
    ws.Range("A4").Value = "Expert3"
    ws.Range("B2:E2").Value = Array(3, 4, 2, 1)
    ws.Range("B3:E3").Value = Array(1, 4, 2, 3)
    ws.Range("B4:E4").Value = Array(4, 3, 1, 2)
    
    ' ===== ЛИСТ 2: SAATY (эксперт 1: A4 > A3 > A1 > A2) =====
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Saaty"
    ws.Range("A1").Value = "SAATY METHOD (expert 1: A4>A3>A1>A2)"
    ws.Range("B3").Value = "A1"
    ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3"
    ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "A1"
    ws.Range("A5").Value = "A2"
    ws.Range("A6").Value = "A3"
    ws.Range("A7").Value = "A4"
    
    ' Матрица: A4>A3>A1>A2
    ' A4 vs A3 = 3 (немного), A4 vs A1 = 7 (значительно),
    ' A4 vs A2 = 9 (явно), A3 vs A1 = 5, A3 vs A2 = 7, A1 vs A2 = 3
    ws.Range("B4:E4").Value = Array(1, 1 / 3, 1 / 5, 1 / 7)   ' A1 vs ...
    ws.Range("B5:E5").Value = Array(3, 1, 1 / 7, 1 / 9)        ' A2 vs ...
    ws.Range("B6:E6").Value = Array(5, 7, 1, 1 / 3)            ' A3 vs ...
    ws.Range("B7:E7").Value = Array(7, 9, 3, 1)                ' A4 vs ...
    
    ' Цены (среднее геометрическое)
    ws.Range("F3").Value = "Prices"
    ws.Range("G3").Formula = "=(B4*C4*D4*E4)^(1/4)"
    ws.Range("G4").Formula = "=(B5*C5*D5*E5)^(1/4)"
    ws.Range("G5").Formula = "=(B6*C6*D6*E6)^(1/4)"
    ws.Range("G6").Formula = "=(B7*C7*D7*E7)^(1/4)"
    ws.Range("F7").Value = "Sum"
    ws.Range("G7").Formula = "=SUM(G3:G6)"
    
    ' Веса
    ws.Range("F2").Value = "Weights"
    ws.Range("H3").Formula = "=G3/$G$7"
    ws.Range("H4").Formula = "=G4/$G$7"
    ws.Range("H5").Formula = "=G5/$G$7"
    ws.Range("H6").Formula = "=G6/$G$7"
    
    ' Проверка
    ws.Range("A8").Value = "Col sums"
    ws.Range("B8").Formula = "=SUM(B4:B7)"
    ws.Range("C8").Formula = "=SUM(C4:C7)"
    ws.Range("D8").Formula = "=SUM(D4:D7)"
    ws.Range("E8").Formula = "=SUM(E4:E7)"
    
    ws.Range("F10").Value = "Lambda ="
    ws.Range("G10").Formula = "=SUMPRODUCT(B8:E8,H3:H6)"
    ws.Range("F11").Value = "IS ="
    ws.Range("G11").Formula = "=(G10-4)/3"
    ws.Range("F12").Value = "C1C ="
    ws.Range("G12").Value = 0.9
    ws.Range("F13").Value = "OS ="
    ws.Range("G13").Formula = "=G11/G12"
    ws.Range("F14").Value = "Result:"
    ws.Range("G14").Formula = "=IF(G13<0.2,""OK"",""NeedFix"")"
    
    ' ===== ЛИСТ 3: PREFERENCE =====
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Preference"
    ws.Range("A1").Value = "PREFERENCE METHOD (3 experts)"
    ws.Range("B3").Value = "A1"
    ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3"
    ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "E1"
    ws.Range("A5").Value = "E2"
    ws.Range("A6").Value = "E3"
    ws.Range("B4:E4").Value = Array(3, 4, 2, 1)
    ws.Range("B5:E5").Value = Array(1, 4, 2, 3)
    ws.Range("B6:E6").Value = Array(4, 3, 1, 2)
    
    ws.Range("A8").Value = "Transform B = 4 - X"
    ws.Range("B9").Value = "A1"
    ws.Range("C9").Value = "A2"
    ws.Range("D9").Value = "A3"
    ws.Range("E9").Value = "A4"
    ws.Range("A10").Value = "E1"
    ws.Range("A11").Value = "E2"
    ws.Range("A12").Value = "E3"
    ws.Range("B10:E10").Formula = "=4-B4"
    ws.Range("B11:E11").Formula = "=4-B5"
    ws.Range("B12:E12").Formula = "=4-B6"
    
    ws.Range("A13").Value = "Cj:"
    ws.Range("B13").Formula = "=SUM(B10:B12)"
    ws.Range("C13").Formula = "=SUM(C10:C12)"
    ws.Range("D13").Formula = "=SUM(D10:D12)"
    ws.Range("E13").Formula = "=SUM(E10:E12)"
    
    ws.Range("F14").Value = "Sum C ="
    ws.Range("G14").Formula = "=SUM(B13:E13)"
    
    ws.Range("A15").Value = "Weights Vj:"
    ws.Range("B15").Formula = "=B13/$G$14"
    ws.Range("C15").Formula = "=C13/$G$14"
    ws.Range("D15").Formula = "=D13/$G$14"
    ws.Range("E15").Formula = "=E13/$G$14"
    
    ws.Range("A17").Value = "Concordance check"
    ws.Range("A18").Value = "Sj:"
    ws.Range("B18").Formula = "=SUM(B4:B6)"
    ws.Range("C18").Formula = "=SUM(C4:C6)"
    ws.Range("D18").Formula = "=SUM(D4:D6)"
    ws.Range("E18").Formula = "=SUM(E4:E6)"
    
    ws.Range("A19").Value = "A ="
    ws.Range("B19").Formula = "=3*(4+1)/2"
    ws.Range("A20").Value = "S ="
    ws.Range("B20").Formula = "=(B18-B19)^2+(C18-B19)^2+(D18-B19)^2+(E18-B19)^2"
    ws.Range("A21").Value = "W ="
    ws.Range("B21").Formula = "=12*B20/(3^2*4*(4^2-1))"
    ws.Range("A22").Value = "Result:"
    ws.Range("B22").Formula = "=IF(B21>0.5,""OK"",""NeedFix"")"
    
    ' ===== ЛИСТ 4: RANK (10-балльная шкала) =====
    ' Логика: 1-е место = 10, 2-е = 8, 3-е = 5, 4-е = 2
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Rank"
    ws.Range("A1").Value = "RANK METHOD (3 experts, 10 points)"
    ws.Range("B3").Value = "A1"
    ws.Range("C3").Value = "A2"
    ws.Range("D3").Value = "A3"
    ws.Range("E3").Value = "A4"
    ws.Range("A4").Value = "E1"
    ws.Range("A5").Value = "E2"
    ws.Range("A6").Value = "E3"
    ' Э1: A4=10, A3=8, A1=5, A2=2
    ' Э2: A1=10, A3=8, A4=5, A2=2
    ' Э3: A3=10, A4=8, A2=5, A1=2
    ws.Range("B4:E4").Value = Array(5, 2, 8, 10)
    ws.Range("B5:E5").Value = Array(10, 2, 8, 5)
    ws.Range("B6:E6").Value = Array(2, 5, 10, 8)
    
    ws.Range("A7").Value = "Cj:"
    ws.Range("B7").Formula = "=SUM(B4:B6)"
    ws.Range("C7").Formula = "=SUM(C4:C6)"
    ws.Range("D7").Formula = "=SUM(D4:D6)"
    ws.Range("E7").Formula = "=SUM(E4:E6)"
    
    ws.Range("A8").Value = "Sum C:"
    ws.Range("B8").Formula = "=SUM(B7:E7)"
    
    ws.Range("A9").Value = "Weights Vj:"
    ws.Range("B9").Formula = "=B7/$B$8"
    ws.Range("C9").Formula = "=C7/$B$8"
    ws.Range("D9").Formula = "=D7/$B$8"
    ws.Range("E9").Formula = "=E7/$B$8"
    
    ws.Range("A11").Value = "Means Xj:"
    ws.Range("B11").Formula = "=B7/3"
    ws.Range("C11").Formula = "=C7/3"
    ws.Range("D11").Formula = "=D7/3"
    ws.Range("E11").Formula = "=E7/3"
    
    ws.Range("G3").Value = "Disp experts:"
    ws.Range("G4").Formula = "=((B4-$B$11)^2+(C4-$C$11)^2+(D4-$D$11)^2+(E4-$E$11)^2)/3"
    ws.Range("G5").Formula = "=((B5-$B$11)^2+(C5-$C$11)^2+(D5-$D$11)^2+(E5-$E$11)^2)/3"
    ws.Range("G6").Formula = "=((B6-$B$11)^2+(C6-$C$11)^2+(D6-$D$11)^2+(E6-$E$11)^2)/3"
    
    ws.Range("A13").Value = "Disp alternatives:"
    ws.Range("B13").Formula = "=((B4-$B$11)^2+(B5-$B$11)^2+(B6-$B$11)^2)/2"
    ws.Range("C13").Formula = "=((C4-$C$11)^2+(C5-$C$11)^2+(C6-$C$11)^2)/2"
    ws.Range("D13").Formula = "=((D4-$D$11)^2+(D5-$D$11)^2+(D6-$D$11)^2)/2"
    ws.Range("E13").Formula = "=((E4-$E$11)^2+(E5-$E$11)^2+(E6-$E$11)^2)/2"
    
    ThisWorkbook.Sheets("Data").Activate
    MsgBox "Done! Variant 8 sheets created."
End Sub
