Sub SozdatLab4
    Dim oDoc As Object
    Dim oSheets As Object
    Dim oSheet As Object
    
    oDoc = ThisComponent
    oSheets = oDoc.Sheets
    
    ' Удаляем старые листы, если есть
    Dim names(3) As String
    names(0) = "Данные"
    names(1) = "Саати"
    names(2) = "Предпочтения"
    names(3) = "Ранг"
    
    Dim i As Integer
    For i = 0 To 3
        If oSheets.hasByName(names(i)) Then
            oSheets.removeByName(names(i))
        End If
    Next i
    
    ' ===== ЛИСТ 1: ДАННЫЕ =====
    oSheets.insertNewByName("Данные", 0)
    oSheet = oSheets.getByName("Данные")
    oSheet.getCellByPosition(0,0).setString("Альтернатива")
    oSheet.getCellByPosition(1,0).setString("A1")
    oSheet.getCellByPosition(2,0).setString("A2")
    oSheet.getCellByPosition(3,0).setString("A3")
    oSheet.getCellByPosition(4,0).setString("A4")
    oSheet.getCellByPosition(0,1).setString("Эксперт 1")
    oSheet.getCellByPosition(0,2).setString("Эксперт 2")
    oSheet.getCellByPosition(0,3).setString("Эксперт 3")
    ' Места (1 = лучший)
    oSheet.getCellByPosition(1,1).setValue(2)
    oSheet.getCellByPosition(2,1).setValue(1)
    oSheet.getCellByPosition(3,1).setValue(4)
    oSheet.getCellByPosition(4,1).setValue(3)
    oSheet.getCellByPosition(1,2).setValue(1)
    oSheet.getCellByPosition(2,2).setValue(2)
    oSheet.getCellByPosition(3,2).setValue(4)
    oSheet.getCellByPosition(4,2).setValue(3)
    oSheet.getCellByPosition(1,3).setValue(1)
    oSheet.getCellByPosition(2,3).setValue(2)
    oSheet.getCellByPosition(3,3).setValue(3)
    oSheet.getCellByPosition(4,3).setValue(4)
    
    ' ===== ЛИСТ 2: СААТИ =====
    oSheets.insertNewByName("Саати", 1)
    oSheet = oSheets.getByName("Саати")
    oSheet.getCellByPosition(0,0).setString("МЕТОД СААТИ (эксперт 1)")
    oSheet.getCellByPosition(0,1).setString("Матрица парных сравнений")
    ' Заголовки
    oSheet.getCellByPosition(1,2).setString("A1")
    oSheet.getCellByPosition(2,2).setString("A2")
    oSheet.getCellByPosition(3,2).setString("A3")
    oSheet.getCellByPosition(4,2).setString("A4")
    oSheet.getCellByPosition(0,3).setString("A1")
    oSheet.getCellByPosition(0,4).setString("A2")
    oSheet.getCellByPosition(0,5).setString("A3")
    oSheet.getCellByPosition(0,6).setString("A4")
    ' Матрица (мнение 1-го эксперта: A2 > A1 > A4 > A3)
    oSheet.getCellByPosition(1,3).setValue(1)
    oSheet.getCellByPosition(2,3).setValue(1/3)
    oSheet.getCellByPosition(3,3).setValue(7)
    oSheet.getCellByPosition(4,3).setValue(3)
    
    oSheet.getCellByPosition(1,4).setValue(3)
    oSheet.getCellByPosition(2,4).setValue(1)
    oSheet.getCellByPosition(3,4).setValue(9)
    oSheet.getCellByPosition(4,4).setValue(5)
    
    oSheet.getCellByPosition(1,5).setValue(1/7)
    oSheet.getCellByPosition(2,5).setValue(1/9)
    oSheet.getCellByPosition(3,5).setValue(1)
    oSheet.getCellByPosition(4,5).setValue(1/3)
    
    oSheet.getCellByPosition(1,6).setValue(1/3)
    oSheet.getCellByPosition(2,6).setValue(1/5)
    oSheet.getCellByPosition(3,6).setValue(3)
    oSheet.getCellByPosition(4,6).setValue(1)
    
    ' Цены альтернатив (среднее геометрическое)
    oSheet.getCellByPosition(5,2).setString("Цены")
    oSheet.getCellByPosition(6,3).setFormula("=(B4*C4*D4*E4)^(1/4)")
    oSheet.getCellByPosition(6,4).setFormula("=(B5*C5*D5*E5)^(1/4)")
    oSheet.getCellByPosition(6,5).setFormula("=(B6*C6*D6*E6)^(1/4)")
    oSheet.getCellByPosition(6,6).setFormula("=(B7*C7*D7*E7)^(1/4)")
    oSheet.getCellByPosition(5,6).setString("Сумма")
    oSheet.getCellByPosition(6,6).setFormula("=SUM(G4:G7)")
    
    ' Веса
    oSheet.getCellByPosition(5,1).setString("Веса")
    oSheet.getCellByPosition(7,3).setFormula("=G4/$G$7")
    oSheet.getCellByPosition(7,4).setFormula("=G5/$G$7")
    oSheet.getCellByPosition(7,5).setFormula("=G6/$G$7")
    oSheet.getCellByPosition(7,6).setFormula("=G7/$G$7")
    
    ' Проверка на непротиворечивость
    oSheet.getCellByPosition(0,7).setString("Суммы столбцов")
    oSheet.getCellByPosition(1,7).setFormula("=SUM(B4:B7)")
    oSheet.getCellByPosition(2,7).setFormula("=SUM(C4:C7)")
    oSheet.getCellByPosition(3,7).setFormula("=SUM(D4:D7)")
    oSheet.getCellByPosition(4,7).setFormula("=SUM(E4:E7)")
    
    oSheet.getCellByPosition(5,9).setString("λ =")
    oSheet.getCellByPosition(6,9).setFormula("=SUMPRODUCT(B8:E8,H4:H7)")
    oSheet.getCellByPosition(5,10).setString("ИС =")
    oSheet.getCellByPosition(6,10).setFormula("=(G10-4)/3")
    oSheet.getCellByPosition(5,11).setString("C1C =")
    oSheet.getCellByPosition(6,11).setValue(0.9)
    oSheet.getCellByPosition(5,12).setString("ОС =")
    oSheet.getCellByPosition(6,12).setFormula("=G11/G12")
    oSheet.getCellByPosition(5,13).setString("Вывод:")
    oSheet.getCellByPosition(6,13).setFormula("=IF(G13<0.2,""Непротиворечиво"",""Требует уточнения"")")
    
    ' ===== ЛИСТ 3: ПРЕДПОЧТЕНИЯ =====
    oSheets.insertNewByName("Предпочтения", 2)
    oSheet = oSheets.getByName("Предпочтения")
    oSheet.getCellByPosition(0,0).setString("МЕТОД ПРЕДПОЧТЕНИЙ (3 эксперта)")
    oSheet.getCellByPosition(0,1).setString("Таблица мест")
    oSheet.getCellByPosition(1,2).setString("A1")
    oSheet.getCellByPosition(2,2).setString("A2")
    oSheet.getCellByPosition(3,2).setString("A3")
    oSheet.getCellByPosition(4,2).setString("A4")
    oSheet.getCellByPosition(0,3).setString("Э1")
    oSheet.getCellByPosition(0,4).setString("Э2")
    oSheet.getCellByPosition(0,5).setString("Э3")
    ' Места
    oSheet.getCellByPosition(1,3).setValue(2)
    oSheet.getCellByPosition(2,3).setValue(1)
    oSheet.getCellByPosition(3,3).setValue(4)
    oSheet.getCellByPosition(4,3).setValue(3)
    oSheet.getCellByPosition(1,4).setValue(1)
    oSheet.getCellByPosition(2,4).setValue(2)
    oSheet.getCellByPosition(3,4).setValue(4)
    oSheet.getCellByPosition(4,4).setValue(3)
    oSheet.getCellByPosition(1,5).setValue(1)
    oSheet.getCellByPosition(2,5).setValue(2)
    oSheet.getCellByPosition(3,5).setValue(3)
    oSheet.getCellByPosition(4,5).setValue(4)
    
    ' Преобразование B = 4 - X
    oSheet.getCellByPosition(0,7).setString("Преобразование B = 4 - X")
    oSheet.getCellByPosition(1,8).setString("A1")
    oSheet.getCellByPosition(2,8).setString("A2")
    oSheet.getCellByPosition(3,8).setString("A3")
    oSheet.getCellByPosition(4,8).setString("A4")
    oSheet.getCellByPosition(0,9).setString("Э1")
    oSheet.getCellByPosition(0,10).setString("Э2")
    oSheet.getCellByPosition(0,11).setString("Э3")
    oSheet.getCellByPosition(1,9).setFormula("=4-B4")
    oSheet.getCellByPosition(2,9).setFormula("=4-C4")
    oSheet.getCellByPosition(3,9).setFormula("=4-D4")
    oSheet.getCellByPosition(4,9).setFormula("=4-E4")
    oSheet.getCellByPosition(1,10).setFormula("=4-B5")
    oSheet.getCellByPosition(2,10).setFormula("=4-C5")
    oSheet.getCellByPosition(3,10).setFormula("=4-D5")
    oSheet.getCellByPosition(4,10).setFormula("=4-E5")
    oSheet.getCellByPosition(1,11).setFormula("=4-B6")
    oSheet.getCellByPosition(2,11).setFormula("=4-C6")
    oSheet.getCellByPosition(3,11).setFormula("=4-D6")
    oSheet.getCellByPosition(4,11).setFormula("=4-E6")
    
    ' Суммы по столбцам
    oSheet.getCellByPosition(0,12).setString("Cj:")
    oSheet.getCellByPosition(1,12).setFormula("=SUM(B10:B12)")
    oSheet.getCellByPosition(2,12).setFormula("=SUM(C10:C12)")
    oSheet.getCellByPosition(3,12).setFormula("=SUM(D10:D12)")
    oSheet.getCellByPosition(4,12).setFormula("=SUM(E10:E12)")
    
    oSheet.getCellByPosition(5,13).setString("Сумма C =")
    oSheet.getCellByPosition(6,13).setFormula("=SUM(B13:E13)")
    
    oSheet.getCellByPosition(0,14).setString("Веса Vj:")
    oSheet.getCellByPosition(1,14).setFormula("=B13/$G$14")
    oSheet.getCellByPosition(2,14).setFormula("=C13/$G$14")
    oSheet.getCellByPosition(3,14).setFormula("=D13/$G$14")
    oSheet.getCellByPosition(4,14).setFormula("=E13/$G$14")
    
    ' Проверка согласованности
    oSheet.getCellByPosition(0,16).setString("Проверка согласованности")
    oSheet.getCellByPosition(0,17).setString("Sj (суммы исходных):")
    oSheet.getCellByPosition(1,17).setFormula("=SUM(B4:B6)")
    oSheet.getCellByPosition(2,17).setFormula("=SUM(C4:C6)")
    oSheet.getCellByPosition(3,17).setFormula("=SUM(D4:D6)")
    oSheet.getCellByPosition(4,17).setFormula("=SUM(E4:E6)")
    
    oSheet.getCellByPosition(0,18).setString("A =")
    oSheet.getCellByPosition(1,18).setFormula("=3*(4+1)/2")
    
    oSheet.getCellByPosition(0,19).setString("S =")
    oSheet.getCellByPosition(1,19).setFormula("=(B18-B19)^2+(C18-B19)^2+(D18-B19)^2+(E18-B19)^2")
    
    oSheet.getCellByPosition(0,20).setString("W =")
    oSheet.getCellByPosition(1,20).setFormula("=12*B20/(3^2*4*(4^2-1))")
    
    oSheet.getCellByPosition(0,21).setString("Вывод:")
    oSheet.getCellByPosition(1,21).setFormula("=IF(B21>0.5,""Согласовано"",""Требует уточнения"")")
    
    ' ===== ЛИСТ 4: РАНГ =====
    oSheets.insertNewByName("Ранг", 3)
    oSheet = oSheets.getByName("Ранг")
    oSheet.getCellByPosition(0,0).setString("МЕТОД РАНГА (3 эксперта, 10 баллов)")
    oSheet.getCellByPosition(0,1).setString("Таблица баллов")
    oSheet.getCellByPosition(1,2).setString("A1")
    oSheet.getCellByPosition(2,2).setString("A2")
    oSheet.getCellByPosition(3,2).setString("A3")
    oSheet.getCellByPosition(4,2).setString("A4")
    oSheet.getCellByPosition(0,3).setString("Э1")
    oSheet.getCellByPosition(0,4).setString("Э2")
    oSheet.getCellByPosition(0,5).setString("Э3")
    ' Баллы
    oSheet.getCellByPosition(1,3).setValue(9)
    oSheet.getCellByPosition(2,3).setValue(10)
    oSheet.getCellByPosition(3,3).setValue(2)
    oSheet.getCellByPosition(4,3).setValue(8)
    oSheet.getCellByPosition(1,4).setValue(10)
    oSheet.getCellByPosition(2,4).setValue(9)
    oSheet.getCellByPosition(3,4).setValue(2)
    oSheet.getCellByPosition(4,4).setValue(8)
    oSheet.getCellByPosition(1,5).setValue(10)
    oSheet.getCellByPosition(2,5).setValue(9)
    oSheet.getCellByPosition(3,5).setValue(8)
    oSheet.getCellByPosition(4,5).setValue(2)
    
    ' Суммы
    oSheet.getCellByPosition(0,6).setString("Cj:")
    oSheet.getCellByPosition(1,6).setFormula("=SUM(B4:B6)")
    oSheet.getCellByPosition(2,6).setFormula("=SUM(C4:C6)")
    oSheet.getCellByPosition(3,6).setFormula("=SUM(D4:D6)")
    oSheet.getCellByPosition(4,6).setFormula("=SUM(E4:E6)")
    
    oSheet.getCellByPosition(0,7).setString("Сумма C:")
    oSheet.getCellByPosition(1,7).setFormula("=SUM(B7:E7)")
    
    oSheet.getCellByPosition(0,8).setString("Веса Vj:")
    oSheet.getCellByPosition(1,8).setFormula("=B7/$B$8")
    oSheet.getCellByPosition(2,8).setFormula("=C7/$B$8")
    oSheet.getCellByPosition(3,8).setFormula("=D7/$B$8")
    oSheet.getCellByPosition(4,8).setFormula("=E7/$B$8")
    
    ' Средние по альтернативам
    oSheet.getCellByPosition(0,10).setString("Средние Xj:")
    oSheet.getCellByPosition(1,10).setFormula("=B7/3")
    oSheet.getCellByPosition(2,10).setFormula("=C7/3")
    oSheet.getCellByPosition(3,10).setFormula("=D7/3")
    oSheet.getCellByPosition(4,10).setFormula("=E7/3")
    
    ' Дисперсии по экспертам
    oSheet.getCellByPosition(6,2).setString("Дисперсии экспертов:")
    oSheet.getCellByPosition(6,3).setFormula("=((B4-$B$11)^2+(C4-$C$11)^2+(D4-$D$11)^2+(E4-$E$11)^2)/3")
    oSheet.getCellByPosition(6,4).setFormula("=((B5-$B$11)^2+(C5-$C$11)^2+(D5-$D$11)^2+(E5-$E$11)^2)/3")
    oSheet.getCellByPosition(6,5).setFormula("=((B6-$B$11)^2+(C6-$C$11)^2+(D6-$D$11)^2+(E6-$E$11)^2)/3")
    
    ' Дисперсии по альтернативам
    oSheet.getCellByPosition(0,12).setString("Дисперсии альтернатив:")
    oSheet.getCellByPosition(1,12).setFormula("=((B4-$B$11)^2+(B5-$B$11)^2+(B6-$B$11)^2)/2")
    oSheet.getCellByPosition(2,12).setFormula("=((C4-$C$11)^2+(C5-$C$11)^2+(C6-$C$11)^2)/2")
    oSheet.getCellByPosition(3,12).setFormula("=((D4-$D$11)^2+(D5-$D$11)^2+(D6-$D$11)^2)/2")
    oSheet.getCellByPosition(4,12).setFormula("=((E4-$E$11)^2+(E5-$E$11)^2+(E6-$E$11)^2)/2")
    
    MsgBox "Готово! Созданы 4 листа с расчётами для Варианта 1."
End Sub
