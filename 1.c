Sub SozdatLab4_V8
    Dim oDoc As Object
    Dim oSheets As Object
    Dim oSheet As Object
    Dim i As Integer
    
    oDoc = ThisComponent
    oSheets = oDoc.Sheets
    
    ' Удаляем старые листы
    Dim names(3) As String
    names(0) = "Data"
    names(1) = "Saaty"
    names(2) = "Preference"
    names(3) = "Rank"
    For i = 0 To 3
        If oSheets.hasByName(names(i)) Then
            oSheets.removeByName(names(i))
        End If
    Next i
    
    ' ===== ЛИСТ 1: DATA =====
    oSheets.insertNewByName("Data", 0)
    oSheet = oSheets.getByName("Data")
    oSheet.getCellByPosition(0,0).setString("Alternative")
    oSheet.getCellByPosition(1,0).setString("A1")
    oSheet.getCellByPosition(2,0).setString("A2")
    oSheet.getCellByPosition(3,0).setString("A3")
    oSheet.getCellByPosition(4,0).setString("A4")
    oSheet.getCellByPosition(0,1).setString("Expert1")
    oSheet.getCellByPosition(0,2).setString("Expert2")
    oSheet.getCellByPosition(0,3).setString("Expert3")
    ' Места (1 = лучший): Вариант 8
    oSheet.getCellByPosition(1,1).setValue(3)
    oSheet.getCellByPosition(2,1).setValue(4)
    oSheet.getCellByPosition(3,1).setValue(2)
    oSheet.getCellByPosition(4,1).setValue(1)
    oSheet.getCellByPosition(1,2).setValue(1)
    oSheet.getCellByPosition(2,2).setValue(4)
    oSheet.getCellByPosition(3,2).setValue(2)
    oSheet.getCellByPosition(4,2).setValue(3)
    oSheet.getCellByPosition(1,3).setValue(4)
    oSheet.getCellByPosition(2,3).setValue(3)
    oSheet.getCellByPosition(3,3).setValue(1)
    oSheet.getCellByPosition(4,3).setValue(2)
    
    ' ===== ЛИСТ 2: SAATY =====
    oSheets.insertNewByName("Saaty", 1)
    oSheet = oSheets.getByName("Saaty")
    oSheet.getCellByPosition(0,0).setString("SAATY METHOD (expert 1)")
    oSheet.getCellByPosition(1,2).setString("A1")
    oSheet.getCellByPosition(2,2).setString("A2")
    oSheet.getCellByPosition(3,2).setString("A3")
    oSheet.getCellByPosition(4,2).setString("A4")
    oSheet.getCellByPosition(0,3).setString("A1")
    oSheet.getCellByPosition(0,4).setString("A2")
    oSheet.getCellByPosition(0,5).setString("A3")
    oSheet.getCellByPosition(0,6).setString("A4")
    ' Матрица: A4>A3>A1>A2 (Вариант 8, эксперт 1)
    oSheet.getCellByPosition(1,3).setValue(1)
    oSheet.getCellByPosition(2,3).setValue(1/3)
    oSheet.getCellByPosition(3,3).setValue(1/5)
    oSheet.getCellByPosition(4,3).setValue(1/7)
    oSheet.getCellByPosition(1,4).setValue(3)
    oSheet.getCellByPosition(2,4).setValue(1)
    oSheet.getCellByPosition(3,4).setValue(1/7)
    oSheet.getCellByPosition(4,4).setValue(1/9)
    oSheet.getCellByPosition(1,5).setValue(5)
    oSheet.getCellByPosition(2,5).setValue(7)
    oSheet.getCellByPosition(3,5).setValue(1)
    oSheet.getCellByPosition(4,5).setValue(1/3)
    oSheet.getCellByPosition(1,6).setValue(7)
    oSheet.getCellByPosition(2,6).setValue(9)
    oSheet.getCellByPosition(3,6).setValue(3)
    oSheet.getCellByPosition(4,6).setValue(1)
    
    ' Цены (столбец G, строки 4-7)
    oSheet.getCellByPosition(5,2).setString("Prices")
    oSheet.getCellByPosition(6,3).setFormula("=(B4*C4*D4*E4)^(1/4)")
    oSheet.getCellByPosition(6,4).setFormula("=(B5*C5*D5*E5)^(1/4)")
    oSheet.getCellByPosition(6,5).setFormula("=(B6*C6*D6*E6)^(1/4)")
    oSheet.getCellByPosition(6,6).setFormula("=(B7*C7*D7*E7)^(1/4)")
    ' Сумма — G8 (не G7, чтобы не зациклиться)
    oSheet.getCellByPosition(5,7).setString("Sum")
    oSheet.getCellByPosition(6,7).setFormula("=SOMME(G4:G7)")
    
    ' Веса (столбец H, строки 4-7)
    oSheet.getCellByPosition(5,1).setString("Weights")
    oSheet.getCellByPosition(7,3).setFormula("=G4/$G$8")
    oSheet.getCellByPosition(7,4).setFormula("=G5/$G$8")
    oSheet.getCellByPosition(7,5).setFormula("=G6/$G$8")
    oSheet.getCellByPosition(7,6).setFormula("=G7/$G$8")
    
    ' Суммы столбцов — строка 8 (B8:E8)
    oSheet.getCellByPosition(0,7).setString("Col sums")
    oSheet.getCellByPosition(1,7).setFormula("=SOMME(B4:B7)")
    oSheet.getCellByPosition(2,7).setFormula("=SOMME(C4:C7)")
    oSheet.getCellByPosition(3,7).setFormula("=SOMME(D4:D7)")
    oSheet.getCellByPosition(4,7).setFormula("=SOMME(E4:E7)")
    
    ' Проверка (строки 11-14, столбцы F-G)
    oSheet.getCellByPosition(5,10).setString("Lambda =")
    oSheet.getCellByPosition(6,10).setFormula("=SOMMEPROD(B8:E8;H4:H7)")
    oSheet.getCellByPosition(5,11).setString("IS =")
    oSheet.getCellByPosition(6,11).setFormula("=(G11-4)/3")
    oSheet.getCellByPosition(5,12).setString("C1C =")
    oSheet.getCellByPosition(6,12).setValue(0.9)
    oSheet.getCellByPosition(5,13).setString("OS =")
    oSheet.getCellByPosition(6,13).setFormula("=G12/G13")
    oSheet.getCellByPosition(5,14).setString("Result:")
    oSheet.getCellByPosition(6,14).setFormula("=SI(G14<0.2;""OK"";""NeedFix"")")
    
    ' ===== ЛИСТ 3: PREFERENCE =====
    oSheets.insertNewByName("Preference", 2)
    oSheet = oSheets.getByName("Preference")
    oSheet.getCellByPosition(0,0).setString("PREFERENCE METHOD")
    oSheet.getCellByPosition(1,2).setString("A1")
    oSheet.getCellByPosition(2,2).setString("A2")
    oSheet.getCellByPosition(3,2).setString("A3")
    oSheet.getCellByPosition(4,2).setString("A4")
    oSheet.getCellByPosition(0,3).setString("E1")
    oSheet.getCellByPosition(0,4).setString("E2")
    oSheet.getCellByPosition(0,5).setString("E3")
    ' Места
    oSheet.getCellByPosition(1,3).setValue(3)
    oSheet.getCellByPosition(2,3).setValue(4)
    oSheet.getCellByPosition(3,3).setValue(2)
    oSheet.getCellByPosition(4,3).setValue(1)
    oSheet.getCellByPosition(1,4).setValue(1)
    oSheet.getCellByPosition(2,4).setValue(4)
    oSheet.getCellByPosition(3,4).setValue(2)
    oSheet.getCellByPosition(4,4).setValue(3)
    oSheet.getCellByPosition(1,5).setValue(4)
    oSheet.getCellByPosition(2,5).setValue(3)
    oSheet.getCellByPosition(3,5).setValue(1)
    oSheet.getCellByPosition(4,5).setValue(2)
    
    ' Преобразование B = 4 - X (строки 10-12)
    oSheet.getCellByPosition(0,7).setString("Transform B = 4 - X")
    oSheet.getCellByPosition(1,8).setString("A1")
    oSheet.getCellByPosition(2,8).setString("A2")
    oSheet.getCellByPosition(3,8).setString("A3")
    oSheet.getCellByPosition(4,8).setString("A4")
    oSheet.getCellByPosition(0,9).setString("E1")
    oSheet.getCellByPosition(0,10).setString("E2")
    oSheet.getCellByPosition(0,11).setString("E3")
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
    
    ' Суммы Cj — строка 13
    oSheet.getCellByPosition(0,12).setString("Cj:")
    oSheet.getCellByPosition(1,12).setFormula("=SOMME(B10:B12)")
    oSheet.getCellByPosition(2,12).setFormula("=SOMME(C10:C12)")
    oSheet.getCellByPosition(3,12).setFormula("=SOMME(D10:D12)")
    oSheet.getCellByPosition(4,12).setFormula("=SOMME(E10:E12)")
    
    ' Общая сумма C — G14
    oSheet.getCellByPosition(5,13).setString("Sum C =")
    oSheet.getCellByPosition(6,13).setFormula("=SOMME(B13:E13)")
    
    ' Веса — строка 15
    oSheet.getCellByPosition(0,14).setString("Weights Vj:")
    oSheet.getCellByPosition(1,14).setFormula("=B13/$G$14")
    oSheet.getCellByPosition(2,14).setFormula("=C13/$G$14")
    oSheet.getCellByPosition(3,14).setFormula("=D13/$G$14")
    oSheet.getCellByPosition(4,14).setFormula("=E13/$G$14")
    
    ' Проверка согласованности — строки 18-22
    oSheet.getCellByPosition(0,16).setString("Concordance check")
    oSheet.getCellByPosition(0,17).setString("Sj:")
    oSheet.getCellByPosition(1,17).setFormula("=SOMME(B4:B6)")
    oSheet.getCellByPosition(2,17).setFormula("=SOMME(C4:C6)")
    oSheet.getCellByPosition(3,17).setFormula("=SOMME(D4:D6)")
    oSheet.getCellByPosition(4,17).setFormula("=SOMME(E4:E6)")
    oSheet.getCellByPosition(0,18).setString("A =")
    oSheet.getCellByPosition(1,18).setFormula("=3*(4+1)/2")
    oSheet.getCellByPosition(0,19).setString("S =")
    oSheet.getCellByPosition(1,19).setFormula("=(B18-B19)^2+(C18-B19)^2+(D18-B19)^2+(E18-B19)^2")
    oSheet.getCellByPosition(0,20).setString("W =")
    oSheet.getCellByPosition(1,20).setFormula("=12*B20/(3^2*4*(4^2-1))")
    oSheet.getCellByPosition(0,21).setString("Result:")
    oSheet.getCellByPosition(1,21).setFormula("=SI(B21>0.5;""OK"";""NeedFix"")")
    
    ' ===== ЛИСТ 4: RANK =====
    oSheets.insertNewByName("Rank", 3)
    oSheet = oSheets.getByName("Rank")
    oSheet.getCellByPosition(0,0).setString("RANK METHOD (10 points)")
    oSheet.getCellByPosition(1,2).setString("A1")
    oSheet.getCellByPosition(2,2).setString("A2")
    oSheet.getCellByPosition(3,2).setString("A3")
    oSheet.getCellByPosition(4,2).setString("A4")
    oSheet.getCellByPosition(0,3).setString("E1")
    oSheet.getCellByPosition(0,4).setString("E2")
    oSheet.getCellByPosition(0,5).setString("E3")
    ' Баллы: 1-е=10, 2-е=8, 3-е=5, 4-е=2 (Вариант 8)
    oSheet.getCellByPosition(1,3).setValue(5)
    oSheet.getCellByPosition(2,3).setValue(2)
    oSheet.getCellByPosition(3,3).setValue(8)
    oSheet.getCellByPosition(4,3).setValue(10)
    oSheet.getCellByPosition(1,4).setValue(10)
    oSheet.getCellByPosition(2,4).setValue(2)
    oSheet.getCellByPosition(3,4).setValue(8)
    oSheet.getCellByPosition(4,4).setValue(5)
    oSheet.getCellByPosition(1,5).setValue(2)
    oSheet.getCellByPosition(2,5).setValue(5)
    oSheet.getCellByPosition(3,5).setValue(10)
    oSheet.getCellByPosition(4,5).setValue(8)
    
    ' Суммы — строка 7
    oSheet.getCellByPosition(0,6).setString("Cj:")
    oSheet.getCellByPosition(1,6).setFormula("=SOMME(B4:B6)")
    oSheet.getCellByPosition(2,6).setFormula("=SOMME(C4:C6)")
    oSheet.getCellByPosition(3,6).setFormula("=SOMME(D4:D6)")
    oSheet.getCellByPosition(4,6).setFormula("=SOMME(E4:E6)")
    
    ' Общая сумма — B8
    oSheet.getCellByPosition(0,7).setString("Sum C:")
    oSheet.getCellByPosition(1,7).setFormula("=SOMME(B7:E7)")
    
    ' Веса — строка 9
    oSheet.getCellByPosition(0,8).setString("Weights Vj:")
    oSheet.getCellByPosition(1,8).setFormula("=B7/$B$8")
    oSheet.getCellByPosition(2,8).setFormula("=C7/$B$8")
    oSheet.getCellByPosition(3,8).setFormula("=D7/$B$8")
    oSheet.getCellByPosition(4,8).setFormula("=E7/$B$8")
    
    ' Средние — строка 11
    oSheet.getCellByPosition(0,10).setString("Means Xj:")
    oSheet.getCellByPosition(1,10).setFormula("=B7/3")
    oSheet.getCellByPosition(2,10).setFormula("=C7/3")
    oSheet.getCellByPosition(3,10).setFormula("=D7/3")
    oSheet.getCellByPosition(4,10).setFormula("=E7/3")
    
    ' Дисперсии экспертов — G4:G6
    oSheet.getCellByPosition(6,2).setString("Disp experts:")
    oSheet.getCellByPosition(6,3).setFormula("=((B4-$B$11)^2+(C4-$C$11)^2+(D4-$D$11)^2+(E4-$E$11)^2)/3")
    oSheet.getCellByPosition(6,4).setFormula("=((B5-$B$11)^2+(C5-$C$11)^2+(D5-$D$11)^2+(E5-$E$11)^2)/3")
    oSheet.getCellByPosition(6,5).setFormula("=((B6-$B$11)^2+(C6-$C$11)^2+(D6-$D$11)^2+(E6-$E$11)^2)/3")
    
    ' Дисперсии альтернатив — строка 13
    oSheet.getCellByPosition(0,12).setString("Disp alternatives:")
    oSheet.getCellByPosition(1,12).setFormula("=((B4-$B$11)^2+(B5-$B$11)^2+(B6-$B$11)^2)/2")
    oSheet.getCellByPosition(2,12).setFormula("=((C4-$C$11)^2+(C5-$C$11)^2+(C6-$C$11)^2)/2")
    oSheet.getCellByPosition(3,12).setFormula("=((D4-$D$11)^2+(D5-$D$11)^2+(D6-$D$11)^2)/2")
    oSheet.getCellByPosition(4,12).setFormula("=((E4-$E$11)^2+(E5-$E$11)^2+(E6-$E$11)^2)/2")
    
    MsgBox "Done! Variant 8 sheets created."
End Sub
