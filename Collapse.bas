Attribute VB_Name = "Collapse"
Sub TimetableCollapse()

    collapse_day = 1 '(1 Tuesday, 3 Thursday)
    collapse_year = 1 '(0 Juniors, 1 Seniors)
    show_meetings = True
    
    resolution_message = "Cover Required"
    'resolution_message = "Reroomed"
    cover_message = "Cover Teacher"
    'cover_message = "New Room"
    
    years_to_move = 0 'Juniors (lesson)
    'years_to_move = 1 'Seniors (room)
    
    'Set up list of juniors and seniors
    Dim year() As Variant
    ReDim year(1, 2)
    year(0, 0) = "9"
    year(0, 1) = "10"
    year(1, 0) = "11"
    year(1, 1) = "L"
    year(1, 2) = "U"
    
    Dim timetable() As Variant
    Dim teacher() As Variant
    teachers = Range("A4").End(xlDown).Row
    ReDim timetable(teachers - 4, 4, 8)
    ReDim teacher(teachers - 4, 1)
        
    'check for variations between the two weeks and create dictionary
    For c = 1 To 47
        If c >= 3 Then
            d = (c - 3) \ 9 'day (0-4)
            p = (c - 3) Mod 9 'period (0-8)
        End If
        For r = 4 To teachers
            even_week = Worksheets("Even Week").Cells(r, c).Value
            odd_week = Worksheets("Odd Week").Cells(r, c).Value
            If c <= 2 Then
                If even_week = odd_week Then
                    teacher(r - 4, c - 1) = even_week
                Else
                    MsgBox ("Error with row " & r & ": " & even_week & ", " & odd_week)
                    Exit Sub
                End If
            Else
                If even_week = odd_week Then
                    Lessons = even_week
                Else
                    Lessons = even_week & odd_week
                End If
                timetable(r - 4, d, p) = Lessons
            End If
        Next r
    Next c
    
    'make list of senior teachers
    Dim senior() As Variant
    ReDim senior(teachers)
    
    For t = 1 To teachers - 4
        senior_class = False
        
         'Skip if SMT
        initial = teacher(t, 1)
        If initial = "BAHF" Or initial = "JJ" Or initial = "JMH" Or initial = "NXB" Or initial = "GS" Or initial = "MAC" Or initial = "LJG" Or initial = "DID" Or initial = "JWML" Or initial = "CMQ" Or initial = "KMJ" Or initial = "RNS" Or initial = "BCB" Then
            GoTo Done
        End If
        
        For d = 0 To 4
            For p = 0 To 8
                
                y = Left(timetable(t, d, p), 2)
                
                If InStr(y, "9") = 1 Or InStr(y, "10") = 1 Or InStr(y, "11") = 1 Or InStr(y, "L") = 1 Or InStr(y, "U") = 1 Then
                    senior_class = True
                    GoTo Done
                End If
            Next p
        Next d
Done:
        senior(t) = senior_class
    Next t
    
    
    MsgBox ("Step 1 of 3: Timetable, Teachers and Senior Teachers Imported Successfully")
    
    'Tues & Thurs  d = 1 & 3
    'periods 4 & 7  p = 3 & 6
    'periods 5 & 8  p = 4 & 7
    
    Dim Collapse() As Variant
    Dim freestaff() As Variant
    Dim free() As Variant
    Dim clash() As Variant

    'collapse(P4/P5,number)= Teacher
    ReDim Collapse(1, teachers)
    ReDim freestaff(1, teachers)
    ReDim clash(1)
    ReDim free(1)
    
    clash(0) = 0
    clash(1) = 0
    free(0) = 0
    free(1) = 0
    
    For t = 0 To UBound(teacher, 1)
        For i = 0 To 1
            first_period = timetable(t, collapse_day, 3 + i)
            second_period = timetable(t, collapse_day, 6 + i)
            
            'Check if lesson/event in both
            If first_period <> None And second_period <> None Then
                
                'Skip if Games/Off Games and record as free if Games for both
                If InStr(first_period, "Games") <> 0 Or InStr(second_period, "Games") <> 0 Then
                    If InStr(first_period, "Games") <> 0 And InStr(second_period, "Games") <> 0 Then
                        GoTo StaffFree
                    End If
                    GoTo NextIteration
                End If
                
                'Skip if only part time in second timing
                If InStr(second_period, "Part Time") = 1 Then
                    GoTo NextIteration
                End If
                
                'Skip if meeting
                If show_meetings = False Then
                    If InStr(first_period, "Meeting") <> 0 Or InStr(second_period, "Meeting") <> 0 Then
                        GoTo NextIteration
                    End If
                End If
                
                'Check if second lesson will collapse based on year
                For j = 0 To 2
                    year_name = year(collapse_year, j)
                    If year_name <> None And InStr(second_period, year_name) = 1 Then
                        'Add to collapse list
                        Collapse(i, clash(i)) = t
                        clash(i) = clash(i) + 1
                        Exit For
                    End If
                Next j
            
            'If both free or directed from above
            ElseIf first_period = None And second_period = None Then
StaffFree:
                'Check if teacher is in senior
                If senior(t) = True Then
                    freestaff(i, free(i)) = t
                    free(i) = free(i) + 1
                End If
            End If

NextIteration:
        Next i
    Next t
    
    MsgBox ("Step 2 of 3: Clash List made (" & clash(0) & ", " & clash(1) & ")")
    
    'Export Timetable Collapse
    
    For i = 0 To 1
        
        'Check if it already exists
        exists = False
        first_period = "P" & Str(4 + i)
        second_period = "P" & Str(7 + i)
        sheet_name = "Collapse " & first_period
        For s = 1 To Worksheets.Count
            If Worksheets(s).Name = sheet_name Then
                exists = True
            End If
        Next s
        
        'If not, create collapse export sheet
        If exists = False Then
            Sheets.Add.Name = sheet_name
            'titles
            Worksheets(sheet_name).Cells(1, 2).Value = first_period
            Worksheets(sheet_name).Cells(1, 3).Value = second_period
            Worksheets(sheet_name).Cells(1, 4).Value = resolution_message
            Worksheets(sheet_name).Cells(1, 5).Value = cover_message
            Worksheets(sheet_name).Cells(1, 7).Value = "Free Teachers"
            
            'work through collapse
            For t = 0 To teachers
                teacher_number = Collapse(i, t)
                If teacher_number = Empty Then
                    Exit For
                End If
                teacher_name = teacher(teacher_number, 0)
                first_period = timetable(teacher_number, collapse_day, 3 + i)
                second_period = timetable(teacher_number, collapse_day, 6 + i)
                
                Worksheets(sheet_name).Cells(2 + t, 1).Value = teacher(teacher_number, 0)
                Worksheets(sheet_name).Cells(2 + t, 2).Value = Lesson(first_period)
                Worksheets(sheet_name).Cells(2 + t, 3).Value = Lesson(second_period)
                
                'create resolutions
                'Cancel meetings
                If InStr(first_period, "Meeting") <> 0 Or InStr(second_period, "Meeting") <> 0 Then
                    Worksheets(sheet_name).Cells(2 + t, 4).Value = "Cancel Meeting"
                    Worksheets(sheet_name).Cells(2 + t, 5).Value = "None Required"
                
                'if one if not senior move other
                ElseIf InStr(Lesson(first_period), "Yr") = 1 Then
                    Worksheets(sheet_name).Cells(2 + t, 4).Value = Lesson(second_period)
                
                'check which one needs to move 0 Juniors or 1 Seniors
                Else
                    For j = 0 To 2
                        If InStr(first_period, year(years_to_move, j)) = 1 Then
                            Worksheets(sheet_name).Cells(2 + t, 4).Value = Lesson(first_period)
                        ElseIf InStr(second_period, year(years_to_move, j)) = 1 Then
                            Worksheets(sheet_name).Cells(2 + t, 4).Value = Lesson(second_period)
                        End If
                    Next j
                
                End If
            
            Next t
            
            'List free teachers
            For t = 0 To teachers
            
            
                If freestaff(i, t) = "" Then
                    Exit For
                Else
                    teacher_name = teacher(freestaff(i, t), 0)
                End If
                Worksheets(sheet_name).Cells(2 + t, 7).Value = teacher_name
                Next t
        
        Else
            MsgBox ("Sheet for " & sheet_name & " already exists.")
        End If
        
        'Format the sheet
        FormatClashes (sheet_name)
    Next i
    
    MsgBox ("Step 3 of 3: Clashes Successfully exported")

End Sub

Function Lesson(s)
    Dim s1, s2, subj, y As String
    
    If InStr(s, "Meeting") = 0 Then
        'split by delimiter
        s1 = Split(s, "/")(1)
        
        'split by new line
        s2 = Split(s1, Chr(10))(0)
        
        'split to only 2 character subj code
        subj = Left(s2, 2)
        
        If subj = "Bi" Then
            subj = "Biology"
        ElseIf subj = "Ch" Then
            subj = "Chemistry"
        ElseIf subj = "Ph" Then
            subj = "Physics"
        ElseIf subj = "Ma" Then
            subj = "Maths"
        ElseIf subj = "Cs" Then
            subj = "Computer Science"
        ElseIf subj = "Sc" Then
            subj = "Science"
        ElseIf subj = "Es" Then
            subj = "ESS"
        ElseIf subj = "En" Then
            subj = "English"
        ElseIf subj = "Gm" Then
            subj = "German"
        ElseIf subj = "Sp" Then
            subj = "Spanish"
        ElseIf subj = "Fr" Then
            subj = "French"
        ElseIf subj = "Hi" Then
            subj = "History"
        ElseIf subj = "Gg" Then
            subj = "Geography"
        ElseIf subj = "Gp" Then
            subj = "Global Perspectives"
        ElseIf subj = "Dt" Then
            subj = "DT"
        ElseIf subj = "Dr" Then
            subj = "Drama"
        ElseIf subj = "Px" Then
            subj = "PE"
        ElseIf subj = "Pp" Then
            subj = "Philosophy"
        ElseIf subj = "Py" Then
            subj = "Psychology"
        ElseIf subj = "Tk" Then
            subj = "ToK"
        ElseIf subj = "Po" Then
            subj = "Politics"
        ElseIf subj = "So" Then
            subj = "Sociology"
        ElseIf subj = "Dp" Then
            subj = "Divinity & Philosophy"
        ElseIf subj = "Rs" Then
            subj = "Religous Studies"
        ElseIf subj = "Bm" Then
            subj = "Business Management"
        ElseIf subj = "Bs" Then
            subj = "Business"
        ElseIf subj = "Ec" Then
            subj = "Economics"
        ElseIf subj = "Ss" Then
            subj = "Supervised Study"
        
        End If
        
        'get level
        L = Right(Left(s, 3), 1) & "L "
        If L <> "HL " And L <> "SL " Then
            L = ""
        End If
        
        'get course/set
        c = " " & Right(s2, Len(s2) - 2)
        
        If Right(c, 1) = "T" Then
            c = " Triple Set " & Left(c, Len(c) - 1)
        ElseIf Right(c, 1) = "D" Then
            c = " Dual Set " & Left(c, Len(c) - 1)
        ElseIf Right(c, 1) = "S" Then
            c = " Single Set " & Left(c, Len(c) - 1)
        ElseIf c <> " " Then
            c = " Set" & c
        Else
            c = ""
        End If
        
        
        'get room (Only so complicated to overcome lack of room problem)
        r_len = Len(s1) - Len(s2) - 1
        If r_len >= 0 Then
            r = Chr(10) & "(" & Right(s, r_len) & ")"
        Else
            r = Chr(10) & "(Room Unspecified)"
        End If
            
        If InStr(s, "9") = 1 Then
            y = "Shell " & subj & c & r
        ElseIf InStr(s, "10") = 1 Then
            y = "Remove " & subj & c & r
        ElseIf InStr(s, "11") = 1 Then
            y = "Fifth " & subj & c & r
        ElseIf InStr(s, "LI") = 1 Then
            y = "LVI IB " & L & subj & c & r
        ElseIf InStr(s, "LA") = 1 Then
            If Right(s2, 1) = "B" Then
                y = "LVI Btec " & subj & r
            Else
                y = "LVI A Level " & subj & c & r
            End If
        ElseIf InStr(s, "UI") = 1 Then
            y = "UVI IB " & L & subj & c & r
        ElseIf InStr(s, "UA") = 1 Then
            If Right(s2, 1) = "B" Then
                y = "UVI Btec " & subj & r
            Else
                y = "UVI A Level " & subj & c & r
            End If
        Else
            y = "Yr" & Left(s, 1) & " " & subj & c & r
        End If
    Else
        y = s
    End If
    
    Lesson = y

End Function

Sub FormatClashes(n)

    Worksheets(n).Range("B1:F100").HorizontalAlignment = xlCenter
    Worksheets(n).Range("A1:G100").VerticalAlignment = xlCenter
    
    With Worksheets(n).Columns("A")
        .ColumnWidth = 20
        .Font.Bold = True
    End With
    
    Worksheets(n).Columns("B").ColumnWidth = 25
    Worksheets(n).Columns("C").ColumnWidth = 25
    Worksheets(n).Columns("D").ColumnWidth = 25
    Worksheets(n).Columns("E").ColumnWidth = 20
    Worksheets(n).Columns("F").ColumnWidth = 6
    Worksheets(n).Columns("G").ColumnWidth = 20
    
    Worksheets(n).Rows(1).Font.Bold = True
    
    For r = 2 To 20
        Worksheets(n).Rows(r).RowHeight = 25
    Next r
    
End Sub



