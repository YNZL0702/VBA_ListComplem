Function ListComplement(Array_A, new_Array)

Dim i As Integer
Dim rn As Integer: rn = Range("A65536").End(xlUp).Row   '
Dim Comp_str As String
Dim Comp_arr

If (Not IsEmpty(Array_A)) Then                      '当A 为 非空的时候

    If (Not IsEmpty(new_Array)) Then                    '当new 为 非空的时候
    
        For i = 1 To UBound(new_Array)

            If IsError(Application.Match(new_Name(i), Array_A, 0)) = True Then    ' 在A 数组里， 寻找new(i)是否存在
    
                Comp_str = Comp_str & new_Name(i) & ","
    
            End If
    
        Next
                If Len(Comp_str) >= 1 Then
                    Comp_arr = Left(Comp_str, Len(Comp_str) - 1)
                    Comp_arr = Split(Comp_arr, ",")                   'Split() 生成的数组，index从0 开始
                    Comp_str = ""                                      
                    Range("B" & rn + 1).Resize(UBound(Comp_arr) - LBound(Comp_arr) + 1, 1) = Application.WorksheetFunction.Transpose(Comp_arr)
                Else
                    MsgBox "added Documents are already in the list"
                End If
    Else
        If IsEmpty(new_Array) Then                '当new 为 空的时候
            
                Exit Function
        End If
        
    End If

Else                                             '当A 为 空的时候

    If (Not IsEmpty(new_Array)) Then              '当new 为 非空的时候
            Comp_arr = new_Array
            Range("B" & rn + 1).Resize(UBound(Comp_arr) - LBound(Comp_arr) + 1, 1) = Application.WorksheetFunction.Transpose(Comp_arr)

    Else: Exit Function
    End If
            
End If

End Function