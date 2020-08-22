Public Function ToBoolean(ByVal ToConvert As Variant) As Boolean
    'Check type of the variable
    Select Case VarType(ToConvert)
        'Semantically 'empty' types
        Case 0, 1, 8192, 8193
            ToBoolean = False
        'Numbers
        Case 2, 3, 4, 5, 6, 7, 14, 17, 20
            'Convert to an Integer first
            If CInt(ToConvert) > 0 Then
                ToBoolean = True
            Else
                ToBoolean = False
            End If
        'Strings or types that can be directly converted to string
        Case 8, 10, 11
            On Error Resume Next
            ToConvert = CStr(ToConvert)
            If Err.Number <> 0 Then
                'Means we might have an onject with default property
                If ToConvert Is Nothing Then
                    ToBoolean = False
                Else
                    ToBoolean = True
                End If
            Else
                Select Case LCase(ToConvert)
                    Case "true", "èñòèíà", "yes", "äà", "ya"
                        ToBoolean = True
                    Case "false", "ëîæü", "no", "íåò", "nein", "", "n/a", "null", "nan"
                        ToBoolean = False
                    Case Else
                        'Doing this to avoid type mismatch error
                        If IsNumeric(ToConvert) Then
                            If CInt(ToConvert) > 0 Then
                                ToBoolean = True
                            Else
                                ToBoolean = False
                            End If
                        Else
                            'As some other programming languages treat a non empty string as True
                            ToBoolean = True
                        End If
                End Select
            End If
            On Error GoTo 0
        'Objects
        Case 9, 13
            If ToConvert Is Nothing Then
                ToBoolean = False
            Else
                ToBoolean = True
            End If
        'Arrays
        Case Else
            '8192 is added to any regular value identifying array of respective types. 8192 and 8193 would mean arrays of 'empty' elements, so we exclude them
            If VarType(ToConvert) = 12 Or VarType(ToConvert) = 36 Or VarType(ToConvert) >= 8194 Then
                If UBound(ToConvert) - LBound(ToConvert) + 1 < 1 Then
                    ToBoolean = False
                Else
                    ToBoolean = True
                End If
            Else
                ToBoolean = False
            End If
    End Select
End Function
