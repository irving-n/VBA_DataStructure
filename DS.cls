VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TERMS:
'
'   DataStructure: A value-variable or object-variable which holds a set of other values/objects

Public Enum eDataStructureFilterMode
    eFilterTrap
    eFilterOut
End Enum

Private Sub Class_Initialize()
End Sub


'PROPERTIES


'METHODS

Public Function Append(ByRef body As Variant, ByRef appendix As Variant) As Variant
    Dim output_type As String
    Dim arr As Variant, pend As Variant
    Dim current_ubound As Long, new_ubound As Long, k As Long
    Dim newkey As String
    
    
On Error GoTo Append_clean_up

    If isDataStructure(body) Then
        output_type = TypeName(body)
        Select Case output_type
            Case "Dictionary"
                k = 0
                newkey = CStr(k)
                While body.Exists(newkey)
                    k = k + 1
                    newkey = CStr(k)
                Wend
                body.Add Key:=newkey, Item:=appendix
                    
            Case "Collection"
                body.Add appendix
            Case Else
                On Error Resume Next
                
                current_ubound = UBound(body)
                new_ubound = current_ubound + 1
                
                If Err.Number <> 0 Then Err.Clear
                ReDim Preserve body(new_ubound)
                
                If IsObject(appendix) Then
                    Set body(new_ubound) = appendix
                  Else
                    body(new_ubound) = appendix
                End If
                
        End Select
    
    End If
    
Append_clean_up:
    If Err.Number <> 0 Then
        Debug.Print "Error [" & Err.Number & "]: " & Err.Description
        Err.Clear
    End If
    
    If IsObject(body) Then
        Set Append = body
      Else
        Append = body
    End If
End Function


Public Function Apply(DataStructure_arr As Variant, funcName As String, elem_arg_idx_pos As Integer, ParamArray args() As Variant) As Variant
'Note - this function is experimental
    Dim ub As Integer, i As Integer, j As Integer, arg_count As Integer, total_args_count As Integer, xfer_ceiling As Integer
    Dim funcString As String
    Dim rtn As Variant, args_xfer As Variant, val As Variant
    Dim compile_col As Collection, funcPart_col As Collection
    
    If Not IsArray(DataStructure_arr) Then
        Apply = Empty
        Exit Function
    End If
    
    If UBound(args) >= 0 Then
        If UBound(args) = 0 Then 'Fix off by 1 error
            xfer_ceiling = 0
          Else
            xfer_ceiling = UBound(args) + 1
        End If
        total_args_count = 1 + UBound(args) + 1
        ReDim args_xfer(total_args_count - 1)
        args_xfer(elem_arg_idx_pos) = Empty 'This will be replaced in a loop
        j = 0 'position of the input arguments - increment by 1 each time one of them is assigned, pseudo-stack
'        total_args_count = 1 + UBound(args) + 1
        For i = 0 To (total_args_count - 1)
            If j <= xfer_ceiling Then
                If i <> elem_arg_idx_pos Then
                    On Error Resume Next
                    args_xfer(i) = args(j)
                    If Err.Number <> 0 Then
                        Err.Clear
                        Set args_xfer(i) = args(j)
                    End If
                    j = j + 1
                    On Error GoTo 0
                End If
              Else
'                args_xfer(i) =
            End If
        Next i
        
'        ReDim args_xfer(1 + UBound(args))
      Else
        ReDim args_xfer(0)
        total_args_count = 1
    End If
    
'    For i = 0 To 5 'can always add more if needed
    
    ub = UBound(DataStructure_arr)
    ReDim rtn(ub)
    
    arg_count = UBound(args)
    Set compile_col = New Collection
    For i = 0 To ub
        On Error Resume Next
        args_xfer(elem_arg_idx_pos) = DataStructure_arr(i)
        If Err.Number <> 0 Then
            Err.Clear
            Set args_xfer(elem_arg_idx_pos) = DataStructure_arr(i)
        End If
        On Error GoTo 0
        
        On Error Resume Next
        Select Case total_args_count  ' warning - ugly; it needs to be like this because the arguments supplied can't be filled with a default Missing reference
            Case 1: compile_col.Add Item:=Application.run(funcName, args_xfer(0))
            Case 2: compile_col.Add Item:=Application.run(funcName, args_xfer(0), args_xfer(1))
            Case 3: compile_col.Add Item:=Application.run(funcName, args_xfer(0), args_xfer(1), args_xfer(2))
            Case 4: compile_col.Add Item:=Application.run(funcName, args_xfer(0), args_xfer(1), args_xfer(2), args_xfer(3))
            Case 5: compile_col.Add Item:=Application.run(funcName, args_xfer(0), args_xfer(1), args_xfer(2), args_xfer(3), args_xfer(4))
            Case 6: compile_col.Add Item:=Application.run(funcName, args_xfer(0), args_xfer(1), args_xfer(2), args_xfer(3), args_xfer(4), args_xfer(5))
            Case Else: Debug.Print "Please edit code to add more arguments, follow the same pattern and it should work."
        End Select
        If Err.Number <> 0 Then
            funcString = funcName & "("
            Set funcPart_col = New Collection
            For Each val In args_xfer
                If TypeName(val) = "String" Then
                    funcPart_col.Add "'" & val & "'"
                  ElseIf IsObject(val) Then
                    'uh oh
                  ElseIf IsNull(val) Then
                    val = "Null"
                  Else
                    If Not IsEmpty(val) Then funcPart_col.Add val
                End If
            Next val
            funcString = funcString & Join(CollectionToArray(funcPart_col), ", ") & ")"
            compile_col.Add Item:=Application.Evaluate(funcString)
            Err.Clear
        End If
        On Error GoTo 0
'        rtn(i) = Application.run(funcName, args(0))
    Next i
    Apply = CollectionToArray(compile_col)
End Function


Public Function CharacterArray(ByVal text As String) As Variant

    Dim strlen As Integer, i As Integer
    Dim rtn() As Variant
    strlen = Len(text)
    If strlen > 0 Then
        ReDim rtn(strlen - 1)
        For i = 0 To Len(text) - 1
            rtn(i) = Mid(text, i + 1, 1)
        Next i
    End If
    CharacterArray = rtn
    
End Function


Public Function Convert(DataStructure As Variant, Optional ByVal OutputType As String, Optional ConversionOptions As Variant) As Variant
    Dim typeNm As String
    Dim tmpArr As Variant 'temporary array
    Dim rtn As Variant
    Dim col As Collection
    
    typeNm = TypeName(DataStructure)
    If Not isDataStructure(typeNm, True) Then Set Convert = Nothing: Exit Function
    If Not isDataStructure(OutputType, True) Then Set Convert = Nothing: Exit Function
    
    Select Case typeNm
        Case "Dictionary"
            tmpArr = Me.Zip(DataStructure)
        Case "Collection"
            Set col = DataStructure
            tmpArr = CollectionToArray(col)
        Case Else 'Some form of array
            tmpArr = DataStructure
    End Select
    
    If TypeName(tmpArr) = OutputType Then 'no conversion necessary
        rtn = tmpArr
        GoTo Convert_clean_up
    End If
    
    Select Case OutputType
        Case "Dictionary"
            If IsArray(tmpArr(0)) Then
                Set rtn = ArrayToDictionary(tmpArr, True)
              Else
                Set rtn = ArrayToDictionary(tmpArr, False)
            End If
        Case "Collection"
            If IsArray(tmpArr(0)) Then
                Set rtn = ArrayToCollection(tmpArr, True)
              Else
                Set rtn = ArrayToCollection(tmpArr, False)
            End If
        Case "Variant()"
            rtn = Me.Apply(tmpArr, "CVar", 0)
        Case "Byte()"
            rtn = Me.Apply(tmpArr, "CByte", 0)
        Case "Integer()"
            rtn = Me.Apply(tmpArr, "CInt", 0)
        Case "Double()"
            rtn = Me.Apply(tmpArr, "CDbl", 0)
        Case "Single()"
            rtn = Me.Apply(tmpArr, "CSng", 0)
        Case "Long()"
            rtn = Me.Apply(tmpArr, "CLng", 0)
        Case "Boolean()"
            rtn = Me.Apply(tmpArr, "CBool", 0)
        Case Else
            'uh oh
    End Select

Convert_clean_up:
    If Err.Number <> 0 Then Debug.Print "Error [" & Err.Number & "] in method [Convert] in class DS: " & Err.Description: Err.Clear
    
    Set col = Nothing
    
    If IsObject(rtn) Then
        Set Convert = rtn
      Else
        Convert = rtn
    End If
    
End Function


Public Function Copy(DataStructure As Variant, Optional ByVal failOnNestedObjects As Boolean = False) As Variant
    'Considering renaming or having aliases for this method: Clone, Duplicate - might better convey the doubling of memory behind the scenes
    '
    'Creates and returns a new copy without references to original
    'Will not handle nested objects for now; possible to handle nested DataStructures recursively in the future
    Dim rtn As Variant
    Dim thing As Variant, comp As Variant, span As Variant
    Dim curkey As String, curitem As Variant
    Dim i As Long, lb As Long, ub As Long
    
    On Error GoTo Copy_clean_up
    If isDataStructure(DataStructure) Then
        Select Case TypeName(DataStructure)
            Case "Dictionary"
                Set rtn = New Scripting.Dictionary
                For Each comp In Zip(DataStructure.Keys, DataStructure.Items)
                    curkey = comp(0)
                    If IsObject(comp(1)) Then
                        If failOnNestedObjects Then Err.Raise 450
                        Set curitem = DS.Copy(comp(1)) 'Warning: recursive
                      Else
                        curitem = comp(1)
                    End If
                    rtn.Add Key:=curkey, Item:=curitem
                Next comp
                
            Case "Collection"
                Set rtn = New Collection
                For Each thing In DataStructure
                    If IsObject(thing) Then
                        If failOnNestedObjects Then
                            Err.Raise 450
                          'no need to have an Else clause as the "add" method will accept both values & references
                        End If
                    End If
                    rtn.Add thing
                Next thing
                
            Case Else
                'Catch uninitialized arrays
                span = -1
                
                On Error Resume Next
                lb = LBound(DataStructure)
                ub = UBound(DataStructure)
                
                span = ub - lb
                
                If Err.Number = 0 Then
                    ReDim rtn(LBound(DataStructure) To UBound(DataStructure))
                  Else
                    GoTo Copy_clean_up
                End If
                
                On Error GoTo Copy_clean_up
                
                For i = lb To ub
                    If IsObject(DataStructure(i)) Then
                        If failOnNestedObjects Then
                            Err.Raise 450
                          Else
                            Set rtn(i) = Me.Copy(DataStructure(i)) 'Warning: recursive
                        End If
                    End If
                Next i
        End Select
    End If
                
Copy_clean_up:

    If Err.Number <> 0 Then
        Debug.Print "Error [" & Err.Number & "]: " & Err.Description
        Err.Clear
    End If
    
    If IsObject(rtn) Then
        Set Copy = rtn
      Else
        Copy = rtn
    End If
    
End Function


Public Function Range(ByVal start_value As Variant, ByVal end_value As Variant, Optional ByVal step As Variant = 1, Optional ByVal ovf_limit As Long = 500) As Variant
On Error GoTo Range_clean_up

    Dim ovf_buffer As Iterator, elem As Iterator, repit As Iterator
    Dim col As Collection
    Dim threshold_breached As Boolean
    Dim altitude As Integer '0 = flat, -1 = descending, 1 = ascending
    Dim rtn As Variant, num As Variant
    Dim repeats As Long
    Dim factor As Long, compare_factor As Long
    Dim maxval As Variant, minval As Variant
    
    Set ovf_buffer = New Iterator
    Set elem = New Iterator
    Set repit = New Iterator
    Set col = New Collection
    rtn = Array()
    
    If (Not IsNumeric(start_value)) Or (Not IsNumeric(end_value)) Or (Not IsNumeric(step)) Then GoTo Range_clean_up
    altitude = Switch(step = 0, 0, _
                            end_value > start_value, 1, _
                            end_value < start_value, -1, _
                            end_value = start_value, 0) '0 = flat, -1 = descending, 1 = ascending
                            
    If CBool(altitude) Then
        factor = 1
        For Each num In Array(start_value, end_value, step)
            If InStr(1, CStr(num), ".") > 0 Then
                compare_factor = 10 ^ (Len(CStr(Abs(num))) - InStr(1, CStr(Abs(num)), "."))
                If compare_factor > factor Then factor = compare_factor
            End If
        Next num
        repeats = (Abs(((end_value * factor) - (start_value * factor)) / (step * factor))) + 1
        
        If ((end_value > start_value) And step < 0) _
            Or ((end_value < start_value) And step > 0) _
            Or (step = 0) Then
                Debug.Print "Warning! Supplied parameters will not converge. Loop pre-emptively limited."
        End If
      Else
        repeats = 1
    End If
    
    ovf_buffer.stairs 1, 1
    repit.stairs 1, 1
    elem.stairs start_value, step_:=step
    
    Do:
        col.Add elem.Next_
'        threshold_breached = Choose(altitude + 2, _
'                                                        (elem.value <= end_value), _
'                                                        (elem.value = end_value), _
'                                                        (elem.value >= end_value)) '+2 because choose is index base 1
'        If elem.value > 0.8 And elem.value < 0.9 Then Stop
        If VarType(elem.value) = vbDouble Then 'Fixes fuzzy floating point math
            elem.value = CDbl(CStr(elem.value))
        End If
    Loop While (repit.Next_ < repeats) And (ovf_buffer.Next_ < ovf_limit)
    
    If ovf_buffer.value > ovf_limit Then
        Debug.Print "Warning! Loop pre-emptively stopped at ovf_limit [" & ovf_limit & "]"
    End If
    
    
    With DS
        maxval = .Maximum(Array(start_value, end_value))
        minval = .Minimum(Array(start_value, end_value))
        
        rtn = .Convert(col, "Variant()")
        rtn = .Filter(rtn, "<=", maxval)
        rtn = .Filter(rtn, ">=", minval)
    End With
Range_clean_up:
    If Err.Number <> 0 Then
        Debug.Print "Class [" & TypeName(Me) & "] method [Range]"
        Debug.Print Tab(10); "Error [" & Err.Number & "]: " & Err.description
        Err.Clear
    End If
    
    Set ovf_buffer = Nothing
    Set elem = Nothing
    Set col = Nothing
    Range = rtn
End Function


Public Function Reverse(DataStructure As Variant, Optional ByVal propagate As Boolean = False) As Variant
    Dim typeNm As String
    Dim rtn As Variant, DataStructure_arr As Variant, col As Collection
    Dim lb As Integer, ub As Integer, i As Integer
    
    If Not isDataStructure(DataStructure) Then
        If IsObject(DataStructure) Then
            Set Reverse = DataStructure
          Else
            Reverse = DataStructure
        End If
        Exit Function
    End If
    
'Initialize return container to appropriate type
    typeNm = TypeName(DataStructure)
    Select Case typeNm
        Case "Dictionary"
            Set rtn = New Scripting.Dictionary
        Case "Collection"
            Set rtn = New Collection
        Case Else
            ReDim rtn(UBound(DataStructure))
    End Select
    
'Convert input into array for easy indexing
    Select Case typeNm
        Case "Dictionary"
            DataStructure_arr = Zip(DataStructure) 'Creates nested array containing key-item pairs
        Case "Collection"
            Set col = DataStructure
            DataStructure_arr = CollectionToArray(col)
        Case Else
            DataStructure_arr = DataStructure 'redundant, but uniform
    End Select
    
'Reverse elements
    ub = UBound(DataStructure_arr)
    lb = LBound(DataStructure_arr)
    For i = lb To ub
        Select Case typeNm
            Case "Dictionary"
                If propagate And (isDataStructure(DataStructure_arr(ub - i)(1))) Then
                    rtn.Add Key:=DataStructure_arr(ub - i)(0), Item:=Reverse(DataStructure_arr(ub - i)(1), propagate:=True)
                  Else
                    rtn.Add Key:=DataStructure_arr(ub - i)(0), Item:=DataStructure_arr(ub - i)(1)
                End If
            Case "Collection"
                If propagate And (isDataStructure(DataStructure_arr(ub - i))) Then
                    rtn.Add Item:=Reverse(DataStructure_arr(ub - i))
                  Else
                    rtn.Add Item:=DataStructure_arr(ub - i)
                End If
            Case Else
                If propagate And (isDataStructure(DataStructure_arr(ub - i))) Then
                    rtn(i) = Reverse(DataStructure_arr(ub - i))
                  Else
                    rtn(i) = DataStructure_arr(ub - i)
                End If
        End Select
    Next i
    
    If IsObject(rtn) Then
        Set Reverse = rtn
      Else
        Reverse = rtn
    End If
End Function


Public Function Exists(needle As Variant, haystack As Variant, Optional ByVal wildcard_needle As Boolean = False, Optional ByVal wildcard_haystack As Boolean = False) As Boolean
    Dim match_found As Boolean, typeNm As String
    Dim search_arr As Variant
    Dim col As Collection
    Dim itm As Variant
On Error GoTo Exists_clean_up
    match_found = False
    If Not isDataStructure(haystack) Then Exists = match_found: Exit Function
    
'Standardize various possible input types for variable haystacks - normalize as array
    typeNm = TypeName(haystack)
    Select Case typeNm
        Case "Dictionary" 'Dictionary has built in method for exists, but this function can expand the method to include wildcards on either the find/needle side or the search side
            search_arr = haystack.Items
        Case "Collection"
            Set col = haystack 'reference assignment is necessary here because byref parameter can't handle a reference to a 2nd reference, to a 3rd reference, etc...
            search_arr = CollectionToArray(col)
        Case Else 'Array
            search_arr = haystack
    End Select
    
    For Each itm In search_arr
        If wildcard_needle And (Not IsObject(itm)) And (Not IsObject(needle)) Then
            If CStr(itm) Like ("*" & CStr(needle) & "*") Then
                match_found = True
                Exit For
            End If
        End If
        If wildcard_haystack And (Not IsObject(itm)) And (Not IsObject(needle)) Then
            If CStr(needle) Like ("*" & CStr(itm) & "*") Then
                match_found = True
                Exit For
            End If
        End If
        If IsObject(itm) And IsObject(needle) Then
            If itm Is needle Then
                match_found = True
                Exit For
            End If
        End If
        If (Not IsObject(itm)) And (Not IsObject(needle)) Then
            If itm = needle Then
                match_found = True
                Exit For
            End If
        End If
        If ((IsObject(itm)) And (Not IsObject(needle))) _
            Or ((Not IsObject(itm)) And (IsObject(needle))) Then
            match_found = False 'explicit, noting here that the search DataStructure might be composed of BOTH objects and non-objects,
                                                '   and that the intent is to continue searching through the remainder of the for-loop
        End If
    Next itm
    
Exists_clean_up:
    If Err.Number <> 0 Then Debug.Print "Error [" & Err.Number & "] in method [Exists] in class DS: " & Err.Description: Err.Clear
    Set col = Nothing
    
    Exists = match_found
    
End Function


Public Function Enumerate(enumerable As Variant, Optional ByVal starting_idx As Variant, Optional ByVal increment As Variant) As Variant
'I.N. 2/23/22
    'Enumerable types: Array, Collection, Dictionary
    'P/N: Drops keys with dictionaries & collections; adds only the items

On Error GoTo enumerate_clean_up

    Dim typed_name As String
    Dim val As Variant, i As Integer
    Dim tmp_col As Collection
    Dim rtn As Variant

    Set tmp_col = New Collection

    If IsMissing(starting_idx) Then starting_idx = 0
    i = starting_idx
    typed_name = TypeName(enumerable)

    If Not isDataStructure(typed_name, True) Then GoTo enumerate_clean_up
    
    Select Case typed_name
        Case "Variant()", "String()", "Integer()", "Double()", _
                "Single()", "Byte()", "Long()", "Collection":
            For Each val In enumerable
                tmp_col.Add Item:=Array(i, val)
                i = i + 1
            Next val
        Case "Dictionary":
            For Each val In enumerable.Items
                tmp_col.Add Item:=Array(i, val)
                i = i + 1
            Next val
        Case Else:
            'incompatible
    End Select

    rtn = CollectionToArray(tmp_col)

enumerate_clean_up:
    If Err.Number <> 0 Then
        Debug.Print "Error [" & Err.Number & "]: " & Err.Description
        Debug.Print "   Context: " & "Function [enumerate] "
        Err.Clear
        rtn = Array()
    End If

    Set tmp_col = Nothing
    Enumerate = rtn
    
End Function


Public Function fill(ByRef container As Variant, ByRef stuff As Variant, _
                            Optional ByVal extra_serving_size As Long = -1) As Variant
    Dim rtn As Variant
    Dim keyv As Variant
    Dim i As Integer
    Dim init_dim As Variant, add_dim As Variant
    Dim container_ As Variant
    
On Error GoTo Fill_clean_up
    If Not isDataStructure(container, accept_typename:=False) Then
        On Error Resume Next
        If Not IsObject(container) Then
            If IsNumeric(container) Then
                ReDim container_(val(container) - 1)
              Else
                rtn = False
                GoTo Fill_clean_up
            End If
          Else
            rtn = False
            GoTo Fill_clean_up
        End If
'        rtn = False
'        GoTo Fill_clean_up
      Else 'DataStructure
        If IsObject(container) Then
            Set container_ = container
          Else
            container_ = container
        End If
    End If
On Error GoTo Fill_clean_up
    
    Select Case TypeName(container_)
        Case "Dictionary"
            If extra_serving_size > 0 Then
'                container.Add key:=???, item:= stuff
              Else
                
                For Each keyv In container_.Keys
                    On Error Resume Next
                    container_(keyv) = stuff
                    If Err.Number <> 0 Then
                        Err.Clear
                        On Error GoTo Fill_clean_up
                        Set container_(keyv) = stuff
                    End If
                Next keyv
            End If
        
        Case "Collection"
            If extra_serving_size > 0 Then
                For i = 1 To extra_serving_size
                    container_.Add stuff
                Next i
              Else
                For i = 1 To container_.Count
                    If IsObject(stuff) Then
                        Set container_(i) = stuff
                      Else
                        container_(i) = stuff
                    End If
                Next i
            End If

        Case Else:
            If extra_serving_size > 0 Then
                init_dim = UBound(container_)
                add_dim = init_dim + extra_serving_size
                ReDim Preserve container_(add_dim)
                For i = init_dim + 1 To add_dim
                    If IsObject(stuff) Then
                        Set container_(i) = stuff
                      Else
                        container_(i) = stuff
                    End If
                Next i
                
              Else
                For i = 0 To UBound(container_)
                    If IsObject(stuff) Then
                        Set container_(i) = stuff
                      Else
                        container_(i) = stuff
                    End If
                Next i
            End If

    End Select
    
Fill_clean_up:
    If Err.Number <> 0 Then
        Debug.Print "Error [" & Err.Number & "]: " & Err.Description
        rtn = False
      Else
        rtn = True
    End If
    
    If isDataStructure(container) Then
        'Save changes to byref container and return a boolean
        If IsObject(container_) Then
            Set container = container_
          Else
            container = container_
        End If
        
      Else
        'Return a filled array if given a value instead of an actual container to fill
        rtn = container_
    End If
    fill = rtn
End Function


Public Function Filter(DataStructure As Variant, ByVal operator_expression As String, compare_against As Variant, _
                                Optional ByRef other_paired_expressions As Variant, _
                                Optional ByVal FilterMode As eDataStructureFilterMode = eFilterTrap) As Variant
On Error GoTo Filter_clean_up
    Dim COLander As Collection
    Dim arr As Variant, val As Variant, thing As Variant
    Dim rtn As Variant
    Dim operand_isDataStructure As Boolean
    
    If Not isDataStructure(DataStructure, accept_typename:=False) Then
        GoTo Filter_clean_up
      Else
        If TypeName(DataStructure) = "Dictionary" Then
            'more complicated
          ElseIf TypeName(DataStructure) = "Collection" Then

            arr = Me.Convert(DataStructure, "Variant()")
          Else
            arr = DataStructure
        End If
    End If
    
    operand_isDataStructure = isDataStructure(compare_against)
    
    operator_expression = Trim(operator_expression)
    If Me.Exists(LCase(operator_expression), Array("in", "is in", "not in", "is not in", "inside", "between", "out", "outside", "beyond")) Then
        If operand_isDataStructure Then
            operand_isDataStructure = False 'Though technically a DataStructure operator, the intention is not to iterate on each individual element, but to check the operand as a whole
        End If
    End If
    
    Set COLander = New Collection
    Select Case LCase(operator_expression)
        Case "=", _
                "<>", _
                "is", _
                "is not", "isn't", _
                ">", ">=", _
                "<", "<=", _
                "like", _
                "not like", "liken't", _
                "in", "is in", "not in", "is not in", _
                "inside", "between", _
                "out", "outside", "beyond"
            
            If operand_isDataStructure Then
                'Check against many
                If FilterMode = eFilterTrap Then
                    'Keep going
                  Else
                    'Reverse all operands
                    'TO DO!
                End If
              Else
                If FilterMode = eFilterTrap Then
                    'Keep going
                  ElseIf FilterMode = eFilterOut Then
                    operator_expression = Match(LCase(operator_expression), "=", _
                                                                                                        "=", "<>", _
                                                                                                        "<>", "=", _
                                                                                                        "is", "is not", _
                                                                                                        "is not", "is", _
                                                                                                        ">", "<=", _
                                                                                                        "<=", ">", _
                                                                                                        "<", ">=", _
                                                                                                        ">=", "<", _
                                                                                                        "like", "not like", _
                                                                                                        "not like", "like", _
                                                                                                        "in", "out", _
                                                                                                        "out", "in")
                    'Reverse
                End If
                For Each thing In DataStructure
                    If Resolve(thing, LCase(operator_expression), compare_against) Then
                        COLander.Add thing
                      Else
                        'Don't add
                    End If
                Next thing
            End If
'
'        Case "<>"
'        Case "is", "Is", "IS" 'Objects only
'
'        Case "is not", "Is Not", "IS NOT"
'        Case ">"
'        Case ">="
'        Case "<"
'        Case "<="
'        Case "like", "Like", "LIKE"
'        Case "not like", "Not Like", "NOT LIKE"
        Case "like[or]" ', "Like[Or]", "LIKE[OR]"
        Case "like[and]" ', "Like[And]", "LIKE[AND]"
        Case "like[xor]" ', "Like[Xor]", "LIKE[XOR]"
        Case "like[xand]" ', "Like[Xand]", "LIKE[XAND]"
'        Case "in", "In", "IN"
'        Case "not in", "Not In", "NOT IN"
'        Case "between", "inside"
'        Case "beyond", "out", "outside"
        Case Else
    End Select
    
    If IsObject(DataStructure) Then
        Set rtn = Convert(COLander, TypeName(DataStructure))
      Else
        rtn = Convert(COLander, TypeName(DataStructure))
    End If
    
Filter_clean_up:
    If Err.Number <> 0 Then Debug.Print "Error [" & Err.Number & "] in method [Filter] in class DS: " & Err.description: Err.Clear
    Set COLander = Nothing
    If IsObject(rtn) Then
        Set Filter = rtn
      Else
        Filter = rtn
    End If
End Function


Public Function Flatten(embedded_DataStructure As Variant, Optional ByVal OutputType As String = "") As Variant
    Dim recursive_col As Collection
    Dim rtn As Variant
    Dim thing As Variant, validated_DataStructure As Variant
    Dim rtn_type As String
    
    If isDataStructure(embedded_DataStructure) Then
        If isDataStructure(OutputType) Then
            rtn_type = OutputType
          Else
            rtn_type = TypeName(embedded_DataStructure)
        End If
      Else
        Set Flatten = Nothing
        Exit Function
    End If
    
    Set recursive_col = New Collection
    
    If TypeName(embedded_DataStructure) = "Dictionary" Then
        validated_DataStructure = Zip(embedded_DataStructure)
      Else
        On Error Resume Next
        validated_DataStructure = embedded_DataStructure
        If Err.Number <> 0 Then
            Err.Clear
            Set validated_DataStructure = embedded_DataStructure
        End If
    End If
    
    For Each thing In validated_DataStructure
        If isDataStructure(thing) Then
            Set recursive_col = Merge(recursive_col, Flatten(thing))
          Else
            recursive_col.Add Item:=thing
        End If
    Next thing
    
    If TypeName(recursive_col) <> rtn_type Then
        On Error Resume Next
        Set rtn = Convert(recursive_col, rtn_type)
        If Err.Number <> 0 Then
            Err.Clear
            rtn = Convert(recursive_col, rtn_type)
        End If
      Else
        Set rtn = Convert(recursive_col, rtn_type)
    End If
    
End Function

Public Function GuiltyByAssociation(associated_DataStructure As Variant, DataStructure As Variant, ByVal operator_expression As String, compare_against As Variant, ParamArray other_args() As Variant) As Variant

End Function


Public Function Intersection(ParamArray DataStructures() As Variant) As Variant
    'Get common elements between 2 or more DataStructure items
    'Dictionaries & collections will be compared by items and not by keys
    'Caution - potentially expensive operation
    'Requires homogeneous parameters, i.e., same type
    
    Dim master_collection As Collection
    Dim dict As Scripting.Dictionary
    Dim DataStructure As Variant, thing As Variant
    Dim i As Long, j As Long
    Dim arr As Variant, structure As Variant, main_structure As Variant
    Dim homCol As Collection
    Dim rtn As Variant
    
On Error GoTo Intersection_clean_up
    Set homCol = New Collection
    For Each thing In DataStructures
        homCol.Add thing
    Next thing
    
    Set master_collection = New Collection 'compile master_collection as a group of like-formatted data structures - copied!! treats as immutable
    If Homogeneous(homCol) = "False" Then
        GoTo Intersection_clean_up
      Else
        For Each structure In homCol
            master_collection.Add Me.Copy(structure) 'Note the copy call here - this allows the object reference to remain unmutated
        Next structure
    End If
    
    If master_collection.Count <= 1 Then 'Haven't decided what to do about calling an intersection on a singular item
        'if preventing user from calling method on a singular item, should change the parameters to DataStructure1, DataStructure2, paramarray others() as variant
        GoTo Intersection_clean_up
    End If
    For i = 1 To master_collection.Count
        Set main_structure = master_collection(i)
        For Each thing In main_structure
            'if thing exists in each one (use the resolve method) then add it to the yes pile, otherwise add it to the no pile
            
            'then remove the thing from all structures in master_collection
            For j = (i + 1) To master_collection.Count
                
            Next j
        Next thing
    Next i
'    Select Case Homogeneous(homCol)
'        Case "False"
'            GoTo Intersection_clean_up
'        Case "Dictionary", "Collection"
'            'Proper format already
'            'COPY EACH DataStructure
'            For Each structure In homCol
'                Set thing = Me.Copy(structure)
'                master_collection.Add thing
'            Next structure
'
'        Case Else
'            For Each thing In homCol
'                master_collection.Add Me.Copy(thing)
'            Next thing
'
'    End Select

'Once all data structures are compiled, go through each one; sorting them into two separate piles
'Pile 1 contains items that exist in ALL of the data structures; Pile 2 contains items that don't exist in all data structures
'The main master_collection group should be un-stacked with each pass through in order to reduce the total iterations

       
Intersection_clean_up:
    If Err.Number <> 0 Then
        Debug.Print "Error [" & Err.Number & "]: " & Err.Description
        Err.Clear
    End If
    
    If IsObject(rtn) Then
        Set Intersection = rtn
      Else
        Intersection = rtn
    End If
    'Add to a collection of dictionaries so that the dictionaries can be paired down
End Function


Public Function Homogeneous(DataStructure As Variant, ParamArray others() As Variant) As Variant
'Check if all elements of the provided arguments are same-typed
'Returns the typename if uniform, False otherwise
    Dim rtn As Variant
    Dim val As Variant
    Dim tn As String
    Dim types As Scripting.Dictionary, curtype As Variant
    Dim all_args As Collection, thing As Variant
    
On Error GoTo Homogeneous_clean_up
    Set all_args = New Collection
    
    If UBound(others) >= 0 Then
        
        all_args.Add DataStructure
        For Each thing In others
            all_args.Add thing
        Next thing
    
      ElseIf isDataStructure(DataStructure) Then
        For Each thing In DataStructure
            all_args.Add thing
        Next thing
      ElseIf Not IsObject(DataStructure) Then
        all_args.Add DataStructure
      Else
        rtn = False
        GoTo Homogeneous_clean_up
    End If
    
    rtn = True 'uniform until proven otherwise
    Set types = New Scripting.Dictionary
    For Each thing In all_args
        curtype = TypeName(thing)
        If Not types.Exists(curtype) Then
            types.Add Key:=curtype, Item:=curtype
        End If
        
        If types.Count > 1 Then
            rtn = False
            Exit For
          Else
            rtn = curtype
        End If
    Next thing

Homogeneous_clean_up:
    If Err.Number <> 0 Then
        Debug.Print "Error [" & Err.Number & "]: " & Err.Description
        rtn = False
        Err.Clear
    End If
    
    Set types = Nothing
    Set all_args = Nothing
    
    Homogeneous = rtn
End Function


Public Function Merge(ParamArray DataStructures() As Variant) As Variant
    Dim merge_dictionary As Scripting.Dictionary
    Dim submerge As Variant, submerge_arr As Variant, comp As Variant
    Dim copy_DataStructures As Variant
    Dim output_type As String
    Dim i As Integer, j As Integer
    Dim basekey As String, key_ As String, addkey As String
    Dim dct As Variant, part As Variant
    
    Set merge_dictionary = New Scripting.Dictionary
    
    output_type = TypeName(DataStructures(0))
    If Not isDataStructure(output_type, accept_typename:=True) Then Exit Function
    
    For Each submerge In DataStructures
        i = 1
        If TypeName(submerge) = "Dictionary" And output_type = "Dictionary" Then
            'Splice together keys and items to the main dictionary while retaining keys
            For Each comp In Zip(submerge)
                j = 0
                basekey = comp(0)
                addkey = basekey
                While merge_dictionary.Exists(addkey)
                    addkey = basekey & CStr(j)
                    j = j + 1
                Wend
                merge_dictionary.Add Key:=addkey, Item:=comp(1)
            Next comp
          Else
            'Convert everything to an array to add to dictionary
            If TypeName(submerge) = "Dictionary" Then
                submerge_arr = submerge.Items
              Else
                submerge_arr = submerge
            End If
            
            For Each part In submerge_arr
                merge_dictionary.Add Key:=i, Item:=part
                i = i + 1
            Next part
            
        End If
    Next submerge
    
    Select Case output_type
        Case "Dictionary"
            Set Merge = merge_dictionary
        Case "Collection"
            Set Merge = ArrayToCollection(merge_dictionary.Items)
        Case Else
            Merge = merge_dictionary.Items
    End Select

End Function

Public Function Minimum(DataStructure As Variant, Optional ByVal string_support As Boolean = False) As Variant
'String support will enable string length comparisons string lengths of provided elements
    Dim val As Variant
    Dim col As Collection
    Dim numeric_composition As Boolean
    Dim string_composition As Boolean
    Dim lowest As Variant
    Dim DataStructure_copy As Variant
    
    If Not isDataStructure(DataStructure) Then GoTo Minimum_clean_up
    If DataStructureCount(DataStructure) < 1 Then GoTo Minimum_clean_up
    
    Select Case TypeName(DataStructure)
        Case "Dictionary"
            DataStructure_copy = DataStructure.Items
        Case "Collection"
            Set col = DataStructure
            DataStructure_copy = CollectionToArray(col)
        Case Else
            DataStructure_copy = DataStructure
    End Select
    
    For Each val In DataStructure_copy
        If Not IsNumeric(val) Then
            numeric_composition = False
            Exit For
        End If
    Next val
    
    lowest = DataStructure_copy(0)
    If numeric_composition Then
        For Each val In DataStructure_copy
            If val < lowest Then
                lowest = val
            End If
        Next val
    End If
    
Minimum_clean_up:
    If Err.Number <> 0 Then Debug.Print "Error [" & Err.Number & "]: " & Err.Description: Err.Clear
    Set col = Nothing
    Minimum = lowest
End Function

Public Function Maximum(DataStructure As Variant) As Variant
'String support will enable string length comparisons string lengths of provided elements
    Dim val As Variant
    Dim col As Collection
    Dim numeric_composition As Boolean
    Dim string_composition As Boolean
    Dim highest As Variant
    Dim DataStructure_copy As Variant
    
    If Not isDataStructure(DataStructure) Then GoTo Maximum_clean_up
    If DataStructureCount(DataStructure) < 1 Then GoTo Maximum_clean_up
    
    Select Case TypeName(DataStructure)
        Case "Dictionary"
            DataStructure_copy = DataStructure.Items
        Case "Collection"
            Set col = DataStructure
            DataStructure_copy = CollectionToArray(col)
        Case Else
            DataStructure_copy = DataStructure
    End Select
    
    For Each val In DataStructure_copy
        If Not IsNumeric(val) Then
            numeric_composition = False
            Exit For
        End If
    Next val
    
    highest = DataStructure_copy(0)
    If numeric_composition Then
        For Each val In DataStructure_copy
            If val > highest Then
                highest = val
            End If
        Next val
    End If
    
Maximum_clean_up:
    If Err.Number <> 0 Then Debug.Print "Error [" & Err.Number & "]: " & Err.Description: Err.Clear
    Set col = Nothing
    Maximum = highest

End Function

Public Function Ones(ParamArray dimensions() As Variant) As Variant
'Returns an array with size based on the dimensions provided
'With square brackets denoting an array...
    'Ex: Ones(2, 4, 2) = [ _
                                        [ _
                                            [1, 1], _
                                            [1, 1], _
                                            [1, 1], _
                                            [1, 1] _
                                        ], _
                                        [ _
                                            [1, 1], _
                                            [1, 1], _
                                            [1, 1], _
                                            [1, 1] _
                                        ] _
                                    ]
        'i.e., 2 Arrays both comprised of 4 arrays, each of which contains 2 1-elements
'    while me.Fill
End Function


Public Function Outersection(ParamArray DataStructures() As Variant) As Variant
    'Returns the uncommon elements
End Function


Public Function Pop(ByRef DataStructure As Variant, Optional ByVal Index As Long) As Variant
    'Return the item indicated by the identifier; mutate the reference to exclude the element denoted by identifier
    Dim rtn As Variant, mutated As Variant
    Dim arr As Variant, comp As Variant
    Dim i As Integer, j As Integer, keyrem As Variant
    
    If Not isDataStructure(DataStructure) Then
        GoTo Pop_clean_up
    End If
    
    Select Case TypeName(DataStructure)

        Case "Dictionary"
            If IsMissing(Index) Then Index = 0
            If Index = -1 Then Index = DataStructure.Count - 1 'Last element
            
            If DataStructure.Count >= Index + 1 Then
                arr = Me.Zip(DataStructure.Keys, DataStructure.Items)
                keyrem = arr(Index)(0)
                If IsObject(arr(Index)(1)) Then
                    Set rtn = arr(Index)(1)
                  Else
                    rtn = arr(Index)(1)
                End If
                DataStructure.Remove Key:=keyrem
            End If
            
        Case "Collection"
            If IsMissing(Index) Then Index = 1
            If Index = -1 Then Index = DataStructure.Count
            If Index = 0 Then Index = 1 'Collections are base 1
            If DataStructure.Count >= Index Then
                If IsObject(DataStructure(Index)) Then
                    Set rtn = DataStructure(Index)
                  Else
                    rtn = DataStructure(Index)
                End If
            End If
            
            If DataStructure.Count > 0 Then
                Set DataStructure = Me.Remove(DataStructure, Index - 1) 'Method [Remove] uses index base 0
            End If

        Case Else 'Array
            If IsMissing(Index) Then Index = 0
            If Index = -1 Then Index = UBound(DataStructure)
            If UBound(DataStructure) = 0 Then
                rtn = DataStructure(0)
                DataStructure = Array()
                GoTo Pop_clean_up
            End If
            ReDim mutated(UBound(DataStructure) - 1)
            rtn = DataStructure(Index)
            j = 0 'index of last addition to mutated array, starting from 0
            For i = 0 To UBound(DataStructure)
                If i <> Index Then
                    mutated(j) = DataStructure(i)
                    j = j + 1
                End If
            Next i
            If j <> i - 1 Then
                'uh oh
                'don't overwrite
              Else
                DataStructure = mutated
            End If
    End Select
    
Pop_clean_up:
    If Err.Number <> 0 Then
        Debug.Print "Error [" & Err.Number & "]: " & Err.Description
        Err.Clear
    End If
    
    If IsObject(rtn) Then
        Set Pop = rtn
      Else
        Pop = rtn
    End If

End Function

Public Function Remove(DataStructure As Variant, Optional ByVal Index As Integer = 0) As Variant
    Dim col As Collection
    Dim i As Integer
    Dim arr As Variant, rtn As Variant
    Dim rkey As String
    
    If Not isDataStructure(DataStructure) Then
        On Error Resume Next
        Remove = DataStructure
        If Err.Number <> 0 Then Err.Clear: Set Remove = DataStructure
        Exit Function
    End If
    
    Select Case TypeName(DataStructure)
        Case "Dictionary"
            rkey = DataStructure.Keys(Index)
            Set rtn = DataStructure.Remove(rkey)
        Case "Collection"
            Set col = DataStructure
            arr = CollectionToArray(col)
            Set rtn = ArrayToCollection(Me.Remove(arr, Index))
        Case Else
            Set col = New Collection
            
            For i = 0 To UBound(DataStructure)
                If i <> Index Then col.Add DataStructure(i)
            Next i
            rtn = CollectionToArray(col)
    End Select
    
    If IsObject(rtn) Then
        Set Remove = rtn
      Else
        Remove = rtn
    End If
End Function

Public Function Transpose(ParamArray DataStructures() As Variant) As Variant

End Function

Public Function Zip(ParamArray DataStructures() As Variant) As Variant
    On Error GoTo Zip_clean_up
    Dim rtn As Variant, copy_DataStructures As Variant, DataStructure As Variant, thing As Variant
    Dim output_col As Collection, tcol As Collection
    Dim i As Integer, j As Integer, thisCount As Integer, lastCount As Integer, itCount As Integer

'Make sure all parameters are the same count/size
    lastCount = -2 'arbitrary, but not -1 since ubound of empty array is -1
    For Each DataStructure In DataStructures
        thisCount = DataStructureCount(DataStructure)
        
        If lastCount = -2 Then
            lastCount = thisCount
        End If
        
        If thisCount <> lastCount Then
            'Incompatible
            Zip = Empty
            Exit Function
        End If
    Next DataStructure
    itCount = thisCount 'Use later to output and size the return array to the span (itCount) of the input DataStructures
    
    'Special case for standalone dictionaries; dictionaries are compacted and returned
    If UBound(DataStructures) = 0 Then
        If TypeName(DataStructures(0)) = "Dictionary" Then
            With DataStructures(0)
                rtn = zipArray_(.Keys, .Items)
            End With
            Zip = rtn
            Exit Function
        End If
    End If
    
    ReDim copy_DataStructures(UBound(DataStructures))
    For i = 0 To UBound(DataStructures)
        On Error Resume Next
        copy_DataStructures(i) = DataStructures(i)
        If Err.Number <> 0 Then
            Err.Clear
            If TypeName(DataStructures(i)) = "Collection" Then
                Set copy_DataStructures(i) = DataStructures(i)
              ElseIf TypeName(DataStructures(i)) = "Dictionary" Then
                copy_DataStructures(i) = Zip(DataStructures(i))
              Else
                Debug.Print "Error in Zip (not a standard raised error): the indexed retrieval errored out, and the object passed doesn't qualify as a Collection or Dictionary"
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next i
    
    On Error GoTo Zip_clean_up
    Set output_col = New Collection
    
    ReDim rtn(itCount - 1)
    
    For j = 0 To itCount - 1
        Set tcol = New Collection
        For Each thing In copy_DataStructures
            Select Case TypeName(thing)
                Case "Collection"
                    tcol.Add Item:=thing(j + 1)
                Case "Dictionary" 'uh oh
                    tcol.Add Item:=thing(thing.Keys(j))
                Case Else
                    tcol.Add Item:=thing(j)
            End Select
'            Stop
        Next thing
        rtn(j) = CollectionToArray(tcol)
    Next j
Zip_clean_up:
    If Err.Number <> 0 Then Debug.Print "Error [" & Err.Number & "]: " & Err.Description
    Zip = rtn
End Function


'PRIVATE AUXILIARY METHODS
Private Function DataStructureCount(ByRef DataStructure As Variant) As Integer
On Error GoTo DataStructureCount_clean_up
    Dim i As Integer, val As Variant
    i = 0
    For Each val In DataStructure
        i = i + 1
    Next val
DataStructureCount_clean_up:
    If Err.Number <> 0 Then Err.Clear
'    If Err.number <> 0 Then Debug.Print "Error [" & Err.number & "]: " & Err.description
    DataStructureCount = i
End Function


Private Function DataStructureMaxSpan(ByRef DataStructure As Variant) As Long

End Function


Private Function DataStructureToArray(DataStructure As Variant) As Integer

End Function


Private Function isDataStructure(DataStructure As Variant, Optional ByVal accept_typename As Boolean = True) As Boolean
    Dim typnm As String
    Dim rtn As Boolean: rtn = False
'Note:
'       Accepts Syntax:
'1.           isDataStructure("Variant()")
'2.          isDataStructure(object)
    If accept_typename And TypeName(DataStructure) = "String" Then
        typnm = DataStructure
      Else
        typnm = TypeName(DataStructure)
    End If
    Select Case typnm 'Accepts string name of DataStructure type as well as the DataStructure itself
        Case "Variant()", "Long()", _
                    "Single()", "Byte()", _
                    "Integer()", "Double()", _
                    "String()", "Boolean()", _
                    "Collection", "Dictionary"
            rtn = True
        Case Else
            rtn = False
    End Select
    isDataStructure = rtn
End Function


Private Function CollectionToArray(DataStructure_col As Collection) As Variant
    Dim arr As Variant
    Dim i As Integer
    If DataStructure_col.Count > 0 Then
        ReDim arr(DataStructure_col.Count - 1)
        For i = 0 To UBound(arr)
            If IsObject(DataStructure_col(i + 1)) Then
                Set arr(i) = DataStructure_col(i + 1)
              Else
                arr(i) = DataStructure_col(i + 1)
            End If
        Next i
      Else
        arr = Array()
    End If
    CollectionToArray = arr
End Function


Private Function DictionaryToArray(DataStructure_dictionary As Scripting.Dictionary) As Variant
    Dim comp As Variant, keyarr As Variant, itemarr As Variant
    Dim arr() As Variant
    Dim i As Integer
    If DataStructure_dictionary.Count > 0 Then
        ReDim arr(DataStructure_dictionary.Count - 1)
        keyarr = DataStructure_dictionary.Keys
        itemarr = DataStructure_dictionary.Items
        For i = 0 To (DataStructure_dictionary.Count - 1)
            arr(i) = Array(keyarr(i), itemarr(i))
        Next i
    End If
    DictionaryToArray = arr
End Function


Private Function ArrayToCollection(DataStructure As Variant, Optional ByVal populateKeys As Boolean = True) As Collection
    Dim rtn_col As Collection
    Dim val As Variant
    Set rtn_col = New Collection
    
    If IsArray(DataStructure) Then
        On Error Resume Next
        If UBound(DataStructure) >= 0 Then
            If Err.Number <> 0 Then
                'No items to add
              Else
                For Each val In DataStructure
                    rtn_col.Add val
                Next val
            End If
        End If
    End If
    
    Set ArrayToCollection = rtn_col
End Function


Private Function ArrayToDictionary(DataStructure As Variant, Optional ByVal populateKeys As Boolean = True) As Scripting.Dictionary
    Dim rtn_dictionary As Scripting.Dictionary
    Dim nextkey As String, basekey As String
    Dim val As Variant, comp As Variant
    Dim keyidx_suffix As Integer, i As Integer
    Set rtn_dictionary = New Scripting.Dictionary
    
    If IsArray(DataStructure) Then
        If IsArray(DataStructure(0)) Then
            If populateKeys Then
                For Each comp In DataStructure
                    keyidx_suffix = 0
                    basekey = CStr(comp(0))
                    nextkey = basekey
                    While rtn_dictionary.Exists(nextkey)
                        nextkey = basekey & CStr(keyidx_suffix)
                        keyidx_suffix = keyidx_suffix + 1
                    Wend
                    rtn_dictionary.Add Key:=nextkey, Item:=Remove(comp, Index:=0)
                Next comp
              Else 'ignore keys, use regular indexing
                i = 0
                For Each val In DataStructure
                    rtn_dictionary.Add Key:=i, Item:=val
                    i = i + 1
                Next val
            End If 'populate keys
          Else
            i = 0
            For Each val In DataStructure
                rtn_dictionary.Add Key:=i, Item:=val
                i = i + 1
            Next val
        End If 'isArray
      Else
        'input isn't an array - don't do anything
    End If '

ArrayToDictionary_clean_up:
    If Err.Number <> 0 Then Debug.Print "Error [" & Err.Number & "]: " & Err.Description
    
    Set ArrayToDictionary = rtn_dictionary
    
End Function

Private Function zipArray_(ParamArray args() As Variant) As Variant()
'Released I.N. - 2/23/22

    Dim arg As Variant, sub_arg As Variant
    Dim ub_arg As Long, ub_args As Integer, lb_arg As Long, lb_args As Integer
    Dim sz_of_sub_rtn As Integer, sz_of_rtn As Long
    Dim rtn() As Variant, sub_rtn() As Variant
    Dim i As Long, j As Integer
    Dim valid_args As Boolean: valid_args = True
    
'Example usage:
'Input:
'   A:= array(1, 2, 3)
'   B:= array(4, 5, 6)
'   C:= array(7, 8, 9)

'Output:
'   Return= array(array(1, 4, 7), array(2, 5, 8), array(3, 6, 9))
    
'Check first if they're all the same size
    For Each arg In args
        valid_args = IsArray(arg)
        If valid_args Then If UBound(args) = -1 Then valid_args = False Else valid_args = True
        If Not valid_args Then Debug.Print "Argument is not an array.": zipArray_ = Array(Null): Exit Function
    Next arg
    
'Number of params provided in array
    ub_args = UBound(args)
    lb_args = LBound(args)
    sz_of_sub_rtn = ub_args - lb_args
    
'Check if sub-arrays are the same size
    ub_arg = UBound(args(0))
    lb_arg = LBound(args(0))
    sz_of_rtn = ub_arg - lb_arg
    
    For Each arg In args
        If UBound(arg) <> ub_arg Or LBound(arg) <> lb_arg Then valid_args = False
        If Not valid_args Then
            Debug.Print "Array bounds do not match."
            zipArray_ = Array(Null)
            Exit Function
        End If
    Next arg
    
'Arguments are valid - continue...

'Resize sub arrays and array that will be returned
    ReDim rtn(sz_of_rtn)
    ReDim sub_rtn(sz_of_sub_rtn)

'Combine the 1st element of each array, then the 2nd... etc..
    For i = 0 To sz_of_rtn
        For j = 0 To sz_of_sub_rtn

On Error Resume Next
'Try value assignment
            sub_rtn(j) = args(j)(i)
            If Err.Number <> 0 Then
                Err.Clear
'Try reference object
                Set sub_rtn(j) = args(j)(i)
            End If
'Finally
On Error GoTo 0

        Next j
        
    'Add the combined array to the stack of arrays
        rtn(i) = sub_rtn
    Next i
    
'Return the result
    zipArray_ = rtn
End Function

Private Function Resolve(left_side As Variant, operator As String, right_side As Variant) As Boolean
On Error GoTo Resolve_clean_up
    Dim vali As Variant
    Dim lowercase_operator As String, mini_ As Variant, maxi_ As Variant
    lowercase_operator = LCase(operator)
    Dim rtn As Boolean
    
    rtn = False
    
    Select Case lowercase_operator
        Case "="
            rtn = (left_side = right_side)
        Case "<>"
            rtn = (left_side <> right_side)
        Case "is", "Is", "IS" 'Objects only
            If IsObject(left_side) And IsObject(right_side) Then
                rtn = (left_side Is right_side)
            End If
        Case "is not"
            If IsObject(left_side) And IsObject(right_side) Then
                rtn = (left_side Is Not right_side)
            End If
        Case ">"
            If IsNumeric(left_side) And IsNumeric(right_side) Then
                rtn = (left_side > right_side)
            End If
        Case ">="
            If IsNumeric(left_side) And IsNumeric(right_side) Then
                rtn = (left_side >= right_side)
            End If
        Case "<"
            If IsNumeric(left_side) And IsNumeric(right_side) Then
                rtn = (left_side < right_side)
            End If
        Case "<="
            If IsNumeric(left_side) And IsNumeric(right_side) Then
                rtn = (left_side <= right_side)
            End If
        Case "in", "is in"
            rtn = Exists(left_side, right_side)
        Case "like"
            rtn = (CStr(left_side) Like ("*" & CStr(right_side) & "*"))
        Case "not like"
            rtn = (Not (CStr(left_side) Like ("*" & CStr(right_side) & "*")))
        Case "between"
            If Not IsArray(right_side) Then GoTo Resolve_clean_up
            If UBound(right_side) <> 1 Then GoTo Resolve_clean_up
            rtn = ((left_side >= Minimum(right_side)) And (left_side <= Maximum(right_side)))
        Case "outside"
            If Not IsArray(right_side) Then GoTo Resolve_clean_up
            If UBound(right_side) <> 1 Then GoTo Resolve_clean_up
            rtn = ((left_side < Minimum(right_side)) Or (left_side > Maximum(right_side)))
        Case Else
    End Select
Resolve_clean_up:

Resolve = rtn
End Function

