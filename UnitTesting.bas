Attribute VB_Name = "UnitTesting"
Option Compare Database
Option Explicit

Sub UnitTest_DS_Append()
    Dim Test As Scripting.Dictionary
    Dim test_container As Collection
    
    Dim arr As Variant, arr1 As Variant, arr2 As Variant
    Dim col_3 As Object, col_6 As Collection, col_7 As Collection, dict As Scripting.Dictionary, dict1 As Scripting.Dictionary, subdict As Scripting.Dictionary
    Dim appendix As Variant, appended_obj As Object, nothing_obj As Variant
    Dim obj_8 As Object
    Dim dict_10_1 As Scripting.Dictionary, dict_10_2 As Scripting.Dictionary
    Dim initial_ub As Long
    Dim final_ub As Long
    Dim test_number As Integer
    Dim test_description As String
    Dim test_cases As Variant
    Dim expected As Variant
    Dim expected_as_str
    
    Dim some_integer As Integer, i As Integer, j As Integer
    Dim output As Variant, test_result As Variant, result As Variant

    Set obj_8 = New Scripting.Dictionary
    Set Test = New Scripting.Dictionary
    Set appended_obj = New Collection: appended_obj.Add "d"
    Set col_3 = New Collection: col_3.Add "this is an object"
    Set col_6 = New Collection: col_6.Add 1: col_6.Add 2: col_6.Add 3
    Set col_7 = New Collection: col_7.Add "a"
    Set dict_10_1 = New Scripting.Dictionary: dict_10_1.Add "Base Dictionary Test 10", 10
    Set dict_10_2 = New Scripting.Dictionary: dict_10_2.Add "Appended Dictionary Test 10", 10
    Set dict = New Scripting.Dictionary: dict.Add "a", 1: dict.Add "b", 2: dict.Add "c", 3
    Set subdict = New Scripting.Dictionary: subdict.Add Key:="a", Item:=1
    Set dict1 = New Scripting.Dictionary: dict1.Add Key:="1", Item:=subdict
    Set nothing_obj = Nothing
    some_integer = 1
    
    test_cases = Array( _
                            Array("Test Number", "Description", "Given Data Structure", "Appendix", "Initial Count", "Final Count", "<Optional> Expected Equivalent"), _
 _
                            Array(1, "Append a number to the end of an array", Array(1, 2, 3, 4), 5, 4, 5, Array(1, 2, 3, 4, 5)), _
                            Array(2, "Append a string to an empty array", Array(), "lone string in empty array", 0, 1, Array("lone string in empty array")), _
                            Array(3, "Append an object (collection) to an array", Array("a", "b", "c"), col_3, 3, 4, Array("a", "b", "c", col_3)), _
                            Array(4, "Append an empty string to an array", Array("a", "b", "c"), "", 3, 4, Array("a", "b", "c", "")), _
                            Array(5, "Append an array to an array", Array(1, 2, 3), Array(4, 5, 6), 3, 4, Array(1, 2, 3, Array(4, 5, 6))), _
                            Array(6, "Append a number to a Collection", col_6, 4, 3, 4), _
                            Array(7, "Append Nothing to a collection", col_7, nothing_obj, 1, 2), _
                            Array(8, "Append Something to Nothing", nothing_obj, obj_8, 0, 0, nothing_obj), _
                            Array(9, "Append 1 2-element array to a dictionary", dict, Array("d", 4), 3, 4), _
                            Array(10, "Append a dictionary to another dictionary", dict_10_1, dict_10_2, 1, 2), _
                            Array(11, "Append a number to a non-data structure", some_integer, 1, 0, 0) _
                        )
    Set test_container = New Collection
    
    For i = 1 To UBound(test_cases)
        Set Test = New Scripting.Dictionary
        For j = 0 To UBound(test_cases(i))
            Select Case j
                Case 0, 1 'Title, Description
                    result = DS.Copy(i)(j)
                Case 2 'Given data structure
                    If IsObject(test_cases(i)(j)) Then
                        Set result = DS.Copy(test_cases(i)(j))
                      Else
                        result = DS.Copy(test_cases(i)(j))
                    End If
                Case 3 'Appendix
                    If IsObject(test_cases(i)(j)) Then
                        Set result = test_cases(i)(j)
                      Else
                        result = test_cases(i)(j)
                    End If
                Case 4 'Initial element count
                    Select Case TypeName(test_cases(i)(2))
                        Case "Dictionary", "Collection"
                            result = test_cases(i)(2).Count
                        Case "Variant()"
                            result = UBound(test_cases(i)(2)) + 1
                        Case Else
                            result = 0
                    End Select
                Case 5 'Final element count
                    DS.Append test_cases(i)(2), test_cases(i)(3)
                    Select Case TypeName(test_cases(i)(2))
                        Case "Dictionary", "Collection"
                            result = test_cases(i)(2).Count
                        Case "Variant()"
                            result = UBound(test_cases(i)(2)) + 1
                        Case Else
                            result = 0
                    End Select
                Case 6 'Expected result (Optional)
                    result = test_cases(i)(j)
            End Select
            
            Test.Add Key:=Test(0)(j), Item:=result
        Next j
        test_container.Add Item:=Test
        
    Next i
'("Test Number", "Description", "Given Data Structure", "Appendix", "Initial Count", "Final Count", "<Optional> Expected Equivalent")

    Dim printer_stack As Collection
    Dim things_left As Integer, needs_flattening As Boolean

    
    For Each test_result In test_container
        Select Case (TypeName(test_result("Given Data Structure")))
            Case "Collection", "Dictionary"
                things_left = test_result("Given Data Structure").Count
            Case "Variant()"
                things_left = UBound(test_result("Given Data Structure")) + 1
            Case Else
                things_left = 0
        End Select
        
        Debug.Print "[" & test_result("Test Number") & "] - " & test_result("Description")
        Set printer_stack = New Collection
        
'        While (things_left > 0) Or needs_flattening
'            If needs_flattening Then
                
        Select Case TypeName(test_result("Given Data Structure"))
            Case "Dictionary"
                
            Case "Collection"
            Case "Variant()"
            Case Else
        End Select
        
'        For Each thing In test_result("Given Data Structure")
'        While UBound(things_to_string) > -1
    Next test_result
        
        
        
        
        
'    test.Add Key:="Number", Item:=i
'    test.Add Key:="Description", Item:=descriptions(i)
'    test.Add Key:="Given data structure", Item:=data_structures(i)
'    test.Add Key:="Appendix", Item:=appendices(i)
'    test.Add Key:="Initial Upper Bound", Item:=upper_ubounds(i)
'    test.Add Key:="Final Upper Bound", Item:=final_ubounds(i)
    
'Test Case:
    test_number = 1
    test_description = "Append a number to the end of an array, mutate-in-place"
    Debug.Print "Test Case [" & test_number & "]"
    Debug.Print Tab(10); test_description
    arr1 = Array(1, 2, 3, 4)
    initial_ub = UBound(arr1)
    appendix = 5
    DS.Append arr1, appendix
    final_ub = UBound(arr1)
    Debug.Assert ((UBound(arr1) = initial_ub + 1) = final_ub)
    Debug.Assert (arr1(final_ub) = appendix)
    Debug.Assert DS.Equivalent(arr1, Array(1, 2, 3, 4, 5))
    Debug.Print "Result: Passed!"
    
'Test Case:
    test_number = 2
    test_description = "Append a number to the end of an array, return the array (still mutates)"
    Debug.Print "Test Case [" & test_number & "]"
    Debug.Print Tab(10); test_description
    arr1 = Array(1, 2, 3, 4, 5, 6)
    initial_ub = UBound(arr1)
    appendix = 5
    arr2 = DS.Append(arr1, appendix)
    final_ub = UBound(arr2)
    Debug.Assert ((UBound(arr2) = initial_ub + 1) = final_ub)
    Debug.Assert (arr2(final_ub) = appendix)
    Debug.Assert DS.Equivalent(arr2, arr1)
    Debug.Print "Test [" & test_number & "]: " & test_description & " - Passed."
        
'Test Case:
    test_number = 1
    test_description = "Append a number to the end of an array, mutate-in-place"
    Debug.Print "Test Case [" & test_number & "]"
    Debug.Print Tab(10); test_description
    arr = Array(1, 2, 3, 4)
    initial_ub = UBound(arr)
    appendix = 5
    DS.Append arr, appendix
    final_ub = UBound(arr)
    Debug.Assert ((UBound(arr) = initial_ub + 1) = final_ub)
    Debug.Assert (arr(final_ub) = appendix)
    Debug.Assert DS.Equivalent(arr, Array(1, 2, 3, 4, 5))
    Debug.Print "Test [" & test_number & "]: " & test_description & " - Passed."
End Sub

Sub UnitTest_DS_Apply()
Dim arr As Variant
Dim splits As Variant
Dim obj As Object

arr = Array("abc", "def", "ghi", "jkl")

'splits = DS.Apply(arr, "DS.CharacterArray", 0)
Stop
End Sub

Sub Test()
    
    Dim oSC As Object
    Dim abc As String
    Dim def As String
    Dim arr1 As Variant, piece As Variant
    Dim arr_str As String
    Dim carry As Variant
    Set oSC = CreateObjectx86("ScriptControl")
    oSC.Language = "VBScript"
    abc = "abc"
    def = "def"
    Dim str_col As Collection
    Set str_col = New Collection
    
    For Each piece In Array(abc, def)
        carry = DS.CharacterArray(piece)
        carry = DS.Apply(carry, "CStr", 0)
        carry = DS.Apply(carry, "Asc", 0)
        carry = DS.Prefixed(carry, "Chr(")
        carry = DS.Postfixed(carry, ")")
        str_col.Add Join(carry, " & ")
    Next piece

    arr1 = DS.Convert(str_col, "Variant()")
    Debug.Print oSC.Eval("Join(Array(" & Join(arr1, ", ") & "), Empty)")
    
End Sub

Function CreateObjectx86(Optional sProgID)
Static oWnd As Object
Static bRunning As Boolean
#If Win64 Then
    bRunning = InStr(TypeName(oWnd), "HTMLWindow") > 0
    Select Case True
        Case IsMissing(sProgID)
            If bRunning Then oWnd.lost = False
            Exit Function
        Case Not bRunning
            Set oWnd = CreateWindow()
            oWnd.execScript "Function CreateObjectx86(sProgID): Set CreateObjectx86 = CreateObject(sProgID) End Function", "VBScript"
            oWnd.execScript "var Lost, App;": Set oWnd.App = Application
            oWnd.execScript "Sub Check(): On Error Resume Next: Lost = True: App.Run(""CreateObjectx86""): If Lost And (Err.Number = 1004 Or Err.Number = 0) Then close: End If End Sub", "VBScript"
            oWnd.execScript "setInterval('Check();', 500);"
    End Select
    Set CreateObjectx86 = oWnd.CreateObjectx86(sProgID)
#Else
    Set CreateObjectx86 = CreateObject(sProgID)
#End If

End Function

Function CreateWindow()
    
    ' source http://forum.script-coding.com/viewtopic.php?pid=75356#p75356
    Dim sSignature, oShellWnd, oProc
    
    On Error Resume Next
    Do Until Len(sSignature) = 32
        sSignature = sSignature & Hex(Int(Rnd * 16))
    Loop
    CreateObject("WScript.Shell").Run "%systemroot%\syswow64\mshta.exe about:""<head><script>moveTo(-32000,-32000);document.title='x86Host'</script><hta:application showintaskbar=no /><object id='shell' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shell.putproperty('" & sSignature & "',document.parentWindow);</script></head>""", 0, False
    Do
        For Each oShellWnd In CreateObject("Shell.Application").Windows
            Set CreateWindow = oShellWnd.GetProperty(sSignature)
            If Err.Number = 0 Then Exit Function
            Err.Clear
        Next
    Loop
    
End Function

Sub UnitTest_CharacterArray()

End Sub


Sub UnitTest_DS_Equivalent()
    Dim thing1 As Variant, thing2 As Variant
    Dim test_number As Integer
    Dim test_description As String
    Dim common_object1 As Object, common_object2 As Object, common_object3 As Object
    
'Test Case:
    test_number = 1
    test_description = "Standard equal arrays"
    Debug.Print "Test Case [" & test_number & "]"
    Debug.Print Tab(10); test_description
    thing1 = Array(1, 2, 3, 4)
    thing2 = Array(1, 2, 3, 4)
    Debug.Assert DS.Equivalent(thing1, thing2)
    Debug.Print "Result: Passed!"
    
'Test Case:
    test_number = 2
    test_description = "Standard Non-equal arrays"
    Debug.Print "Test Case [" & test_number & "]"
    Debug.Print Tab(10); test_description
    thing1 = Array(1, 2, 3, 4)
    thing2 = Array(1, 2, 5, 6)
    Debug.Assert Not DS.Equivalent(thing1, thing2)
    Debug.Print "Result: Passed!"
    
'Test Case:
    test_number = 3
    test_description = "Standard Different sized arrays"
    Debug.Print "Test Case [" & test_number & "]"
    Debug.Print Tab(10); test_description
    thing1 = Array(1, 2, 3, 4)
    thing2 = Array(1, 2, 3, 4, 5)
    Debug.Assert Not DS.Equivalent(thing1, thing2)
    Debug.Print "Result: Passed!"
    
'Test Case:
    test_number = 4
    test_description = "Different type array-collection"
    Debug.Print "Test Case [" & test_number & "]"
    Debug.Print Tab(10); test_description
    thing1 = Array(1, 2, 3, 4)
    Set thing2 = New Collection
    thing2.Add 1: thing2.Add 2: thing2.Add 3: thing2.Add 4
    Debug.Assert Not DS.Equivalent(thing1, thing2)
    Debug.Print "Result: Passed!"
    
'Test Case:
    test_number = 5
    test_description = "Equivalent Dictionaries"
    Debug.Print "Test Case [" & test_number & "]"
    Debug.Print Tab(10); test_description
    Set thing1 = New Scripting.Dictionary
    Set thing2 = New Scripting.Dictionary
    thing1.Add Key:="Key1", Item:=1: thing1.Add Key:="Key2", Item:=2: thing1.Add Key:="Key3", Item:=3
    thing2.Add Key:="Key1", Item:=1: thing2.Add Key:="Key2", Item:=2: thing2.Add Key:="Key3", Item:=3
    Debug.Assert DS.Equivalent(thing1, thing2)
    Debug.Print "Result: Passed!"
    
'Test Case:
    test_number = 6
    test_description = "Non-Equivalent Dictionaries with differing keys"
    Debug.Print "Test Case [" & test_number & "]"
    Debug.Print Tab(10); test_description
    Set thing1 = New Scripting.Dictionary
    Set thing2 = New Scripting.Dictionary
    thing1.Add Key:="Key1", Item:=1: thing1.Add Key:="Key2", Item:=2: thing1.Add Key:="Key3", Item:=3
    thing2.Add Key:="key2", Item:=1: thing2.Add Key:="Key3", Item:=2: thing2.Add Key:="Key4", Item:=3
    Debug.Assert Not DS.Equivalent(thing1, thing2)
    Debug.Print "Result: Passed!"
    
'Test Case:
    test_number = 7
    test_description = "Non-Equivalent Dictionaries with differing items"
    Debug.Print "Test Case [" & test_number & "]"
    Debug.Print Tab(10); test_description
    Set thing1 = New Scripting.Dictionary
    Set thing2 = New Scripting.Dictionary
    thing1.Add Key:="Key1", Item:=1: thing1.Add Key:="Key2", Item:=2: thing1.Add Key:="Key3", Item:=3
    thing2.Add Key:="key1", Item:=2: thing2.Add Key:="Key2", Item:=3: thing2.Add Key:="Key3", Item:=4
    Debug.Assert Not DS.Equivalent(thing1, thing2)
    Debug.Print "Result: Passed!"
    
'Test Case:
    test_number = 8
    test_description = "Equivalent arrays with common element object references"
    Debug.Print "Test Case [" & test_number & "]"
    Debug.Print Tab(10); test_description
    Set common_object1 = New Scripting.Dictionary
    Set common_object2 = common_object1
    thing1 = Array(common_object1, common_object2)
    thing2 = Array(common_object1, common_object2)
    Debug.Assert DS.Equivalent(thing1, thing2)
    Debug.Print "Result: Passed!"
    
'Test Case:
    test_number = 8
    test_description = "Equivalent arrays with non-common element object references"
    Debug.Print "Test Case [" & test_number & "]"
    Debug.Print Tab(10); test_description
    Set common_object1 = New Scripting.Dictionary
    Set common_object2 = New Scripting.Dictionary
    Set common_object3 = New Scripting.Dictionary
    thing1 = Array(common_object1, common_object2)
    thing2 = Array(common_object1, common_object3)
    Debug.Assert Not DS.Equivalent(thing1, thing2)
    Debug.Print "Result: Passed!"
    

Set thing1 = Nothing
Set thing2 = Nothing
Set common_object1 = Nothing
Set common_object2 = Nothing
Set common_object3 = Nothing
End Sub


Sub UnitTest_DS_Flatten()
    Dim nested_arr As Variant, flattened_arr As Variant
    Dim nested_col As Collection, flattened_col_as_arr As Variant
    Dim subnested_col As Collection
    Dim nested_dict As Scripting.Dictionary, flattened_dict_as_arr As Variant
    Dim elem As Variant
    Dim output As Variant
    
    nested_arr = Array( _
                            Array("a", "b", "c", "d"), _
                            Array("e", "f", "g", "h"), _
                            Array( _
                                Array("i", "j", "k", "l"), _
                                Array("m", "n", "o", "p") _
                                    ), _
                            Array("q", "r", "s", "t"), _
                            Array("u"), _
                            Array("v"), _
                            Array(Array(Array("w", "x"), "y"), "z") _
                                )
    flattened_arr = DS.Flatten(nested_arr)
    Debug.Print "Elements in output: " & UBound(flattened_arr) + 1
    Debug.Print "As a string: " & Join(flattened_arr, " ")
    
    Debug.Print "Reversed elements: " & Join(DS.Reverse(flattened_arr), " ")
    Stop
    
    Set nested_col = New Collection
    Debug.Print ""
    Set subnested_col = New Collection
    With subnested_col
        .Add 13: .Add 14: .Add 15: .Add 16
        .Add Array(17, 18, 19)
    End With
    Asc ("a")
    
    nested_col.Add flattened_arr
    nested_col.Add DS.Reverse(flattened_arr)
    With nested_col
        .Add 1: .Add 2: .Add 3
        .Add Array(4, 5, 6)
        .Add Array(Array(7, 8, 9), Array(10, 11))
        .Add 12
        .Add subnested_col
    End With
    nested_col.Add 20
    flattened_col_as_arr = DS.Flatten(nested_col)
    Debug.Print "Elements in nested collection: " & Join(DS.Apply(flattened_col_as_arr, "CStr", 0), " ")
    Stop
    Set nested_col = Nothing
    Set subnested_col = Nothing
    
End Sub

Sub UnitTest_DS_Fill()

    Dim arr As Variant
    Dim col As Collection
    Dim rtn As Variant
    Dim t1 As Variant, t2 As Variant, t3 As Variant, t4 As Variant
    Dim test_number As Integer
    Dim test_description As String
    Dim test_subj As Variant, test_operator As String

'Test Case:
    test_number = 1
    test_description = "No container given - Expect new array of ubound " & t1 & " with " & t1 + 1 & " elements, Filled with " & t2
    t1 = 9: t2 = 5
    arr = DS.fill(t1, t2)
    Debug.Assert (UBound(arr) = t1)
    Debug.Assert DS.Equivalent(arr, Array(t2, t2, t2, t2, t2, _
                                                                t2, t2, t2, t2, t2))
    Debug.Print "Test [" & test_number & "]: " & test_description & " - Passed."

'Test Case:
    test_number = 2
    test_description = "Predefined Array of Size " & t1 & " with " & t1 - 1 & " elements, Filled with " & t2
    t1 = 10 - 1: t2 = 5
    ReDim arr(t1)
    DS.fill arr, t2
    Debug.Assert (UBound(arr) = t1)
    Debug.Assert DS.Equivalent(arr, Array(t2, t2, t2, t2, t2, _
                                                                t2, t2, t2, t2, t2))
    Debug.Print "Test [" & test_number & "]: " & test_description & " - Passed."
    
'Test Case:
    test_number = 3
    test_description = "Empty Collection filled to count " & t1 & " with " & t2
    t1 = 10: t2 = 5
    Set col = New Collection
    DS.fill col, t2, t1
    Stop
    Debug.Assert col.Count = t1
    Debug.Assert DS.Equivalent(DS.Convert(col, "Variant()"), Array(5, 5, 5, 5, 5, _
                                                                                                    5, 5, 5, 5, 5))
    Debug.Print "Test [" & test_number & "]: " & test_description & " - Passed."

'Test Case:
    test_number = 4
'    test_description = "Empty Collection filled to count " & t1 & " with " & t2
'    t1 = 10: t2 = 5
'    Set col = New Collection
'    DS.fill col, t2, t1
'    Stop
'    Debug.Assert col.Count = t1
'    Debug.Assert DS.Equivalent(DS.Convert(col, "Variant()"), Array(5, 5, 5, 5, 5, _
'                                                                                                    5, 5, 5, 5, 5))
'    Debug.Print "Test [" & test_number & "]: " & test_description & " - Passed."
    
'Test Case:
'    test_number = 4
'    test_description = "Empty Collection filled to count " & t1 & " with " & t2
'    t1 = 10: t2 = 5
'    Set col = New Collection
'    DS.fill col, t2, t1
'    Stop
'    Debug.Assert col.Count = t1
'    Debug.Assert DS.Equivalent(DS.Convert(col, "Variant()"), Array(5, 5, 5, 5, 5, _
'                                                                                                    5, 5, 5, 5, 5))
'    Debug.Print "Test [" & test_number & "]: " & test_description & " - Passed."

End Sub

Sub UnitTest_Template()
    Dim arr As Variant
    Dim col As Collection
    Dim rtn As Variant
    Dim t1 As Variant, t2 As Variant, t3 As Variant, t4 As Variant
    Dim test_number As Integer
    Dim test_description As String
    Dim test_subj As Variant, test_operator As String
    
End Sub
Sub UnitTest_DS_Range()
    Dim arr As Variant
    Dim test_number As Integer
    Dim test_description As String
    
'Test Case:
    test_number = 1
    test_description = "Increasing Integers from 0 to 10"
    arr = DS.Range(start_value:=0, end_value:=10, step:=1)
    Debug.Assert (UBound(arr) = 10)
    Debug.Assert DS.Equivalent(arr, Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    Debug.Print "Test [" & test_number & "]: " & test_description & " - Passed."
    
'Test Case:
    test_number = 2
    test_description = "Decreasing Integers from 10 to 0"
    arr = DS.Range(start_value:=10, end_value:=0, step:=-1)
    Debug.Assert DS.Equivalent(arr, Array(10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0))
    Debug.Print "Test [" & test_number & "]: " & test_description & " - Passed."
    
'Test Case:
    test_number = 3
    test_description = "Increasing Decimals from -2.5 to 2.5, at 0.5 increments"
    arr = DS.Range(start_value:=-2.5, end_value:=2.5, step:=0.5)
    Debug.Assert DS.Equivalent(arr, Array(-2.5, -2, -1.5, -1, -0.5, 0, 0.5, 1, 1.5, 2, 2.5))
    Debug.Print "Test [" & test_number & "]: " & test_description & " - Passed."
        
'Test Case:
    test_number = 4
    test_description = "Decreasing Decimals from 2.5 to -2.5"
    arr = DS.Range(start_value:=2.5, end_value:=-2.5, step:=-0.5)
    Debug.Assert DS.Equivalent(arr, Array(2.5, 2, 1.5, 1, 0.5, 0, -0.5, -1, -1.5, -2, -2.5))
    Debug.Print "Test [" & test_number & "]: " & test_description & " - Passed."
    
'Test Case:
    test_number = 5
    test_description = "Increasing Decimals with Non-intersecting Bounds"
    arr = DS.Range(start_value:=-2, end_value:=5, step:=1.5)
    Debug.Assert DS.Equivalent(arr, Array(-2, -0.5, 1, 2.5, 4))
    Debug.Print "Test [" & test_number & "]: " & test_description & " - Passed."
    
'Test Case:
    test_number = 6
    test_description = "Decreasing Decimals with Non-intersecting Bounds"
    arr = DS.Range(start_value:=2, end_value:=-3, step:=-1.5)
    Debug.Assert DS.Equivalent(arr, Array(2, 0.5, -1, -2.5))
    Debug.Print "Test [" & test_number & "]: " & test_description & " - Passed."


End Sub

Sub UnitTest_DS_Match()
    Dim arr As Variant
    Dim col As Collection
    Dim rtn As Variant
    Dim t1 As Variant, t2 As Variant, t3 As Variant, t4 As Variant, t5 As Variant
    Dim test_number As Integer
    Dim test_description As String
    Dim test_subj As Variant, test_operator As String
    
'Test Case:
    test_number = 1
    test_description = "Single Subject Value Equality Match"
    test_subj = 1
    test_operator = "="
    t1 = 0: t2 = 1: t3 = 2: t4 = 3
    rtn = DS.Match(test_subj, test_operator, _
                                    t1, "a", _
                                    t2, "b", _
                                    t3, "c", _
                                    t4, "d")
    Debug.Assert (rtn = "b")
    Debug.Print "Test passed: " & test_description
'    Stop
    
'Test Case:
    test_number = 2
    test_description = "Single Subject Value Inequality Match"
    test_subj = "a"
    test_operator = "<>"
    t1 = "a": t2 = "a": t3 = "b": t4 = "a": t5 = 1
    rtn = DS.Match(test_subj, test_operator, _
                                    t1, "a", _
                                    t2, "b", _
                                    t3, "c", _
                                    t4, "d", _
                                    t5, "e")
    Debug.Assert (rtn = "c")
    Debug.Print "Test passed: " & test_description
    
'Test Case:
    test_number = 3
    test_description = "Single Subject Value Comparison Match"
    test_subj = 12
    test_operator = "<"
    t1 = 5: t2 = 7: t3 = 9: t4 = 11: t5 = 13
    rtn = DS.Match(test_subj, test_operator, _
                                    t1, "a", _
                                    t2, "b", _
                                    t3, "c", _
                                    t4, "d", _
                                    t5, "e")
    Debug.Assert (rtn = "e")
    Debug.Print "Test passed: " & test_description
    
'Test Case:
    test_number = 4
    test_description = "Single Subject No-match"
    test_subj = "1"
    test_operator = "="
    t1 = "a": t2 = "a": t3 = "b": t4 = "a": t5 = 0
    rtn = DS.Match(test_subj, test_operator, _
                                    t1, "a", _
                                    t2, "b", _
                                    t3, "c", _
                                    t4, "d", _
                                    t5, "e")
    Debug.Assert (IsEmpty(rtn))
    Debug.Print "Test passed: " & test_description
End Sub
