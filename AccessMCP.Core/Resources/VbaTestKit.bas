Attribute VB_Name = "VbaTestKit"
Option Explicit

' Injected by VbaTestHost. Not written by a developer.
Private Const RESULT_PATH As String = "@RESULT_PATH@"

Private Sub Emit(ByVal Line As String)
    Dim f As Integer
    f = FreeFile
    Open RESULT_PATH For Append As #f
    Print #f, Line
    Close #f
End Sub

Private Function J(ByVal Text As String) As String
    Dim i As Long
    Dim ch As Long
    Dim out As String
    Dim bs As String
    Dim code As String
    bs = Chr$(92)
    For i = 1 To Len(Text)
        ch = AscW(Mid$(Text, i, 1))
        If ch < 0 Then ch = ch + 65536
        Select Case ch
            Case 34
                out = out & bs & Chr$(34)
            Case 92
                out = out & bs & bs
            Case 8
                out = out & bs & "b"
            Case 9
                out = out & bs & "t"
            Case 10
                out = out & bs & "n"
            Case 12
                out = out & bs & "f"
            Case 13
                out = out & bs & "r"
            Case Is < 32, Is > 126
                code = Hex$(ch)
                out = out & bs & "u" & String$(4 - Len(code), "0") & code
            Case Else
                out = out & Chr$(ch)
        End Select
    Next i
    J = out
End Function

Public Sub Log(ByVal Text As String)
    Emit "{""event"":""log"",""text"":""" & J(CStr(Text)) & """}"
End Sub

Public Sub EmitValue(ByVal Value As Variant)
    Emit "{""event"":""value"",""type"":""" & J(TypeName(Value)) & """,""text"":""" & J(ValueText(Value)) & """}"
End Sub

Public Sub EmitError(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    Emit "{""event"":""error"",""number"":" & CStr(Number) & ",""source"":""" & J(Source) & """,""description"":""" & J(Description) & """}"
End Sub

' Fails the running test. The sentinel source lets the host tell an assertion
' failure apart from any other runtime error.
Public Sub Fail(ByVal Message As String)
    EmitError vbObjectError + 2001, "VbaTestKit", Message
    Err.Raise vbObjectError + 2001, "VbaTestKit", Message
End Sub

Public Sub Check(ByVal Condition As Boolean, Optional ByVal Message As String = "Condition was false.")
    If Not Condition Then Fail Message
End Sub

Public Sub AreEqual(ByVal Expected As Variant, ByVal Actual As Variant, Optional ByVal Message As String = "")
    If Not ValuesEqual(Expected, Actual) Then
        Fail Compose(Message, "Expected <" & ValueText(Expected) & "> but was <" & ValueText(Actual) & ">.")
    End If
End Sub

Public Sub AreNotEqual(ByVal NotExpected As Variant, ByVal Actual As Variant, Optional ByVal Message As String = "")
    If ValuesEqual(NotExpected, Actual) Then
        Fail Compose(Message, "Did not expect <" & ValueText(Actual) & ">.")
    End If
End Sub

Public Sub IsTrue(ByVal Condition As Boolean, Optional ByVal Message As String = "")
    If Not Condition Then
        Fail Compose(Message, "Expected True but was False.")
    End If
End Sub

Public Sub IsFalse(ByVal Condition As Boolean, Optional ByVal Message As String = "")
    If Condition Then
        Fail Compose(Message, "Expected False but was True.")
    End If
End Sub

' NOTE: no assertion may share a name with a VBA built-in. `IsNull` is one,
' so the null checks are spelled IsNullValue / IsNotNullValue; a module-level
' procedure shadows the built-in project-wide.
Public Sub IsNullValue(ByVal Value As Variant, Optional ByVal Message As String = "")
    If Not IsNull(Value) Then
        Fail Compose(Message, "Expected Null but was <" & ValueText(Value) & ">.")
    End If
End Sub

Public Sub IsNotNullValue(ByVal Value As Variant, Optional ByVal Message As String = "")
    If IsNull(Value) Then
        Fail Compose(Message, "Did not expect Null.")
    End If
End Sub

Public Sub Contains(ByVal Haystack As String, ByVal Needle As String, Optional ByVal Message As String = "")
    If InStr(1, Haystack, Needle, vbTextCompare) = 0 Then
        Fail Compose(Message, "Expected <" & Haystack & "> to contain <" & Needle & ">.")
    End If
End Sub

Public Sub Throws(ByVal ProcName As String, Optional ByVal Message As String = "")
    ' Standalone VBA has no host Application object, so a procedure cannot be
    ' invoked by name from inside the module. The host runs each test itself
    ' and reports the error it caught; this helper exists for source that
    ' expects the name. It always fails to make that limitation visible.
    Fail Compose(Message, "Throws is not available in the standalone host.")
End Sub

Private Function ValuesEqual(ByVal A As Variant, ByVal B As Variant) As Boolean
    If IsObject(A) Or IsObject(B) Then
        ValuesEqual = (A Is B)
        Exit Function
    End If
    If IsNull(A) Or IsEmpty(A) Then
        ValuesEqual = (IsNull(B) Or IsEmpty(B))
        Exit Function
    End If
    If IsNull(B) Or IsEmpty(B) Then
        ValuesEqual = False
        Exit Function
    End If
    ValuesEqual = (CStr(A) = CStr(B))
End Function

Private Function ValueText(ByVal Value As Variant) As String
    If IsObject(Value) Then
        ValueText = "[Object " & TypeName(Value) & "]"
    ElseIf IsNull(Value) Then
        ValueText = "Null"
    ElseIf IsEmpty(Value) Then
        ValueText = "Empty"
    Else
        ValueText = CStr(Value)
    End If
End Function

Private Function Compose(ByVal Message As String, ByVal Detail As String) As String
    If Len(Message) > 0 Then
        Compose = Message & " " & Detail
    Else
        Compose = Detail
    End If
End Function
