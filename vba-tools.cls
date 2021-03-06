VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "underscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' forEach
' map
' filter
' contains
' indexOf
' pluck
' max
' min
' sort
' sortBy
' group
' groupBy
' size
' slice
' splice
' append
' without
' zip
' insert
' remove
' fromRange
' fromTableRange
' toArray
' writeTable
' searchForTableRange

' mapDict
' extend
' pick
' without
' sortKeys
' orderKeys
' has
' values
' keys

' isString
' isNumber
' isCollection
' isDictionary
' isBoolean
' isNull
' isEmpty
' isArray
' isFunction
' isObject
Option Explicit

'**********************************************
' Collection Methods
'**********************************************

Sub forEach(c As Collection, f As String)
    Dim i As Integer
    For i = 1 To c.count
        Application.Run f, c(i), i, c
    Next i
End Sub

Function map(c As Collection, f As String) As Collection
    Dim i As Integer
    For i = 1 To c.count
        map.Add (Application.Run(f, c(i), i, c))
    Next i
End Function

Function filter(c As Collection, f As String) As Collection
    Dim i As Integer
    For i = 1 To c.count
        If Application.Run(f, c(i), i) = True Then filter.Add c(i)
    Next i
End Function

Function contains(c As Collection, v As Variant) As Boolean
    Dim i As Variant
    contains = False
    For Each i In c
        If i = v Then
            contains = True
            Exit Function
        End If
    Next i
End Function

Function indexOf(c As Collection, v As Variant) As Integer
    Dim i As Integer
    indexOf = -1
    For i = 1 To c.count
        If c(i) = v Then
            indexOf = i
            Exit Function
        End If
    Next i
End Function

Function pluck(c As Collection, key As String) As Collection
    Dim i As Integer
    For i = 1 To c.count
        If isDictionary(c(i)) Then
            pluck.Add c(i)(key)
        Else
            MsgBox "pluck expects a collection of dictionaries to work"
            Exit Function
        End If
    Next i
End Function

Function max(c As Collection) As Variant
    Dim i As Integer
    max = c(1)
    For i = 2 To c.count
        If c(i) > c(i - 1) Then
            max = c(i)
        End If
    Next i
End Function

Function min(c As Collection) As Variant
    Dim i As Integer
    min = c(1)
    For i = 2 To c.count
        If c(i) < c(i - 1) Then
            min = c(i)
        End If
    Next i
End Function

Function isSorted(c As Collection) As Boolean
    Dim i As Integer
    isSorted = True
    For i = 2 To c.count
        If c(i - 1) > c(i) Then
            isSorted = False
            Exit Function
        End If
    Next i
End Function

Function sort(c As Collection, Optional lo As Long = -1, Optional hi As Long = -1) As Collection
    ' QuickSort algorithm
    If lo = -1 Then lo = 1
    If hi = -1 Then hi = c.count
    
    Dim p As Long
    If lo < hi Then
        p = partition(c, lo, hi)
        sort c, lo, p
        sort c, p + 1, hi
    End If
End Function

Function sortBy(ByVal c As Collection, key As String, Optional lo As Long = -1, Optional hi As Long = -1) As Collection
    ' QuickSort algorithm
    If lo = -1 Then lo = 1
    If hi = -1 Then hi = c.count
    
    Dim p As Long
    If lo < hi Then
        p = partitionDict(c, key, lo, hi)
        sortBy c, key, lo, p
        sortBy c, key, p + 1, hi
    End If
End Function

Function group(c As Collection, f As String) As Dictionary
    Dim v As Variant, _
        key As String
        
    For Each v In c
        key = Application.Run(f, v)
        If has(group, key) Then
        Else
            group.Add key, New Collection
        End If
        group.Item(key).Add v
    Next v
    
End Function

Function groupBy(c As Collection, key As String) As Dictionary
    
End Function

Function size(o As Object) As Collection
    If isRange(o) Then
        size.Add o.Rows.count
        size.Add o.Columns.count
    ElseIf isCollection(o) Then
        size.Add o.count
        If isCollection(o(1)) Then
            size.Add o(1).count
        Else
            size.Add 0
        End If
    ElseIf isDictionary(o) Then
        size.Add o.count
    End If
End Function

' slice
Function slice(c As Collection, a As Long, count As Long) As Collection
    Dim i As Long
    For i = a To a + count
        slice.Add c(i)
    Next i
End Function

' splice
Function splice(c As Collection, a As Long, count As Long, Optional replace As Collection) As Collection
    Dim i As Long
    For i = a To a + count
        splice.Add c(a)
        remove c, a
    Next i
End Function

' append

' without
' zip
' insert
' remove
' fromRange
' fromTableRange
' toArray
' writeTable
' searchForTableRange

Private Function partition(c As Collection, lo As Long, hi As Long) As Long
    Dim pivot As Variant
    
    If isObject(c(lo)) Then
        Set pivot = c(lo)
    Else
        pivot = c(lo)
    End If
    
    Dim i As Long, j As Long
    i = lo - 1
    j = hi + 1
    While True
        Do
            i = i + 1
        Loop While c(i) < pivot
        Do
            j = j - 1
        Loop While c(j) > pivot
        If i >= j Then
            partition = j
            Exit Function
        End If
        swap c, i, j
    Wend
End Function

Private Function partitionDict(c As Collection, k As String, lo As Long, hi As Long) As Long
    Dim pivot As Variant
    Set pivot = c(lo)
    
    Dim i As Long, j As Long
    i = lo - 1
    j = hi + 1
    While True
        Do
            i = i + 1
        Loop While c(i)(k) < pivot(k)
        Do
            j = j - 1
        Loop While c(j)(k) > pivot(k)
        If i >= j Then
            partitionDict = j
            Exit Function
        End If
        swap c, i, j
    Wend
End Function

Private Function swap(c As Collection, a As Long, b As Long)
    Dim i As Variant, j As Variant
    If isObject(c(a)) Then
        Set i = c(a)
    Else
        i = c(a)
    End If
    
    If isObject(c(b)) Then
        Set j = c(b)
    Else
        j = c(b)
    End If
    
    insert c, j, a
    insert c, i, b + 1
    remove c, b + 2
    remove c, a + 1
    
End Function

Function insert(c As Collection, v As Variant, i As Long)
    c.Add v, , i
End Function

Function remove(c As Collection, i As Long)
    c.remove i
End Function

Function fromRange(r As Range) As Collection
    Dim a As Variant, _
        is2D As Boolean, _
        i As Long, _
        j As Long, _
        c As New Collection
        
    a = r.value
    is2D = r.Rows.count > 1 And r.Columns.count > 1
    
    If is2D Then
        For i = 1 To r.Rows.count
            c.Add New Collection
            For j = 1 To r.Columns.count
                c(i).Add a(i, j)
            Next j
        Next i
    Else
        For i = 1 To r.Rows.count
            For j = 1 To r.Columns.count
                c.Add a(i, j)
            Next j
        Next i
    End If
    Set fromRange = c
    
End Function

Function fromTableRange(rTopLeft As Range) As Collection
    Dim head As Collection
    Set head = fromRange(rTopLeft.Resize(1, rTopLeft.End(xlToRight).Column - rTopLeft.Column + 1))
    
    Dim a As Variant, _
        i As Long, _
        ii As Long, _
        j As Long, _
        jj As Long, _
        c As New Collection
        
    Set a = searchForTableRange(rTopLeft)
    i = a.Rows.count
    j = a.Columns.count
    
    a = a.value
    
    For ii = 2 To i
        c.Add New Dictionary
        For jj = 1 To j
            c(ii - 1).Add head(jj), a(ii, jj)
        Next jj
    Next ii
    
    Set fromTableRange = c
End Function

Function toArray(c As Collection, Optional includeHead As Boolean = False) As Variant
    Dim a As Variant
    
    Dim i As Long, _
        j As Long, _
        k As Collection
'        offset4Head As Integer
'        offset4Head = 0
        
    If isDictionary(c(1)) Then
        Set k = keys(c(1))
        ReDim a(1 To c.count + 1, 1 To c(1).count) As Variant
'        If includeHead Then
'            ReDim a(1 To c.Count + 1, 1 To c(1).Count) As Variant
'            For i = 1 To k.Count
'                a(1, i) = k(i)
'            Next i
'            offset4Head = 1
'        End If
        
        For i = 1 To c.count
            For j = 1 To k.count
                a(i, j) = c(i).Item(k(j))
            Next j
        Next i
    ElseIf isCollection(c(1)) Then
        ReDim a(1 To c.count + 1, 1 To c(1).count) As Variant
        For i = 1 To c.count
            For j = 1 To c(i).count
                a(i, j) = c(i)(j)
            Next j
        Next i
    Else
        ReDim a(1 To c.count + 1) As Variant
        For i = 1 To c.count
            a(i) = c(i)
        Next i
    End If

    toArray = a
End Function

Sub writeTable(c As Collection, r As Range)
    r.Resize(1, c(1).count).value = toArray(keys(c(1)))
    r.Offset(1, 0).Resize(c.count, c(1).count).value = toArray(c)
End Sub

Function searchForTableRange(rTopLeft As Range) As Range
    Set searchForTableRange = rTopLeft.Resize(rTopLeft.End(xlDown).Row - rTopLeft.Row + 1, rTopLeft.End(xlToRight).Column - rTopLeft.Column + 1)
End Function

'**********************************************
' Dictionary Methods
'**********************************************

' mapDict
' extend
' pick
' without
' sortKeys
' orderKeys
' has
' values
' keys

Function mapDict()

End Function

Function extend()

End Function

Function pick()

End Function

Function without()

End Function

Function sortKeys()

End Function

Function orderKeys()

End Function

Function has(d As Dictionary, k As String) As Boolean
    has = d.Exists(k)
End Function

Function values(d As Dictionary) As Collection
    Dim a As Variant, c As New Collection
    For Each a In d.Items
        c.Add a
    Next a
    Set values = c
    Set c = Nothing
End Function

Function keys(d As Dictionary) As Collection
    Dim a As Variant, c As New Collection
    For Each a In d.keys
        c.Add a
    Next a
    Set keys = c
    Set c = Nothing
End Function

'**********************************************
' Utility Methods
'**********************************************

Function isString(v As Variant) As Boolean
    isString = (VarType(v) = vbString)
End Function

Function isNumber(v As Variant) As Boolean
    Dim t As VbVarType
    t = VarType(v)
    isNumber = (t = vbDecimal Or t = vbDouble Or t = vbInteger Or t = vbLong Or t = vbLong)
End Function

Function isCollection(v As Variant) As Boolean
    isCollection = (TypeOf v Is Collection)
End Function

Function isDictionary(v As Variant) As Boolean
    isDictionary = (TypeOf v Is Dictionary)
End Function

Function isBoolean(v As Variant) As Boolean
    isBoolean = (VarType(v) = vbBoolean)
End Function

Function isNull(v As Variant) As Boolean
    isNull = (VarType(v) = vbNull)
End Function

Function isEmpty(v As Variant) As Boolean
    isEmpty = (VarType(v) = vbEmpty)
End Function

Function isArray(v As Variant) As Boolean
    isArray = (VarType(v) = vbArray)
End Function

'Function isFunction(v As String) As Boolean
'    On Error Resume Next
'    Dim m As Object
'    Set m = ActiveWorkbook.VBProject.vbComponents(v).CodeModule
'    If Err.Number <> 0 Then
'        isFunction = False
'    Else
'        isFunction = True
'    End If
'End Function

Function isObject(v As Variant) As Boolean
    isObject = (VarType(v) = vbObject)
End Function

Function isRange(r As Variant) As Boolean
    isRange = (TypeOf v Is Range)
End Function

