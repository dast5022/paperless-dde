Attribute VB_Name = "PaperlessDocumentImporter"
Private p&, token, dic
Global UserToken As String
Const pageSize As Integer = 200
Const ProgrammVersion As String = "v2.0.0"


Sub GetDocumentTypes()
    Dim firstrowlocal As Integer
    Dim paperlessUrl As String
    Dim wsResult As Object
    
    ' Get API-Url
    paperlessAPIUrl = ThisWorkbook.Names("paperlessAPIUrl").RefersToRange.Value
    
    ' Get Token for login
    If UserToken = "" Then
        UserToken = GetToken()
    End If
    
    ' Set workbook names
    Set wsResult = ThisWorkbook.Sheets("document_types")
    Set overview = ThisWorkbook.Sheets("overview")
    
    ' Build url for query
    paperlessUrl = paperlessAPIUrl & "document_types/"
    
    ' Set first row of target sheet to fill in data
    firstrowlocal = 2
    
    ' read data
    result = GetQuery(paperlessUrl, wsResult, firstrowlocal)
    
    ' Jump back to first sheet
    overview.Activate
    
End Sub


Sub GetCustomFields()
    Dim firstrowlocal As Integer
    Dim paperlessUrl As String
    Dim wsResult As Object
    
    ' Get API-Url
    paperlessAPIUrl = ThisWorkbook.Names("paperlessAPIUrl").RefersToRange.Value
    
    ' Get Token for login
    If UserToken = "" Then
        UserToken = GetToken()
    End If
    
    ' Set workbook names
    Set wsResult = ThisWorkbook.Sheets("custom_fields")
    Set overview = ThisWorkbook.Sheets("overview")
    
    ' Build url for query
    paperlessUrl = paperlessAPIUrl & "custom_fields/"
    
    ' Set first row of target sheet to fill in data
    firstrowlocal = 2
    
    ' read data
    result = GetQuery(paperlessUrl, wsResult, firstrowlocal)
    
    ' Jump back to first sheet
    overview.Activate
    
End Sub


Sub GetCorrespondents()
    Dim firstrowlocal As Integer
    Dim paperlessUrl As String
    Dim wsResult As Object
    
    ' Get API-Url
    paperlessAPIUrl = ThisWorkbook.Names("paperlessAPIUrl").RefersToRange.Value
    
    ' Get Token for login
    If UserToken = "" Then
        UserToken = GetToken()
    End If
        
    ' Set workbook names
    Set wsResult = ThisWorkbook.Sheets("correspondents")
    Set overview = ThisWorkbook.Sheets("overview")
    
    ' Set first row of target sheet to fill in data
    firstrowlocal = 2

    ' Build url for query
    paperlessUrl = paperlessAPIUrl & "correspondents/"
    
    ' read data
    result = GetQuery(paperlessUrl, wsResult, firstrowlocal)
    
    ' Jump back to first sheet
    overview.Activate
    
End Sub

Sub GetTags()
    Dim firstrowlocal As Integer
    Dim paperlessUrl As String
    Dim wsResult As Object

    ' Get API-Url
    paperlessAPIUrl = ThisWorkbook.Names("paperlessAPIUrl").RefersToRange.Value
    
    ' Get Token for login
    If UserToken = "" Then
        UserToken = GetToken()
    End If
        
    ' Set workbook names
    Set wsResult = ThisWorkbook.Sheets("tags")
    Set overview = ThisWorkbook.Sheets("overview")

    ' Set first row of target sheet to fill in data
    firstrowlocal = 2
    
    ' Build url for query
    paperlessUrl = paperlessAPIUrl & "tags/"
    
    ' read data
    result = GetQuery(paperlessUrl, wsResult, firstrowlocal)
    
    ' Jump back to first sheet
    overview.Activate
    
End Sub


Sub GetUsers()
    Dim firstrowlocal As Integer
    Dim paperlessUrl As String
    Dim wsResult As Object
    
    ' Get API-Url
    paperlessAPIUrl = ThisWorkbook.Names("paperlessAPIUrl").RefersToRange.Value
    
    ' Get Token for login
    If UserToken = "" Then
        UserToken = GetToken()
    End If
        
    ' Set workbook names
    Set wsResult = ThisWorkbook.Sheets("users")
    Set overview = ThisWorkbook.Sheets("overview")

    ' Set first row of target sheet to fill in data
    firstrowlocal = 2
    
    ' Build url for query
    paperlessUrl = paperlessAPIUrl & "users/"
    
    ' read data
    result = GetQuery(paperlessUrl, wsResult, firstrowlocal)
    
    ' Jump back to first sheet
    overview.Activate
    
End Sub


Sub GetStoragePaths()
    Dim firstrowlocal As Integer
    Dim paperlessUrl As String
    Dim wsResult As Object
    
    ' Get API-Url
    paperlessAPIUrl = ThisWorkbook.Names("paperlessAPIUrl").RefersToRange.Value
    
    ' Get Token for login
    If UserToken = "" Then
        UserToken = GetToken()
    End If
        
    ' Set workbook names
    Set wsResult = ThisWorkbook.Sheets("storage_paths")
    Set overview = ThisWorkbook.Sheets("overview")

    ' Set first row of target sheet to fill in data
    firstrowlocal = 2
    
    ' Build url for query
    paperlessUrl = paperlessAPIUrl & "storage_paths/"
    
    ' read data
    result = GetQuery(paperlessUrl, wsResult, firstrowlocal)
    
    ' Jump back to first sheet
    overview.Activate
    
End Sub


Sub GetDocuments()
    Dim firstrowlocal As Integer
    Dim paperlessUrl As String
    Dim wsResult As Object
    Dim Query As String
    
    ' Get API-Url
    paperlessAPIUrl = ThisWorkbook.Names("paperlessAPIUrl").RefersToRange.Value
    
    ' Get Token for login
    If UserToken = "" Then
        UserToken = GetToken()
    End If
    
    ' set workbook names
    Set wsResult = ThisWorkbook.Sheets("documents")
    Set overview = ThisWorkbook.Sheets("overview")

    ' Set first row of target sheet to fill in data
    firstrowlocal = 2
    
    ' Get query from input
    overview.Activate
    Query = ThisWorkbook.Names("document_query").RefersToRange.Value
     
    ' execute query
    paperlessUrl = paperlessAPIUrl & "documents/?" & Query
    result = GetQuery(paperlessUrl, wsResult, firstrowlocal)
    
    ' Jump back to first sheet
    overview.Activate

End Sub

Sub ReplaceColumns()
    Dim target_sheet As Object
    Dim col As Integer
    Dim source_sheet As Object
    Dim firstrowlocal As Integer
    Dim o As Integer
    
    Set target_sheet = ThisWorkbook.Sheets("documents")
    
    answer = MsgBox("Replace ids by names in sheet " & target_sheet.Name & "?", vbOKCancel)
    If answer = vbCancel Then
        Exit Sub
    End If

    On Error GoTo ErrorRaise
    
    ' replace all values in column correspondents
    Set source_sheet = ThisWorkbook.Sheets("correspondents")
    target_field = "correspondent"
    col = Application.WorksheetFunction.match(target_field, target_sheet.Rows("1:1"), 0)
    firstrowlocal = 2
    result = ReplaceIDs(target_sheet, col, firstrowlocal, source_sheet)
    
    ' replace all values in column document_type
    Set source_sheet = ThisWorkbook.Sheets("document_types")
    target_field = "document_type"
    col = Application.WorksheetFunction.match(target_field, target_sheet.Rows("1:1"), 0)
    firstrowlocal = 2
    result = ReplaceIDs(target_sheet, col, firstrowlocal, source_sheet)
    
    ' replace all values in storage path
    Set source_sheet = ThisWorkbook.Sheets("storage_paths")
    target_field = "storage_path"
    col = Application.WorksheetFunction.match(target_field, target_sheet.Rows("1:1"), 0)
    firstrowlocal = 2
    result = ReplaceIDs(target_sheet, col, firstrowlocal, source_sheet)
    
    ' replace all values in column owner
    Set source_sheet = ThisWorkbook.Sheets("users")
    target_field = "owner"
    col = Application.WorksheetFunction.match(target_field, target_sheet.Rows("1:1"), 0)
    firstrowlocal = 2
    result = ReplaceIDs(target_sheet, col, firstrowlocal, source_sheet)

    
    ' replace all values in first column tags
    Set source_sheet = ThisWorkbook.Sheets("tags")
    target_field = "tagStart"
    col = Application.WorksheetFunction.match(target_field, target_sheet.Rows("1:1"), 0)
    firstrowlocal = 2
    result = ReplaceIDs(target_sheet, col, firstrowlocal, source_sheet)
    coloftagStart = col
    On Error GoTo 0
    
    target_field = "tag"
    For o = coloftagStart + 1 To 100
        If target_sheet.Cells(1, o).Value = "tag" Then
            result = ReplaceIDs(target_sheet, o, firstrowlocal, source_sheet)
        Else
            Exit For
        End If
    Next o
    
    MsgBox "Done", vbInformation
    
    Exit Sub
ErrorRaise:
    field = source_sheet.Name
    MsgBox "Error: field " & target_field & " not found on sheet " & target_sheet.Name, vbInformation
    
    
End Sub

Sub ClearAll()

    answer = MsgBox("Delete all data from all sheets?", vbOKCancel)
    If answer = vbCancel Then
        Exit Sub
    End If

    Sheets("storage_paths").Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("correspondents").Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("document_types").Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("tags").Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("users").Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("documents").Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("custom_fields").Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    
    ' Jump back to first sheet
    Sheets("overview").Activate
    
End Sub




Function ReplaceIDs(target_sheet As Object, col As Integer, firstrowlocal As Integer, source_sheet As Object)
    Dim search As Variant
    Dim result As String

    CountofRows = target_sheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    For i = firstrowlocal To CountofRows
        search = target_sheet.Cells(i, col).Value
        If IsNumeric(search) And search > 0 Then
            result = Application.WorksheetFunction.VLookup(search, source_sheet.Range("A:B"), 2, False)
            target_sheet.Cells(i, col).Value = result
        End If
    Next i
    
End Function


Function GetQuery(paperlessUrl As String, wsResult As Object, firstrowlocal As Integer)
    Dim jsonResponse As String
    Dim http As Object
    Dim col As Integer
    Dim i As Long
    Dim Filter As String
    Dim check As Variant
    Dim cols As Variant
    Dim partdic As Variant
    Dim num1 As Double
    Dim num2 As Integer
    Dim numbersOfRuns As Integer
    
    ' add page-size = 1 to query, just to get the count of result
    If Right(paperlessUrl, 1) = "/" Then
        queryUrl = paperlessUrl & "?page_size=1"
    Else
        queryUrl = paperlessUrl & "&page_size=1"
    End If
    
    ' Make the first API request
    Debug.Print queryUrl
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", queryUrl, False
    http.setRequestHeader "Authorization", "Token " & UserToken
    http.send
    
    ' Get the JSON response
    jsonResponse = http.responsetext
    Set dic = ParseJSON(jsonResponse)
    
    ' Get count
    Count = CInt(dic("fields.count"))
    
    ' Ask user to go on
    If Count > 0 Then
        answer = MsgBox("Got " & Count & " rows from paperless. Go on and overwrite data in sheet?", vbOKCancel)
        If answer = vbCancel Then
            Exit Function
        End If
    End If
    
    ' Clear all lines from declared first row on
    wsResult.Activate
    wsResult.Rows(CStr(firstrowlocal) & ":" & CStr(firstrowlocal)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    wsResult.Select
    CountofCols = wsResult.Cells(1, wsResult.Columns.Count).End(xlToLeft).Column
    Debug.Print "Anzahl Spalten: " & CountofCols
    

    ' calculate numbers of queries
    num1 = Count / pageSize
    num2 = Count Mod pageSize
    If num2 = 0 Then
        numbersOfRuns = num1
    Else
        numbersOfRuns = WorksheetFunction.RoundUp(num1, 0)
    End If
    Debug.Print "Anzahl Abfragen: " + CStr(numbersOfRuns)

    z = 0
    For Page = 1 To numbersOfRuns
    
        ' add page-size = 10000 to query, just to get the count of result
        If Right(paperlessUrl, 1) = "/" Then
            queryUrl = paperlessUrl & "?page_size=" & pageSize & "&page=" & Page
        Else
            queryUrl = paperlessUrl & "&page_size=" & pageSize & "&page=" & Page
        End If
            
        
        ' Make the second API request
        Debug.Print queryUrl
        http.Open "GET", queryUrl, False
        http.setRequestHeader "Authorization", "Token " & UserToken
        http.send
        
        ' Get the JSON response
        jsonResponse = http.responsetext
        Set dic = ParseJSON(jsonResponse)
        'Debug.Print ListPaths(dic)
         
        
        ' Write results to cells
        i = 0
        Do
        
            ' Loop through each column in the result sheet to find matching fields
            For col = 1 To CountofCols
                check = ""
                cols = Null
                NewValue = ""
                VersionTable = ""
                fieldName = wsResult.Cells(1, col).Value
    
        
                Filter = "fields.results(" & i & ")." & CStr(fieldName)
                'Debug.Print "sarching for " & Filter & " for column " & col & " (" & fieldName & ")"
                
                Select Case fieldName
                        
                    ' tags field: seperate column for each tag
                    Case "tagStart"
                        Filter = "fields.results(" & i & ").tags"
                        check = dic(Filter)
                        If check <> "[]" Then
                            cols = Array(Filter & "(*)", "")
                            VersionTable = GetFilteredTable(dic, cols)
                            If IsArray(VersionTable) Then
                                Length = UBound(VersionTable)
                                col = col - 1
                                For y = 1 To Length
                                    col = col + 1
                                    If (wsResult.Cells(firstrowlocal - 1, col).Value = fieldName Or wsResult.Cells(firstrowlocal - 1, col).Value = "tag") Then
                                        wsResult.Cells(firstrowlocal + z + i, col).Value = CStr(VersionTable(y, 1))
                                    Else
                                        y = Length
                                        col = col - 1
    
                                    End If
                                    
                                Next y
                            Else
                                wsResult.Cells(firstrowlocal + z + i, col).Value = "null"
                            End If
                        Else
                            wsResult.Cells(firstrowlocal + z + i, col).Value = "null"
                        End If
                        
                    ' content field: set a prefix to ensure, that content is always text
                    Case "content"
                        If dic(Filter) <> "null" Then
                            wsResult.Cells(firstrowlocal + z + i, col).Value = "'" & dic(Filter)
                        Else
                            wsResult.Cells(firstrowlocal + z + i, col).Value = ""
                        End If
                        
                    ' normal fields
                    Case "name", "username", "path", "id", "correspondent", "document_type", "storage_path", "title", "created", "created_date", "modified", "added", "deleted_at", "archive_serial_number", "original_file_name", "archived_file_name", "owner", "user_can_change", "is_shared_by_requester", "page_count"
                        If dic(Filter) <> "null" Then
                            wsResult.Cells(firstrowlocal + z + i, col).Value = dic(Filter)
                        Else
                            wsResult.Cells(firstrowlocal + z + i, col).Value = ""
                        End If
                    
                    ' do nothing, already done before
                    Case "tag"
                        
                        
                    ' notes
                    Case "notes"

                        u = 0
                        notesstring = ""
                        Do
                            ' check note has id
                            Filter = "fields.results(" & i & ").notes(" & u & ").id"
                            field_id = dic(Filter)
                            
                            ' check note was not deleted
                            Filter = "fields.results(" & i & ").notes(" & u & ").deleted_at"
                            field_deleted_at = dic(Filter)

                            ' check note was not restored
                            Filter = "fields.results(" & i & ").notes(" & u & ").restored_at"
                            field_restored_at = dic(Filter)

                            ' then, read value
                            Filter = "fields.results(" & i & ").notes(" & u & ").note"

                            If CInt(field_id) > 0 And (field_deleted_at = "null" Or (field_deleted_at <> "null" And field_restored_at <> "null")) Then
                                If u > 0 Then
                                    notesstring = notesstring & Chr(10)
                                End If
                                notesstring = notesstring + "[" & u + 1 & "] " & dic(Filter)
                            End If
                            u = u + 1
                        Loop While field_id > 0
                        
                        wsResult.Cells(firstrowlocal + z + i, col).Value = notesstring
                        
                                          
                    ' perhabs a custom field?
                    Case Else
                        ' check if field name of sheet documents exits on sheet custom_fields
                        On Error Resume Next
                        test = 0
                        test = Application.WorksheetFunction.VLookup(fieldName, ThisWorkbook.Sheets("custom_fields").Range("A:B"), 2, False)
                        On Error GoTo 0
                        
                        ' when found
                        If test > 0 Then
                            
                            ' iterate through all custom fields in json-structure
                            u = 0
                            Do
                                ' check if field matches
                                Filter = "fields.results(" & i & ").custom_fields(" & u & ").field"
                                field = dic(Filter)

                                ' then, read value
                                If CInt(field) = CInt(test) Then
                                    Filter = "fields.results(" & i & ").custom_fields(" & u & ").value"
                                    wsResult.Cells(firstrowlocal + z + i, col).Value = dic(Filter)
                                End If
                                u = u + 1
                            Loop While field > 0
                        End If
                                          
                                          
                End Select
                
            Next col
            
            i = i + 1
            ' check if next row also exits
            Filter = "fields.results(" & i & ").id"
            
        ' only go on, when next row has id
        Loop While CInt(dic(Filter)) > 0 And i < pageSize

        z = z + pageSize
        
    Next Page
    
    MsgBox "Done", vbInformation
    Exit Function
    
noresult:
    answer = MsgBox("no results", vbOKCancel)
    
    Exit Function
    

End Function


Function GetToken()
    If UserToken = "" Then
        UserToken = CStr(InputBox("Please enter Token to acccess your paperless-ngx instance.", "Token"))
    End If
    GetToken = UserToken
    
End Function


'##################################################################################################################
' Helper-functions for parsing JSON
'##################################################################################################################


Function ParseJSON(json$, Optional key$ = "fields") As Object
    p = 1
    token = Tokenize(json)
    Set dic = CreateObject("Scripting.Dictionary")
    If token(p) = "{" Then ParseObj key Else ParseArr key
    Set ParseJSON = dic
End Function
Function ParseObj(key$)
    Do: p = p + 1
        Select Case token(p)
            Case "]"
            Case "[":  ParseArr key
            Case "{"
                       If token(p + 1) = "}" Then
                           p = p + 1
                           dic.Add key, "null"
                       Else
                           ParseObj key
                       End If
                
            Case "}":  key = ReducePath(key): Exit Do
            Case ":":  key = key & "." & token(p - 1)
            Case ",":  key = ReducePath(key)
            Case Else: If token(p + 1) <> ":" Then dic.Add key, token(p)
        End Select
    Loop
End Function
Function ParseArr(key$)
    Dim e&
    Do: p = p + 1
        Select Case token(p)
            Case "}"
            Case "{":  ParseObj key & ArrayID(e)
            Case "[":  ParseArr key
            Case "]":  Exit Do
            Case ":":  key = key & ArrayID(e)
            Case ",":  e = e + 1
            Case Else: dic.Add key & ArrayID(e), token(p)
        End Select
    Loop
End Function


Function Tokenize(s$)
    Const Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:0|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
    Tokenize = RExtract(s, Pattern, True)
End Function


Function RExtract(s$, Pattern, Optional bGroup1Bias As Boolean, Optional bGlobal As Boolean = True)
  Dim c&, m, n, v
  With CreateObject("vbscript.regexp")
    .Global = bGlobal
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = Pattern
    If .test(s) Then
      Set m = .Execute(s)
      ReDim v(1 To m.Count)
      For Each n In m
        c = c + 1
        v(c) = n.Value
        If bGroup1Bias Then If Len(n.submatches(0)) Or n.Value = """""" Then v(c) = n.submatches(0)
      Next
    End If
  End With
  RExtract = v
End Function
Function ArrayID$(e)
    ArrayID = "(" & e & ")"
End Function
Function ReducePath$(key$)
    If InStr(key, ".") Then ReducePath = Left(key, InStrRev(key, ".") - 1) Else ReducePath = key
End Function
Function ListPaths(dic)
    Dim s$, v
    For Each v In dic
        s = s & v & " --> " & dic(v) & vbLf
    Next
    Debug.Print s
End Function
Function GetFilteredValues(dic, match)
    Dim c&, i&, v, w
    v = dic.keys
    ReDim w(1 To dic.Count)
    For i = 0 To UBound(v)
        If v(i) Like match Then
            c = c + 1
            w(c) = dic(v(i))
        End If
    Next
    If c = 0 Then
    c = 1
    End If
    ReDim Preserve w(1 To c)
    GetFilteredValues = w
End Function
Function GetFilteredTable(dic, cols)
    Dim c&, i&, j&, v, w, z
    v = dic.keys
    z = GetFilteredValues(dic, cols(0))
    ReDim w(1 To UBound(z), 1 To UBound(cols) + 1)
    For j = 1 To UBound(cols) + 1
         z = GetFilteredValues(dic, cols(j - 1))
         For i = 1 To UBound(z)
            w(i, j) = z(i)
         Next
    Next
    GetFilteredTable = w
End Function
Function OpenTextFile$(f)
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Open
        .LoadFromFile f
        OpenTextFile = .ReadText
    End With
End Function


'##################################################################################################################
' Function for building the application the first time
'##################################################################################################################

Sub A_BuildApplication()

    On Error GoTo startBuild
    ThisWorkbook.Sheets("overview").Select
    MsgBox ("The applications already exits")
    Exit Sub
    
startBuild:
    On Error GoTo 0
    ActiveSheet.Name = "overview"
    Cells.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Columns("A:A").ColumnWidth = 31
    Columns("B:B").ColumnWidth = 31
    Columns("D:D").ColumnWidth = 31
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1. Set API-Url"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "http://your-ip-to-paperless/api/"
    
    Range("B2").Select
    ActiveWorkbook.Names.Add Name:="paperlessAPIUrl", RefersToR1C1:= _
        "=overview!R2C2"
    
    Range("B2").Select
    Selection.Hyperlinks.Delete
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "2. Get correspondent names"
    Range("B4").Select
    ActiveSheet.Buttons.Add(175.2, 43.2, 141.6, 14.4).Select
    Selection.OnAction = "GetCorrespondents"
    Selection.Characters.Text = "Get correspondents"
    With Selection.Characters(start:=1, Length:=18).Font
        .Name = "Calibri"
        .FontStyle = "Standard"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "3. Get document type names"
    Range("B6").Select
    ActiveSheet.Buttons.Add(175.2, 72, 141.6, 14.4).Select
    Selection.OnAction = "GetDocumentTypes"
    Selection.Characters.Text = "Get document types"
    With Selection.Characters(start:=1, Length:=18).Font
        .Name = "Calibri"
        .FontStyle = "Standard"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("A8").Select
    ActiveCell.FormulaR1C1 = "4. Get tag names"
    Range("B8").Select
    ActiveSheet.Buttons.Add(175.2, 100.8, 141.6, 14.4).Select
    Selection.OnAction = "GetTags"
    Selection.Characters.Text = "Get tags"
    With Selection.Characters(start:=1, Length:=8).Font
        .Name = "Calibri"
        .FontStyle = "Standard"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("A10").Select
    ActiveCell.FormulaR1C1 = "5. Get user names"
    Range("B10").Select
    ActiveSheet.Buttons.Add(175.2, 129.6, 141.6, 14.4).Select
    Selection.OnAction = "GetUsers"
    Selection.Characters.Text = "Get users"
    With Selection.Characters(start:=1, Length:=9).Font
        .Name = "Calibri"
        .FontStyle = "Standard"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("A12").Select
    ActiveCell.FormulaR1C1 = "6. Get storage path names"
    Range("B12").Select
    ActiveSheet.Buttons.Add(175.2, 158.4, 141.6, 14.4).Select
    Selection.OnAction = "GetStoragePaths"
    Selection.Characters.Text = "Get storage paths"
    With Selection.Characters(start:=1, Length:=17).Font
        .Name = "Calibri"
        .FontStyle = "Standard"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("A14").Select
    ActiveCell.FormulaR1C1 = "7. Get custom field names"
    Range("B14").Select
    ActiveSheet.Buttons.Add(175.2, 187.2, 141.6, 14.4).Select
    Selection.OnAction = "GetCustomFields"
    Selection.Characters.Text = "Get custom fields"
    With Selection.Characters(start:=1, Length:=17).Font
        .Name = "Calibri"
        .FontStyle = "Standard"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("A16").Select
    ActiveCell.FormulaR1C1 = "8. Get documents"
    Range("B16").Select
    ActiveSheet.Buttons.Add(175.2, 216, 141.6, 14.4).Select
    Selection.OnAction = "GetDocuments"
    Selection.Characters.Text = "Get documents"
    With Selection.Characters(start:=1, Length:=13).Font
        .Name = "Calibri"
        .FontStyle = "Standard"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("D15").Select
    ActiveCell.FormulaR1C1 = "Query for documents"
    Range("D16").Select
    ActiveCell.FormulaR1C1 = "tags__id=1"
    Range("D16").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    Range("D16").Select
    ActiveWorkbook.Names.Add Name:="document_query", RefersToR1C1:= _
        "=overview!R16C4"
        
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A18").Select
    ActiveCell.FormulaR1C1 = "9. Replace ids by names"
    Range("B18").Select
    ActiveSheet.Buttons.Add(175.2, 244.8, 141.6, 14.4).Select
    Selection.OnAction = "ReplaceColumns"
    Range("D28").Select
    ActiveSheet.Shapes.Range(Array("Button 8")).Select
    Selection.Characters.Text = "Replace ids"
    With Selection.Characters(start:=1, Length:=11).Font
        .Name = "Calibri"
        .FontStyle = "Standard"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("A20").Select
    ActiveCell.FormulaR1C1 = "Clear all (reset)"
    Range("B20").Select
    ActiveSheet.Buttons.Add(175.2, 273.6, 141.6, 14.4).Select
    Selection.OnAction = "ClearAll"
    Selection.Characters.Text = "Clear all"
    With Selection.Characters(start:=1, Length:=9).Font
        .Name = "Calibri"
        .FontStyle = "Standard"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With

    Range("F15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Example queries:"
    Range("F16").Select
    ActiveCell.FormulaR1C1 = "tags__id=1"
    Range("F17").Select
    ActiveCell.FormulaR1C1 = "correspondent__id=1"
    Range("F18").Select
    ActiveCell.FormulaR1C1 = "correspondent__id=1&document_type__id=1"
    Range("F19").Select
    Range("F15").Select
    Selection.Font.Bold = True
    Range("D15").Select
    Selection.Font.Bold = True
    Range("A16").Select
    Selection.Font.Bold = True
    ActiveWorkbook.Save

    ' build sheet "documents"
    ThisWorkbook.Sheets.Add after:=Sheets(Worksheets.Count)
    ThisWorkbook.ActiveSheet.Name = "documents"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "id"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "correspondent"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "document_type"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "storage_path"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "title"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "content"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "tagStart"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "tag"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "tag"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "tag"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "tag"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "created"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "created_date"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "modified"
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "added"
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "deleted_at"
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "archive_serial_number"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "original_file_name"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "archived_file_name"
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "owner"
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "user_can_change"
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "is_shared_by_requester"
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "notes"
    Range("X1").Select
    ActiveCell.FormulaR1C1 = "CustomFieldName"
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "page_count"
    
    Range("A1:Y1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("B2").Select
    ActiveWindow.FreezePanes = True
    
    
    ' build sheet "correspondents"
    ThisWorkbook.Sheets.Add after:=Sheets(Worksheets.Count)
    ThisWorkbook.ActiveSheet.Name = "correspondents"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "id"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "name"
    Range("A1:B1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("B2").Select
    ActiveWindow.FreezePanes = True
    
    
    ' build sheet "document types"
    ThisWorkbook.Sheets.Add after:=Sheets(Worksheets.Count)
    ThisWorkbook.ActiveSheet.Name = "document_types"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "id"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "name"
    Range("A1:B1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("B2").Select
    ActiveWindow.FreezePanes = True
    
    
    ' build sheet "tags"
    ThisWorkbook.Sheets.Add after:=Sheets(Worksheets.Count)
    ThisWorkbook.ActiveSheet.Name = "tags"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "id"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "name"
    Range("A1:B1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("B2").Select
    ActiveWindow.FreezePanes = True
    Range("B2").Select
    ActiveWindow.FreezePanes = True


    ' build sheet "user names"
    ThisWorkbook.Sheets.Add after:=Sheets(Worksheets.Count)
    ThisWorkbook.ActiveSheet.Name = "users"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "id"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "username"
    Range("A1:B1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("B2").Select
    ActiveWindow.FreezePanes = True
    Range("B2").Select
    ActiveWindow.FreezePanes = True
    

    ' build sheet "storage paths"
    ThisWorkbook.Sheets.Add after:=Sheets(Worksheets.Count)
    ThisWorkbook.ActiveSheet.Name = "storage_paths"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "id"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "name"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "path"
    Range("A1:C1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("B2").Select
    ActiveWindow.FreezePanes = True
    

    ' build sheet "custom_fields"
    ThisWorkbook.Sheets.Add after:=Sheets(Worksheets.Count)
    ThisWorkbook.ActiveSheet.Name = "custom_fields"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "name"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "id"
    Range("A1:B1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("B2").Select
    ActiveWindow.FreezePanes = True
    
    ThisWorkbook.Sheets("overview").Select
    
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(correspondents!C[-2],"">0"")"
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(document_types!C[-2],"">0"")"
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(tags!C[-2],"">0"")"
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(users!C[-2],"">0"")"
    Range("C12").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(storage_paths!C[-2],"">0"")"
    Range("C14").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(custom_fields!C[-1],"">0"")"
    Range("C16").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(documents!C[-2],"">0"")"
    Range("C17").Select
    
        Columns("C:C").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Paperless Document Importer, " & ProgrammVersion
    Range("A1").Select
    Rows("1:1").RowHeight = 52.2
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
    Selection.Font.Size = 16
    
    
    
    Range("B2").Select

End Sub

Sub A_RemovePersonalInformation()
    ' Helper-function to remove all personal Information from file before uploading to Github
    ThisWorkbook.RemoveDocumentInformation (xlRDIRemovePersonalInformation)
End Sub
