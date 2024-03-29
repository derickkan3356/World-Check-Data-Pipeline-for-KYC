Option Base 1
Option Explicit

Public Const ConnString As String = "driver={SQL server};Server=#####;Database=#####;uid=#####;Password=#####"
Private Const inputTemplatePath As String = "Template\WC Input template.xlsx"
Private Const outputWCPath As String = "Template\WC upload template.xlsx"
Private Const outputTemplatePath As String = "Template\WC output template.xlsx"
Global selectedColumn As Integer

Sub ImportWCdata()
    Dim currentWorkbook As Workbook
    Dim indStartRow As Long, indLastRow As Long, indLastColumn As Long, indLastColumnStr As String, indClearColumnStr As String
    Dim entStartRow As Long, entLastRow As Long, entLastColumn As Long, entLastColumnStr As String, entClearColumnStr As String
    Dim startRow As Long, lastRow As Long, lastColumn As Long
    Dim mandatoryColumn As String
    Dim entityNameColumn As String, fundNameColumn As String, clientNameColumn As String
    Dim cell As Range, header As Range, row As Integer, sheet As Worksheet
    Dim entityMapping As Object, fundMapping As Object, clientMapping As Object
    Dim noIndividualData As Boolean, noEntityData As Boolean
    Dim findMandatory As Variant
    Dim indUpdatedCount As Integer, indInsertedCount As Integer, entUpdatedCount As Integer, entInsertedCount As Integer

    Dim conn As adodb.Connection
    Dim cmd As adodb.Command
    Dim rs As adodb.Recordset, result As adodb.Recordset
    Dim SQLquery As String

    Dim givenName As String, familyName As String, chineseName As String, gender As String, placeOfBirth As String
    Dim dateOfBirth As String, citizenship As String, cityLocation As String, regionLocation As String, countryLocation As String
    Dim entityID As String, capacity As String, caseID As String, lastScreenDate As String, registeredCountry As String, entityName As String
    Dim fundName As String, fundID As String, clientName As String, clientID As String, screeningFlagPep As String, detailsFlag As String, archive As String

    Set currentWorkbook = ActiveWorkbook
    Set entityMapping = CreateObject("Scripting.Dictionary")
    Set fundMapping = CreateObject("Scripting.Dictionary")
    Set clientMapping = CreateObject("Scripting.Dictionary")

    ' create an input template workbook if invalid active workbook
    If Not SheetExist("Individual screening", currentWorkbook) Or Not SheetExist("Entity screening", currentWorkbook) Then
        OpenTemplateWorkbook(inputTemplatePath)
        Exit Sub
    End If

    With currentWorkbook.Sheets("Individual Screening")
        indStartRow = .Columns("B").Find(What:="Given Name(s)", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).row + 1
        indLastRow = .Range("B9999").End(xlUp).row
        If indLastRow = indStartRow - 1 Then
            indLastRow = indStartRow
        End If
        indLastColumn = .Range("B" & indStartRow - 1).End(xlToRight).Column
        indLastColumnStr = Split(Cells(1, indLastColumn).Address, "$")(1)
        indClearColumnStr = Split(Cells(1, indLastColumn + 1).Address, "$")(1)
        entityNameColumn = Split(.Rows(indStartRow - 1).Find(What:="Entity Name", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Address, "$")(1)

        ' check the sheet have data
        noIndividualData = Application.WorksheetFunction.CountA(.Range("B" & indStartRow & ":" & indLastColumnStr & indLastRow)) = 0

        If Not noIndividualData Then
            ' check the mandatory fields
            For Each header In .Range("B" & indStartRow - 1 & ":" & indLastColumnStr & indStartRow - 1)
                If header Like "*[*]*" Then
                    mandatoryColumn = Split(header.Address, "$")(1)
                    For Each cell In .Range(mandatoryColumn & indStartRow & ":" & mandatoryColumn & indLastRow)
                        If cell.Value2 = "" Then
                            MsgBox header & " is a mandatory field. Please fill in all its row."
                            Exit Sub
                        End If
                    Next cell
                End If
            Next header
        End If
    End With

    With currentWorkbook.Sheets("Entity Screening")
        entStartRow = .Columns("B").Find(What:="Entity Name", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).row + 1
        entLastRow = .Range("B9999").End(xlUp).row
        If entLastRow = entStartRow - 1 Then
            entLastRow = entStartRow
        End If
        entLastColumn = .Range("B" & entStartRow - 1).End(xlToRight).Column
        entLastColumnStr = Split(Cells(1, entLastColumn).Address, "$")(1)
        entClearColumnStr = Split(Cells(1, entLastColumn + 1).Address, "$")(1)

        ' check the sheet have data
        noEntityData = Application.WorksheetFunction.CountA(.Range("B" & entStartRow & ":" & entLastColumnStr & entLastRow)) = 0

        If Not noEntityData Then
            ' check the mandatory fields
            For Each header In .Range("B" & entStartRow - 1 & ":" & entLastColumnStr & entStartRow - 1)
                If header.Value2 Like "*[*]*" Then
                    mandatoryColumn = Split(header.Address, "$")(1)
                    For Each cell In .Range(mandatoryColumn & entStartRow & ":" & mandatoryColumn & entLastRow)
                        If cell.Value2 = "" Then
                            MsgBox header & " is a mandatory field. Please fill in all its row."
                            .Activate
                            Exit Sub
                        End If
                    Next cell
                End If
            Next header
        End If
    End With

    On Error GoTo Error_Handler
    Application.ScreenUpdating = False
    Set conn = New adodb.Connection
    conn.Open ConnString
    Set cmd = New adodb.Command
    Set rs = New adodb.Recordset
    indInsertedCount = 0
    indUpdatedCount = 0
    entInsertedCount = 0
    entUpdatedCount = 0

    For Each sheet In currentWorkbook.Worksheets
        If (sheet.Name = "Individual screening" And Not noIndividualData) Or (sheet.Name = "Entity screening" And Not noEntityData) Then
            If sheet.Name = "Individual screening" Then
                startRow = indStartRow
                lastRow = indLastRow
                lastColumn = indLastColumn
            ElseIf sheet.Name = "Entity screening" Then
                startRow = entStartRow
                lastRow = entLastRow
                lastColumn = entLastColumn
            End If
            With sheet
                fundNameColumn = Split(.Rows(startRow - 1).Find(What:="Fund Name", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Address, "$")(1)
                clientNameColumn = Split(.Rows(startRow - 1).Find(What:="Client Name", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Address, "$")(1)

                ' fundID mapping
                SQLquery = "SELECT FundID, FundName FROM tblcorePartyTypeFund"
                rs.Open SQLquery, conn, adLockReadOnly
                Do While Not rs.EOF
                    fundMapping(rs.Fields("FundName").Value) = rs.Fields("FundID").Value
                    rs.MoveNext
                Loop
                rs.Close
                For Each cell In .Range(fundNameColumn & startRow & ":" & fundNameColumn & lastRow)
                    If cell.Value2 <> "" Then
                        If fundMapping.Exists(cell.Value2) Then
                            .Cells(cell.row, lastColumn + 2).Value2 = fundMapping(cell.Value2)  ' fundID put in the 2nd column after last column
                        Else
                            MsgBox "The Fund Name: '" & cell.Value2 & "' does not exist in EMS. Please check your input."
                            GoTo ExitSub
                        End If
                    End If
                Next cell

                ' clientID mapping
                SQLquery = "SELECT ClientID, ClientName FROM tblcorePartyTypeClient"
                rs.Open SQLquery, conn, adLockReadOnly
                Do While Not rs.EOF
                    clientMapping(rs.Fields("clientName").Value) = rs.Fields("clientID").Value
                    rs.MoveNext
                Loop
                rs.Close
                For Each cell In .Range(clientNameColumn & startRow & ":" & clientNameColumn & lastRow)
                    If cell.Value2 <> "" Then
                        If clientMapping.Exists(cell.Value2) Then
                            .Cells(cell.row, lastColumn + 3).Value2 = clientMapping(cell.Value2)  ' clientID put in the 3rd column after last column
                        Else
                            MsgBox "The Client Name: '" & cell.Value2 & "' does not exist in EMS. Please check your input."
                            GoTo ExitSub
                        End If
                    End If
                Next cell
            End With
        End If
    Next sheet

    With currentWorkbook.Sheets("Entity screening")
        If Not noEntityData Then
            Set cmd = New adodb.Command
            'import to SQL server
            For row = entStartRow To entLastRow
                On Error GoTo CannotFindColumn
                entityName = .Cells(row, .Rows(entStartRow - 1).Find(What:="Entity Name", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                chineseName = .Cells(row, .Rows(entStartRow - 1).Find(What:="Chinese Name", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                registeredCountry = .Cells(row, .Rows(entStartRow - 1).Find(What:="Registered Country", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                caseID = .Cells(row, .Rows(entStartRow - 1).Find(What:="Case ID", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                lastScreenDate = Format(.Cells(row, .Rows(entStartRow - 1).Find(What:="Last screen date", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value, "DD MMMM YYYY")
                fundID = .Cells(row, Split(Cells(row, entLastColumn + 2).Address, "$")(1)).Value2
                clientID = .Cells(row, Split(Cells(row, entLastColumn + 3).Address, "$")(1)).Value2
                If InStr(currentWorkbook.Name, "Output") Then
                    screeningFlagPep = .Cells(row, .Rows(indStartRow - 1).Find(What:="Potential PEP", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                    If screeningFlagPep = "Y" Then screeningFlagPep = "1"
                    If screeningFlagPep = "N" Then screeningFlagPep = "0"
                    detailsFlag = .Cells(row, .Rows(indStartRow - 1).Find(What:="Details Flag", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                    archive = .Cells(row, .Rows(entStartRow - 1).Find(What:="Archive", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                    If archive = "Y" Then archive = "1"
                    If archive = "N" Then archive = "0"
                End If
                On Error GoTo Error_Handler
                Set cmd = New adodb.Command
                With cmd
                    .ActiveConnection = conn
                    .CommandText = "importEntityScreening" 'Procedure name
                    .CommandType = adCmdStoredProc
                    .NamedParameters = True
                    .Parameters.Append .CreateParameter("@EntityName", adVarChar, adParamInput, 128, entityName)
                    .Parameters.Append .CreateParameter("@ChineseName", adVarWChar, adParamInput, 128, chineseName)
                    .Parameters.Append .CreateParameter("@RegCountry", adVarChar, adParamInput, 128, registeredCountry)
                    If fundID = "" Then
                        .Parameters.Append .CreateParameter("@FundID", adSmallInt, adParamInput, , Null)
                    Else
                        .Parameters.Append .CreateParameter("@FundID", adSmallInt, adParamInput, , fundID)
                    End If
                    If clientID = "" Then
                        .Parameters.Append .CreateParameter("@ClientID", adSmallInt, adParamInput, , Null)
                    Else
                        .Parameters.Append .CreateParameter("@ClientID", adSmallInt, adParamInput, , clientID)
                    End If
                    If caseID = "" Then
                        .Parameters.Append .CreateParameter("@CaseID", adVarChar, adParamInput, 128, Null)
                    Else
                        .Parameters.Append .CreateParameter("@CaseID", adVarChar, adParamInput, 128, caseID)
                    End If
                    If lastScreenDate = "" Then
                        .Parameters.Append .CreateParameter("@LastScreenDate", adVarChar, adParamInput, 128, Null)
                    Else
                        .Parameters.Append .CreateParameter("@LastScreenDate", adVarChar, adParamInput, 128, lastScreenDate)
                    End If
                    If screeningFlagPep = "" Then
                        .Parameters.Append .CreateParameter("@PotentialPep", adBoolean, adParamInput, , Null)
                    Else
                        .Parameters.Append .CreateParameter("@PotentialPep", adBoolean, adParamInput, , screeningFlagPep)
                    End If
                    If detailsFlag = "" Then
                        .Parameters.Append .CreateParameter("@DetailsFlag", adVarChar, adParamInput, 4000, Null)
                    Else
                        .Parameters.Append .CreateParameter("@DetailsFlag", adVarChar, adParamInput, 4000, detailsFlag)
                    End If
                    If archive = "" Then
                        .Parameters.Append .CreateParameter("@Archive", adBoolean, adParamInput, , False)
                    Else
                        .Parameters.Append .CreateParameter("@Archive", adBoolean, adParamInput, , archive)
                    End If

                    ' collect the count of records update or insert
                    Set result = New adodb.Recordset
                    result.Open cmd
                    entInsertedCount = entInsertedCount + result.Fields("InsertCount").Value
                    entUpdatedCount = entUpdatedCount + result.Fields("UpdateCount").Value
                End With
            Next row
            ' delete records after import
            .Range("B" & entStartRow & ":" & entLastColumnStr & entLastRow).ClearContents
        End If
    End With

    With currentWorkbook.Sheets("Individual screening")
        If Not noIndividualData Then
            ' EntityID mapping
            SQLquery = "SELECT EntityID, EntityName FROM tblEntityScreening"
            rs.Open SQLquery, conn, adLockReadOnly
            Do While Not rs.EOF
                entityMapping(rs.Fields("EntityName").Value) = rs.Fields("EntityID").Value
                rs.MoveNext
            Loop
            rs.Close
            For Each cell In .Range(entityNameColumn & indStartRow & ":" & entityNameColumn & indLastRow)
                If cell.Value2 <> "" Then
                    If entityMapping.Exists(cell.Value2) Then
                        .Cells(cell.row, indLastColumn + 1).Value2 = entityMapping(cell.Value2)  ' entityID put in the 1st column after last column
                    Else
                        MsgBox "The Entity Name: '" & cell.Value2 & "' does not exist in EMS. Please import this entity first."
                        GoTo ExitSub
                    End If
                End If
            Next cell

            'import to SQL server
            For row = indStartRow To indLastRow
                On Error GoTo CannotFindColumn
                givenName = .Cells(row, .Rows(indStartRow - 1).Find(What:="Given Name", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                familyName = .Cells(row, .Rows(indStartRow - 1).Find(What:="Family Name", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                chineseName = .Cells(row, .Rows(indStartRow - 1).Find(What:="Chinese Name", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value
                gender = .Cells(row, .Rows(indStartRow - 1).Find(What:="Gender", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                placeOfBirth = .Cells(row, .Rows(indStartRow - 1).Find(What:="Place of Birth", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                dateOfBirth = .Cells(row, .Rows(indStartRow - 1).Find(What:="Date of birth", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value
                citizenship = .Cells(row, .Rows(indStartRow - 1).Find(What:="Citizenship", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                cityLocation = .Cells(row, .Rows(indStartRow - 1).Find(What:="City Location", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                regionLocation = .Cells(row, .Rows(indStartRow - 1).Find(What:="Region Location", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                countryLocation = .Cells(row, .Rows(indStartRow - 1).Find(What:="Country Location", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                entityID = .Cells(row, Split(Cells(row, indLastColumn + 1).Address, "$")(1)).Value2
                capacity = .Cells(row, .Rows(indStartRow - 1).Find(What:="Capacity", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                caseID = .Cells(row, .Rows(indStartRow - 1).Find(What:="Case ID", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                lastScreenDate = Format(.Cells(row, .Rows(indStartRow - 1).Find(What:="Last screen date", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value, "DD MMMM YYYY")
                fundID = .Cells(row, Split(Cells(row, indLastColumn + 2).Address, "$")(1)).Value2
                clientID = .Cells(row, Split(Cells(row, indLastColumn + 3).Address, "$")(1)).Value2
                If InStr(currentWorkbook.Name, "Output") Then
                    screeningFlagPep = .Cells(row, .Rows(indStartRow - 1).Find(What:="Potential PEP", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                    If screeningFlagPep = "Y" Then screeningFlagPep = "1"
                    If screeningFlagPep = "N" Then screeningFlagPep = "0"
                    detailsFlag = .Cells(row, .Rows(indStartRow - 1).Find(What:="Details Flag", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                    archive = .Cells(row, .Rows(indStartRow - 1).Find(What:="Archive", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Column).Value2
                    If archive = "Y" Then archive = "1"
                    If archive = "N" Then archive = "0"
                End If
                On Error GoTo Error_Handler
                Set cmd = New adodb.Command
                With cmd
                    .ActiveConnection = conn
                    .CommandText = "importIndividualScreening" 'Procedure name
                    .CommandType = adCmdStoredProc
                    .NamedParameters = True
                    .Parameters.Append .CreateParameter("@GivenName", adVarChar, adParamInput, 128, givenName)
                    .Parameters.Append .CreateParameter("@FamilyName", adVarChar, adParamInput, 128, familyName)
                    .Parameters.Append .CreateParameter("@ChineseName", adVarWChar, adParamInput, 128, chineseName)
                    .Parameters.Append .CreateParameter("@Gender", adVarChar, adParamInput, 6, gender)
                    .Parameters.Append .CreateParameter("@PlaceOfBirth", adVarChar, adParamInput, 128, placeOfBirth)
                    If dateOfBirth = "" Then
                        .Parameters.Append .CreateParameter("@DateOfBirth", adVarChar, adParamInput, 128, Null)
                    Else
                        .Parameters.Append .CreateParameter("@DateOfBirth", adVarChar, adParamInput, 128, dateOfBirth)
                    End If
                    .Parameters.Append .CreateParameter("@Citizenship", adVarChar, adParamInput, 128, citizenship)
                    .Parameters.Append .CreateParameter("@CityLocation", adVarChar, adParamInput, 128, cityLocation)
                    .Parameters.Append .CreateParameter("@RegionLocation", adVarChar, adParamInput, 128, regionLocation)
                    .Parameters.Append .CreateParameter("@CountryLocation", adVarChar, adParamInput, 128, countryLocation)
                    If entityID = "" Then
                        .Parameters.Append .CreateParameter("@EntityID", adSmallInt, adParamInput, , Null)
                    Else
                        .Parameters.Append .CreateParameter("@EntityID", adSmallInt, adParamInput, , entityID)
                    End If
                    If fundID = "" Then
                        .Parameters.Append .CreateParameter("@FundID", adSmallInt, adParamInput, , Null)
                    Else
                        .Parameters.Append .CreateParameter("@FundID", adSmallInt, adParamInput, , fundID)
                    End If
                    If clientID = "" Then
                        .Parameters.Append .CreateParameter("@ClientID", adSmallInt, adParamInput, , Null)
                    Else
                        .Parameters.Append .CreateParameter("@ClientID", adSmallInt, adParamInput, , clientID)
                    End If
                    If capacity = "" Then
                        .Parameters.Append .CreateParameter("@Capacity", adVarChar, adParamInput, 128, Null)
                    Else
                        .Parameters.Append .CreateParameter("@Capacity", adVarChar, adParamInput, 128, capacity)
                    End If
                    If caseID = "" Then
                        .Parameters.Append .CreateParameter("@CaseID", adVarChar, adParamInput, 128, Null)
                    Else
                        .Parameters.Append .CreateParameter("@CaseID", adVarChar, adParamInput, 128, caseID)
                    End If
                    If lastScreenDate = "" Then
                        .Parameters.Append .CreateParameter("@LastScreenDate", adVarChar, adParamInput, 128, Null)
                    Else
                        .Parameters.Append .CreateParameter("@LastScreenDate", adVarChar, adParamInput, 128, lastScreenDate)
                    End If
                    If screeningFlagPep = "" Then
                        .Parameters.Append .CreateParameter("@PotentialPep", adBoolean, adParamInput, , Null)
                    Else
                        .Parameters.Append .CreateParameter("@PotentialPep", adBoolean, adParamInput, , screeningFlagPep)
                    End If
                    If detailsFlag = "" Then
                        .Parameters.Append .CreateParameter("@DetailsFlag", adVarChar, adParamInput, 4000, Null)
                    Else
                        .Parameters.Append .CreateParameter("@DetailsFlag", adVarChar, adParamInput, 4000, detailsFlag)
                    End If
                    If archive = "" Then
                        .Parameters.Append .CreateParameter("@Archive", adBoolean, adParamInput, , False)
                    Else
                        .Parameters.Append .CreateParameter("@Archive", adBoolean, adParamInput, , archive)
                    End If

                    ' collect the count of records update or insert
                    Set result = New adodb.Recordset
                    result.Open cmd
                    indInsertedCount = indInsertedCount + result.Fields("InsertCount").Value
                    indUpdatedCount = indUpdatedCount + result.Fields("UpdateCount").Value
                End With
            Next row
            ' delete records after import
            .Range("B" & indStartRow & ":" & indLastColumnStr & indLastRow).ClearContents
        End If
    End With

    If noEntityData And noIndividualData Then
        MsgBox "No data is found"
    Else
        MsgBox indInsertedCount & " individuals have been created." & vbNewLine & indUpdatedCount & " individuals have been updated." & _
                vbNewLine & entInsertedCount & " entities have been created." & vbNewLine & entUpdatedCount & " entities have been updated."
    End If

    ExitSub:
        conn.Close
        Set conn = Nothing
        Set rs = Nothing
        Set cmd = Nothing
        ' Clear mapping
        With currentWorkbook.Sheets("Individual Screening")
            .Range(indClearColumnStr & indStartRow & ":Z999").ClearContents
        End With
        With currentWorkbook.Sheets("Entity Screening")
            .Range(entClearColumnStr & entStartRow & ":Z999").ClearContents
        End With

        Application.ScreenUpdating = True
        Exit Sub

    Error_Handler:
        MsgBox Err.Description
        GoTo ExitSub

    CannotFindColumn:
        MsgBox "Column name in this workbook is not updated." & vbCrLf & "Please generate a new workbook."
        GoTo ExitSub

End Sub

Sub ExportWCdata()

    Dim currentWorkbook As Workbook
    Dim IDs As New Collection, id As Variant
    Dim exportAll As Integer
    Dim dateColumn As String
    Dim cell As Range

    Dim conn As adodb.Connection
    Dim cmd As adodb.Command
    Dim rs As adodb.Recordset
    Dim baseSQL As String, SQL As String

    Set currentWorkbook = ActiveWorkbook
    ' create a output WC workbook if invalid active workbook
    If InStr(currentWorkbook.Name, "Output for WC") = 0 Then
        OpenTemplateWorkbook (outputWCPath)
        If InStr(ActiveWorkbook.Name, "Output for WC") = 0 Then Exit Sub
        Set currentWorkbook = ActiveWorkbook
    End If

    On Error GoTo Error_Handler

    Set conn = New adodb.Connection
    conn.Open ConnString

    exportAll = MsgBox("Would you like to export all the individual and entity?", vbOKCancel)
    If exportAll = 1 Then
        SQL = "Select * From exportScreeningWCByClient(-1)"
    Else
        ' search engine to get rows by multiple client or fund names
        Set IDs = searchEngine_ClientFund(conn)
        If IDs.Count = 0 Then GoTo ExitSub
        SQL = ""
        For Each id In IDs
            Select Case selectedColumn
                Case 2
                    baseSQL = "Select * From exportScreeningWCByClient(" & id & ")"
                Case 6
                    baseSQL = "Select * From exportScreeningWCByFund(" & id & ")"
            End Select
            ' union sql query for each id
            If SQL = "" Then
                SQL = baseSQL
            Else
                SQL = SQL & " UNION " & baseSQL
            End If
        Next id
    End If
    Application.ScreenUpdating = False
    ' extract data from SQL server
    Set rs = New adodb.Recordset
    rs.Open SQL, conn, adOpenForwardOnly, adLockReadOnly

    ' check if the recordset is empty
    If rs.EOF Then
        MsgBox "The data with that Fund/Client name was not found. Please try again."
        GoTo ExitSub
    End If

    With currentWorkbook.Sheets(1)
        .Range("A2").CopyFromRecordset rs
        ' date formatting
        dateColumn = Split(.Rows(1).Find(What:="Date Of Birth", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Address, "$")(1)
        For Each cell In .Range(dateColumn & "2:" & dateColumn & "9999")
            Select Case True
                Case cell.Value = "01/01/1900"  'null value
                    cell.Value = ""
                Case cell.Value Like "??/????" Or cell.Value Like "?/????"   'mm/yyyy or m/yyyy
                    cell.NumberFormat = "mmm-yyyy"
                    cell.Value = CDate(cell.Value)
                Case cell.Value Like "*/*/????"  'dd/mm/yyyy or dd/m/yyyy or d/mm/yyyy or d/m/yyyy
                    cell.NumberFormat = "dd-mmm-yyyy"
                    cell.Value = CDate(cell.Value)
            End Select
        Next cell
        .Range(dateColumn & "2:" & dateColumn & "9999").HorizontalAlignment = xlRight
    End With

    ExitSub:
        conn.Close
        Set conn = Nothing
        Set rs = Nothing
        Set cmd = Nothing
        Application.ScreenUpdating = True
        Exit Sub

    Error_Handler:
        MsgBox Err.Description
        GoTo ExitSub

End Sub

Sub ExportWCtemplate()
    Dim currentWorkbook As Workbook
    Dim IDs As New Collection, id As Variant
    Dim exportAll As Integer
    Dim noIndividualData As Boolean, noEntityData As Boolean
    Dim lastRow As Integer
    Dim sheet As Worksheet
    Dim cell As Range
    Dim screeningFlagColumn As String, archiveColumn As String

    Dim conn As adodb.Connection
    Dim cmd As adodb.Command
    Dim rs As adodb.Recordset
    Dim baseSQLindividual As String, baseSQLentity As String, SQLindividual As String, SQLentity As String

    Set currentWorkbook = ActiveWorkbook
    ' create an output template workbook if invalid active workbook
    If Not SheetExist("Individual screening", currentWorkbook) Or Not SheetExist("Entity screening", currentWorkbook) Or InStr(currentWorkbook.Name, "Output - Screening name template") = 0 Then
        OpenTemplateWorkbook (outputTemplatePath)
        If InStr(ActiveWorkbook.Name, "Output - Screening name template") = 0 Then Exit Sub
        Set currentWorkbook = ActiveWorkbook
    End If

    On Error GoTo Error_Handler

    Set conn = New adodb.Connection
    conn.Open ConnString

    exportAll = MsgBox("Would you like to export all the individual and entity?", vbOKCancel)
    If exportAll = 1 Then
        SQLindividual = "Select * From exportIndScreeningTempByClient(-1)"
        SQLentity = "Select * From exportEntScreeningTempByClient(-1)"
    Else
        ' search engine to get rows by multiple client or fund names
        Set IDs = searchEngine_ClientFund(conn)
        If IDs.Count = 0 Then GoTo ExitSub
        SQLindividual = ""
        SQLentity = ""
        For Each id In IDs
            Select Case selectedColumn
                Case 2
                    baseSQLindividual = "Select * From exportIndScreeningTempByClient(" & id & ")"
                    baseSQLentity = "Select * From exportEntScreeningTempByClient(" & id & ")"
                Case 6
                    baseSQLindividual = "Select * From exportIndScreeningTempByFund(" & id & ")"
                    baseSQLentity = "Select * From exportEntScreeningTempByFund(" & id & ")"
            End Select
            ' union sql query for each id
            If SQLindividual = "" Then
                SQLindividual = baseSQLindividual
            Else
                SQLindividual = SQLindividual & " UNION " & baseSQLindividual
            End If
            If SQLentity = "" Then
                SQLentity = baseSQLentity
            Else
                SQLentity = SQLentity & " UNION " & baseSQLentity
            End If
        Next id
    End If

    Application.ScreenUpdating = False
    ' extract data from SQL server
    For Each sheet In currentWorkbook.Sheets
        With sheet
            Set rs = New adodb.Recordset
            If sheet.Name = "Individual screening" Then
                rs.Open SQLindividual, conn, adOpenForwardOnly, adLockReadOnly
            Else
                rs.Open SQLentity, conn, adOpenForwardOnly, adLockReadOnly
            End If
            ' check if the recordset is empty
            If rs.EOF Then
                If sheet.Name = "Individual screening" Then noIndividualData = True Else noEntityData = True
            End If

            .Range("B4").CopyFromRecordset rs
            Set rs = Nothing
        End With
    Next sheet

    If noIndividualData And noEntityData Then
        MsgBox "The data with that Fund/Client name was not found. Please try again."
        GoTo ExitSub
    End If

    ' Format Screening Flag PEP column
    For Each sheet In currentWorkbook.Sheets
        With sheet
            lastRow = .Range("B9999").End(xlUp).row
            If lastRow > 3 Then
                screeningFlagColumn = Split(.Rows(3).Find(What:="Potential PEP", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Address, "$")(1)
                archiveColumn = Split(.Rows(3).Find(What:="Archive", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Address, "$")(1)
                For Each cell In .Range(screeningFlagColumn & "4:" & screeningFlagColumn & lastRow & "," & archiveColumn & "4:" & archiveColumn & lastRow)
                    If cell.Value2 = "True" Then cell.Value2 = "Y"
                    If cell.Value2 = "False" Then cell.Value2 = "N"
                Next cell
                With .Range(screeningFlagColumn & "4:" & screeningFlagColumn & lastRow & "," & archiveColumn & "4:" & archiveColumn & lastRow).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Y,N"
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowError = True
                End With
            End If
        End With
    Next sheet

    ExitSub:
        conn.Close
        Set conn = Nothing
        Set rs = Nothing
        Set cmd = Nothing
        Application.ScreenUpdating = True
        Exit Sub

    Error_Handler:
        MsgBox Err.Description
        GoTo ExitSub

End Sub

Public Function searchEngine_ClientFund(conn As adodb.Connection) As Collection
' this map either the selected client or fund name into ID
' and return a collection of selected IDs

    Dim baseSQL As String, SQL As String
    Dim rs As adodb.Recordset
    Dim searchKeyword As String
    Dim wbSelector As Workbook
    Dim selectedRange As Range, IDcell As Range
    Dim IDs As New Collection
    Dim vaildRange As Boolean

    ' allow user to search by client name or fund name
    baseSQL = "SELECT c.ClientID, c.ClientName, ClientIndividualCount, ClientEntityCount, f.FundID, f.FundName, FundIndividualCount, FundEntityCount " & _
                "FROM tblcorePartyTypeClient AS c " & _
                "INNER JOIN tblcorePartyTypeMainFund AS m ON m.ClientID = c.ClientID " & _
                "FULL OUTER JOIN tblcorePartyTypeFund AS f ON f.MainFundID = m.MainFundID " & _
                "LEFT JOIN (SELECT ClientID, COUNT(IndividualID) AS ClientIndividualCount FROM tblIndividualScreening GROUP BY ClientID) AS ci ON ci.ClientID = c.ClientID " & _
                "LEFT JOIN (SELECT ClientID, COUNT(EntityID) AS ClientEntityCount FROM tblEntityScreening GROUP BY ClientID) AS ce ON ce.ClientID = c.ClientID " & _
                "LEFT JOIN (SELECT FundID, COUNT(IndividualID) AS FundIndividualCount FROM tblIndividualScreening GROUP BY FundID) AS fi ON fi.FundID = f.FundID " & _
                "LEFT JOIN (SELECT FundID, COUNT(EntityID) AS FundEntityCount FROM tblEntityScreening GROUP BY FundID) AS fe ON fe.FundID = f.FundID " & _
                "WHERE (f.FundName LIKE '%SEARCH_KEYWORD%' OR c.ClientName LIKE '%SEARCH_KEYWORD%') And (ClientIndividualCount Is Not Null Or ClientEntityCount Is Not Null Or FundIndividualCount Is Not Null Or FundEntityCount Is Not Null)"

    searchKeyword = Application.InputBox("Please type in keyword of the fund name or client name to be exported:" & vbCrLf & _
                                        "(Click Cancel if you would like to choose from list of all funds or clients instead)", Type:=2)
    If searchKeyword = "False" Then searchKeyword = vbNullString
    SQL = Replace(baseSQL, "SEARCH_KEYWORD", searchKeyword)

    Set rs = New adodb.Recordset
    rs.Open SQL, conn, adLockReadOnly

    ' check if the recordset is empty
    If rs.EOF Then
        MsgBox "The Fund/Client name was not found." & vbCrLf & "Please check your input and try again."
        Exit Function
    End If

    ' create new workbook to display the list of searched fund name
    Set wbSelector = Workbooks.Add
    With wbSelector.Sheets(1)
        .Range("B1").Value2 = "Client Name"
        .Range("C1").Value2 = "Individual count"
        .Range("D1").Value2 = "Entity count"
        .Range("F1").Value2 = "Fund Name"
        .Range("G1").Value2 = "Individual count"
        .Range("H1").Value2 = "Entity count"
        .Range("A2").CopyFromRecordset rs
        .Columns("A").Hidden = True
        .Columns("E").Hidden = True
        .Columns("B:D").AutoFit
        .Columns("F:H").AutoFit
        With .Range("B1:H1")
            .Interior.ColorIndex = 48
            .Font.Bold = True
        End With
        .Columns("F").Borders(xlEdgeLeft).LineStyle = xlContinuous

        ' choose name from the list and return ID, either client or fund
        vaildRange = False
        Do While Not vaildRange
            On Error Resume Next
            Set selectedRange = Application.InputBox("Select one or more from either client or fund names:", Type:=8)
            On Error GoTo 0
            If selectedRange Is Nothing Then
                wbSelector.Close SaveChanges:=False
                Exit Function
            ElseIf selectedRange.Columns.Count > 1 Then
                MsgBox "Please only select either client or fund names"
            ElseIf selectedRange.Column <> 2 And selectedRange.Column <> 6 Then
                MsgBox "Please only select the column of client or fund name"
            Else
                vaildRange = True
            End If
        Loop
        ' assign ID to a collection
        For Each IDcell In selectedRange.Offset(0, -1)
            IDs.Add IDcell.Value2
        Next IDcell
    End With
    selectedColumn = selectedRange.Column
    wbSelector.Close SaveChanges:=False
    Set searchEngine_ClientFund = IDs
End Function

Public Sub OpenTemplateWorkbook(templatePath)
    Dim tempWorkbook As Workbook
    Dim savePath As Variant

    Set tempWorkbook = Workbooks.Open(templatePath)
    savePath = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Select a location to save As")
    If savePath <> False Then
        On Error GoTo CloseWorkbook
        tempWorkbook.SaveAs Filename:=savePath
        Exit Sub
    End If

    CloseWorkbook:
        If Not tempWorkbook Is Nothing Then
            tempWorkbook.Close False
        End If
        MsgBox "You must select a location to save the workbook"
End Sub