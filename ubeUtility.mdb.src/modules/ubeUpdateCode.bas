Option Compare Database
Option Explicit
'
' MS-Access Back End Update Utility by Peter Hibbs
'
'   Originally found at:
'       http://www.rogersaccesslibrary.com/forum/back-end-update-utility_topic410.html
'
'   This branch is hosted at:
'       https://github.com/A9G-Data-Droid/MS-Access-Update-Back-end-Utility
'
Public Const gRefTable As String = "Settings"             'defines name of the table in backend to hold the 'ubeVersion'

Const adhcErrObjectExists As Long = 3012                 'see adhGetProp, etc procedures
Const adhcErrNotInCollection As Long = 3270
Const adhcErrInvalidType As Long = 30001

Private beDB As Database
Private thisDb As Database

''' Get BE DB in a reliable manner
Private Property Get backendDB() As Database
    If (beDB Is Nothing) Then
        ' Get link to back-end file and create reference table if missing
        Set beDB = OpenDatabase(GetBEDBPath(gRefTable))
    End If
    
    Set backendDB = beDB
End Property

''' This allows us to only call "CurrentDb" once.
Private Property Get thisDatabase() As Database
    If (thisDb Is Nothing) Then Set thisDb = CurrentDb
    
    Set thisDatabase = thisDb
End Property



''' Get and set the back end version
Public Property Get beVersion() As Long

    On Error GoTo ErrorCode
    
    beVersion = CLng(Nz(DLookup("[DataValue]", gRefTable, "[Setting] = 'ubeVersion'")))
    
ErrorCode:
    If Err.Number > 0 Then beVersion = 0 ' Table not found = must be zero
End Property

Public Property Let beVersion(ByVal newVersion As Long)
    
    On Error GoTo ErrorCode
    
    AddFieldDataToRecord gRefTable, "Setting", "ubeVersion", "DataValue", newVersion
       
ErrorCode:
    If Err.Number > 0 Then
        MsgBox "ERROR: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Unhandled exception"
    End If
    
End Property


''' Push a value in to a field, for a given record
Public Sub AddFieldDataToRecord(ByVal TableName As String, ByVal recordIDName As String, ByVal recordValue As String, ByVal FieldName As String, ByVal fieldValue As String)
    With backendDB.OpenRecordset(TableName, dbOpenDynaset)
        .FindFirst "[" & recordIDName & "] = '" & recordValue & "'"
        If .NoMatch Then
            .AddNew
            .Fields.Item(recordIDName).Value = recordValue
        Else
            .Edit
        End If
        
        .Fields.Item(FieldName).Value = fieldValue
        .Update
    End With
End Sub


''' Update selected back-end file with required changes
'''Entry  (vDeveloper) = TRUE if called from Update form, = FALSE if called from user start up form
'''       (gRefTable) = Name of table in back-end file which holds the 'ubeVersion' reference field
'''Exit   (UpdateBackEndFile) = False if error or True if OK
Public Function UpdateBackEndFile(ByVal vDeveloper As Boolean) As Boolean

    On Error GoTo ErrorCode
    
    'fetch last Version number
    Dim vVersion As Long
    vVersion = beVersion
    
    'if User mode and updates available then
    If vVersion < DMax("ID", "ubeUpdate") Then
        If vDeveloper = False Then
            If MsgBox("WARNING: Your database back-end file requires an update. " _
                    & "Click Yes to continue or No to quit.", vbQuestion + vbYesNo, "Update Pending") = vbNo Then
                Application.Quit
            End If
        End If
        
        BackupDatabase GetBEDBPath
        UpdateBackEndFile = SchemaUpdate(vVersion, vDeveloper)
    End If
    
ErrorCode:
    If Err.Number > 0 Then
        MsgBox "ERROR: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Unhandled exception"
    End If

End Function


''' Creates a backup in a subfolder .\Backups from the cited DB path
''' Entry (dbFullPath) Requires fully qualified path to the DB you want to back up
''' Backup files use ISO Compliant timestamps to avoid collisions.
'''     (Can only be called once per second without error)
Public Sub BackupDatabase(ByVal dbFullPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FilesystemObject")
    
    ' Define and create the backup path
    Dim backupPath As String
    backupPath = fso.BuildPath(fso.GetParentFolderName(dbFullPath), "Backups")
    If Not fso.FolderExists(backupPath) Then fso.CreateFolder backupPath

    ' Define the backup filename
    Dim timeStamp As String
    timeStamp = "_" & Format$(Now, "yyyy-MM-ddThhmmss") & "."  ' ISO Compliant file timestamp
    Dim backupDestination As String
    backupDestination = fso.BuildPath(backupPath, fso.GetBaseName(dbFullPath) & timeStamp & fso.GetExtensionName(dbFullPath))
    
    ' Create the backup
    DBEngine.CompactDatabase dbFullPath, backupDestination
End Sub


''' Returns full path to the backend
'''     Optionally will use the table cited to pick a specific back end by link
'''     When no table is specified it uses the first linked table it finds
Private Function GetBEDBPath(Optional ByVal knownTable As String = vbNullString) As String

    On Error GoTo ErrorCode
    
    Dim vPathname As String
    
    If knownTable = vbNullString Then
        Dim tdf As TableDef
        For Each tdf In thisDatabase.TableDefs          'loop thru all tables
            If (tdf.Attributes And dbSystemObject) = 0 Then 'skip system tables
                If Nz(tdf.Connect) <> vbNullString Then 'if table is linked then
                    vPathname = tdf.Connect          'fetch connection string
                    Exit For                         'short circuit
                End If
            End If
        Next tdf
    Else
        vPathname = GetConnectFromTable(knownTable)
    End If
    
    ' remove everything before ';DATABASE='
    GetBEDBPath = Mid$(vPathname, InStrRev(vPathname, "=") + 1)
    
ErrorCode:
    If Err.Number > 0 Then
        MsgBox "ERROR: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Unhandled exception"
    End If
End Function

''' Returns a connection string using the name of a linked table passed in.
Private Function GetConnectFromTable(ByVal knownTable As String) As String

    On Error GoTo ErrorCode

ResumeError:
    
    ' fetch connect def
    GetConnectFromTable = thisDatabase.TableDefs.Item(knownTable).Connect
    
ErrorCode:
    
    If Err.Number = 3265 Or Err.Number = 3078 Then
        ' Create if table does not exist
        If AddReferenceTable(knownTable) Then Resume ResumeError
    End If
End Function


''' This is the main procedure that does all the work on the back end.
'''     It will update the backend to the version passed in.
'''     Needs to know if you are running in developer mode.
Private Function SchemaUpdate(ByVal vVersion As Long, ByVal vDeveloper As Boolean) As Boolean

    On Error GoTo ErrorCode
    
    Dim vID As Variant
    vID = "Unknown"
    
    Dim vPathname As String
    vPathname = GetBEDBPath(gRefTable)
    
    Dim updateList As DAO.Recordset
    Set updateList = thisDatabase.OpenRecordset("SELECT * FROM ubeUpdate WHERE ID > " & vVersion & " ORDER BY ID") 'make list of updates required
    
    Dim vTableName As String
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Do Until updateList.EOF                             'step thru list
        vID = updateList.Fields().Item("ID").Value                             'fetch ID for error message
        
        DoCmd.OpenForm "ubeUpdating"             'show Updating Back End message
        With Forms.Item("ubeUpdating").Controls
            .Item("ShowFileName").Caption = backendDB.Name
            .Item("WaitLabel").Caption = "Step " & vID
        End With
        
        DoEvents
                
        Select Case updateList.Fields().Item("Action").Value                   'select Action type
        Case "Make Table"
            backendDB.Execute "CREATE TABLE [" & updateList.Fields().Item("TableName").Value & "] ([" & updateList.Fields().Item("FieldName").Value & "] " & updateList.Fields().Item("FieldType").Value & ")" 'create table with one field
            backendDB.TableDefs.Refresh                 'refresh table collection
            If TableExists(updateList.Fields().Item("TableName").Value) = True Then DoCmd.DeleteObject acTable, updateList.Fields().Item("TableName") 'if link exists then delete Link
            DoCmd.TransferDatabase acLink, "Microsoft Access", vPathname, acTable, updateList.Fields().Item("TableName"), updateList.Fields().Item("TableName") 'and re-link to new table in BE
            NewFieldDefaults updateList.Fields().Item("TableName").Value, updateList.Fields().Item("FieldName").Value, updateList.Fields().Item("FieldType").Value, Nz(updateList.Fields().Item("Description")) 'always set some properties
            SetProperties updateList.Fields().Item("TableName").Value, updateList.Fields().Item("FieldName").Value, Nz(updateList.Fields().Item("Constraint")), Nz(updateList.Fields().Item("Misc")), Nz(updateList.Fields().Item("Description")) 'set field property (if any)

        Case "Copy Table"
            vTableName = updateList.Fields().Item("TableName").Value            'fetch table name
            If TableExists("ube" & vTableName) = False Then                     'if ube'Table' not exists then
                DoCmd.CopyObject vPathname, , acTable, vTableName               'copy table to back-end file
                DoCmd.Rename "ube" & vTableName, acTable, vTableName            'prefix table name with 'ube'
                'DoCmd.TransferDatabase acLink, "Microsoft Access", vPathname, acTable, vTableName, vTableName 'set link to new table
            Else                                 'if ube'Table' exists then
                DoCmd.CopyObject vPathname, vTableName, acTable, "ube" & vTableName 'copy table to back-end and rename
            End If
            
            If TableExists(vTableName) = True Then DoCmd.DeleteObject acTable, vTableName 'if link exists then delete Link
            DoCmd.TransferDatabase acLink, "Microsoft Access", vPathname, acTable, vTableName, vTableName 'and re-link to new table in BE
            backendDB.TableDefs.Refresh                 'refresh table collection

        Case "Remove Table"
            backendDB.Execute "DROP TABLE [" & updateList.Fields().Item("TableName").Value & "]" 'delete table from back-end file
            If TableExists("ube" & updateList.Fields().Item("TableName").Value) = True Then 'if ube'Table' exists then
                thisDatabase.Execute "DROP TABLE [" & "ube" & updateList.Fields().Item("TableName").Value & "]" 'delete 'ube' table also
            End If
            
            DoCmd.DeleteObject acTable, updateList.Fields().Item("TableName") 'and delete table Link
            backendDB.TableDefs.Refresh                 'refresh table collection

        Case "New Field"
            If updateList.Fields().Item("FieldType").Value = "ATTACHMENT" Then 'if field type = 'ATTACHMENT'
                Set tdf = backendDB.TableDefs.Item(updateList.Fields().Item("TableName")) 'set ref to specified table
                tdf.Fields.Append tdf.CreateField(updateList.Fields().Item("FieldName"), 101) 'add Attachment type field (101 = dbAttachment)
                
                Dim dbLocal As DAO.Database
                Set dbLocal = CurrentDb()        'refresh links to BE (due to bug in A2007)
                For Each tdf In dbLocal.TableDefs 'loop through all tables
                    If tdf.Name = updateList.Fields().Item("TableName").Value Then 'skip if not current table
                        tdf.Connect = ";DATABASE=" & vPathname 'set pathname + filename of back-end
                        tdf.RefreshLink          'and make link to back end
                    End If
                Next tdf
                
                Set tdf = Nothing
            Else
                If updateList.Fields().Item("FieldType").Value = "HYPERLINK" Then 'if field type = 'HYPERLINK'
                    Set tdf = backendDB.TableDefs.Item(updateList.Fields().Item("TableName")) 'set ref to curretn table
                    Set fld = tdf.CreateField(updateList.Fields().Item("FieldName"), dbMemo) 'add Memo field first
                    fld.Attributes = dbHyperlinkField 'set attribute to Hyperlink
                    tdf.Fields.Append fld        'and append field to table
                    tdf.Fields.Refresh
                    Set tdf = Nothing
                Else
                    backendDB.Execute "ALTER TABLE [" & updateList.Fields().Item("TableName").Value & "] ADD [" & updateList.Fields().Item("FieldName").Value & "] " & updateList.Fields().Item("FieldType").Value 'add new field to table
                End If
            End If
            
            backendDB.TableDefs.Refresh                 'refresh table collection
            NewFieldDefaults updateList.Fields().Item("TableName").Value, updateList.Fields().Item("FieldName").Value, updateList.Fields().Item("FieldType").Value, Nz(updateList.Fields().Item("Description")) 'always set some properties
            SetProperties updateList.Fields().Item("TableName").Value, updateList.Fields().Item("FieldName").Value, Nz(updateList.Fields().Item("Constraint")), Nz(updateList.Fields().Item("Misc")), Nz(updateList.Fields().Item("Description")) 'set other field property (if any)

        Case "Delete Field"
            backendDB.Execute "ALTER TABLE [" & updateList.Fields().Item("TableName").Value & "] DROP [" & updateList.Fields().Item("FieldName").Value & "]" 'delete field
            backendDB.TableDefs.Refresh                 'refresh table collection

        Case "Change Type"
            backendDB.Execute "ALTER TABLE [" & updateList.Fields().Item("TableName").Value & "] ALTER [" & updateList.Fields().Item("FieldName").Value & "] " & updateList.Fields().Item("FieldType").Value 'change field type
            backendDB.TableDefs.Refresh                 'refresh table collection
            SetProperties updateList.Fields().Item("TableName").Value, updateList.Fields().Item("FieldName").Value, Nz(updateList.Fields().Item("Constraint")), Nz(updateList.Fields().Item("Misc")), Nz(updateList.Fields().Item("Description")) 'set field property

        Case "Set Property"
            backendDB.TableDefs.Refresh                 'refresh table collection
            SetProperties updateList.Fields().Item("TableName").Value, updateList.Fields().Item("FieldName").Value, Nz(updateList.Fields().Item("Constraint")), Nz(updateList.Fields().Item("Misc")), Nz(updateList.Fields().Item("Description")) 'set field property

        Case "Set Relationship"
            CreateRelationship updateList.Fields().Item("TableName").Value, updateList.Fields().Item("FieldName").Value, updateList.Fields().Item("Constraint").Value, updateList.Fields().Item("Misc").Value, updateList.Fields().Item("Description").Value 'create a Relationship

        Case "Clear Relationship"
            DeleteRelationship updateList.Fields().Item("TableName").Value, updateList.Fields().Item("FieldName").Value, updateList.Fields().Item("Misc").Value, updateList.Fields().Item("Description").Value 'delete a Relationship

        Case "Run Query"
            thisDatabase.Execute updateList.Fields().Item("TableName").Value   'execute UPDATE query

        Case "Execute Code"
            Run updateList.Fields().Item("TableName").Value                    'execute VBA Code

        Case "Run Macro"
            DoCmd.RunMacro updateList.Fields().Item("TableName")         'execute Macro
        
        End Select
        
        beVersion = updateList.Fields().Item("ID").Value                       'copy current ID to ubeVersion field"
        updateList.MoveNext
    Loop
    
    SchemaUpdate = True                          'return Updates OK code

ErrorCode:
    DoCmd.Close acForm, "ubeUpdating"            'close message form
    If Not updateList Is Nothing Then
        updateList.Close
        Set updateList = Nothing
    End If
    
    If Not beDB Is Nothing Then
        beDB.Close
        Set beDB = Nothing
    End If
    
    If Err.Number > 0 Then
        If vDeveloper = True Then                    'if developer mode then
            MsgBox Err.Description & "  (Reference No = " & vID & ")" 'show error + Ref No and exit with False
            Stop
        Else
            MsgBox "ERROR. Back-End Update Failed. (Reference No = " & vID & "). Application will now be shut down to prevent any damage to back-end data file.", vbOKOnly, "Update Fail"
            Application.Quit                         'if update error on user DB then quit
        End If
    End If

End Function


''' Checks if table exists
'''     Entry  (vTableName) = Name of table
'''     Exit   (TableExists) = True if table exists, = False if not
Private Function TableExists(ByRef vTableName As String) As Boolean

    On Error GoTo ErrorCode                      'trap error if next line fails
    
    ' try to read table name from TableDefs
    TableExists = (thisDatabase.TableDefs.Item(vTableName).Name = vTableName)
    Exit Function

ErrorCode:
    TableExists = False                          'TableExists = False if not successful

End Function


''' Checks if relation exists
'''     Entry  (relationName) = Name of table
'''     Exit   (RelationExists) = True if relation exists, = False if not
Public Function RelationExists(ByRef relationName As String) As Boolean

    On Error GoTo ErrorCode                      'trap error if next line fails
    
    ' try to read table name from TableDefs
    RelationExists = (backendDB.Relations.Item(relationName).Name = relationName)
    Exit Function

ErrorCode:
    RelationExists = False                          'TableExists = False if not successful

End Function


'''Change or add a field property
'''Entry  (vTableName) = name of table to change
'''       (vFieldName) = name of field to change
'''       (vPropertyType) = name of field property to be changed (if NULL then just change Description property, if any)
'''       (vParameters) = any required parameters (i.e. Field default value, New field name or Ordinal position, etc)
'''       (vDescription) = text for description column of specified field or other data
'''Exit   Specified property changed
'''       Any errors handled by main UpdateBackEndFile routine
Private Sub SetProperties(ByRef vTableName As String, ByRef vFieldName As String, ByVal vPropertyType As String, ByRef vParameters As String, ByRef vDescription As String)

    Dim vStatus As Boolean
    Dim vR As Variant
    Dim vRTF As Long

    Select Case vPropertyType
    Case "Text Field Size ="
        If backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName).Type = dbText Then 'if Field type = TEXT then
            backendDB.Execute "ALTER TABLE [" & vTableName & "] ALTER COLUMN [" & vFieldName & "] TEXT (" & vParameters & ")" 'change field size
        End If
    Case "Set Compression"
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "UniCodeCompression", True) 'set UniCode Compression
    Case "Required ="
        If vParameters = "Yes" Or vParameters = "True" Then vStatus = True 'convert Yes/True to Boolean
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "Required", vStatus) 'set set Required Yes/No
    Case "Allow Zero Len ="
        If vParameters = "Yes" Or vParameters = "True" Then vStatus = True 'convert Yes/True to Boolean
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "AllowZeroLength", vStatus) 'set Allow Zero Length
    Case "Default Value ="
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "DefaultValue", vParameters) 'set Default Value
    Case "Input Mask ="
        vR = adhSetProp(thisDatabase.TableDefs.Item(vTableName).Fields.Item(vFieldName), "InputMask", vParameters) 'set Input Mask Value
    Case "Format ="
        vR = adhSetProp(thisDatabase.TableDefs.Item(vTableName).Fields.Item(vFieldName), "Format", vParameters) 'set Format type
    Case "Decimal Places ="
        vR = adhSetProp(thisDatabase.TableDefs.Item(vTableName).Fields.Item(vFieldName), "DecimalPlaces", vParameters) 'set Decimal Places
    Case "Validation Rule ="
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "ValidationRule", vParameters) 'set Validation Rule
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "ValidationText", vDescription) 'set Validation Text
    Case "New Field Name ="
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "Name", vParameters) 'set/change Field Name
    Case "Description ="
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "Description", vDescription) 'set Description field
    Case "Ordinal Position ="
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "OrdinalPosition", vParameters) 'set Field Ordinal posn
    Case "Set Primary Key ="
        If vParameters <> vbNullString Then vFieldName = vFieldName & "," & vParameters 'add in extra fields (if any)
        vR = adhCreatePrimaryKey(vTableName, "PrimaryKey", vFieldName) 'set Primary Key/s
    Case "Indexed (No)"
        vR = FindIndex(vTableName, vFieldName)   'find Index name
        If vR <> vbNullString Then backendDB.Execute "DROP INDEX [" & vR & "] ON [" & vTableName & "]" 'remove Index (if any)
    Case "Indexed (Dup OK)"
        vR = FindIndex(vTableName, vFieldName)   'find Index name
        If vR <> vbNullString Then backendDB.Execute "DROP INDEX [" & vR & "] ON [" & vTableName & "]" 'remove Index (if any)
        backendDB.Execute "CREATE INDEX [" & vFieldName & "] ON [" & vTableName & "] ([" & vFieldName & "])" 'set Index (use Field name as Index name)
    Case "Indexed (No Dup)"
        vR = FindIndex(vTableName, vFieldName)   'find Index name
        If vR <> vbNullString Then backendDB.Execute "DROP INDEX [" & vR & "] ON [" & vTableName & "]" 'remove Index (if any)
        backendDB.Execute "CREATE UNIQUE INDEX [" & vFieldName & "] ON [" & vTableName & "] ([" & vFieldName & "])" 'set unique Index (use Field name as Index name)
    Case "Fill With ="
        FillField vTableName, vFieldName, vParameters 'copy data to field
    Case "Rich Text ="
        If vParameters = "Yes" Or vParameters = "True" Then vRTF = 1 Else vRTF = 0 'convert Yes/True to 1/0
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "TextFormat", vRTF) 'set/clear Rich Text format
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "Description", vDescription) 'set Description field
    Case "Caption Name ="
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "Caption", vParameters) 'set Caption property
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "Description", vDescription) 'set Description field
    Case "Smart Tags ="
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "SmartTags", vParameters) 'set Smart Tags property
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "Description", vDescription) 'set Description field
    End Select

End Sub


'''Fill a field with data for all records
'''Entry  (vTableName) = name of table to fill
'''       (vFieldName) = name of field to fill
'''       (vParameter) = data to be copied to table
'''       (db) = Database object referenced to back-end file
'''Exit   Specified field in all records filled with specified value (Note. In Text/Memo fields any double quotes replaced with two single quotes)
'''       Any errors handled by main UpdateBackEndFile routine
Private Sub FillField(ByRef vTableName As String, ByRef vFieldName As String, ByRef vParameter As String)

    Dim vFieldType As Long
    Dim vData As String
    Const QUOTE As String = """"                           'Used in place of Double Quotes

    vFieldType = backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName).Type 'fetch field type
    Select Case vFieldType
    Case dbText, dbMemo                          'if Text or Memo then
        vParameter = Replace(Nz(vParameter), """", "''")                                    'replace any Double Quotes with two Single Quotes
        vData = QUOTE & vParameter & QUOTE       'surround with double quotes
        backendDB.Execute "UPDATE [" & vTableName & "] SET [" & vFieldName & "] = " & vData & " WHERE [" & vFieldName & "] IS NULL" 'fill field for all blank records
    Case dbDate                                  'if date or time
        vData = "#" & Format$(vParameter, "mm\/dd\/yyyy") & "#" 'reformat date for US mode
        backendDB.Execute "UPDATE [" & vTableName & "] SET [" & vFieldName & "] = " & vData & " WHERE [" & vFieldName & "] IS NULL" 'fill field for all blank records
    Case dbBoolean
        backendDB.Execute "UPDATE [" & vTableName & "] SET [" & vFieldName & "] = " & vParameter 'set field to True or False
    Case Else
        vData = vParameter                       'numeric values need no changes
        backendDB.Execute "UPDATE [" & vTableName & "] SET [" & vFieldName & "] = " & vData & " WHERE [" & vFieldName & "] IS NULL" 'fill field for all blank records
    End Select

End Sub


'''Set some properties for new fields regardless
'''Entry  (vTableName) = name of table to change
'''       (vFieldName) = name of field to change
'''       (vFieldType) = field property type
'''       (vDescription) = text for description column of specified field
'''       (db) = Database object referenced to back-end file
'''Exit   Specified field properties set (delete any you don't want)
'''       Any errors handled by main UpdateBackEndFile routine
Private Sub NewFieldDefaults(ByRef vTableName As String, ByRef vFieldName As String, ByVal vFieldType As String, ByRef vDescription As String)

    Dim vR As Variant
    Dim fld As DAO.Field

    vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "Description", vDescription) 'set Description field (if any)
    Select Case vFieldType
    Case "BYTE", "SHORT", "LONG", "SINGLE", "DOUBLE", "CURRENCY" 'select Number fields only
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "DefaultValue", 0) 'set Number fields Default Value to 0
    Case "TEXT", "MEMO", "HYPERLINK"             'select TEXT & MEMO & HYPERLINK fields
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "UniCodeCompression", True) 'always set UniCode Compression on Text/Memo
        vR = adhSetProp(backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName), "AllowZeroLength", True) 'always set AllowZeroLength on Text/Memo
    Case "YESNO"                                 'select Yes/No fields
        Set fld = backendDB.TableDefs.Item(vTableName).Fields.Item(vFieldName) 'change YesNo field format
        SetPropertyDAO fld, "DisplayControl", dbInteger, CInt(acCheckBox) 'to Check Box type
    End Select

End Sub


'''Set some properties for new fields regardless
'''Entry  (vTableName) = name of table with indexed field
'''       (vFieldName) = name of indexed field
'''       (db) = Database object referenced to back-end file
'''Exit   Index name for selected table/field returned or "" if none
'''       Any errors handled by main UpdateBackEndFile routine
Private Function FindIndex(ByRef vTableName As String, ByVal vFieldName As String) As String

    Dim idx As DAO.Index
    Dim tdf As DAO.TableDef
    
    With backendDB.TableDefs
        Set tdf = .Item(vTableName)           'define required table
    End With
    
    For Each idx In tdf.Indexes                  'search Indexes
        If InStr(1, idx.Fields, "+" & vFieldName) > 0 Then 'if index field holds ("+" & field name) then
            FindIndex = idx.Name                 'fetch Index name and
            Exit Function                        'return with Index name
        End If
    Next idx

End Function


'''Create or change a relationship between two tables
'''Entry  (vPKTableName) = Name of table for Primary Key
'''       (vPKFieldName) = Name of Primary Key in primary table
'''       (vFKTableName) = Name of table for Foreign Key
'''       (vFKFieldName) = Name of Foreign Key field in Foreign table
'''Exit   (CreateRelationship) = True if Relationship created or = False if error
Private Function CreateRelationship(ByVal vPKTableName As String, ByVal vPKFieldName As String, ByVal vRelationshipType As String, ByVal vFKTableName As String, ByVal vFKFieldName As String) As Boolean
    
    ' From Access 2000 Developer's Handbook by Litwin, Getz, Gilbert (Sybex) Copyright 1999.  All rights reserved.
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim vRelationshipName As String

    On Error GoTo CreateRelationship_Err
    
    Dim vRelType As Long
    Select Case vRelationshipType                'convert Relationship type to Long Integer
    Case "1-1 Not Enforced"                      'One to One (Not Enforced)
        vRelType = 1
    Case "1-1 Casc Updates"                      'One to One (Cascade Updates)
        vRelType = 1 + 256
    Case "1-1 Casc Deletes"                      'One to One (Cascade Deletes)
        vRelType = 1 + 4096
    Case "1-1 Casc Upd/Del"                      'One to One (Cascade Updates and Deletes)
        vRelType = 1 + 4096 + 256
    Case "1-n Not Enforced"                      'One to Many (Not Enforced)
        vRelType = 2
    Case "1-n Casc Updates"                      'One to Many (Cascade Updates)
        vRelType = 256
    Case "1-n Casc Deletes"                      'One to Many (Cascade Deletes)
        vRelType = 4096
    Case "1-n Casc Upd/Del"                      'One to Many (Cascade Updates and Deletes)
        vRelType = 4096 + 256
    End Select
        
    vRelationshipName = vPKTableName & vFKTableName  'relationship name from both table names, like access does
    
    If RelationExists(vRelationshipName) Then
    
        ' Collect existing fields
        Dim existingFields As Object
        Set existingFields = CreateObject("Scripting.Dictionary")
        Dim aField As Field
        For Each aField In backendDB.Relations.Item(vRelationshipName).Fields
            existingFields.Add aField.Name, aField.ForeignName
        Next aField
        
        ' We can't add to existing relationship so delete existing relation and make new relation
        backendDB.Relations.Delete vRelationshipName
        Set rel = backendDB.CreateRelation(vRelationshipName, vPKTableName, vFKTableName, vRelType)
        
        ' Add all the old fields back in
        Dim savedField As Variant
        For Each savedField In existingFields.Keys
            Set fld = rel.CreateField(existingFields(savedField))   'Set the relation's field collection.
            fld.ForeignName = existingFields.Item(savedField)     'set Foreign table field
            rel.Fields.Append fld                    'append foreign field
        Next savedField
    
    Else  ' create relationship link
        Set rel = backendDB.CreateRelation(vRelationshipName, vPKTableName, vFKTableName, vRelType)
    End If
    
    Set fld = rel.CreateField(vPKFieldName)      'Set the relation's field collection.
    fld.ForeignName = vFKFieldName               'set Foreign table field
    rel.Fields.Append fld                        'append foreign field
    backendDB.Relations.Append rel               'append relationship definition to collection
    
    CreateRelationship = True                    'return True
    
CreateRelationship_Exit:
    Exit Function

CreateRelationship_Err:
    Select Case Err.Number
'    Case adhcErrObjectExists                     'If the relationship already exists, just delete it, and then try to append it again.
'        backendDB.Relations.Delete rel.Name
'        Resume
    Case Else
        MsgBox "Error: " & Err.Description & _
               " (" & Err.Number & ")"
        CreateRelationship = False
        Resume CreateRelationship_Exit
    End Select

End Function


'''Delete a relationship between two tables
'''Entry  (vPKTableName) = Name of table for Primary Key
'''       (vPKFieldName) = Name of Primary Key in primary table
'''       (vFKTableName) = Name of table for Foreign Key
'''       (vFKFieldName) = Name of Foreign Key field in Foreign table
'''       (db) = Database object referenced to back-end file
'''Exit   Relationship deleted
'''       Any errors handled by main UpdateBackEndFile routine
Private Sub DeleteRelationship(ByVal vPKTableName As String, ByVal vPKFieldName As String, ByVal vFKTableName As String, ByVal vFKFieldName As String)

    Dim selectedRelation As DAO.Recordset

    Set selectedRelation = backendDB.OpenRecordset("SELECT szRelationship FROM MSysRelationships " _
                             & "WHERE szReferencedObject = '" & vPKTableName & "' " _
                             & "AND szReferencedColumn = '" & vPKFieldName & "' " _
                             & "AND szObject = '" & vFKTableName & "' " _
                             & "AND szColumn = '" & vFKFieldName & "' " _
                             & "OR szReferencedObject = '" & vFKTableName & "' " _
                             & "AND szReferencedColumn = '" & vFKFieldName & "' " _
                             & "AND szObject = '" & vPKTableName & "' " _
                             & "AND szColumn = '" & vPKFieldName & "'") 'fetch Relationship name
                             
    ' if relationship exists then delete relationship
    If Not selectedRelation.BOF And Not selectedRelation.EOF Then
        backendDB.Relations.Delete selectedRelation.Fields().Item("szRelationship").Value
    End If
    
    selectedRelation.Close
    Set selectedRelation = Nothing

End Sub



'Private Function adhGetProp(ByVal obj As Object, ByRef strName As String) As Variant
'
'    '***** This function is not actually used but has been left in just in case it is needed !! ******
'
'    ' From Access 2000  Developer's Handbook
'    ' by Getz, Litwin, and Gilbert. (Sybex)
'    ' Copyright 1999. All Rights Reserved.
'    ' Get the value of a property
'    ' If there isn't a property of the name passed in,
'    ' return an error value, otherwise, return the
'    ' value of the property.
'
'    ' In:
'    '    obj: An object reference
'    '       (db.TableDefs("tblCustomers"), for example)
'    '    strName: Name for the property to get
'    ' Out:
'    '    Return value:
'    '      If an error occurred, an error value (use IsError() to check)
'    '      If not, the value of the property.
'
'    On Error GoTo adhGetProp_Err
'
'    adhGetProp = obj.Properties(strName)
'
'    Exit Function
'
'adhGetProp_Err:
'    If Err.Number <> adhcErrNotInCollection Then
'        MsgBox "Error: " & Err.Description & " (" & Err.Number & ")"
'    End If
'
'    adhGetProp = CVErr(Err)
'
'End Function


''' Set the value of a property. If it's not there, attemp to append it to the collection of properties.
''' Returns the previous value of the property, or an error value.
''' (use IsError() to check) if there was a problem.
''
''' In:
'''    obj: An object reference
'''       (db.TableDefs("tblCustomers"), for example)
'''    strName: Name for the property to set
'''    varValue: value for the property
''' Out:
'''    Return value:
'''      If an error occurred, an error value (use IsError() to check)
'''      If not, the old value of the property.
Private Function adhSetProp(ByVal obj As Object, ByRef strName As String, ByRef varValue As Variant) As Variant
    
    ' From Access 2000 Developer's Handbook
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    Dim varOldValue As Variant
    Dim varRetval As Variant
    
    On Error GoTo adhSetProp_Err

    varOldValue = obj.Properties(strName)
    obj.Properties(strName) = varValue
    adhSetProp = varOldValue
    
adhSetProp_Err:
    Select Case Err.Number
    Case adhcErrNotInCollection
        varRetval = AddProp(obj, strName, varValue)
        If IsError(varRetval) Then
            adhSetProp = varRetval
        Else
            Resume Next
        End If
    Case Else
        adhSetProp = CVErr(Err.Number)
    End Select

End Function


''' Attempt to add a property to obj.
''' If this succeeds, it returns True. If not, it returns an error, which the caller should
''' check for with IsError().
Private Function AddProp(ByVal obj As Object, ByRef strName As String, ByRef varValue As Variant) As Variant
    
    ' From Access 2000 Developer's Handbook
    ' by Litwin, Getz, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    Dim varRetval As Variant
    Dim prp As DAO.Property
    
    On Error GoTo AddProp_Err
    
    varRetval = adhGetDAOType(varValue)
    If IsError(varRetval) Then
        ' Bubble the error on up a level.
        ' Calling the Error subroutine triggers
        ' the error you request, and passes that
        ' on back up one level here.
        Err.Raise CInt(varRetval)
    Else
        Set prp = obj.CreateProperty(strName, varRetval, varValue)
        obj.Properties.Append prp
        AddProp = True
    End If
    
AddProp_Exit:
    Exit Function
    
AddProp_Err:
    AddProp = CVErr(Err.Number)
    Resume AddProp_Exit

End Function


''' Return the DAO type corresponding to the VarType
''' of the variant value passed in. If there's no
''' correspondence, return the user-defined error
''' adhcErrInvalidType.  Use IsErr() to check that out
''' in the caller.
Private Function adhGetDAOType(ByRef varValue As Variant) As Variant
    
    ' From Access 2000 Developer's Handbook
    ' by Litwin, Getz, Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    Dim intType As Long
    
    On Error GoTo GetDAOType_Err
    
    Select Case VarType(varValue)
    Case vbEmpty, vbNull, vbString
        ' If they sent a null or empty value,
        ' who knows WHAT they want?  Use a string.
        intType = dbText
    Case vbInteger, vbError
        intType = dbInteger
    Case vbLong
        intType = dbLong
    Case vbSingle
        intType = dbSingle
    Case vbDouble
        intType = dbDouble
    Case vbCurrency
        intType = dbCurrency
    Case vbDate
        intType = dbDate
    Case vbBoolean
        intType = dbBoolean
    Case vbObject, vbDataObject
        intType = dbLongBinary
    Case vbByte
        intType = dbByte
    Case Is >= vbArray
        ' No way to store arrays! Trigger a runtime error.
        Err.Raise adhcErrInvalidType
    End Select
    adhGetDAOType = intType
    
GetDAOType_Exit:
    Exit Function
    
GetDAOType_Err:
    adhGetDAOType = CVErr(Err.Number)
    Resume GetDAOType_Exit

End Function


''' Create a new primary key and its index for a table. If the table already has a primary key, remove it.
''' In:
'''     strTableName: name of the table with which to work
'''     strKeyName: name of the index to create
'''     varFields: one or more fields passed as a list of strings to add to the collection of fields in the index.
'''     db Database referenced elsewhere
''' Out:
'''     Return value: True on success, False otherwise.
Private Function adhCreatePrimaryKey(ByRef strTableName As String, ByVal strKeyName As String, ByVal strFields As String) As Boolean
    
    ' From Access 2000 Developer's Handbook by Litwin, Getz, and Gilbert (Sybex) Copyright 1999.  All rights reserved.
    Dim idx As DAO.Index
    Dim tdf As DAO.TableDef
    Dim varPK As Variant
    Dim idxs As DAO.Indexes

    On Error GoTo CreatePrimaryKey_Err
    
    Set tdf = backendDB.TableDefs.Item(strTableName)
    Set idxs = tdf.Indexes

    ' Find out if the table currently has a primary key.
    ' If so, delete it now.
    varPK = FindPrimaryKey(tdf)
    If Not IsNull(varPK) Then idxs.Delete varPK

    ' Create the new index object.
    Set idx = tdf.CreateIndex(strKeyName)
    
    ' Set the new index up as the primary key.
    ' This will also set:
    '   IgnoreNulls property to False,
    '   Required property to True,
    '   Unique property to True.
    idx.Primary = True

    ' Now create the fields that make up the index, and append
    ' each to the collection of fields.
    Dim theFields() As String
    theFields = Split(strFields, ",")
    Dim I As Long
    For I = 0 To UBound(theFields)
        AddField idx, theFields(I)
    Next I
    
    ' Now append the index to the TableDef's index collection
    idxs.Append idx
    adhCreatePrimaryKey = True

    Exit Function

CreatePrimaryKey_Err:
    MsgBox "Error: " & Err.Description & _
           " (" & Err.Number & ")"
    adhCreatePrimaryKey = False

End Function


''' Given a particular tabledef, find the primary key name, if it exists.
''' Return the name of the primary key's index, if it exists, or Null if there wasn't a primary key.
Private Function FindPrimaryKey(ByVal tdf As DAO.TableDef) As Variant

    ' From Access 2000 Developer's Handbook by Litwin, Getz, and Gilbert (Sybex) Copyright 1999.  All rights reserved.
    Dim idx As DAO.Index

    For Each idx In tdf.Indexes
        If idx.Primary Then
            FindPrimaryKey = idx.Name
            Exit Function
        End If
    Next idx
    
    FindPrimaryKey = Null

End Function


''' Given an index object, and a field name, add the field to the index.
''' Return True on success, False otherwise.
Private Function AddField(ByVal idx As DAO.Index, ByVal varIdx As Variant) As Boolean
    
    ' From Access 2000 Developer's Handbook by Litwin, Getz, Gilbert (Sybex) Copyright 1999.  All rights reserved.
    Dim fld As DAO.Field
    
    On Error GoTo AddIndex_Err
    If Len(varIdx & vbNullString) > 0 Then
        Set fld = idx.CreateField(varIdx)
        idx.Fields.Append fld
    End If
    
    AddField = True
    
AddIndex_Exit:
    Exit Function
    
AddIndex_Err:
    AddField = False
    Resume AddIndex_Exit

End Function


'''Purpose:   Set a property for an object, creating if necessary. (Supplied by Allen Browne)
'''Arguments: obj = the object whose property should be set.
'''           strPropertyName = the name of the property to set.
'''           intType = the type of property (needed for creating)
'''           varValue = the value to set this property to.
'''           strErrMsg = string to append any error message to.
Private Function SetPropertyDAO(ByVal obj As Object, ByRef strPropertyName As String, ByRef intType As Long, ByRef varValue As Variant, Optional ByRef strErrMsg As String) As Boolean

    On Error GoTo ErrHandler

    If HasProperty(obj, strPropertyName) Then
        obj.Properties(strPropertyName) = varValue
    Else
        obj.Properties.Append obj.CreateProperty(strPropertyName, intType, varValue)
    End If
    
    SetPropertyDAO = True

    Exit Function

ErrHandler:
    strErrMsg = strErrMsg & obj.Name & "." & strPropertyName & _
                " not set to " & varValue & ". Error " & Err.Number & " - " & _
                Err.Description & vbCrLf

End Function


'Purpose:   Return true if the object has the property. (Supplied by Allen Browne)
Private Function HasProperty(ByVal obj As Object, ByRef strPropName As String) As Boolean

    On Error Resume Next
    
    Debug.Print obj.Properties(strPropName)
    HasProperty = (Err.Number = 0)
    
    On Error GoTo 0

End Function


'''Add the specified reference table to the back-end file
'''Entry  (vRefTable) = name of table to add
'''Exit   (AddReferenceTable) = True if add table was successful
Private Function AddReferenceTable(ByRef vRefTable As String) As Boolean
    
    On Error GoTo ErrorCode
    
    Dim bePathname As String
    bePathname = GetBEDBPath
    Dim localDB As DAO.Database
    Set localDB = OpenDatabase(bePathname)
    With localDB
        .Execute "CREATE TABLE [" & vRefTable & "] (Setting CHAR, DataValue CHAR)" 'create table Setting and Value
        .TableDefs.Refresh                         'refresh table collection
        .Close
    End With
    
    Set localDB = Nothing
    If TableExists(vRefTable) = False Then       'if vRefTable not exists then
        DoCmd.TransferDatabase acLink, "Microsoft Access", bePathname, acTable, vRefTable, vRefTable 'and re-link to new table in BE
    End If
    
    'thisDatabase.Execute "INSERT INTO [" & vRefTable & "] (ubeVersion) VALUES (0)"                 'add one record, set ubeVersion = 0
    AddReferenceTable = True                     'update succeeded
    
ErrorCode:

End Function