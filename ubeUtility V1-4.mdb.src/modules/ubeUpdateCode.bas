Option Compare Database
Option Explicit

Public Const gRefTable = "tblGeneral"               'defines name of single record table in back end with 'ubeVersion' field

Const adhcErrObjectExists = 3012                    'see adhGetProp, etc procedures
Const adhcErrNotInCollection = 3270
Const adhcErrInvalidType = 30001

Dim db As DAO.Database                              'db sets reference for back-end database file
Dim tdf As DAO.TableDef
Dim dbLocal As DAO.Database
Dim fld As DAO.Field

Public Function UpdateBackEndFile(vDeveloper As Boolean) As Boolean

'Update selected back-end file with required changes
'Entry  (vDeveloper) = TRUE if called from Update form, = FALSE if called from user start up form
'       (gRefTable) = Name of table in back-end file which holds the 'ubeVersion' reference field
'Exit   (UpdateBackEndFile) = False if error or True if OK

Dim rst As DAO.Recordset
Dim vVersion As Long
Dim vPathname As String, vTableName As String
Dim vID As Variant

    On Error GoTo ErrorCode

    vID = "Unknown"                                                                                                     'error = 'Unknown'
ResumeError:
    vPathname = CurrentDb.TableDefs(gRefTable).Connect                                                                  'fetch connect def
    vPathname = Right(vPathname, Len(vPathname) - 10)                                                                   'and remove ';DATABASE=' string
    vVersion = Nz(DLookup("ubeVersion", gRefTable))                                                                     'fetch last Version number

    If vVersion < DMax("ID", "ubeUpdate") And vDeveloper = False Then                                                   'if User mode and updates available then
        If MsgBox("WARNING. Your back-end file is about to be updated with new tables and/or fields. You should first make a " _
        & "back-up copy of your data file, click Yes to continue or No to quit database program.", vbQuestion + vbYesNo, "Update Warning") = vbNo Then
            Application.Quit                                                                                            'quit application if user chooses No
        End If
    End If

    Set db = OpenDatabase(vPathname)                                                                                    'link to back-end file using db
    Set rst = CurrentDb.OpenRecordset("SELECT * FROM ubeUpdate WHERE ID > " & vVersion & " ORDER BY ID")                'make list of updates required
    Do Until rst.EOF                                                                                                    'step thru list
        DoCmd.OpenForm "ubeUpdating"                                                                                    'show Updating Back End message
        DoEvents
        vID = rst!ID                                                                                                    'fetch ID for error message
        Select Case rst!Action                                                                                          'select Action type
            Case "Make Table"
                db.Execute "CREATE TABLE [" & rst!TableName & "] ([" & rst!FieldName & "] " & rst!FieldType & ")"       'create table with one field
                db.TableDefs.Refresh                                                                                    'refresh table collection
                If TableExists(rst!TableName) = True Then DoCmd.DeleteObject acTable, rst!TableName                     'if link exists then delete Link
                DoCmd.TransferDatabase acLink, "Microsoft Access", vPathname, acTable, rst!TableName, rst!TableName     'and re-link to new table in BE
                NewFieldDefaults rst!TableName, rst!FieldName, rst!FieldType, Nz(rst!Description)                       'always set some properties
                SetProperties rst!TableName, rst!FieldName, Nz(rst!Constraint), Nz(rst!Misc), Nz(rst!Description)       'set field property (if any)

            Case "Copy Table"
                vTableName = rst!TableName                                                                              'fetch table name
                If TableExists("ube" & vTableName) = False Then                                                         'if ube'Table' not exists then
                    DoCmd.CopyObject vPathname, , acTable, vTableName                                                   'copy table to back-end file
                    DoCmd.Rename "ube" & vTableName, acTable, vTableName                                                'prefix table name with 'ube'
                    DoCmd.TransferDatabase acLink, "Microsoft Access", vPathname, acTable, vTableName, vTableName       'set link to new table
                Else                                                                                                    'if ube'Table' exists then
                    DoCmd.CopyObject vPathname, vTableName, acTable, "ube" & vTableName                                 'copy table to back-end and rename
                End If
                If TableExists(vTableName) = True Then DoCmd.DeleteObject acTable, vTableName                           'if link exists then delete Link
                DoCmd.TransferDatabase acLink, "Microsoft Access", vPathname, acTable, vTableName, vTableName           'and re-link to new table in BE
                db.TableDefs.Refresh                                                                                    'refresh table collection

            Case "Remove Table"
                db.Execute "DROP TABLE [" & rst!TableName & "]"                                                         'delete table from back-end file
                If TableExists("ube" & rst!TableName) = True Then                                                       'if ube'Table' exists then
                    CurrentDb.Execute "DROP TABLE [" & "ube" & rst!TableName & "]"                                      'delete 'ube' table also
                End If
                DoCmd.DeleteObject acTable, rst!TableName                                                               'and delete table Link
                db.TableDefs.Refresh                                                                                    'refresh table collection

            Case "New Field"
                If rst!FieldType = "ATTACHMENT" Then                                                                    'if field type = 'ATTACHMENT'
                    Set tdf = db.TableDefs(rst!TableName)                                                               'set ref to specified table
                    tdf.Fields.Append tdf.CreateField(rst!FieldName, 101)                                               'add Attachment type field (101 = dbAttachment)
                    
                    Set dbLocal = CurrentDb()                                                                           'refresh links to BE (due to bug in A2007)
                    For Each tdf In dbLocal.TableDefs                                                                   'loop through all tables
                        If tdf.Name = rst!TableName Then                                                                'skip if not current table
                            tdf.Connect = ";DATABASE=" & vPathname                                                      'set pathname + filename of back-end
                            tdf.RefreshLink                                                                             'and make link to back end
                        End If
                    Next tdf
                    Set tdf = Nothing
                Else
                    If rst!FieldType = "HYPERLINK" Then                                                                 'if field type = 'HYPERLINK'
                        Set tdf = db.TableDefs(rst!TableName)                                                           'set ref to curretn table
                        Set fld = tdf.CreateField(rst!FieldName, dbMemo)                                                'add Memo field first
                        fld.Attributes = dbHyperlinkField                                                               'set attribute to Hyperlink
                        tdf.Fields.Append fld                                                                           'and append field to table
                        tdf.Fields.Refresh
                        Set tdf = Nothing
                    Else
                        db.Execute "ALTER TABLE [" & rst!TableName & "] ADD [" & rst!FieldName & "] " & rst!FieldType   'add new field to table
                    End If
                End If
                db.TableDefs.Refresh                                                                                    'refresh table collection
                NewFieldDefaults rst!TableName, rst!FieldName, rst!FieldType, Nz(rst!Description)                       'always set some properties
                SetProperties rst!TableName, rst!FieldName, Nz(rst!Constraint), Nz(rst!Misc), Nz(rst!Description)       'set other field property (if any)

            Case "Delete Field"
                db.Execute "ALTER TABLE [" & rst!TableName & "] DROP [" & rst!FieldName & "]"                           'delete field
                db.TableDefs.Refresh                                                                                    'refresh table collection

            Case "Change Type"
                db.Execute "ALTER TABLE [" & rst!TableName & "] ALTER [" & rst!FieldName & "] " & rst!FieldType         'change field type
                db.TableDefs.Refresh                                                                                    'refresh table collection
                SetProperties rst!TableName, rst!FieldName, Nz(rst!Constraint), Nz(rst!Misc), Nz(rst!Description)       'set field property

            Case "Set Property"
                db.TableDefs.Refresh                                                                                    'refresh table collection
                SetProperties rst!TableName, rst!FieldName, Nz(rst!Constraint), Nz(rst!Misc), Nz(rst!Description)       'set field property

            Case "Set Relationship"
                CreateRelationship rst!TableName, rst!FieldName, rst!Constraint, rst!Misc, rst!Description              'create a Relationship

            Case "Clear Relationship"
                DeleteRelationship rst!TableName, rst!FieldName, rst!Misc, rst!Description                              'delete a Relationship

            Case "Run Query"
                CurrentDb.Execute rst!TableName                                                                         'execute UPDATE query

            Case "Execute Code"
                Run rst!TableName                                                                                       'execute VBA Code

            Case "Run Macro"
                DoCmd.RunMacro rst!TableName                                                                            'execute Macro
        
        End Select
        CurrentDb.Execute "UPDATE [" & gRefTable & "] SET ubeVersion = " & rst!ID                                       'copy current ID to ubeVersion field
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
    db.Close
    Set db = Nothing
    UpdateBackEndFile = True                                                                                            'return Updates OK code
    DoCmd.Close acForm, "ubeUpdating"                                                                                   'close message form
    Exit Function

ErrorCode:
    DoCmd.Close acForm, "ubeUpdating"                                                                                   'close message form
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    
    If Err = 3265 Or Err = 3078 Then                                                                                    'if table does not exist then
        If AddReferenceTable(gRefTable) = False Then Exit Function                                                      'allow user to create one or abort if returns False
        GoTo ResumeError                                                                                                'continue with other updates (if any)
    End If

    If vDeveloper = True Then                                                                                           'if developer mode then
        MsgBox Err.Description & "  (Reference No = " & vID & ")"                                                       'show error + Ref No and exit with False
    Else
        MsgBox "ERROR. Back-End Update Failed. (Reference No = " & vID & "). Application will now be shut down to prevent any damage to back-end data file.", vbOKOnly, "Update Fail"
        Application.Quit                                                                                                'if update error on user DB then quit
    End If

End Function

Public Function TableExists(vTableName As String) As Boolean

'Checks if table exists
'Entry  (vTableName) = Name of table
'Exit   (TableExists) = True if table exists, = False if not

Dim vName As String

    On Error GoTo ErrorCode                                     'trap error if next line fails
    
    vName = CurrentDb.TableDefs(vTableName).Name                'try to read table name from TableDefs
    TableExists = True                                          'TableExists = True if successful
    Exit Function

ErrorCode:
    TableExists = False                                         'TableExists = False if not successful

End Function

Public Sub SetProperties(vTableName As String, vFieldName As String, vPropertyType As String, vParameters As String, vDescription As String)

'Change or add a field property
'Entry  (vTableName) = name of table to change
'       (vFieldName) = name of field to change
'       (vPropertyType) = name of field property to be changed (if NULL then just change Description property, if any)
'       (vParameters) = any required parameters (i.e. Field default value, New field name or Ordinal position, etc)
'       (vDescription) = text for description column of specified field or other data
'Exit   Specified property changed
'       Any errors handled by main UpdateBackEndFile routine

Dim vStatus As Boolean
Dim vR As Variant
Dim vRTF As Long

    Select Case vPropertyType
        Case "Text Field Size ="
            If db.TableDefs(vTableName).Fields(vFieldName).Type = dbText Then                                               'if Field type = TEXT then
                db.Execute "ALTER TABLE [" & vTableName & "] ALTER COLUMN [" & vFieldName & "] TEXT (" & vParameters & ")"  'change field size
            End If
        Case "Set Compression"
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "UniCodeCompression", True)                        'set UniCode Compression
        Case "Required ="
            If vParameters = "Yes" Or vParameters = "True" Then vStatus = True                                              'convert Yes/True to Boolean
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "Required", vStatus)                               'set set Required Yes/No
        Case "Allow Zero Len ="
            If vParameters = "Yes" Or vParameters = "True" Then vStatus = True                                              'convert Yes/True to Boolean
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "AllowZeroLength", vStatus)                        'set Allow Zero Length
        Case "Default Value ="
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "DefaultValue", vParameters)                       'set Default Value
        Case "Input Mask ="
            vR = adhSetProp(CurrentDb.TableDefs(vTableName).Fields(vFieldName), "InputMask", vParameters)                   'set Input Mask Value
        Case "Format ="
            vR = adhSetProp(CurrentDb.TableDefs(vTableName).Fields(vFieldName), "Format", vParameters)                      'set Format type
        Case "Decimal Places ="
            vR = adhSetProp(CurrentDb.TableDefs(vTableName).Fields(vFieldName), "DecimalPlaces", vParameters)               'set Decimal Places
        Case "Validation Rule ="
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "ValidationRule", vParameters)                     'set Validation Rule
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "ValidationText", vDescription)                    'set Validation Text
        Case "New Field Name ="
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "Name", vParameters)                               'set/change Field Name
        Case "Description ="
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "Description", vDescription)                       'set Description field
        Case "Ordinal Position ="
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "OrdinalPosition", vParameters)                    'set Field Ordinal posn
        Case "Set Primary Key ="
            If vParameters <> "" Then vFieldName = vFieldName & "," & vParameters                                           'add in extra fields (if any)
            vR = adhCreatePrimaryKey(vTableName, "PrimaryKey", vFieldName)                                                  'set Primary Key/s
        Case "Indexed (No)"
            vR = FindIndex(vTableName, vFieldName)                                                                          'find Index name
            If vR <> "" Then db.Execute "DROP INDEX [" & vR & "] ON [" & vTableName & "]"                                   'remove Index (if any)
        Case "Indexed (Dup OK)"
            vR = FindIndex(vTableName, vFieldName)                                                                          'find Index name
            If vR <> "" Then db.Execute "DROP INDEX [" & vR & "] ON [" & vTableName & "]"                                   'remove Index (if any)
            db.Execute "CREATE INDEX [" & vFieldName & "] ON [" & vTableName & "] ([" & vFieldName & "])"                   'set Index (use Field name as Index name)
        Case "Indexed (No Dup)"
            vR = FindIndex(vTableName, vFieldName)                                                                          'find Index name
            If vR <> "" Then db.Execute "DROP INDEX [" & vR & "] ON [" & vTableName & "]"                                   'remove Index (if any)
            db.Execute "CREATE UNIQUE INDEX [" & vFieldName & "] ON [" & vTableName & "] ([" & vFieldName & "])"            'set unique Index (use Field name as Index name)
        Case "Fill With ="
            FillField vTableName, vFieldName, vParameters                                                                   'copy data to field
        Case "Rich Text ="
            If vParameters = "Yes" Or vParameters = "True" Then vRTF = 1 Else vRTF = 0                                      'convert Yes/True to 1/0
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "TextFormat", vRTF)                                'set/clear Rich Text format
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "Description", vDescription)                       'set Description field
        Case "Caption Name ="
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "Caption", vParameters)                            'set Caption property
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "Description", vDescription)                       'set Description field
        Case "Smart Tags ="
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "SmartTags", vParameters)                          'set Smart Tags property
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "Description", vDescription)                       'set Description field
    End Select

End Sub

Public Sub FillField(vTableName As String, vFieldName As String, vParameter As String)

'Fill a field with data for all records
'Entry  (vTableName) = name of table to fill
'       (vFieldName) = name of field to fill
'       (vParameter) = data to be copied to table
'       (db) = Database object referenced to back-end file
'Exit   Specified field in all records filled with specified value (Note. In Text/Memo fields any double quotes replaced with two single quotes)
'       Any errors handled by main UpdateBackEndFile routine

Dim vFieldType As Long
Dim vData As String
Const QUOTE = """"                                                                              'Used in place of Double Quotes

    vFieldType = db.TableDefs(vTableName).Fields(vFieldName).Type                               'fetch field type
    Select Case vFieldType
        Case dbText, dbMemo                                                                     'if Text or Memo then
            vParameter = Replace(Nz(vParameter), """", "''")                                    'replace any Double Quotes with two Single Quotes
            vData = QUOTE & vParameter & QUOTE                                                  'surround with double quotes
            db.Execute "UPDATE [" & vTableName & "] SET [" & vFieldName & "] = " & vData & " WHERE [" & vFieldName & "] IS NULL" 'fill field for all blank records
        Case dbDate                                                                             'if date or time
            vData = "#" & Format(vParameter, "mm\/dd\/yyyy") & "#"                              'reformat date for US mode
            db.Execute "UPDATE [" & vTableName & "] SET [" & vFieldName & "] = " & vData & " WHERE [" & vFieldName & "] IS NULL" 'fill field for all blank records
        Case dbBoolean
            db.Execute "UPDATE [" & vTableName & "] SET [" & vFieldName & "] = " & vParameter   'set field to True or False
        Case Else
            vData = vParameter                                                                  'numeric values need no changes
            db.Execute "UPDATE [" & vTableName & "] SET [" & vFieldName & "] = " & vData & " WHERE [" & vFieldName & "] IS NULL" 'fill field for all blank records
    End Select

End Sub

Public Sub NewFieldDefaults(vTableName As String, vFieldName As String, vFieldType As String, vDescription As String)

'Set some properties for new fields regardless
'Entry  (vTableName) = name of table to change
'       (vFieldName) = name of field to change
'       (vFieldType) = field property type
'       (vDescription) = text for description column of specified field
'       (db) = Database object referenced to back-end file
'Exit   Specified field properties set (delete any you don't want)
'       Any errors handled by main UpdateBackEndFile routine

Dim vR As Variant
Dim fld As DAO.Field

    vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "Description", vDescription)           'set Description field (if any)
    Select Case vFieldType
        Case "BYTE", "SHORT", "LONG", "SINGLE", "DOUBLE", "CURRENCY"                                    'select Number fields only
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "DefaultValue", 0)             'set Number fields Default Value to 0
        Case "TEXT", "MEMO", "HYPERLINK"                                                                'select TEXT & MEMO & HYPERLINK fields
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "UniCodeCompression", True)    'always set UniCode Compression on Text/Memo
            vR = adhSetProp(db.TableDefs(vTableName).Fields(vFieldName), "AllowZeroLength", True)       'always set AllowZeroLength on Text/Memo
        Case "YESNO"                                                                                    'select Yes/No fields
            Set fld = db.TableDefs(vTableName).Fields(vFieldName)                                       'change YesNo field format
            Call SetPropertyDAO(fld, "DisplayControl", dbInteger, CInt(acCheckBox))                     'to Check Box type
    End Select

End Sub

Public Function FindIndex(vTableName As String, vFieldName As String) As String

'Set some properties for new fields regardless
'Entry  (vTableName) = name of table with indexed field
'       (vFieldName) = name of indexed field
'       (db) = Database object referenced to back-end file
'Exit   Index name for selected table/field returned or "" if none
'       Any errors handled by main UpdateBackEndFile routine

Dim idx As DAO.Index
Dim tdf As DAO.TableDef

    Set tdf = db.TableDefs(vTableName)                                  'define required table
    For Each idx In tdf.Indexes                                         'search Indexes
        If InStr(1, idx.Fields, "+" & vFieldName) > 0 Then              'if index field holds ("+" & field name) then
            FindIndex = idx.Name                                        'fetch Index name and
            Exit Function                                               'return with Index name
        End If
    Next idx

End Function

Function CreateRelationship(vPKTableName As String, vPKFieldName As String, vRelationshipType As String, vFKTableName As String, vFKFieldName As String) As Boolean

' From Access 2000 Developer's Handbook by Litwin, Getz, Gilbert (Sybex) Copyright 1999.  All rights reserved.

'Create or change a relationship between two tables
'Entry  (vPKTableName) = Name of table for Primary Key
'       (vPKFieldName) = Name of Primary Key in primary table
'       (vFKTableName) = Name of table for Foreign Key
'       (vFKFieldName) = Name of Foreign Key field in Foreign table
'Exit   (CreateRelationship) = True if Relationship created or = False if error

Dim rel As DAO.Relation
Dim fld As DAO.Field
Dim vRelType As Long
Dim vRelationshipName As String

    On Error GoTo CreateRelationship_Err

    Select Case vRelationshipType                                                           'convert Relationship type to Long Integer
        Case "1-1 Not Enforced"                                                             'One to One (Not Enforced)
            vRelType = 1
        Case "1-1 Casc Updates"                                                             'One to One (Cascade Updates)
            vRelType = 1 + 256
        Case "1-1 Casc Deletes"                                                             'One to One (Cascade Deletes)
            vRelType = 1 + 4096
        Case "1-1 Casc Upd/Del"                                                             'One to One (Cascade Updates and Deletes)
            vRelType = 1 + 4096 + 256
        Case "1-n Not Enforced"                                                             'One to Many (Not Enforced)
            vRelType = 2
        Case "1-n Casc Updates"                                                             'One to Many (Cascade Updates)
            vRelType = 256
        Case "1-n Casc Deletes"                                                             'One to Many (Cascade Deletes)
            vRelType = 4096
        Case "1-n Casc Upd/Del"                                                             'One to Many (Cascade Updates and Deletes)
            vRelType = 4096 + 256
    End Select

    vRelationshipName = vPKTableName & vFKTableName                                         'create a relationship name from both tables
    Set rel = db.CreateRelation(vRelationshipName, vPKTableName, vFKTableName, vRelType)    'create relationship link
    Set fld = rel.CreateField(vPKFieldName)                                                 'Set the relation's field collection.
    fld.ForeignName = vFKFieldName                                                          'set Foreign table field
    rel.Fields.Append fld                                                                   'append foreign field
    db.Relations.Append rel                                                                 'append relationship definition to collection
    CreateRelationship = True                                                               'return True
    
CreateRelationship_Exit:
    Exit Function

CreateRelationship_Err:
    Select Case Err.Number
        Case adhcErrObjectExists                        'If the relationship already exists, just delete it, and then try to append it again.
            db.Relations.Delete rel.Name
            Resume
        Case Else
            MsgBox "Error: " & Err.Description & _
             " (" & Err.Number & ")"
            CreateRelationship = False
            Resume CreateRelationship_Exit
    End Select

End Function

Public Sub DeleteRelationship(vPKTableName As String, vPKFieldName As String, vFKTableName As String, vFKFieldName As String)

'Delete a relationship between two tables
'Entry  (vPKTableName) = Name of table for Primary Key
'       (vPKFieldName) = Name of Primary Key in primary table
'       (vFKTableName) = Name of table for Foreign Key
'       (vFKFieldName) = Name of Foreign Key field in Foreign table
'       (db) = Database object referenced to back-end file
'Exit   Relationship deleted
'       Any errors handled by main UpdateBackEndFile routine

Dim rst As DAO.Recordset

    Set rst = db.OpenRecordset("SELECT szRelationship FROM MSysRelationships " _
    & "WHERE szReferencedObject = '" & vPKTableName & "' " _
    & "AND szReferencedColumn = '" & vPKFieldName & "' " _
    & "AND szObject = '" & vFKTableName & "' " _
    & "AND szColumn = '" & vFKFieldName & "' " _
    & "OR szReferencedObject = '" & vFKTableName & "' " _
    & "AND szReferencedColumn = '" & vFKFieldName & "' " _
    & "AND szObject = '" & vPKTableName & "' " _
    & "AND szColumn = '" & vPKFieldName & "'")                                      'fetch Relationship name
    If Not rst.BOF And Not rst.EOF Then                                             'if relationship exists then
        db.Relations.Delete rst!szRelationship                                      'delete relationship (if any)
    End If
    rst.Close
    Set rst = Nothing

End Sub

Public Function adhGetProp(obj As Object, strName As String) As Variant

'***** This function is not actually used but has been left in just in case it is needed !! ******

' From Access 2000  Developer's Handbook
' by Getz, Litwin, and Gilbert. (Sybex)
' Copyright 1999. All Rights Reserved.
' Get the value of a property
' If there isn't a property of the name passed in,
' return an error value, otherwise, return the
' value of the property.

' In:
'    obj: An object reference
'       (db.TableDefs("tblCustomers"), for example)
'    strName: Name for the property to get
' Out:
'    Return value:
'      If an error occurred, an error value (use IsError() to check)
'      If not, the value of the property.
    
    On Error GoTo adhGetProp_Err
    
    adhGetProp = obj.Properties(strName)
    
adhGetProp_Exit:
    Exit Function
    
adhGetProp_Err:
    Select Case Err.Number
       Case adhcErrNotInCollection
          ' Do nothing
       Case Else
          MsgBox "Error: " & Err.Description & _
           " (" & Err.Number & ")"
    End Select
    adhGetProp = CVErr(Err)
    Resume adhGetProp_Exit

End Function

Public Function adhSetProp(obj As Object, strName As String, varValue As Variant) As Variant
    
' From Access 2000 Developer's Handbook
' by Getz, Litwin, and Gilbert (Sybex)
' Copyright 1999.  All rights reserved.

' Set the value of a property.
' If it's not there, attemp to append it to the
' collection of properties.
' Guess on the data type based on the value passed in.
' Returns the previous value of the property, or an error value
' (use IsError() to check) if there was a problem.

' In:
'    obj: An object reference
'       (db.TableDefs("tblCustomers"), for example)
'    strName: Name for the property to set
'    varValue: value for the property
' Out:
'    Return value:
'      If an error occurred, an error value (use IsError() to check)
'      If not, the old value of the property.
          
Dim varOldValue As Variant
Dim prp As DAO.Property
Dim varRetval As Variant
    
    On Error GoTo adhSetProp_Err

    varOldValue = obj.Properties(strName)
    obj.Properties(strName) = varValue
    adhSetProp = varOldValue
    
adhSetPropExit:
    Exit Function
    
adhSetProp_Err:
    Select Case Err.Number
        Case adhcErrNotInCollection
            varRetval = AddProp(obj, strName, varValue)
            If IsError(varRetval) Then
                adhSetProp = varRetval
                Resume adhSetPropExit
            Else
                Resume Next
            End If
        Case Else
            adhSetProp = CVErr(Err.Number)
            Resume adhSetPropExit
    End Select

End Function

Private Function AddProp(obj As Object, strName As String, varValue As Variant) As Variant
    
' From Access 2000 Developer's Handbook
' by Litwin, Getz, and Gilbert (Sybex)
' Copyright 1999.  All rights reserved.

' Attempt to add a property to obj.
' If this succeeds, it returns True.
' If not, it returns an error, which the caller should
' check for with IsError().

Dim varRetval As Variant
Dim prp As DAO.Property
    
    On Error GoTo AddProp_Err
    
    varRetval = adhGetDAOType(varValue)
    If IsError(varRetval) Then
        ' Bubble the error on up a level.
        ' Calling the Error subroutine triggers
        ' the error you request, and passes that
        ' on back up one level here.
        Error CInt(varRetval)
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

Private Function adhGetDAOType(varValue As Variant) As Variant
    
' From Access 2000 Developer's Handbook
' by Litwin, Getz, Gilbert (Sybex)
' Copyright 1999.  All rights reserved.

' Return the DAO type corresponding to the VarType
' of the variant value passed in. If there's no
' correspondence, return the user-defined error
' adhcErrInvalidType.  Use IsErr() to check that out
' in the caller.

Dim intType As Integer
    
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
            Error adhcErrInvalidType
    End Select
    adhGetDAOType = intType
    
GetDAOType_Exit:
    Exit Function
    
GetDAOType_Err:
    adhGetDAOType = CVErr(Err.Number)
    Resume GetDAOType_Exit

End Function

Public Function adhCreatePrimaryKey(strTableName As String, strKeyName As String, strFields As String) As Boolean
    
' From Access 2000 Developer's Handbook by Litwin, Getz, and Gilbert (Sybex) Copyright 1999.  All rights reserved.

' Create a new primary key and its index for a table. If the table already has a primary key, remove it.
' In:
'     strTableName: name of the table with which to work
'     strKeyName: name of the index to create
'     varFields: one or more fields passed as a list of strings to add to the collection of fields in the index.
'     db Database referenced elsewhere
' Out:
'     Return value: True on success, False otherwise.

Dim idx As DAO.Index
Dim tdf As DAO.TableDef
Dim fld As DAO.Field
Dim varPK As Variant
Dim varIdx As Variant
Dim idxs As DAO.Indexes
'Dim db As DAO.Database                                 'NOT REQUIRED !!
Dim I As Integer

    On Error GoTo CreatePrimaryKey_Err
    
'    Set db = CurrentDb()                               'NOT REQUIRED !!
    Set tdf = db.TableDefs(strTableName)
    Set idxs = tdf.Indexes

    ' Find out if the table currently has a primary key.
    ' If so, delete it now.
    varPK = FindPrimaryKey(tdf)
    If Not IsNull(varPK) Then
        idxs.Delete varPK
    End If
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
    For I = 0 To UBound(Split(strFields, ","))
        AddField idx, Split(strFields, ",")(I)
    Next I
    
    ' Now append the index to the TableDef's index collection
    idxs.Append idx
    adhCreatePrimaryKey = True

CreatePrimaryKey_Exit:
    Exit Function

CreatePrimaryKey_Err:
    MsgBox "Error: " & Err.Description & _
     " (" & Err.Number & ")"
    adhCreatePrimaryKey = False
    Resume CreatePrimaryKey_Exit

End Function

Private Function FindPrimaryKey(tdf As DAO.TableDef) As Variant

' From Access 2000 Developer's Handbook by Litwin, Getz, and Gilbert (Sybex) Copyright 1999.  All rights reserved.

' Given a particular tabledef, find the primary key name, if it exists.
' Return the name of the primary key's index, if it exists, or Null if there wasn't a primary key.

Dim idx As DAO.Index

    For Each idx In tdf.Indexes
        If idx.Primary Then
            FindPrimaryKey = idx.Name
            Exit Function
        End If
    Next idx
    FindPrimaryKey = Null

End Function

Private Function AddField(idx As DAO.Index, varIdx As Variant) As Boolean
    
' From Access 2000 Developer's Handbook by Litwin, Getz, Gilbert (Sybex) Copyright 1999.  All rights reserved.

' Given an index object, and a field name, add the field to the index.
' Return True on success, False otherwise.

Dim fld As DAO.Field
    
    On Error GoTo AddIndex_Err
    If Len(varIdx & "") > 0 Then
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

Public Function SetPropertyDAO(obj As Object, strPropertyName As String, intType As Integer, varValue As Variant, Optional strErrMsg As String) As Boolean
    
'Purpose:   Set a property for an object, creating if necessary. (Supplied by Allen Browne)
'Arguments: obj = the object whose property should be set.
'           strPropertyName = the name of the property to set.
'           intType = the type of property (needed for creating)
'           varValue = the value to set this property to.
'           strErrMsg = string to append any error message to.
    
    On Error GoTo ErrHandler

    If HasProperty(obj, strPropertyName) Then
        obj.Properties(strPropertyName) = varValue
    Else
        obj.Properties.Append obj.CreateProperty(strPropertyName, intType, varValue)
    End If
    SetPropertyDAO = True
ExitHandler:
    Exit Function

ErrHandler:
    strErrMsg = strErrMsg & obj.Name & "." & strPropertyName & _
" not set to " & varValue & ". Error " & Err.Number & " - " & _
Err.Description & vbCrLf
    Resume ExitHandler

End Function

Public Function HasProperty(obj As Object, strPropName As String) As Boolean

'Purpose:   Return true if the object has the property. (Supplied by Allen Browne)

Dim varDummy As Variant

    On Error Resume Next
    
    varDummy = obj.Properties(strPropName)
    HasProperty = (Err.Number = 0)

End Function

Public Function AddReferenceTable(vRefTable As String) As Boolean

'Add the specified reference table to the back-end file
'Entry  (vRefTable) = name of table to add
'Exit   (AddReferenceTable) = True if user opted to add table and which was successful or False if user quit or update failed

Dim db As DAO.Database
Dim tdf As TableDef
Dim vPathname As String
Dim I As Integer
    
    On Error GoTo ErrorCode
    
    If MsgBox("ERROR. The Back-End Updater program cannot find table '" & vRefTable & "' in the back-end file," _
    & "do you want to create a new table now so that the back-end file can be updated with new tables " _
    & "and fields?", vbQuestion + vbYesNo, "Add Update Reference Table") = vbNo Then Exit Function
        
    Set db = CurrentDb()                                                                        'find any table in BE to get connection path
    For I = 0 To db.TableDefs.Count - 1                                                         'loop thru all tables
        Set tdf = db.TableDefs(I)                                                               'fetch table definition
        If (tdf.Attributes And dbSystemObject) = 0 Then                                         'skip system tables
            If Nz(tdf.Connect) <> "" Then                                                       'if table is linked then
                vPathname = tdf.Connect                                                         'fetch connection string
                Exit For                                                                        'exit loop (as one found)
            End If
        End If
    Next I
    vPathname = Right(vPathname, Len(vPathname) - 10)                                           'and remove ';DATABASE=' string
    Set db = OpenDatabase(vPathname)                                                            'link to back-end file using db
    db.Execute "CREATE TABLE [" & vRefTable & "] (ubeVersion LONG)"                             'create table with one field
    db.TableDefs.Refresh                                                                        'refresh table collection
    If TableExists(vRefTable) = False Then                                                          'if vRefTable not exists then
        DoCmd.TransferDatabase acLink, "Microsoft Access", vPathname, acTable, vRefTable, vRefTable 'and re-link to new table in BE
    End If
    CurrentDb.Execute "INSERT INTO [" & vRefTable & "] (ubeVersion) VALUES (0)"                 'add one record, set ubeVersion = 0
    AddReferenceTable = True                                                                    'update succeeded
    Exit Function
    
ErrorCode:
End Function