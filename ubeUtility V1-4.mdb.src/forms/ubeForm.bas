Version =19
VersionRequired =19
Begin Form
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    OrderByOn = NotDefault
    ScrollBars =2
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14193
    DatasheetFontHeight =10
    ItemSuffix =25
    Left =1455
    Top =1065
    Right =17220
    Bottom =10590
    DatasheetGridlinesColor =12632256
    OrderBy ="ID"
    RecSrcDt = Begin
        0xd16b7770d086e240
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="ubeUpdate"
    Caption ="Update Back-End File"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnGotFocus ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =645
            BackColor =6697728
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =2130
                    Top =405
                    Width =1605
                    Height =240
                    ForeColor =16777215
                    Name ="Label1"
                    Caption ="Table/Query Name"
                End
                Begin Label
                    OverlapFlags =85
                    Left =4572
                    Top =405
                    Width =855
                    Height =240
                    ForeColor =16777215
                    Name ="Label2"
                    Caption ="Field Name"
                End
                Begin Label
                    OverlapFlags =85
                    Left =6456
                    Top =405
                    Width =810
                    Height =240
                    ForeColor =16777215
                    Name ="Label3"
                    Caption ="Field Type"
                End
                Begin Label
                    OverlapFlags =85
                    Left =546
                    Top =405
                    Width =525
                    Height =240
                    ForeColor =16777215
                    Name ="Label4"
                    Caption ="Action"
                End
                Begin Label
                    OverlapFlags =85
                    Left =8151
                    Top =405
                    Width =705
                    Height =240
                    ForeColor =16777215
                    Name ="Label5"
                    Caption ="Property"
                End
                Begin Label
                    OverlapFlags =85
                    Left =9801
                    Top =405
                    Width =1170
                    Height =240
                    ForeColor =16777215
                    Name ="Label6"
                    Caption ="Additional Data"
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =405
                    Width =330
                    Height =240
                    ForeColor =16777215
                    Name ="Label10"
                    Caption ="Ref"
                End
                Begin Label
                    OverlapFlags =85
                    Left =45
                    Top =30
                    Width =2505
                    Height =315
                    FontSize =11
                    FontWeight =700
                    ForeColor =65535
                    Name ="Label14"
                    Caption ="Update Back-End File "
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =13281
                    Top =60
                    Width =800
                    FontWeight =700
                    ForeColor =255
                    Name ="txtLastRef"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =11820
                            Top =60
                            Width =1410
                            Height =240
                            ForeColor =16777215
                            Name ="Label18"
                            Caption ="Last Update Ref "
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    Left =7186
                    Top =60
                    Width =2750
                    TabIndex =1
                    ForeColor =65535
                    Name ="txtDate"
                    Format ="d mmmm yyyy"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5550
                            Top =60
                            Width =1590
                            Height =240
                            ForeColor =16777215
                            Name ="Label20"
                            Caption ="Date Item Created "
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =11721
                    Top =405
                    Width =1245
                    Height =240
                    ForeColor =16777215
                    Name ="Label22"
                    Caption ="Field Description"
                End
            End
        End
        Begin Section
            Height =240
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =9816
                    Width =1881
                    TabIndex =6
                    LeftMargin =29
                    Name ="Misc"
                    ControlSource ="Misc"
                    StatusBarText ="Other required information"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =2
                    ListRows =14
                    ListWidth =1871
                    Left =6444
                    Width =1671
                    TabIndex =4
                    Name ="FieldType"
                    ControlSource ="FieldType"
                    RowSourceType ="Value List"
                    RowSource ="AUTOINCREMENT;TEXT;DATETIME;BYTE;SHORT;LONG;SINGLE;DOUBLE;CURRENCY;YESNO;MEMO;OL"
                        "EOBJECT;HYPERLINK;ATTACHMENT"
                    StatusBarText ="Type of field"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    ListRows =12
                    ListWidth =1871
                    Left =516
                    Width =1581
                    TabIndex =1
                    Name ="Action"
                    ControlSource ="Action"
                    RowSourceType ="Value List"
                    RowSource ="Make Table;Copy Table;Remove Table;New Field;Delete Field;Change Type;Set Proper"
                        "ty;Set Relationship;Clear Relationship;Run Query;Execute Code;Run Macro"
                    StatusBarText ="Action to be taken (Make Table, Copy Table, Remove Table, Add Field, etc)"
                    ValidationRule ="Is Not Null"
                    ValidationText ="There must be some action code"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =2
                    ListRows =20
                    ListWidth =1985
                    Left =8115
                    TabIndex =5
                    Name ="Constraint"
                    ControlSource ="Constraint"
                    RowSourceType ="Value List"
                    StatusBarText ="Field property"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Width =516
                    ColumnWidth =495
                    BorderColor =4210752
                    Name ="ID"
                    ControlSource ="ID"
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =1
                    Left =11697
                    Width =2496
                    TabIndex =7
                    LeftMargin =29
                    Name ="Description"
                    ControlSource ="Description"
                    StatusBarText ="Description of field"
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =2
                    Left =13707
                    Width =426
                    TabIndex =8
                    ForeColor =255
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    StatusBarText ="Date update made"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =1
                    ListWidth =3402
                    Left =2097
                    Width =2436
                    TabIndex =2
                    Name ="TableName"
                    ControlSource ="TableName"
                    RowSourceType ="Value List"
                    StatusBarText ="Name of table, query, procedure or Macro to add, delete, alter"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =1
                    Left =4533
                    Width =1911
                    TabIndex =3
                    Name ="FieldName"
                    ControlSource ="FieldName"
                    RowSourceType ="Value List"
                    StatusBarText ="Name of field to add, delete, alter"
                End
            End
        End
        Begin FormFooter
            Height =510
            BackColor =6697728
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =12585
                    Top =90
                    Width =1511
                    Height =370
                    Name ="btnClose"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Close form"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10922
                    Top =90
                    Width =1511
                    Height =370
                    TabIndex =1
                    Name ="btnUpdate"
                    Caption ="Update Back End"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Update back-end file with new data"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9259
                    Top =90
                    Width =1511
                    Height =370
                    TabIndex =2
                    Name ="btnAddNew"
                    Caption ="Add New Item"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add new object or instruction to list"
                End
                Begin Label
                    OverlapFlags =85
                    Top =120
                    Width =3735
                    Height =240
                    ForeColor =16777215
                    Name ="Label15"
                    Caption ="Version 1.4     January 2018     by  Peter  D  Hibbs"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =4020
                    Top =105
                    Width =3060
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BackColor =32768
                    ForeColor =16777215
                    Name ="lblOK"
                    Caption ="All Updates Completed OK"
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Acknowledgements to the following experts for help with this project :-
'Getz, Litwin and Gilbert (for writing the Access 2000 Developers Handbook)
'Dirk Goldgar and Allen Browne for help with Relationships code

'Copy this line of code into the Open event of your Start Up form
'See Word documentation if using Access 2007 (.accdb mode)
'       Call UpdateBackEndFile(False)

Private Sub Action_AfterUpdate()

    On Error GoTo ErrorCode

    TableName = "": FieldName = "": Constraint = "": Misc = "": FieldType = "": Description = ""            'clear all fields first
    
    SetConstraintSource                                                                                     'change options in Constraint drop-down
    Select Case Action                                                                                      'select Action type and fill in reqd fields
        Case "Make Table"
            TableName = "(Table Name)": FieldName = "(Field Name)": FieldType = "(Field Type)": Description = "(Field Description - Optional)"
        Case "Copy Table"
            TableName = "(Table Name)"
            TableName.RowSource = FetchObjectList(1)
        Case "Remove Table"
            TableName = "(Table Name)"
            TableName.RowSource = FetchObjectList(2)
        Case "New Field"
            TableName = "(Table Name)": FieldName = "(Field Name)": FieldType = "(Field Type)": Description = "(Field Description - Optional)"
            TableName.RowSource = FetchObjectList(2)
        Case "Delete Field"
            TableName = "(Table Name)": FieldName = "(Field Name)"
            TableName.RowSource = FetchObjectList(2)
        Case "Change Type"
            TableName = "(Table Name)": FieldName = "(Field Name)": FieldType = "(Field Type)"
            TableName.RowSource = FetchObjectList(2)
        Case "Set Property"
            TableName = "(Table Name)": FieldName = "(Field Name)": Constraint = "(Property)"
            TableName.RowSource = FetchObjectList(2)
        Case "Set Relationship"
            TableName = "(PK Table Name)": FieldName = "(PK Field Name)": Constraint = "(Relationship Type)": Misc = "(FK Table Name)": Description = "(FK Field Name)"
            TableName.RowSource = FetchObjectList(2)
        Case "Clear Relationship"
            TableName = "(Table Name)": FieldName = "(Field Name)": Misc = "(Table Name)": Description = "(Field Name)"
            TableName.RowSource = FetchObjectList(2)
        Case "Run Query"
            TableName = "(Query Name)"
            TableName.RowSource = FetchObjectList(3)
        Case "Execute Code"
            TableName = "(Procedure Name)"
        Case "Run Macro"
            TableName = "(Macro Name)"
    End Select
    Exit Sub

ErrorCode:
    If Err = 2176 Then                                          'if overflow error then
        MsgBox "WARNING. There are too many table/query names to show in the drop down field, you must enter your table/query name manually.", vbExclamation + vbOKOnly, "List Overflow for Combo Box (Access 2000 Limitation)"
    End If
    Resume Next                                                 'continue

End Sub

Private Sub btnAddNew_Click()

    DoCmd.GoToRecord , , acNewRec                               'start new record and
    Action.SetFocus                                             'move cursor to Action field

End Sub

Private Sub btnClose_Click()
    DoCmd.Close
End Sub

Private Sub btnUpdate_Click()

    lblOK.Visible = False                                       'hide message label (if visible)
    RunCommand acCmdSaveRecord                                  'update ubeUpdate table
    If UpdateBackEndFile(True) = True Then                      'change table structure in back-end file, return True if OK
        lblOK.BackColor = 32768                                 'make label colour Green
        lblOK.Caption = "All Updates Completed OK"              'set OK message
    Else
        lblOK.BackColor = vbRed                                 'change label colour
        lblOK.Caption = "Error Found on Updates"                'display error message if NOT successful
    End If
    lblOK.Visible = True                                        'display message label
    txtLastRef = Nz(DLookup("ubeVersion", gRefTable))           'display new last Ref number
    btnClose.SetFocus                                           'move focus to Close btn
    ButtonCheck                                                 'enable Update button (if reqd)

End Sub

Private Sub Constraint_AfterUpdate()

    If Action <> "Set Relationship" Then                        'if record Action NOT 'Set Relationship' then
        Misc = ""                                               'clear Misc field
        Description = ""                                        'and Description field
    End If
    
    Select Case Constraint                                      'show possible parameter options when Property selected
        Case "Text Field Size ="
            Misc = "(1 to 255)"
        Case "Required ="
            Misc = "(Yes or No)"
        Case "Allow Zero Len ="
            Misc = "(Yes or No)"
        Case "Validation Rule ="
            Misc = "(Validation Rule)"
            Description = "(Validation Text)"
        Case "Default Value ="
            Misc = "(Default Value)"
        Case "New Field Name ="
            Misc = "(New Field Name)"
        Case "Ordinal Position ="
            Misc = "(1 to n)"
        Case "Description ="
            Description = "(Field Description)"
        Case "Set Primary Key ="
            Misc = "(Extra Field Name/s)"
        Case "Input Mask ="
            Misc = "(Input Mask)"
        Case "Format ="
            Misc = "(Field Format)"
        Case "Caption Name ="
            Misc = "(Caption Name)"
        Case "Decimal Places ="
            Misc = "(0-15)"
        Case "Fill With ="
            Misc = "(Field Data)"
        Case "Rich Text ="
            Misc = "(Yes or No)"
        Case "Smart Tags ="
            Misc = "(Smart Tag ID)"
    End Select

End Sub

Private Sub Form_AfterUpdate()

'Check current record for obvious errors

    If Constraint = "New Field Name =" And Nz(Misc) = "" Then MsgBox "ERROR. If you select to change the field name you must enter a name in the Misc field.", vbOKOnly, "Invalid Definition"
    If Description = "(Field Description - Optional)" Then Description = ""
    If Misc = "(Extra Field Name/s)" Then Misc = ""
    
    'add others here (if required)
    ButtonCheck                                                 'enable Update button (if reqd)

End Sub

Private Sub Form_Current()

    On Error GoTo ErrorCode

    SetConstraintSource                                         'set Constraint options drop-down
    txtDate = ChangeDate                                        'display record date
    
    Select Case Action                                          'for Action type fill in reqd combos
        Case "Copy Table"
            TableName.RowSource = FetchObjectList(1)            'add local table names to TableName combo
            FieldName.RowSource = FetchFieldList(TableName)     'add field names for selected table (if any) to field list
        Case "Remove Table", "New Field", "Delete Field", _
            "Change Type", "Set Property", "Set Relationship", _
            "Clear Relationship"
            TableName.RowSource = FetchObjectList(2)            'add linked table names to TableName combo
            FieldName.RowSource = FetchFieldList(TableName)     'add field names for selected table (if any) to field list
        Case "Run Query"
            TableName.RowSource = FetchObjectList(3)            'add action query names to TableName combo
            FieldName.RowSource = ""
        Case Else
            TableName.RowSource = ""                            'clear TableName combo for other types
            FieldName.RowSource = ""                            'clear FieldName combo for other types
    End Select
    Exit Sub

ErrorCode:
    If Err = 2176 Then Resume Next                              'continue if .RowSource overflow error (A2000 only)

End Sub

Private Sub Form_Dirty(Cancel As Integer)

    If Nz(Action) = "" Then Action.SetFocus                     'if Action field left blank then move cursor back
    lblOK.Visible = False                                       'hide message label (if visible)

End Sub

Private Sub Form_Load()
    OrderByOn = True
    btnClose.SetFocus
End Sub

Private Sub Form_Open(Cancel As Integer)

Dim db As DAO.Database
Dim tdf As TableDef
Dim vPathname As String
Dim I As Integer

    On Error GoTo ErrorCode

ResumeError:
    txtLastRef = Nz(DLookup("ubeVersion", gRefTable))                                               'display last used Ref number from Reference table
    ButtonCheck                                                                                     'enable Update button (if reqd)
    Exit Sub

ErrorCode:
    If Err = 3265 Or Err = 3078 Then                                                                'if table does not exist then
        If AddReferenceTable(gRefTable) = False Then Cancel = True: Exit Sub                        'allow user to create one
        GoTo ResumeError                                                                            'continue with other updates (if any)
    End If

End Sub

Private Sub Misc_AfterUpdate()

    If Constraint = "Required =" Or Constraint = "Allow Zero Len =" Or Constraint = "Rich Text =" Then  'if Constraint requires Yes or No then
        Select Case Misc                                                                                'tidy up Yes/No values
            Case "Yes", "Y"
                Misc = "Yes"
            Case "True", "T"
                Misc = "True"
            Case Else                                                                                   'if not Yes or True then
                Misc = "No"                                                                             'must be No
        End Select
    End If

End Sub

Private Sub TableName_AfterUpdate()
    FieldName.RowSource = FetchFieldList(TableName)                                     'add field names for selected table (if any) to field list
End Sub

Private Sub txtLastRef_AfterUpdate()

'If Developer changes LastRef field manually then

    CurrentDb.Execute "UPDATE [" & gRefTable & "] SET ubeVersion = " & txtLastRef       'update ubeVersion field to new value
    lblOK.Visible = False                                                               'hide message label (if visible)
    ButtonCheck                                                                         'and enable Update button (if reqd)

End Sub

Public Sub ButtonCheck()

'Check if all updates have been done and enable/disable Update btn accordingly

    If DMax("ID", "ubeUpdate") > Val(txtLastRef) Then btnUpdate.Enabled = True Else btnUpdate.Enabled = False   'if any outstanding updates then enable btn

End Sub

Public Sub SetConstraintSource()

'Changes list of options in Constraint drop-down if 'Set Relationships' action selected

    If Action = "Set Relationship" Then                         'if record Action = SetRelationship then
        Constraint.RowSource = "1-1 Not Enforced;" _
        & "1-1 Casc Updates;" _
        & "1-1 Casc Deletes;" _
        & "1-1 Casc Upd/Del;" _
        & "1-n Not Enforced;" _
        & "1-n Casc Updates;" _
        & "1-n Casc Deletes;" _
        & "1-n Casc Upd/Del"                                    'change Constraint options to relationship types
    Else                                                        'if Action NOT relationship then
        Constraint.RowSource = "Text Field Size =;" _
        & "Format =;" _
        & "Caption Name =;" _
        & "Decimal Places =;" _
        & "Input Mask =;" _
        & "Default Value =;" _
        & "Validation Rule =;" _
        & "Required =;" _
        & "Allow Zero Len =;" _
        & "New Field Name =;" _
        & "Ordinal Position =;" _
        & "Description =;" _
        & "Set Primary Key =;" _
        & "Indexed (No);" _
        & "Indexed (Dup OK);" _
        & "Indexed (No Dup);" _
        & "Set Compression;" _
        & "Fill With =;" _
        & "Rich Text =;" _
        & "Smart Tags ="                                         'set Constraint options to default values
    End If

End Sub

Private Sub txtLastRef_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) Like "[!0-9]" And KeyAscii <> vbKeyBack Then KeyAscii = 0              'allow keys 0-9 only
End Sub

Public Function FetchObjectList(vType As Long) As String

'Returns list of local tables, linked tables or action queries
'Entry  (vType) = Type of list requested (1=Local Tables, 2=Linked tables, 3=Action Queries)
'Exit   FetchObjectList = List of specified objects (delimited with ;)

Dim db As DAO.Database
Dim tdf As TableDef, qdf As QueryDef
Dim vAttrib As String
Dim I As Integer

    Set db = CurrentDb()
    Select Case vType
        Case 1                                                                              'chk for local tables
            For I = 0 To db.TableDefs.Count - 1
                Set tdf = db.TableDefs(I)
                vAttrib = (tdf.Attributes And dbSystemObject)
                If vAttrib = 0 Then
                    If Left(tdf.Name, 3) <> "ube" And Nz(tdf.Connect) = "" Then             'if not ube.. and not linked then
                        FetchObjectList = FetchObjectList & tdf.Name & ";"                  'add table name to string
                    End If
                End If
            Next I
        Case 2                                                                              'chk for linked tables
            For I = 0 To db.TableDefs.Count - 1
                Set tdf = db.TableDefs(I)
                vAttrib = (tdf.Attributes And dbSystemObject)
                If vAttrib = 0 Then
                    If Left(tdf.Name, 3) <> "ube" And Nz(tdf.Connect) <> "" Then            'if not ube.. and is linked then
                        FetchObjectList = FetchObjectList & tdf.Name & ";"                  'add table name to string
                    End If
                End If
            Next I
        Case 3                                                                              'chk for queries
            For Each qdf In db.QueryDefs                                                    'for each entry in Query Definition list
                If qdf.Type = 32 Or qdf.Type = 48 Or qdf.Type = 64 Or qdf.Type = 80 Then    'if query type = Delete/Update/Append/Make Table then
                    FetchObjectList = FetchObjectList & qdf.Name & ";"                      'add query name to string
                End If
            Next qdf
    End Select
    db.Close
    Set db = Nothing
    If Len(FetchObjectList) > 0 Then FetchObjectList = Left(FetchObjectList, Len(FetchObjectList) - 1)

End Function

Public Function FetchFieldList(vTable As String) As String

'Returns list of fields in specified table
'Entry  (vTable) = Name of table
'Exit   FetchFieldList = List of field names in table (delimited with ;)
    
Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim fld As DAO.Field
    
    On Error GoTo ErrorCode
    
    Set db = CurrentDb()
    Set tdf = db.TableDefs(vTable)
    For Each fld In tdf.Fields
        FetchFieldList = FetchFieldList & fld.Name & ";"
    Next
    If Len(FetchFieldList) > 0 Then FetchFieldList = Left(FetchFieldList, Len(FetchFieldList) - 1)
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

ErrorCode:
    If Err = 3265 Then
        Set db = Nothing
        Exit Function                        'if table does not exist then exit with ""
    Else
        MsgBox Err.Description
    End If

End Function
