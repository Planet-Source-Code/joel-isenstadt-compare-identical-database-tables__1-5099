VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCompareDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Purpose:   To compare two MS Access Databases, two tables
'Created:   6/23/97 Using VB5
'Upgraded and Modified into a class:  12/23/99 Using VB6
'**********************************************************************************
'Demonstrates:
'       (1) ADO Technique to retrieve the Table names from an MS Access Database
'           using the Connection Object OpenSchema Method
'       (2) ADO Techniques for using a recordset to retrieve Columnn Name
'           Column DataType and Size
'Author:    Joel Isenstadt (iteachvb@aol.com)
'           I am a Microsoft Certified Trainer For Visual Basic
'           Employed by Tier Technology Training:  www.tiertraining.com
'**********************************************************************************
'History:
'           This class came as a result of MS Access being so easy to Change.
'           Before I Joined Tier Technology training I was employed by the
'           The United States Navy (DP1).
'           Our Managers decided that the MDB file should not be locked in any way
'           We were requested to train the users what they can do and what
'           They shouldn't do. This gave no Guarrantee though.
'           We had several programs that required
'           That several tables in different databases (IE MDB's for different years)
'           Be identical in every way.  The procedures in question copied information
'           of processed information by Column Index (Column position) or by using
'           For Each (which also relied on the Column Position)
'           Recently I needed this technique to do some debugging on an
'           Application I wrote for Tier Technology Training.
'           So I dug this program up.  Converted it to a Class and converted
'           All DAO code to ADO code (Just to see if I could do this)
'           Also This allow's this class to easily be converted for use
'           With other databases (Other than MS Access).
'**********************************************************************************
'Required Properties and Methods
'   DB1Path, DB2Path:               Sets the Location and name of the MDB
'   DB1TableName,  DB2TableName:    Sets the Table names to compared
'   OutputType:                     Types are jiFile, jiPrintPreview, jiPrinter
'   If OutputType is jiFile then OutputFileName is required
'   CompareTables:                  Compares the two selected tables
'   AreTablesTheSame:               Result of the Comparison (boolean)
'**********************************************************************************
'Additional Properties and Methods
'   GetDB1Tables, GetDB2Tables:     Retrieves a list of Tables from the specified
'                                   Databases (properties: DB1Path, DB2Path)
'                                   Results placed in arrays DB1Tables(), DB2Tables()
'   DB1TableCount, DB2TableCount:   Ubound of the two Arrays
'   GetDB1TableName, GetDB2TableName: Pass an Index, returns the table name
'                                     from the Array
'**********************************************************************************
'Events and Enums: See Below
'**********************************************************************************
'Notes:
'   Print Preview Output:   No Formatting Provided.
'                           Used a very simple technique to display lines of text.
'                           Passes the text back via an Event
'                           It is up to you to display the text (I demonstrated this
'                           using a Text box)
'   File Output:            Same as Print Preview.
'   Printer Output:         Same as Print preview.
'                           It is up to you to set up printer defaults
'   Using Other databases:  You need to modify Providers and Connect Strings
'                           You might also need to modify OpenSchema method
'                           Check VB6/ADO Help for additonal information
'   You might notice some code is very redundent. I created this class from
'                           a form I had created years ago.  I was lazy.  Sorry!
'                           Next time I need this program I'll fix it (maybe)
'**********************************************************************************
Option Explicit

Public Enum OutType
    jiFile = 1
    jiPrintPreview = 2
    jiPrinter = 3
End Enum

Private DB1Tables() As String
Private DB2Tables() As String

Private WithEvents cn1 As ADODB.Connection
Attribute cn1.VB_VarHelpID = -1
Private WithEvents cn2 As ADODB.Connection
Attribute cn2.VB_VarHelpID = -1
Private WithEvents rs1 As ADODB.Recordset
Attribute rs1.VB_VarHelpID = -1
Private WithEvents rs2 As ADODB.Recordset
Attribute rs2.VB_VarHelpID = -1

'Private Property Variables
Private mvarDB1Path As String
Private mvarDB1TablesCount As Integer
Private mvarDB2Path As String
Private mvarDB2TablesCount As Integer
Private mvarDB1TableName As String
Private mvarDB2TableName As String
Private mvarOutputType As OutType
Private mvarOutputFileName As String
Private mvarAreTablesTheSame As Boolean

'Events
Public Event StatusMessage(sMsg As String)  'Used to post current activity
Public Event Error(sMsg As String)          'Used to indicate an error occured
Public Event OutputLine(sMsg As String)     'Used for print preview
Public Property Get AreTablesTheSame() As Boolean
    AreTablesTheSame = mvarAreTablesTheSame
End Property
Private Function GetFilename(sPath As String) As String
'Purpose:   Parse the File name from the Path
'           Also remove the File Extension
'           Really only useful with ISAM databases
'Author:    Joel Isenstadt 12/23/99
'*******************************************************************************
    Dim i As Integer            'Position Indicator
    Dim sFileName As String     'Holds the file name
    
    i = InStrRev(sPath, "\")            'Get the last Backslash
    If i > 0 And i < Len(sPath) Then    'if the backslash was found but is not
                                        'the last character parse the filename
        sFileName = Mid(sPath, i + 1)
        
        i = InStrRev(sFileName, ".")    'Get the last period
        If i > 0 Then                   'If found the parse the filename
            GetFilename = Left(sFileName, i - 1)
        Else
            GetFilename = sFileName
        End If
    Else
        GetFilename = sPath
    End If

End Function
Public Property Let OutputFileName(ByVal vData As String)
    mvarOutputFileName = vData
End Property
Public Property Get OutputFileName() As String
    OutputFileName = mvarOutputFileName
End Property
Public Property Let OutputType(ByVal vData As OutType)
    mvarOutputType = vData
End Property
Public Property Get OutputType() As OutType
    OutputType = mvarOutputType
End Property
Public Sub CompareTables()
'Purpose:   To compare two tables from two different databases
'           Compares the Column Name, DataType and Size
'           Then creates a reports detailing each Column from the tables
'           Lists the Column Index then the column name
'           Then the Datatype and the size
'           If the Columns are different then both columns
'           Are Listed
'********************************************************************
'Author:    Joel Isenstadt  12-23-99
'Required Properties:
'   DB1Path, DB2Path,  DB1TableName,  DB2TableName
'   OutputType         If OutputType is jiFile then OutputFileName
'********************************************************************

    Dim i As Integer
    Dim iType1 As Integer
    Dim lSize1 As Long
    Dim iTotFields1 As Integer
    Dim iType2 As Integer
    Dim lSize2 As Long
    Dim iTotFields2 As Integer
    
    On Error GoTo ErrCompareTables
    
    'verify required property information
    If Len(mvarDB1Path) = 0 Then
        RaiseEvent Error("Missing DB1 Path")
        Exit Sub
    End If
    If Len(mvarDB2Path) = 0 Then
        RaiseEvent Error("Missing DB2 Path")
        Exit Sub
    End If
    If Len(mvarDB1TableName) = 0 Then
        RaiseEvent Error("Missing table from database 1")
        Exit Sub
    End If
    If Len(mvarDB2TableName) = 0 Then
        RaiseEvent Error("Missing table from database 2")
        Exit Sub
    End If

    'Verify Output Type and set up the output
    If mvarOutputType = jiFile Then
        If Len(mvarOutputFileName) = 0 Then
            RaiseEvent Error("Missing Output File Name")
            Exit Sub
        End If
        Open mvarOutputFileName For Output As #1
    End If

    'Start processing
    Screen.MousePointer = vbHourglass
    RaiseEvent StatusMessage("Attempting to Open connections")
    
    'Set Variables and Open connections
    Set cn1 = New ADODB.Connection
    Set rs1 = New ADODB.Recordset
    Set cn2 = New ADODB.Connection
    Set rs2 = New ADODB.Recordset

    With cn1
        .Provider = "Microsoft.Jet.OLEDB.3.51"
        .ConnectionString = "Persist Security Info=False;Data Source=" & mvarDB1Path
        .CursorLocation = adUseClient
        .Open
    End With
    With cn2
        .Provider = "Microsoft.Jet.OLEDB.3.51"
        .ConnectionString = "Persist Security Info=False;Data Source=" & mvarDB2Path
        .CursorLocation = adUseClient
        .Open
    End With
    RaiseEvent StatusMessage("Connections Open")
    
    'Fetch data
    DoEvents
    rs1.ActiveConnection = cn1
    rs1.Open "SELECT * FROM [" & mvarDB1TableName & "]"
    DoEvents
    rs2.ActiveConnection = cn2
    rs2.Open "SELECT * FROM [" & mvarDB2TableName & "]"
    DoEvents
    RaiseEvent StatusMessage("Fetch Completed")
    
    'assume the tables are the same. Change only they are false
    mvarAreTablesTheSame = True
    
    'Display the Header Message
    SendOutput "", False
    SendOutput "Comparing ", False
    SendOutput vbTab & "Table 1:" & GetFilename(mvarDB1Path) & "!" & mvarDB1TableName, False
    SendOutput vbTab & "Table 2:" & GetFilename(mvarDB2Path) & "!" & mvarDB2TableName, False
    SendOutput "", False
    
    'Start Comparing by Index Order
    RaiseEvent StatusMessage("Processing Field Names...")
    iTotFields1 = rs1.Fields.Count
    iTotFields2 = rs2.Fields.Count
    Do Until i = iTotFields1 Or i = iTotFields2
        'get data Types
        iType1 = rs1(i).Type
        lSize1 = rs1(i).DefinedSize
        iType2 = rs2(i).Type
        lSize2 = rs2(i).DefinedSize
        'Setup output
        'Determine sameness
        If (rs1(i).Name = rs2(i).Name) And (iType1 = iType2) And (lSize1 = lSize2) Then
            SendOutput "== " & Format(i, "000") & vbTab & rs1(i).Name, True
            SendOutput "  Type = " & FieldType(iType1), True
            SendOutput "  Size = " & CStr(lSize1), False
        Else
            SendOutput "<> " & Format(i, "000") & vbTab & "Table 1: " & rs1(i).Name, True
            SendOutput "  Type = " & FieldType(iType1), True
            SendOutput "  Size = " & CStr(lSize1), False
            SendOutput vbTab & "Table 2: " & rs2(i).Name, True
            SendOutput "  Type = " & FieldType(iType2), True
            SendOutput "  Size = " & CStr(lSize2), False
            mvarAreTablesTheSame = False
        End If
        i = i + 1
    Loop
    
    'determine if the tables have same amount of fields
    'And display the results
    SendOutput "", False
    If i <> iTotFields1 Then
        SendOutput "Database Table 1 is Finished while Database 2 still has fields", False
        mvarAreTablesTheSame = False
    ElseIf i <> iTotFields2 Then
        SendOutput "Database Table 2 is Finished while Database 1 still has fields", False
        mvarAreTablesTheSame = False
    Else
        SendOutput "Both database Fields have the same amount of Fields", False
    End If
    
    'Display the Trailer Message
    SendOutput "", False
    SendOutput "The two selected tables are ", True
    If mvarAreTablesTheSame Then
        SendOutput "the same!", False
    Else
        SendOutput "NOT the same!", False
    End If
    
    'clean up
    rs1.Close
    DoEvents
    rs2.Close
    DoEvents
    cn1.Close
    DoEvents
    cn2.Close
    DoEvents
    RaiseEvent StatusMessage("Processing Completed. Connections Closed")
    
exitCompareTables:
    If mvarOutputType = jiFile Then
        Close #1
    ElseIf mvarOutputType = jiPrinter Then
        Printer.EndDoc
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrCompareTables:
    RaiseEvent Error("Error: " & CStr(Err.Number) & " " & Err.Description)
    Resume exitCompareTables
    
End Sub

Public Property Let DB2TableName(ByVal vData As String)
    mvarDB2TableName = vData
End Property
Private Function FieldType(iType As Integer) As String
'Purpose:   Convert ado Data Type number (enum) into string
'Author:    Joel Isenstadt 12/23/99
'Reference: ADO Help File
'*******************************************************************************

    Select Case iType
        Case adBigInt
            FieldType = "adBigInt"
        Case adBinary
            FieldType = "adBinary"
        Case adBoolean
            FieldType = "adBoolean"
        Case adBSTR
            FieldType = "adBSTR"
        Case adChar
            FieldType = "adChar"
        Case adCurrency
            FieldType = "adCurrency"
        Case adDate
            FieldType = "adDate"
        Case adDBTimeStamp
            FieldType = "adDBTimeStamp"
        Case adDecimal
            FieldType = "adDecimal"
        Case adDouble
            FieldType = "adDouble"
        Case adEmpty
            FieldType = "adEmpty"
        Case adError
            FieldType = "adError"
        Case adGUID
            FieldType = "adGUID"
        Case adInteger
            FieldType = "adInteger"
        Case adLongVarBinary
            FieldType = "adLongVarBinary"
        Case adLongVarChar
            FieldType = "adLongVarChar"
        Case adLongVarWChar
            FieldType = "adLongVarWChar"
        Case adNumeric
            FieldType = "adNumeric"
        Case adSingle
            FieldType = "adSingle"
        Case adSmallInt
            FieldType = "adSmallInt"
        Case adTinyInt
            FieldType = "adTinyInt"
        Case adVarBinary
            FieldType = "adVarBinary"
        Case adVarChar
            FieldType = "adVarChar"
        Case adVariant
            FieldType = "adVariant"
        Case adVarNumeric
            FieldType = "adVarNumeric"
        Case adVarWChar
            FieldType = "adVarWChar"
        Case adWChar
            FieldType = "adWCharx"
        Case Else
            FieldType = "Unknown"
    End Select

End Function
Public Property Get DB2TableName() As String
    DB2TableName = mvarDB2TableName
End Property
Public Property Let DB1TableName(ByVal vData As String)
    mvarDB1TableName = vData
End Property
Public Property Get DB1TableName() As String
    DB1TableName = mvarDB1TableName
End Property
Public Property Get DB1TableCount() As Integer
    DB1TableCount = mvarDB1TablesCount
End Property
Public Property Get DB2TableCount() As Integer
    DB2TableCount = mvarDB2TablesCount
End Property
Public Property Let DB2Path(ByVal vData As String)
    mvarDB2Path = vData
End Property
Public Property Get DB2Path() As String
    DB2Path = mvarDB2Path
End Property
Public Property Let DB1Path(ByVal vData As String)
    mvarDB1Path = vData
End Property
Public Property Get DB1Path() As String
    DB1Path = mvarDB1Path
End Property
Public Function GetDB1TableName(Index As Integer) As String
'Purpose:   Retrieve one Table name from the private array
'Author:    Joel Isenstadt 12/23/99
'*******************************************************************************
    If mvarDB1TablesCount > 0 Then
        GetDB1TableName = DB1Tables(Index)
    Else
        GetDB1TableName = ""
    End If
End Function
Public Function GetDB2TableName(Index As Integer) As String
'Purpose:   Retrieve one Table name from the private array
'Author:    Joel Isenstadt 12/23/99
'*******************************************************************************
    If mvarDB2TablesCount > 0 Then
        GetDB2TableName = DB2Tables(Index)
    Else
        GetDB2TableName = ""
    End If
End Function
Public Function GetDBTables(iWhichDB As Integer, bGetSystemTables As Boolean) As Boolean
'Purpose:   To Retrieve all of the Table names and place into the
'           private array for later retrieval
'Author:    Joel Isenstadt 12/23/99
'*******************************************************************************
    Dim sTablename As String
    
    If iWhichDB = 1 Then
        If Len(mvarDB1Path) = 0 Then
            RaiseEvent Error("Missing DB Pathname for Database " & CStr(iWhichDB))
            Exit Function
        End If
    Else
        If Len(mvarDB2Path) = 0 Then
            RaiseEvent Error("Missing DB Pathname for Database " & CStr(iWhichDB))
            Exit Function
        End If
    End If
    
    On Error GoTo ErrGetDBTables
    
    Screen.MousePointer = vbHourglass

    RaiseEvent StatusMessage("Attempting to Open Connection")
    Set cn1 = New ADODB.Connection
    With cn1
        'setup the connection
        .Provider = "Microsoft.Jet.OLEDB.3.51"
        If iWhichDB = 1 Then
            .ConnectionString = "Persist Security Info=False;Data Source=" & mvarDB1Path
        Else
            .ConnectionString = "Persist Security Info=False;Data Source=" & mvarDB2Path
        End If
        .CursorLocation = adUseClient
        .Open
        DoEvents
        RaiseEvent StatusMessage("Connection Open. Building Table List...")
        
        'get the list of tables using the  Openschema method
        'The OpenSchema method has three arguments (third is optional)
        'The first argument identifies the type of schema information to return
        '   Identified as the Query Type
        'The second is an array that sets the Constraints and the Column names
        'This array is diffent for each of the Query Types
        'When working with MS Access you need to replace all constraints that
        '   are invalid with empty.  As in this case only the Table Name is needed
        'Once the Recordset is built then the Remaining of the processing
        'remains the same as when processing actual data.
        Set rs1 = .OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
        
        'Setup the Arrays
        If iWhichDB = 1 Then
            ReDim Preserve DB1Tables(0)
        Else
            ReDim Preserve DB2Tables(0)
        End If
        
        'Loop the Recordset and retrieve the table names
        Do While Not rs1.EOF
            sTablename = rs1!TABLE_NAME
            'Option to eliminate System tables (for Access Only)
            If Left(UCase(sTablename), 4) = "MSYS" Or Left(sTablename, 1) = "~" Then
                If bGetSystemTables Then
                    If iWhichDB = 1 Then
                        ReDim Preserve DB1Tables(UBound(DB1Tables) + 1)
                        DB1Tables(UBound(DB1Tables)) = sTablename
                    Else
                        ReDim Preserve DB2Tables(UBound(DB2Tables) + 1)
                        DB2Tables(UBound(DB2Tables)) = sTablename
                    End If
                End If
            Else
                If iWhichDB = 1 Then
                    ReDim Preserve DB1Tables(UBound(DB1Tables) + 1)
                    DB1Tables(UBound(DB1Tables)) = sTablename
                Else
                    ReDim Preserve DB2Tables(UBound(DB2Tables) + 1)
                    DB2Tables(UBound(DB2Tables)) = sTablename
                End If
            End If
            rs1.MoveNext
        Loop
        
        'Get Table count (for the property)
        If iWhichDB = 1 Then
            mvarDB1TablesCount = UBound(DB1Tables)
        Else
            mvarDB2TablesCount = UBound(DB2Tables)
        End If
        
        'Clean Up
        rs1.Close
        DoEvents
        .Close
        DoEvents
    End With
    RaiseEvent StatusMessage("Processing Completed. Connection Closed")
    GetDBTables = True
    
exitGetDBTables:
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrGetDBTables:
    RaiseEvent Error("Error: " & CStr(Err.Number) & " " & Err.Description)
    GetDBTables = False
    Resume exitGetDBTables
    
    
End Function

Private Sub SendOutput(sMsg As String, bSameLine As Boolean)
'Purpose:   Provide very simple output to the correct device
'           by the Property OutputType
'Arguement  sMsg        Text output
'           bSameline   Indicates if the line of text is not completed
'           In that case the output is saved in the static variable sHold
'Author:    Joel Isenstadt 12/23/99
'
'YES there are better ways to do this.  I wanted simplicity
'*******************************************************************************

    Static sHold As String
    
    If bSameLine Then
        Select Case mvarOutputType
            Case jiFile
                Print #1, sMsg;
            Case jiPrintPreview
                sHold = sHold & sMsg
            Case jiPrinter
                Printer.Print sMsg;
        End Select
    Else
        Select Case mvarOutputType
        Case jiFile
            Print #1, sMsg
        Case jiPrintPreview
            RaiseEvent OutputLine(sHold & sMsg)
            sHold = ""
        Case jiPrinter
            Printer.Print sMsg
        End Select
    End If
End Sub
Private Sub Class_Initialize()
    'Set up the Dynamic Arrays
    ReDim DB1Tables(0)
    ReDim DB2Tables(0)
End Sub


