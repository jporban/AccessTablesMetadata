Attribute VB_Name = "TablesMetadata"
'---------------------------------------------------------------------------------------
' Module    : TablesMetadata
' Author    : Jean-Philippe Orban, XCLens sprls
' Website   : http://www.xclens.be
' Purpose   : Tool for populating table field's metadata like descriptions and captions.
'             Used to maintain forms and reports captions at the field description level.
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Variables:
' ~~~~~~~~~~
' mModuleName       Name of this module for debugging purpose
' mModuleErr        Starting error number for this module (increment by 100 for each module)
' mDebug            Switch on or off the debug mode
' mSavedError       Same structure as VB's Err object used to retain a copy of its values
'
' Procdures:
' ~~~~~~~~~~
' Function GetFieldProperty
' Function GetFieldPropertyValue
' Sub LogError
' Sub RaiseError
' Sub SaveError
' Sub SetAction
' Sub SetFieldPropertyValue
' Sub SaveError
' Sub WriteFieldsMetadata
'
' Revision History:
' ~~~~~~~~~~~~
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 0.1       2019-May-14             Initial Release
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Enum mOverWriteOptionsEnum
    mOverWriteDescriptions = 1
    mOverWriteCaptions = 2
End Enum

Public Enum mErrorEnum
    mErrorOk = 0
    mErrorPropertyNotFound = 3270
End Enum

Public Type mSavedError_type
    Number As Long
    Source As String
    Description As String
    HelpFile As String
    HelpContext As Long
    Action As String
End Type

Public Const mModuleName = "TableDefTools"
Public Const mModuleErr = vbObjectError + 0

Public mDebug As Boolean
Public mSavedError As mSavedError_type


'---------------------------------------------------------------------------------------
' Procedure : SaveError
' Author    : Jean-Philippe Orban, XCLens sprls
' Website   : http://www.xclens.be
' Purpose   : Save error object's properties in module variables for later re-raise
'             before the Err object is cleared by changing the "On Errror" behavior
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Parameters:
' ~~~~~~~~~~~
' pAction           Name of the action being made
'
' Return value:
' ~~~~~~~~~~~~~
' Nothing
'
' Usage:
' ~~~~~~
' SaveError "Dividing some number by 0"
'
' Revision History:
' ~~~~~~~~~~~~
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 0.1       2019-May-14             Initial Release
'---------------------------------------------------------------------------------------

Sub SaveError(Optional pAction As String)
    Const lProcName = "SaveError"
    Const lProcErr = mModuleErr + 10
    
    mSavedError.Number = Err.Number
    mSavedError.Source = Err.Source
    mSavedError.Description = Err.Description
    mSavedError.HelpFile = Err.HelpFile
    mSavedError.HelpContext = Err.HelpContext
    If pAction <> "" Then SetAction pAction
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetAction
' Author    : Jean-Philippe Orban, XCLens sprls
' Website   : http://www.xclens.be
' Purpose   : Set the name the action being made and output it to the console.
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Parameters:
' ~~~~~~~~~~~
' pAction           Name of the action being made
'
' Return value:
' ~~~~~~~~~~~~~
' Nothing
'
' Usage:
' ~~~~~~
' SetAction "Doing some action"
'
' Revision History:
' ~~~~~~~~~~~~
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 0.1       2019-May-14             Initial Release
'---------------------------------------------------------------------------------------

Sub SetAction( _
    pAction As String _
    , Optional pProcName As String _
)
    Const lProcName = "SetAction"
    Const lProcErr = mModuleErr + 20
    
    mSavedError.Action = pAction
    Debug.Print "[" & CStr(Now) & "]: " & IIf(pProcName = "", "", pProcName & ": ") & pAction
End Sub

'---------------------------------------------------------------------------------------
' Procedure : LogError
' Author    : Jean-Philippe Orban, XCLens sprls
' Website   : http://www.xclens.be
' Purpose   : Log a non-critical error to the console
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Parameters:
' ~~~~~~~~~~~
' pDescription
'
' Return value:
' ~~~~~~~~~~~~~
' Nothing
'
' Usage:
' ~~~~~~
' LogError "This is a warning"
'
' Revision History:
' ~~~~~~~~~~~~
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 0.1       2019-May-14             Initial Release
'---------------------------------------------------------------------------------------

Sub LogError( _
    Optional pDescription As String _
    , Optional pProcName As String _
)
    Const lProcName = "LogError"
    Const lProcErr = mModuleErr + 30
    
    Dim lDescription As String
    lDescription = "Run-time error '" & Err.Number & "':" _
        & vbCrLf & Err.Description & vbCrLf _
        & IIf(mModuleName = "", "", vbCrLf & "Module: " & mModuleName) _
        & vbCrLf & "Procedure: " & vbCrLf & pProcName _
        & IIf(mSavedError.Action = "", "", vbCrLf & "Action: " & mSavedError.Action) _
        & IIf(pDescription = "", "", vbCrLf & pDescription)
    Select Case Err.Number
        Case mErrorOk
        Case Else
            Debug.Print "[" & CStr(Now()) & "]: " & lDescription
    End Select
    Err.Clear
    SaveError 'Reset saved error properties
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RaiseError
' Author    : Jean-Philippe Orban, XCLens sprls
' Website   : http://www.xclens.be
' Purpose   : Re-raise a previously saved error
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Parameters:
' ~~~~~~~~~~~
' none
'
' Return value:
' ~~~~~~~~~~~~~
' Nothing
'
' Usage:
' ~~~~~~
' On Error Resume Next
' i=1/0
' SaveError "Dividing some number by 0"
' On Error Goto 0
' RaiseError
'
' Revision History:
' ~~~~~~~~~~~~
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 0.1       2019-May-14             Initial Release
'---------------------------------------------------------------------------------------

Sub RaiseError( _
    Optional pErrNumber As Long _
    , Optional pErrSource As String _
    , Optional pErrDescription As String _
    , Optional pErrHelpFile As String _
    , Optional pErrHelpContext As Long _
)
    Const lProcName = "RaiseError"
    Const lProcErr = mModuleErr + 40
    
    Err.Raise Number:=Nz(pErrNumber, mSavedError.Number), _
        Source:=Nz(pErrSource, mSavedError.Source), _
        Description:=Nz(pErrDescription, mSavedError.Description), _
        HelpFile:=Nz(pErrHelpFile, mSavedError.HelpFile), _
        HelpContext:=Nz(pErrHelpContext, mSavedError.HelpContext)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetFieldProperty
' Author    : Jean-Philippe Orban, XCLens sprls
' Website   : http://www.xclens.be
' Purpose   : Returns a field object based on its table, field and property name
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Parameters:
' ~~~~~~~~~~~
' pPropertyName                 Name of the field's property (built-in or custom)
' pFieldObj                     Field object (set back if missing)
' pTableName                    Name of the table (set back if missing)
' pFieldName                    Name of the field (set back if missing)
'
' Return value:
' ~~~~~~~~~~~~~
' Property Object
' Nothing if the property was not found
'
' Usage:
' ~~~~~~
' Set MyProperty = GetFieldProperty( _
'       pPropertyName:="MyCustomPropertyName" _
'       , pTableName:="MyTableName" _
'       , pFieldName:="MyFieldName")
'
' Set MyProperty = GetFieldProperty( _
'       pPropertyName:="Description" _
'       , pFieldObj:=MyFieldObj)
'
' Revision History:
' ~~~~~~~~~~~~
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 0.1       2019-May-14             Initial Release
'---------------------------------------------------------------------------------------

Function GetFieldProperty( _
    ByVal pPropertyName As String _
    , Optional ByRef pFieldObj As DAO.Field _
    , Optional ByRef pTableName As String _
    , Optional ByRef pFieldName As String _
) As DAO.Property

    Const lProcName = "GetFieldProperty"
    Const lProcErr = mModuleErr + 50
    
    Dim myTableDef As DAO.TableDef
    
    On Error GoTo Error_Handler
    
    If pFieldObj Is Nothing Then 'Missing field object
        If pTableName = "" Or pFieldName = "" Then 'Missing table and/or field name
            Err.Raise Number:=lProcErr + 1, _
                Description:="Invalid function usage : both field object and field names are missing" 'Stop
        Else
            'Get tabledef by name
            SetAction pProcName:=lProcName, pAction:="Get tabledef by name"
            On Error Resume Next: Set myTableDef = CurrentDb.QueryDefs(pTableName): SaveError
            On Error GoTo Error_Handler 'Stop ignoring errors
            Select Case mSavedError.Number
                Case mErrorOk 'Ok let's continue
                Case mErrorPropertyNotFound 'Tabledef not found
                    LogError "Table name: " & pTableName 'Non-critical, just log
                Case Else 'Unexpected error
                    RaiseError 'Re-raise and stop
            End Select
            'Get field by name
            SetAction pProcName:=lProcName, pAction:="Get field by name"
            On Error Resume Next: Set pFieldObj = myTableDef.Fields(pFieldName): SaveError
            On Error GoTo Error_Handler 'Stop ignoring errors
            Select Case mSavedError.Number
                Case mErrorOk 'Ok let's continue
                Case mErrorPropertyNotFound 'Field not found
                    LogError "Table name: " & pTableName _
                        & vbCrLf & "Field name :" & pFieldName 'Non-critical, just log
                Case Else 'Unexpected error
                    RaiseError 'Re-raise and stop
            End Select
            'Get property by name
            SetAction pProcName:=lProcName, pAction:="Get property by name"
            On Error Resume Next: Set GetFieldProperty = myTableDef.Fields(pPropertyName): SaveError
            On Error GoTo Error_Handler 'Stop ignoring errors
            Select Case mSavedError.Number
                Case mErrorOk 'Ok let's continue
                Case mErrorPropertyNotFound 'Property not found
                    LogError "Table name: " & pTableName _
                        & vbCrLf & "Field name :" & pFieldName _
                        & vbCrLf & "Property name : " & pPropertyName 'Non-critical, just log
                Case Else 'Unexpected error
                    RaiseError 'Re-raise and stop
            End Select
        End If
    Else 'Set back named arguments
        If Not pFieldObj Is Nothing Then
            pTableName = pFieldObj.SourceTable
            pFieldName = pFieldObj.Name
        Else 'Invalid field object -- this should never happen
            Err.Raise Number:=lProcErr + 1, _
                Description:="Invalid field object" 'Raise and stop
        End If
    End If
    
    On Error Resume Next: Set GetFieldProperty = pFieldObj.Properties(pPropertyName): SaveError
    On Error GoTo Error_Handler 'Stop ignoring errors
    Select Case mSavedError.Number
        Case mErrorOk 'Ok let's continue
        Case mErrorPropertyNotFound 'Property not found -- this should never happen
            RaiseError 'Re-raise and stop
        Case Else 'Unexpected error
            RaiseError 'Re-raise and stop
    End Select

Error_Handler_Exit:
    Exit Function 'Avoid error handler
    
Error_Handler:
    Dim lDescription As String
    lDescription = "Run-time error '" & Err.Number & "':" _
        & vbCrLf & Err.Description & vbCrLf _
        & IIf(mModuleName = "", "", vbCrLf & "Module: " & mModuleName) _
        & vbCrLf & "Procedure: " & vbCrLf & lProcName _
        & IIf(mSavedError.Action = "", "", vbCrLf & "Action: " & mSavedError.Action)
    Select Case Err.Number
        Case mErrorOk
        Case Else
            Debug.Print lDescription
            MsgBox prompt:=lDescription, _
                buttons:=vbOKOnly, _
                title:=mSavedError.Source
    End Select
    Err.Clear
    SaveError 'Reset saved error properties
    Resume Error_Handler_Exit

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFieldPropertyValue
' Author    : Jean-Philippe Orban, XCLens sprls
' Website   : http://www.xclens.be
' Purpose   : Get the value of a fields property
'             - "Description"
'             - "Caption"
'             - Custom property
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Parameters:
' ~~~~~~~~~~~
' pPropertyName                 Name of the field's property (built-in or even custom)
' pFieldObj                     Field object
' pTableName                    Name of the table
' pFieldName                    Name of the field
' pPropertyObj                  Property object (set back by this function)
'
' Return value:
' ~~~~~~~~~~~~~
' Value of the field's property
' "" if the property was not found
'
' Usage:
' ~~~~~~
' MyPropertyValue = GetFieldPropertyValue( _
'       pPropertyName:="Description" _
'       , pTableName:="MyTableName" _
'       , pFieldName:="MyFieldName")
'
' MyPropertyValue = GetFieldPropertyValue( _
'       pPropertyName:="MyCustomPropertyName" _
'       , pFieldObj:=MyFieldObj)
'
' Revision History:
' ~~~~~~~~~~~~
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 0.1       2019-May-14             Initial Release
'---------------------------------------------------------------------------------------

Function GetFieldPropertyValue( _
    ByVal pPropertyName As String _
    , Optional ByRef pFieldObj As DAO.Field _
    , Optional ByRef pTableName As String _
    , Optional ByRef pFieldName As String _
    , Optional ByRef pPropertyObj As DAO.Property _
) As String

    Const lProcName = "GetFieldPropertyValue"
    Const lProcErr = mModuleErr + 60
    
    Dim myProperty As DAO.Property
    
    On Error GoTo Error_Handler
    GetFieldPropertyValue = "" 'Returns empty string if property was not found
    
    If pFieldObj Is Nothing Then
        If pTableName = "" Or pFieldName = "" Then
            Err.Raise Number:=lProcErr + 1, _
                Description:="Both field object and field names are missing"
        Else
            Set myProperty = GetFieldProperty( _
                pPropertyName:=pPropertyName, _
                pTableName:=pTableName, _
                pFieldName:=pFieldName _
            )
        End If
    Else
        Set myProperty = GetFieldProperty( _
            pPropertyName:=pPropertyName, _
            pFieldObj:=pFieldObj _
        )
    End If

    SetAction pProcName:=lProcName, pAction:="Get field object property's value"
    If Not myProperty Is Nothing Then
        On Error Resume Next: GetFieldPropertyValue = myProperty.Value: SaveError
        On Error GoTo Error_Handler
        Select Case mSavedError.Number
            Case mErrorOk
            Case Else
                RaiseError
        End Select
    End If

Error_Handler_Exit:
    Exit Function 'Avoid error handler
    
Error_Handler:
    Dim lDescription As String
    lDescription = "Run-time error '" & Err.Number & "':" _
        & vbCrLf & Err.Description & vbCrLf _
        & IIf(mModuleName = "", "", vbCrLf & "Module: " & mModuleName) _
        & vbCrLf & "Procedure: " & vbCrLf & lProcName _
        & IIf(mSavedError.Action = "", "", vbCrLf & "Action: " & mSavedError.Action)
    Select Case Err.Number
        Case mErrorOk
        Case Else
            Debug.Print lDescription
            MsgBox prompt:=lDescription, _
                buttons:=vbOKOnly, _
                title:=mSavedError.Source
    End Select
    Err.Clear
    SaveError 'Reset saved error properties
    Resume Error_Handler_Exit

End Function

'---------------------------------------------------------------------------------------
' Procedure : SetFieldPropertyValue
' Author    : Jean-Philippe Orban, XCLens sprls
' Website   : http://www.xclens.be
' Purpose   : Set the value of a field property. Create if it doesn't exist yet.
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Parameters:
' ~~~~~~~~~~~
' pValue                        Value for the property
' pPropertyName                 Name of the field's property (built-in or custom)
' pFieldObj                     Field object (set back if missing)
' pTableName                    Name of the table (set back if missing)
' pFieldName                    Name of the field (set back if missing)
' pPropertyObj                  Property object (set back if missing)
'
' Return value:
' ~~~~~~~~~~~~~
' Nothing
'
' Usage:
' ~~~~~~
' SetFieldPropertyValue pPropertyObj:=MyPropertyObj, pValue:="MyValue"
'
' SetFieldPropertyValue _
'   pTableName:="MyTable", _
'   pFieldName:="MyField",_
'   pPropertyName:= "Caption", _
'   pValue:="MyValue"
'
' Revision History:
' ~~~~~~~~~~~~
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 0.1       2019-May-14             Initial Release
'---------------------------------------------------------------------------------------

Sub SetFieldPropertyValue( _
    ByVal pValue As String _
    , Optional ByVal pPropertyName As String _
    , Optional ByRef pFieldObj As DAO.Field _
    , Optional ByVal pTableName As String _
    , Optional ByVal pFieldName As String _
    , Optional ByRef pPropertyObj As DAO.Property _
)

    Const lProcName = "SetFieldPropertyValue"
    Const lProcErr = mModuleErr + 70
    
    Dim myTableDef As DAO.TableDef
    Dim myProperty As DAO.Property
    
    On Error GoTo Error_Handler
    
    If pPropertyObj Is Nothing Then
        If pFieldObj Is Nothing Then
            If pTableName = "" _
                Or pFieldName = "" _
                Or pPropertyName = "" _
            Then 'All optional arguments missing
                Err.Raise _
                    lProcErr + 1 _
                    , Description:="Property object or named properties are all missing"
            Else 'Find table, field and property objects based on their names
                SetAction pProcName:=lProcName, pAction:="Get property object of the field '" & pTableName & "." & pFieldName & "'"
                Set pPropertyObj = GetFieldProperty( _
                    pPropertyName:=pPropertyName _
                    , pTableName:=pTableName _
                    , pFieldName:=pFieldName _
                    , pFieldObj:=pFieldObj _
                )
                If pFieldObj Is Nothing Then 'Field not found
                    Err.Raise _
                        lProcErr + 2 _
                        , Description:= _
                            "A field named '" & pFieldName & _
                            "' for a table named '" & pTableName & _
                            "' could not be found."
                End If
            End If
        Else 'Valid field object
            pTableName = pFieldObj.SourceTable
            pFieldName = pFieldObj.Name
            Set pPropertyObj = GetFieldProperty( _
                pPropertyName:=pPropertyName _
                , pFieldObj:=pFieldObj _
            )
        End If
    Else
        If Not pFieldObj Is Nothing Then
            pTableName = pFieldObj.SourceTable
            pFieldName = pFieldObj.Name
            Set pPropertyObj = GetFieldProperty( _
                pPropertyName:=pPropertyName _
                , pFieldObj:=pFieldObj _
            )
        End If
    End If
    
    If pPropertyObj Is Nothing Then  'The property does not exist yet
        'Create the property and sets its value
        SetAction pProcName:=lProcName, pAction:="Create a property '" & pPropertyName _
            & "' for the field '" & pTableName & "." & pFieldName _
            & "' at sets its value to '" & pValue & "'"
        Set pPropertyObj = pFieldObj.CreateProperty( _
                Name:=pPropertyName _
                , Type:=dbText _
                , Value:=pValue _
        )
        SetAction pProcName:=lProcName, pAction:="Append field property to collection"
        pFieldObj.Properties.Append pPropertyObj
    Else 'The property object already exists
        'Set the property value
        SetAction pProcName:=lProcName, pAction:="Set the value of the property '" & pPropertyName & "' for the field '" & pTableName & "." & pFieldName & "' to '" & pValue & "'"
        pPropertyObj.Value = pValue
    End If

Error_Handler_Exit:
    Exit Sub 'Avoid error handler
    
Error_Handler:
    Dim lDescription As String
    lDescription = "Run-time error '" & Err.Number & "':" _
        & vbCrLf & Err.Description & vbCrLf _
        & IIf(mModuleName = "", "", vbCrLf & "Module: " & mModuleName) _
        & vbCrLf & "Procedure: " & vbCrLf & lProcName _
        & IIf(mSavedError.Action = "", "", vbCrLf & "Action: " & mSavedError.Action)
    Select Case Err.Number
        Case mErrorOk
        Case Else
            Debug.Print lDescription
            MsgBox prompt:=lDescription, _
                buttons:=vbOKOnly, _
                title:=mSavedError.Source
    End Select
    Err.Clear
    SaveError 'Reset saved error properties
    Resume Error_Handler_Exit

End Sub

'---------------------------------------------------------------------------------------
' Procedure : WriteFieldsMetadata
' Author    : Jean-Philippe Orban, XCLens sprls
' Website   : http://www.xclens.be
' Purpose   : Populate the table fields metadata
'             - description
'             - caption
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Parameters:
' ~~~~~~~~~~~
' pOverwriteOptions             Sum of the desired overwrite options (bitmask) :
'   - mOverwriteDescriptions    Replace fields descriptions based on their name
'   - mOverwriteCaptions        Replace fields captions based on their name
' pDebug                        Switch debug on or off
'
' Return value:
' ~~~~~~~~~~~~~
' Nothing
'
' Usage:
' ~~~~~~
' WriteFieldsMetadata pOverwrite:=mOverwriteDescriptions + mOverwriteCaptions, pDebug:=True
' WriteFieldsMetadata 3
'
' Revision History:
' ~~~~~~~~~~~~
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 0.1       2019-May-14             Initial Release
'---------------------------------------------------------------------------------------

Sub WriteFieldsMetadata( _
    Optional pOverwriteOptions As mOverWriteOptionsEnum _
    , Optional pDebug As Boolean = True _
)

    Const lProcName = "WriteFieldsMetadata"
    Const lProcErr = mModuleErr + 80
    
    Dim myTableDef As DAO.TableDef
    Dim myProperty As DAO.Property
    Dim myField As DAO.Field
    Dim myFieldName, myDescription, myCaption As String
    
    mDebug = pDebug 'Switch debug mode on or off
    
    For Each myTableDef In CurrentDb.TableDefs
        If Not (myTableDef.Attributes And dbSystemObject) Then 'Filter out system tables (bitmask)
            SetAction pProcName:=lProcName, pAction:="TableDef '" & myTableDef.Name & "'"
            For Each myField In myTableDef.Fields
                'Read field's name, remove underscores and set proper case
                SetAction pProcName:=lProcName, pAction:="Read field name of '" & myTableDef.Name & "." & myField.Name & "'"
                myFieldName = StrConv(Replace(myField.Name, "_", " "), vbProperCase)
                'Field's description
                If (pOverwriteOptions And mOverWriteDescriptions) Then 'Description overwrite enabled
                    'Set field's description based on its name
                    myDescription = myFieldName
                    SetAction pProcName:=lProcName, pAction:="Set field description of '" & myTableDef.Name & "." & myField.Name & "' to '" & myDescription & "'"
                    SetFieldPropertyValue _
                        pPropertyName:="Description" _
                        , pFieldObj:=myField _
                        , pPropertyObj:=myProperty _
                        , pValue:=myDescription
                Else 'Description overwrite disabled
                    'Get current field's description
                    SetAction pProcName:=lProcName, pAction:="Get the current description for field '" & myTableDef.Name & "." & myField.Name & "'"
                    myDescription = GetFieldPropertyValue( _
                        pFieldObj:=myField _
                        , pPropertyName:="Description" _
                    )
                    If myDescription = "" Then 'Field description not set yet, set it
                        myDescription = myFieldName
                        SetAction pProcName:=lProcName, pAction:="Description for field '" & myTableDef.Name & "." & myField.Name & "' not set yet. Set it to '" & myDescription & "'"
                        SetFieldPropertyValue _
                            pPropertyName:="Description" _
                            , pFieldObj:=myField _
                            , pPropertyObj:=myProperty _
                            , pValue:=myDescription
                    Else 'Field description already set, do nothing
                        SetAction pProcName:=lProcName, pAction:="Description for field '" & myTableDef.Name & "." & myField.Name _
                            & "' already set to '" & myDescription & "' and description overwrite is disabled. Do nothing."
                    End If
                End If
                'Field's caption
                If (pOverwriteOptions And mOverWriteCaptions) Then 'Caption overwrite enabled
                    'Set field's caption to its description
                    myCaption = myDescription
                    SetAction pProcName:=lProcName, pAction:="Set field caption of '" & myTableDef.Name & "." & myField.Name & "' to '" & myCaption & "'"
                    SetFieldPropertyValue _
                        pPropertyName:="Caption" _
                        , pFieldObj:=myField _
                        , pPropertyObj:=myProperty _
                        , pValue:=myCaption
                Else 'Caption overwrite disabled
                    'Get field's caption
                    SetAction pProcName:=lProcName, pAction:="Get the current caption of the field '" & myTableDef.Name & "." & myField.Name & "'"
                    myCaption = GetFieldPropertyValue( _
                        pFieldObj:=myField _
                        , pPropertyName:="Caption" _
                    )
                    If myCaption = "" Then 'Field caption not set yet, set it
                        myCaption = myDescription
                        SetAction pProcName:=lProcName, pAction:="Caption for field '" & myTableDef.Name & "." & myField.Name & "' not set yet. Set it to '" & myCaption & "'"
                        SetFieldPropertyValue _
                            pPropertyName:="Caption" _
                            , pFieldObj:=myField _
                            , pPropertyObj:=myProperty _
                            , pValue:=myCaption
                    Else 'Field caption already set, do nothing
                        SetAction pProcName:=lProcName, pAction:="Caption for field '" & myTableDef.Name & "." & myField.Name _
                            & "' already set to '" & myCaption & "' and caption overwrite is disabled. Do nothing."
                    End If
                End If
            Next
        End If
    Next
    
Error_Handler_Exit:
    Exit Sub 'Avoid error handler
    
Error_Handler:
    Dim lErrDesc As String
    lErrDesc = "Run-time error '" & Err.Number & "':" _
        & vbCrLf & Err.Description & vbCrLf _
        & IIf(mModuleName = "", "", vbCrLf & "Module: " & mModuleName) _
        & vbCrLf & "Procedure: " & vbCrLf & lProcName _
        & IIf(mSavedError.Action = "", "", vbCrLf & "Action: " & mSavedError.Action)
    Select Case Err.Number
        Case mErrorOk
        Case Else
            Debug.Print lErrDesc
            MsgBox prompt:=lErrDesc, _
                buttons:=vbOKOnly, _
                title:=mSavedError.Source
    End Select
    Err.Clear
    SaveError 'Reset saved error properties
    Resume Error_Handler_Exit

End Sub

