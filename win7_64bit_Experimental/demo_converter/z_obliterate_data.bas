Attribute VB_Name = "z_obliterate_data"
Option Compare Database
' Modified from "m_subroutines" in main TM DB

' DOESN'T WORK FOR LINKED TABLE DESIGN
' CURRENTLY IN USE

' IMPORT TO ACTUAL DATA DB AND USE IT THERE
' THEN DELETE MODULE BEFORE TAKING DB LIVE ...

Public db_on As Boolean
    ' for hiding / unhiding db
    ' default should be false
    ' false = hidden
    ' Toggle only works if all hide / unhide operations are performed
    ' from this vb code - it doesn't track
    
Public disableWarningMsgs As Boolean

Sub create_blank_Database()

    Dim strPath As String
        strPath = CurrentProject.FullName
        ' to use current db in commands
        
    MsgBox strPath

    disableWarningMsgs = True
    create_blank_MainDataTables
    clean_up_old_Tables

End Sub

Sub create_blank_MainDataTables()

    ' this code copies current Member and Activity tables to "old"
    ' and creates new blank tables to receive new data
        
    ' Dim directory_ThisDB As String
    '    directory_ThisDB = "D:\MLA_Code_Library\ms_office\db-2009\toastmasters_db\working-db\"
        
    ' directory_ThisDB & "TM-DB4.mdb"
    ' DoCmd.CopyObject , "Member_Records1", acTable = acDefinition, "Member_Records"
    
    msgBoxReturn = MsgBox("You are about to remove all data from the main data records of this database." & vbCrLf & _
                          "After you complete this action, you may not be able to undo it." & vbCrLf & _
                          "Are you sure you wish to continue?", vbYesNo, "Database Developer-Only-Macros Message:")
                          
    ' MsgBox msgBoxReturn
    
    If msgBoxReturn <> 6 Then GoTo skip_this_entire_macro
    
    Dim strPath As String
    ' strPath = "C:\ToastmastersDB\data\TM-DB-Data.accdb"
    strPath = CurrentProject.FullName
            
    clearOut_and_backup_Table "Member_Records", strPath
    clearOut_and_backup_Table "Member_OfficerHistory", strPath
    clearOut_and_backup_Table "Member_Activities", strPath
    clearOut_and_backup_Table "Mentorship", strPath
    clearOut_and_backup_Table "Mentorship_Assignments", strPath
    clearOut_and_backup_Table "Meeting_Notes", strPath
        
    ' Original Code - Depricated for now:
    ' -----------------
    ' msgBoxReturn = MsgBox("Clear sys_tables for value lists too?", vbYesNo, "Database Developer-Only-Macros Message:")
    ' If msgBoxReturn <> 6 Then GoTo skip_this_section
    
    ' MsgBox "comingSoon_MsgBox - Code Missing"
      ' problem with some kind of hidden relationship on these tables?
      ' clearOut_and_backup_Table "Sys_Categories"
      ' clearOut_and_backup_Table "Sys_CategoryItems"
    
    Exit Sub
        
skip_this_section:

    Exit Sub

skip_this_entire_macro:

    Quit acQuitSaveNone

End Sub



'  Modules used in above
' ----------------------------

Private Sub clearOut_and_backup_Table(tableName As String, dbName_str As String)
' clears tableName and creates a backup ending in "_old" with all
' of it's original data

    ' Dim dbName_str As String
    ' dbName_str = CurrentProject.FullName

    Dim tableName1 As String
    tableName1 = tableName & "1"
    Dim tableName_old As String
    tableName_old = tableName & "_old"
    
    DoCmd.TransferDatabase acImport, "Microsoft Access", dbName_str, acTable, tableName, tableName1, True
    ' strPath = database name
    
    DoCmd.Rename tableName_old, acTable, tableName
    DoCmd.Rename tableName, acTable, tableName1

End Sub

Sub clean_up_old_Tables()

    If disableWarningMsg = False Then GoTo skipBeginningTest

    msgBoxReturn = MsgBox("You are about to Delete all data from backups made of this DB's main data tables." & vbCrLf & _
                          "Once you complete this action, you will not be able to undo it." & vbCrLf & _
                          "Depending on context - this could result in the permanent loss of important data." & vbCrLf & _
                          "Are you sure you wish to continue?", vbYesNo, "Database Developer-Only-Macros Message:")
                          
    ' MsgBox msgBoxReturn
    
    If msgBoxReturn <> 6 Then GoTo skip_this_entire_macro
skipBeginningTest:

    ' MsgBox "Tables will now be deleted."
    ' DoCmd.DeleteObject acTable, "Member_Records_old"
    ' DoCmd.DeleteObject acTable, "Member_Activities_old"
    
    delete_TableName_old "Member_Records"
    delete_TableName_old "Member_OfficerHistory"
    delete_TableName_old "Member_Activities"
    delete_TableName_old "Mentorship"
    delete_TableName_old "Mentorship_Assignments"
    delete_TableName_old "Meeting_Notes"
        
    ' msgBoxReturn = MsgBox("Delete sys_itemsTable for value lists too?", vbYesNo, "Database Developer-Only-Macros Message:")
    ' If msgBoxReturn <> 6 Then GoTo skip_this_section
    
    ' delete_TableName_old "Sys_CategoryItems"
    ' delete_TableName_old "Sys_Categories"
    
skip_this_section:
    Exit Sub
    
skip_this_entire_macro:
    MsgBox "Operation Cancelled."
End Sub

Private Sub delete_TableName_old(tableName As String)
On Error GoTo skipThisStep1
    DoCmd.DeleteObject acTable, tableName & "_old"
    Exit Sub
skipThisStep1:
    MsgBox Err.Description & vbCrLf & "Deletion of " & tableName & "_old" & " cancelled.", vbInformation
End Sub

Private Sub backup_Table(tableName As String)
' backs up tableName to a copy ending in "_old"
' original tables left intact.

    Dim strPath As String
    strPath = CurrentProject.FullName

    Dim tableName_old As String
    tableName_old = tableName & "_old"
    
    DoCmd.TransferDatabase acImport, "Microsoft Access", strPath, acTable, tableName, tableName_old, False

End Sub

Sub make_backupCopies_ofMainDataTables_inThisDB()

On Error GoTo error_handling_thisSub
    backup_Table "Member_Records"
    backup_Table "Member_Activities"
    backup_Table "Sys_Categories"
    backup_Table "Sys_CategoryItems"
    
    Exit Sub

error_handling_thisSub:

    MsgBox Err.Description, vbExclamation, "Macro Error:  " & Err.Number

End Sub

Sub create_demoVersion_ThisDB()
On Error GoTo errorHandling_thisSubroutine

    msgBoxReturn = MsgBox("This function replaces all private user data with dumby data." & vbCrLf & _
                          "Once you complete this action, you will not be able to undo it." & vbCrLf & _
                          "You must backup the database before using this function or you will permanently lose data." & vbCrLf & _
                          "Are you sure you wish to continue?", vbYesNo, "Database Developer-Only-Macros Message:")
                              
    If msgBoxReturn <> 6 Then GoTo skip_this_entire_macro

    DoCmd.OpenQuery "Kill_user_private_data"
    
    Exit Sub
    
errorHandling_thisSubroutine:
    MsgBox Err.Description, vbExclamation, "Macro Error:  " & Err.Number

skip_this_entire_macro:

End Sub



