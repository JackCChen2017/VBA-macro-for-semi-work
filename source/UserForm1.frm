VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5484
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7392
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CBAddFunction_Click()

rowCnt = Worksheets("Analysis_Script").Cells(1, 1).CurrentRegion.Rows.Count
Worksheets("Analysis_Script").Cells(rowCnt + 1, 1) = UserForm1.LBFunction.Text
End Sub

Private Sub CBAddSetting_Click()
rowCnt = Worksheets("Analysis_Script").Cells(1, 1).CurrentRegion.Rows.Count
If Worksheets("Analysis_Script").Cells(rowCnt, 2) = "" Then
    Worksheets("Analysis_Script").Cells(rowCnt, 2) = UserForm1.LBParamList.Text & ":" & UserForm1.TBParamSetting.Text
Else
    Worksheets("Analysis_Script").Cells(rowCnt, 2) = Worksheets("Analysis_Script").Cells(rowCnt, 2) & ";" & UserForm1.LBParamList.Text & ":" & UserForm1.TBParamSetting.Text
End If
End Sub



Private Sub CBExecute_Click()

Call AAA_Do_Analysis

End Sub

Private Sub LBFunction_Click()

UserForm1.LBParamList.Clear

Select Case UserForm1.LBFunction.Text
Case "Chart_new"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("SettingWorkbook")
    UserForm1.LBParamList.AddItem ("SettingSheetName")
    UserForm1.LBParamList.AddItem ("ChartSheetPrefix")
    UserForm1.TBInfo.Text = ""
    
Case "Chart_customize_by_title"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("ChartName")
    UserForm1.LBParamList.AddItem ("ChartSetting")
    UserForm1.TBInfo.Text = "ChartSetting: comma change to # and semicolon change to ##" & Chr(10) & "ChartSetting: ChartBy, SeriesBy, XAxisType, YAxisType, XMin, YMin, XMax, YMax, XLabel, YLabel, CrossAtX, CrossAtY, Width, Height, HasGridLineX, HasGridLineY, PlotAreaLine, SaveAsJPGFileName." & Chr(10) & "SaveAsJPGFileName can use special setting: ChartTitle"


Case "Ppt_create"
    UserForm1.LBParamList.AddItem ("aaa")
    UserForm1.TBInfo.Text = ""
    
Case "Ppt_open"
    UserForm1.LBParamList.AddItem ("aaa")
    UserForm1.TBInfo.Text = ""
    
Case "Ppt_close"
    UserForm1.LBParamList.AddItem ("aaa")
    UserForm1.TBInfo.Text = ""
    
Case "Ppt_save"
    UserForm1.LBParamList.AddItem ("aaa")
    UserForm1.TBInfo.Text = ""
    
Case "Ppt_add_slide"
    UserForm1.LBParamList.AddItem ("aaa")
    UserForm1.TBInfo.Text = ""
    
Case "Ppt_slide_changetitle"
    UserForm1.LBParamList.AddItem ("aaa")
    UserForm1.TBInfo.Text = ""
    
Case "Ppt_import_chart"
    UserForm1.LBParamList.AddItem ("aaa")
    UserForm1.TBInfo.Text = ""
    
Case "Ppt_import_picture"
    UserForm1.LBParamList.AddItem ("aaa")
    UserForm1.TBInfo.Text = ""
    
Case "Ppt_import_table"
    UserForm1.LBParamList.AddItem ("aaa")
    UserForm1.TBInfo.Text = ""
    
Case "Data_connection_remove"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.TBInfo.Text = ""
    
Case "Data_retrieval_csv"
    UserForm1.LBParamList.AddItem ("FileList")
    UserForm1.LBParamList.AddItem ("SQLSELECT")
    UserForm1.LBParamList.AddItem ("SQLWhere")
    UserForm1.LBParamList.AddItem ("OutputSheet")
    UserForm1.TBInfo.Text = ""
    
Case "Data_retrieval_lim"
    UserForm1.LBParamList.AddItem ("FileList")
    UserForm1.LBParamList.AddItem ("SQLSELECT")
    UserForm1.LBParamList.AddItem ("SQLWhere")
    UserForm1.LBParamList.AddItem ("OutputSheet")
    UserForm1.TBInfo.Text = ""
    
Case "Data_retrieval_rpt"
    UserForm1.LBParamList.AddItem ("FileList")
    UserForm1.LBParamList.AddItem ("DirFileList")
    UserForm1.TBInfo.Text = ""
    
Case "Data_retrieval_tst"
    UserForm1.LBParamList.AddItem ("Filepath")
    UserForm1.LBParamList.AddItem ("Filename")
    UserForm1.TBInfo.Text = ""
    
Case "Data_retrieval_IVcurveTxt"
    UserForm1.LBParamList.AddItem ("Filepath")
    UserForm1.LBParamList.AddItem ("Filename")
    UserForm1.TBInfo.Text = ""
    
    
Case "Sheet_remove"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.TBInfo.Text = ""
    
Case "Table_collapse_column"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.LBParamList.AddItem ("KeepColName")
    UserForm1.LBParamList.AddItem ("CollapseColName")
    UserForm1.LBParamList.AddItem ("NewColName")
    UserForm1.TBInfo.Text = ""
    
Case "Table_merge"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.LBParamList.AddItem ("WantedHeaderName")
    UserForm1.LBParamList.AddItem ("NonExistFillValue")
    UserForm1.TBInfo.Text = ""
    
Case "Table_split"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("SplitBy")
    UserForm1.LBParamList.AddItem ("SplitColName")
    UserForm1.LBParamList.AddItem ("GroupBy")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.TBInfo.Text = ""
    
Case "Table_split_quick"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("SplitBy")
    UserForm1.LBParamList.AddItem ("SplitColName")
    UserForm1.LBParamList.AddItem ("GroupBy")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.TBInfo.Text = ""
    
Case "Table_sort"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.LBParamList.AddItem ("SortByHeaderName")
    UserForm1.LBParamList.AddItem ("SortByRowOrCol")
    UserForm1.TBInfo.Text = ""
    
Case "Table_add_column"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.LBParamList.AddItem ("NewColName")
    UserForm1.LBParamList.AddItem ("NewColFormula")
    UserForm1.TBInfo.Text = ""
    
Case "Table_stack_column"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.LBParamList.AddItem ("KeepColName")
    UserForm1.LBParamList.AddItem ("StackColName")
    UserForm1.LBParamList.AddItem ("NewLabelColName")
    UserForm1.LBParamList.AddItem ("NewValueColName")
    UserForm1.TBInfo.Text = ""
    
Case "Table_add_row"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.LBParamList.AddItem ("NewRowNum")
    UserForm1.TBInfo.Text = ""
    
Case "Table_del_column"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.LBParamList.AddItem ("DelColName")
    UserForm1.TBInfo.Text = ""
    
Case "Table_del_row"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.LBParamList.AddItem ("DelRowRange")
    UserForm1.TBInfo.Text = ""
    
Case "Table_merge_same"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("ColName")
    UserForm1.LBParamList.AddItem ("RowName")
    UserForm1.LBParamList.AddItem ("Selection_range")
    UserForm1.TBInfo.Text = ""
    
Case "Table_filter_row"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.LBParamList.AddItem ("ColName")
    UserForm1.LBParamList.AddItem ("Criteria")
    UserForm1.TBInfo.Text = ""
    
Case "Table_vlookup"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.LBParamList.AddItem ("LookUpTableWorkbook")
    UserForm1.LBParamList.AddItem ("LookUpTableWorksheet")
    UserForm1.LBParamList.AddItem ("LookupValue")
    UserForm1.LBParamList.AddItem ("ReturnColumnName")
    UserForm1.TBInfo.Text = ""
    
Case "Table_fill_content"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.LBParamList.AddItem ("SheetRange")
    UserForm1.LBParamList.AddItem ("FillContent")
    UserForm1.TBInfo.Text = ""
    
Case "Table_formula_to_value"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.LBParamList.AddItem ("SheetRange")
    UserForm1.TBInfo.Text = ""
    
Case "Table_format"
    UserForm1.LBParamList.AddItem ("DataSheetName")
    UserForm1.LBParamList.AddItem ("OutSheetName")
    UserForm1.LBParamList.AddItem ("SheetRange")
    UserForm1.LBParamList.AddItem ("FormatString")
    UserForm1.TBInfo.Text = ""
    

Case "Xls_open"
    UserForm1.LBParamList.AddItem ("Filepath")
    UserForm1.LBParamList.AddItem ("Filename")
    UserForm1.TBInfo.Text = ""
    
Case "Xls_close"
    UserForm1.LBParamList.AddItem ("Filename")
    UserForm1.TBInfo.Text = ""
    
Case "Xls_sheet_copy"
    UserForm1.LBParamList.AddItem ("SourceWorkbook")
    UserForm1.LBParamList.AddItem ("TargetWorkbook")
    UserForm1.LBParamList.AddItem ("SourceWorksheet")
    UserForm1.TBInfo.Text = ""
    
Case "Xls_sheet_rename"
    UserForm1.LBParamList.AddItem ("SourceWorkbook")
    UserForm1.LBParamList.AddItem ("OldName")
    UserForm1.LBParamList.AddItem ("NewName")
    UserForm1.TBInfo.Text = ""
    
Case "Xls_file_saveas"
    UserForm1.LBParamList.AddItem ("SourceWorkbook")
    UserForm1.LBParamList.AddItem ("SaveAsName")
    UserForm1.TBInfo.Text = ""
    
End Select

End Sub

Private Sub UserForm_Initialize()

SheetExists ("Analysis_Script")
UserForm1.LBFunction.AddItem ("Chart_new")
UserForm1.LBFunction.AddItem ("Chart_customize_by_title")

UserForm1.LBFunction.AddItem ("Data_connection_remove")
UserForm1.LBFunction.AddItem ("Data_retrieval_csv")
UserForm1.LBFunction.AddItem ("Data_retrieval_lim")
UserForm1.LBFunction.AddItem ("Data_retrieval_rpt")
UserForm1.LBFunction.AddItem ("Data_retrieval_tst")
UserForm1.LBFunction.AddItem ("Data_retrieval_IVcurveTxt")

UserForm1.LBFunction.AddItem ("Ppt_create")
UserForm1.LBFunction.AddItem ("Ppt_open")
UserForm1.LBFunction.AddItem ("Ppt_close")
UserForm1.LBFunction.AddItem ("Ppt_save")
UserForm1.LBFunction.AddItem ("Ppt_add_slide")
UserForm1.LBFunction.AddItem ("Ppt_slide_changetitle")
UserForm1.LBFunction.AddItem ("Ppt_import_chart")
UserForm1.LBFunction.AddItem ("Ppt_import_picture")
UserForm1.LBFunction.AddItem ("Ppt_import_table")

UserForm1.LBFunction.AddItem ("Sheet_remove")

UserForm1.LBFunction.AddItem ("Table_collapse_column")
UserForm1.LBFunction.AddItem ("Table_merge")
UserForm1.LBFunction.AddItem ("Table_split")
UserForm1.LBFunction.AddItem ("Table_split_quick")
UserForm1.LBFunction.AddItem ("Table_sort")
UserForm1.LBFunction.AddItem ("Table_add_column")
UserForm1.LBFunction.AddItem ("Table_stack_column")
UserForm1.LBFunction.AddItem ("Table_add_row")
UserForm1.LBFunction.AddItem ("Table_del_column")
UserForm1.LBFunction.AddItem ("Table_del_row")
UserForm1.LBFunction.AddItem ("Table_merge_same")
UserForm1.LBFunction.AddItem ("Table_filter_row")
UserForm1.LBFunction.AddItem ("Table_vlookup")
UserForm1.LBFunction.AddItem ("Table_fill_content")
UserForm1.LBFunction.AddItem ("Table_formula_to_value")
UserForm1.LBFunction.AddItem ("Table_format")

UserForm1.LBFunction.AddItem ("Xls_open")
UserForm1.LBFunction.AddItem ("Xls_close")
UserForm1.LBFunction.AddItem ("Xls_sheet_copy")
UserForm1.LBFunction.AddItem ("Xls_sheet_rename")
UserForm1.LBFunction.AddItem ("Xls_file_saveas")

End Sub
