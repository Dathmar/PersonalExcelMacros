Attribute VB_Name = "macro_saves"
Option Explicit
Sub save_as_xlam()

ActiveWorkbook.SaveAs filename:=ActiveWorkbook.path & "\" & Replace(ActiveWorkbook.Name, "xlsm", "xlam"), FileFormat:=xlOpenXMLAddIn

End Sub
Sub save_GPMs()
Dim save_date As String
Dim loc As String
Dim save_file As String

save_date = Format(Now, "MMddyy")

save_file = "C:\Users\adanner\OneDrive - Educational Testing Service\Asher\Macro backups\GPM\General_Purpose_Macros.xlam"

Call Copy_File_to_Location(save_file, "C:\Users\adanner\OneDrive - Educational Testing Service\Asher\Macro backups\GPM\General_Purpose_Macros_" & save_date & ".xlam")

loc = "\\ETSLAN.ORG\SAO\FS_K12_DATA_03\TestDev\Content Folders\Add-Ins\Excel\General Purpose Macros\General_Purpose_Macros.xlam"
SetAttr loc, vbNormal
Call Copy_File_to_Location(save_file, loc)
SetAttr loc, vbReadOnly

loc = "\\Ets\dfs\FS_SCS_Shared_01\Add-Ins\Excel\General Purpose Macros\General_Purpose_Macros.xlam"
SetAttr loc, vbNormal
Call Copy_File_to_Location(save_file, loc)
SetAttr loc, vbReadOnly

End Sub
Sub save_PARCC()
Dim save_date As String
Dim loc As String
Dim save_file As String

save_date = Format(Now, "MMddyy")

save_file = "C:\Users\adanner\OneDrive - Educational Testing Service\Asher\Macro backups\PARCC\PARCC_FP_Functions.xlam"

Call Copy_File_to_Location(save_file, "C:\Users\adanner\OneDrive - Educational Testing Service\Asher\Macro backups\PARCC\PARCC_FP_Functions_" & save_date & ".xlam")

loc = "\\ETSLAN.ORG\SAO\FS_K12_DATA_03\TestDev\Content Folders\Add-Ins\Excel\PARCC\PARCC_FP_Functions.xlam"
SetAttr loc, vbNormal
Call Copy_File_to_Location(save_file, loc)
SetAttr loc, vbReadOnly

loc = "\\Ets\dfs\FS_SCS_Shared_01\Add-Ins\Excel\PARCC\PARCC_FP_Functions.xlam"
SetAttr loc, vbNormal
Call Copy_File_to_Location(save_file, loc)
SetAttr loc, vbReadOnly

End Sub
Sub save_CollatTemp()
Dim save_date As String
Dim loc As String
Dim save_file As String

save_date = Format(Now, "MMddyy")

save_file = "C:\Users\adanner\OneDrive - Educational Testing Service\Asher\Macro backups\Collat\XPPCollaterals_Functions.xlam"

Call Copy_File_to_Location(save_file, "C:\Users\adanner\OneDrive - Educational Testing Service\Asher\Macro backups\Collat\XPPCollaterals_Functions_" & save_date & ".xlam")

loc = "\\Ets\dfs\FS_SCS_Shared_01\Add-Ins\Excel\XPPCollaterals\XPPCollaterals_Functions.xlam"
SetAttr loc, vbNormal
Call Copy_File_to_Location(save_file, loc)
SetAttr loc, vbReadOnly

End Sub
Sub save_TXSTAAR()
Dim save_date As String
Dim loc As String
Dim save_file As String

save_date = Format(Now, "MMddyy")

save_file = "C:\Users\adanner\OneDrive - Educational Testing Service\Asher\Macro backups\TXSTAAR\TXSTAAR_Functions.xlam"

Call Copy_File_to_Location(save_file, "C:\Users\adanner\OneDrive - Educational Testing Service\Asher\Macro backups\TXSTAAR\TXSTAAR_Functions_" & save_date & ".xlam")

loc = "\\ETSLAN.ORG\SAO\FS_K12_DATA_03\TestDev\Content Folders\Add-Ins\Excel\TXSTAAR\TXSTAAR_Functions.xlam"
SetAttr loc, vbNormal
Call Copy_File_to_Location(save_file, loc)
SetAttr loc, vbReadOnly

loc = "\\Ets\dfs\FS_SCS_Shared_01\Add-Ins\Excel\TXSTAAR\TXSTAAR_Functions.xlam"
SetAttr loc, vbNormal
Call Copy_File_to_Location(save_file, loc)
SetAttr loc, vbReadOnly

End Sub
