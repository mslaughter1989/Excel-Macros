Attribute VB_Name = "SaveWorkbooksModule2"

Sub SaveOpenWorkbooksToMappedFolders()
    Dim wb As Workbook
    Dim fileName As String, fileDate As String
    Dim fileMonth As String, fileYear As String
    Dim folderName As String, fullPath As String
    Dim monthNames As Variant
    Dim regex As Object, matches As Object
    Dim dict As Object
    Dim key As Variant
    Dim fso As Object, logFile As Object
    Dim logPath As String
    Dim saveSuccess As Boolean

    ' Month names array
    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    ' Create dictionary of file name patterns to save paths
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "MarvelHR_723827_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\PrismHR\Marvel HR"
    dict.Add "ResidentialManagementGroup_691345_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Residential Management Group"
    dict.Add "WireMasters_mmddyyyy_FULL.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\WireMasters"
    dict.Add "21stCENTURYEQUIPMENT_324429_mmddyyyy_FULL", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\21st Century Equipment"
    dict.Add "AMERICANYOUTHACADEMY_566161_mmddyy_FULL.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\American Youth Academy"
    dict.Add "APEX Summit ShareWELL_803290_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\APEX-ShareWELL"
    dict.Add "AphenaPharma_yyyymmdd.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Aphena Pharma"
    dict.Add "Competitive Health Elig mmddyyyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\ARM"
    dict.Add "AvantiveSolutions_324998_mmddyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\AvantiveSolutions"
    dict.Add "CLEARWAY_622048_mmddyy_FULL.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Clearway Energy"
    dict.Add "Commcare_311525_yyyymmdd_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\CommCare"
    dict.Add "Cyprus_366105_mmddyy_FULL.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Cyprus Credit Union"
    dict.Add "Dinosaur_Restaurants_625972_mmddyyyy_FULL.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Dinosaur Restaurants"
    dict.Add "DouglasEllimanLLC_691360_mmddyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Douglas Elliman"
    dict.Add "EHS_RECURO_ELIGIBILITY_yyyymmdd.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Edison Health"
    dict.Add "arWestRestaurant_807303_mmddyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Far West Restaurant Group"
    dict.Add "FCFCU_368692_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\FCFCU"
    dict.Add "FirstNorthernBank_656200_mmddyy_FULL.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\First Northern Bank of Dixon"
    dict.Add "RecuroCare - ELIG - mmddyyyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Focus VUC"
    dict.Add "Fraley and Schilling_325381_mmddyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Fraley & Schilling"
    dict.Add "Gardner_314072_mmddyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Gardner Monthly"
    dict.Add "GulfGuaranty_649488_yyyymmdd.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Gulf Guaranty"
    dict.Add "H2OInnovation_768250_mmddyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\H2O Innovation"
    dict.Add "Health Karma_ThinBlueLine_TBL02_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\HealthKarma\Thin Blue Line"
    dict.Add "Health Karma_HKGold_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\HealthKarma\HealthKarma"
    dict.Add "HeartfieldAcademy_575381_mmddyy_FULL.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Hartfield Academy"
    dict.Add "HUB Gulf South SFTP Eligibility yyyymmdd.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Hunt Forest Products"
    dict.Add "IBA_664294_mmddyy_FULL.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\IBA - Focus"
    dict.Add "KIPPMetro_686788_mmddyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\KIPP Metro Atlanta"
    dict.Add "MARS_HILL_UNIVERSITY_680956_mmddyyyy_FULL.CSV", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Mars Hill University"
    dict.Add "Mobilelink_601933_mmddyy_FULL.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Master Mobilelink"
    dict.Add "Recuro_MightyWELL_yyyymmdd.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\MightyWELL"
    dict.Add "Miyamoto_655982_mmddyyyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Miyamoto Intern'l"
    dict.Add "324992_mmddyy_FULL.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\MVT Holdings"
    dict.Add "NHSL_709641_mmddyyyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Naman Howell Smith Lee"
    dict.Add "301448_FULL_yyyymmdd_100554.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Oglethorpe"
    dict.Add "Pediatrics Plus_mmddyyyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\PediatricsPlus and You Scream\You Scream (formerly Pediatrics Plus)"
    dict.Add "PediatricsPlus_yyyymmdd.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\PediatricsPlus and You Scream\PediatricsPlus"
    dict.Add "PF1_RECURO_Eligibility_mmddyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\People First"
    dict.Add "PrecisionToxicology_690856_mmddyy_FULL.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Precision Toxicology"
    dict.Add "RecuroHealth_mmddyyyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Premier Health"
    dict.Add "1_Source_Business_Solutions_733075_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\PrismHR\1Source"
    dict.Add "American Benefits Company_778182_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\PrismHR\American Benefits Company"
    dict.Add "Garyjames_744720_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\PrismHR\Garyjames Inc & Affiliates"
    dict.Add "HR Plus_797011_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\PrismHR\HR Plus LLC"
    dict.Add "Keena_807095_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\PrismHR\Keena"
    dict.Add "DeltaAdmin_757422_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\PrismHR\Delta Administrative"
    dict.Add "RSUtilityStructuresInc_733441_mmddyyyy.csv.pgp.tmp", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\RS Utility Structures"
    dict.Add "SOHO_643475_mmddyyyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\SOHO Studios"
    dict.Add "Summit_Funding_670602_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Summit Funding"
    dict.Add "SureCo File yyyymmdd.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\SureCo"
    dict.Add "TH_MARINE_366105_mmddyy_FULL.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\TH Marine"
    dict.Add "TRU-HealthGroup02262024_mmddyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Tru-Insurance"
    dict.Add "TruityFederalCreditUnion_812887_mmddyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Truity Federal Credit Union"
    dict.Add "USAHaulingRecycling_669506_mmddyyyy_Full.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\USA Hauling & Recycling"
    dict.Add "VOA_663805_mmddyyyy.csv", "C:\Users\MichaelSlaughter\OneDrive - Recuro Health\Documents - Ops\EDI-Eligibility\SFTP - Sharefile\Clients\Volunteers of America"

    ' Create log file in same folder as this workbook
    logPath = ThisWorkbook.Path & "\SaveLog.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.CreateTextFile(logPath, True)

    ' Loop through all open workbooks
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            fileName = wb.Name
            If InStr(fileName, ".") > 0 Then
                fileName = Left(fileName, InStrRev(fileName, ".") - 1)
            End If

            saveSuccess = False

            ' Try to match against each pattern
            For Each key In dict.Keys
                If InStr(key, "_mm") > 0 Then
                    If InStr(fileName, Left(key, InStr(key, "_mm") - 1)) > 0 Then
                        ' Extract 8-digit date
                        Set regex = CreateObject("VBScript.RegExp")
                        regex.pattern = "\d{8}"
                        regex.Global = False
                        regex.IgnoreCase = True

                        If regex.Test(fileName) Then
                            Set matches = regex.Execute(fileName)
                            fileDate = matches(0)
                            fileMonth = Left(fileDate, 2)
                            fileYear = Right(fileDate, 2)

                            folderName = fileMonth & monthNames(CInt(fileMonth) - 1) & fileYear
                            fullPath = dict(key) & "\" & folderName & "\"

                            ' Create folder recursively
                            CreateFoldersRecursively fullPath

                            wb.SaveCopyAs fullPath & wb.Name
                            logFile.WriteLine "Saved: " & wb.Name & " to " & fullPath
                            saveSuccess = True
                            Exit For
                        End If
                    End If
                End If
            Next key

            If Not saveSuccess Then
                logFile.WriteLine "Skipped: " & wb.Name & " (no matching pattern or date)"
            End If
        End If
    Next wb

    logFile.Close
    MsgBox "All matching workbooks have been saved. See SaveLog.txt for details."
End Sub

Sub CreateFoldersRecursively(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub
