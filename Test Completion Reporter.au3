#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#RequireAdmin
;#AutoIt3Wrapper_usex64=n
#include <File.au3>
#include <Array.au3>
#include "Jira.au3"
#include "TestRail.au3"
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <Crypt.au3>
#include <GuiComboBox.au3>

Global $run_ids
Global $html
Global $app_name = "Test Completion Reporter"
Global $ini_filename = @ScriptDir & "\" & $app_name & ".ini"

Global $main_gui = GUICreate("TCR - " & $app_name, 860, 600)

GUICtrlCreateGroup("TestRail Setup", 10, 10, 410, 110)
GUICtrlCreateLabel("TestRail Username", 20, 30, 100, 20)
Global $testrail_username_input = GUICtrlCreateInput(IniRead($ini_filename, "main", "testrailusername", "sgriffin@janison.com"), 140, 30, 250, 20)
GUICtrlCreateLabel("TestRail Password", 20, 50, 100, 20)
Global $testrail_password_input = GUICtrlCreateInput("", 140, 50, 250, 20, $ES_PASSWORD)
GUICtrlCreateLabel("TestRail Project", 20, 70, 100, 20)
Global $testrail_project_combo = GUICtrlCreateCombo("", 140, 70, 250, 20, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))
GUICtrlCreateLabel("TestRail Plan", 20, 90, 100, 20)
Global $testrail_plan_combo = GUICtrlCreateCombo("", 140, 90, 250, 20, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))
GUICtrlCreateGroup("", -99, -99, 1, 1)

Local $testrail_encrypted_password = IniRead($ini_filename, "main", "testrailpassword", "")
Global $testrail_decrypted_password = ""

if stringlen($testrail_encrypted_password) > 0 Then

	$testrail_decrypted_password = _Crypt_DecryptData($testrail_encrypted_password, "applesauce", $CALG_AES_256)
	$testrail_decrypted_password = BinaryToString($testrail_decrypted_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_decrypted_password = ' & $testrail_decrypted_password & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	GUICtrlSetData($testrail_password_input, $testrail_decrypted_password)
Else

	$testrail_decrypted_password = ""
EndIf

GUICtrlCreateGroup("Jira Setup", 440, 10, 410, 110)
GUICtrlCreateLabel("Jira Username", 450, 30, 100, 20)
Global $jira_username_input = GUICtrlCreateInput(IniRead($ini_filename, "main", "jirausername", "sgriffin@janison.com.au"), 570, 30, 250, 20)
GUICtrlCreateLabel("Jira Password", 450, 50, 100, 20)
Global $jira_password_input = GUICtrlCreateInput("", 570, 50, 250, 20, $ES_PASSWORD)
GUICtrlCreateGroup("", -99, -99, 1, 1)

Local $jira_encrypted_password = IniRead($ini_filename, "main", "jirapassword", "")
Global $jira_decrypted_password = ""

if stringlen($jira_encrypted_password) > 0 Then

	$jira_decrypted_password = _Crypt_DecryptData($jira_encrypted_password, "applesauce", $CALG_AES_256)
	$jira_decrypted_password = BinaryToString($jira_decrypted_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jira_decrypted_password = ' & $jira_decrypted_password & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	GUICtrlSetData($jira_password_input, $jira_decrypted_password)
Else

	$jira_decrypted_password = ""
EndIf

GUICtrlCreateLabel("Jira Epic Keys", 20, 130, 100, 20)
Global $epic_key_input = GUICtrlCreateInput(IniRead($ini_filename, "main", "epickeys", "SEAB-4683,SEAB-4687"), 140, 130, 680, 20)
Global $start_button = GUICtrlCreateButton("Start", 10, 160, 100, 20, -1, $BS_DEFPUSHBUTTON)
GUICtrlSetState(-1, $GUI_DISABLE)
Global $display_report_button = GUICtrlCreateButton("Display Report", 120, 160, 100, 20)
GUICtrlSetState(-1, $GUI_DISABLE)
Global $display_data_button = GUICtrlCreateButton("Display Data", 240, 160, 100, 20)
GUICtrlSetState(-1, $GUI_DISABLE)

Global $listview = GUICtrlCreateListView("Epic Key|PID|Extract Status", 10, 200, 410, 300, $LVS_SHOWSELALWAYS)
_GUICtrlListView_SetColumnWidth(-1, 0, 200)
_GUICtrlListView_SetColumnWidth(-1, 1, 50)
_GUICtrlListView_SetColumnWidth(-1, 2, 200)
_GUICtrlListView_SetExtendedListViewStyle($listview, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES))

Global $status_input = GUICtrlCreateInput("Enter the ""Epic Key"" and click ""Start""", 10, 600 - 25, 400, 20, $ES_READONLY, $WS_EX_STATICEDGE)
Global $progress = GUICtrlCreateProgress(420, 600 - 25, 400, 20)

GUISetState(@SW_SHOW, $main_gui)

; Startup SQLite

_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)

FileDelete(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Open(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Exec(-1, "CREATE TABLE report (EpicKey,IssueKey,IssueSummary,IssueType,TestProjectName,TestCaseID,TestCaseName,TestCaseResult,DefectID,DefectDetails,Reason,ImpactActionRequired);") ; CREATE a Table

; Startup TestRail

GUICtrlSetData($status_input, "Starting the TestRail connection ... ")
_TestRailDomainSet("https://janison.testrail.com")
_TestRailLogin(GUICtrlRead($testrail_username_input), GUICtrlRead($testrail_password_input))

if StringLen(GUICtrlRead($testrail_password_input)) > 0 Then

	; Authentication

	GUICtrlSetData($status_input, "Authenticating against TestRail ... ")
	_TestRailAuth()

	GUICtrlSetState($testrail_project_combo, $GUI_DISABLE)
	GUICtrlSetState($testrail_plan_combo, $GUI_DISABLE)

	GUICtrlSetData($status_input, "Querying TestRail Projects ... ")

	Local $project_id_name = _TestRailGetProjectsIDAndNameArray()
	Local $project_id_str = ""

	for $i = 0 to (UBound($project_id_name) - 1)

		if StringLen($project_id_str) > 0 Then

			$project_id_str = $project_id_str & "|"
		EndIf

		$project_id_str = $project_id_str & $project_id_name[$i][0] & " - " & $project_id_name[$i][1]
	Next

	GUICtrlSetData($testrail_project_combo, $project_id_str)
	GUICtrlSetState($testrail_project_combo, $GUI_ENABLE)

	Local $project_to_select = IniRead($ini_filename, "main", "testrailproject", "")
	Local $plan_to_select = IniRead($ini_filename, "main", "testrailplan", "")

	if StringLen($project_to_select) > 0 Then

		Local $index = _GUICtrlComboBox_SelectString($testrail_project_combo, $project_to_select)

		if $index > -1 Then

			query_testrail_plans()

			if StringLen($plan_to_select) > 0 Then

				Local $index = _GUICtrlComboBox_SelectString($testrail_plan_combo, $plan_to_select)

				if $index > -1 Then

					query_testrail_runs()
				EndIf
			EndIf
		EndIf
	EndIf
EndIf

GUICtrlSetData($status_input, "")
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")

; Loop until the user exits.
While 1

	; GUI msg loop...
	$msg = GUIGetMsg()

	Switch $msg

		Case $GUI_EVENT_CLOSE

			IniWrite($ini_filename, "main", "testrailusername", GUICtrlRead($testrail_username_input))
			IniWrite($ini_filename, "main", "testrailproject", GUICtrlRead($testrail_project_combo))
			IniWrite($ini_filename, "main", "testrailplan", GUICtrlRead($testrail_plan_combo))
			IniWrite($ini_filename, "main", "jirausername", GUICtrlRead($jira_username_input))
			IniWrite($ini_filename, "main", "epickeys", GUICtrlRead($epic_key_input))

			$testrail_encrypted_password = _Crypt_EncryptData(GUICtrlRead($testrail_password_input), "applesauce", $CALG_AES_256)
			IniWrite($ini_filename, "main", "testrailpassword", $testrail_encrypted_password)

			$jira_encrypted_password = _Crypt_EncryptData(GUICtrlRead($jira_password_input), "applesauce", $CALG_AES_256)
			IniWrite($ini_filename, "main", "jirapassword", $jira_encrypted_password)


			ExitLoop

		Case $start_button

			_SQLite_Exec(-1, "DELETE FROM report;") ; CREATE a Table

			GUICtrlSetData($progress, 0)
			GUICtrlSetState($epic_key_input, $GUI_DISABLE)
			GUICtrlSetState($start_button, $GUI_DISABLE)
			GUICtrlSetState($display_report_button, $GUI_DISABLE)
			GUICtrlSetState($display_data_button, $GUI_DISABLE)
			GUISetCursor(15, 1, $main_gui)
			_GUICtrlListView_DeleteAllItems($listview)

			; populate listview with epic keys

			Local $epic_key = StringSplit(GUICtrlRead($epic_key_input), ",;|", 2)

			for $each in $epic_key

				Local $pid = ShellExecute(@ScriptDir & "\data_extractor.exe", """" & GUICtrlRead($testrail_username_input) & """ """ & GUICtrlRead($testrail_password_input) & """ """ & $run_ids & """ """ & GUICtrlRead($jira_username_input) & """ """ & GUICtrlRead($jira_password_input) & """ """ & $each & """", "", "", @SW_HIDE)
				GUICtrlCreateListViewItem($each & "|" & $pid & "|In Progress", $listview)
			Next

			While True

				Local $all_epics_done = True

				for $index = 0 to (_GUICtrlListView_GetItemCount($listview) - 1)

					Local $pid = _GUICtrlListView_GetItemText($listview, $index, 1)
					Local $status = _GUICtrlListView_GetItemText($listview, $index, 2)

					if StringCompare($status, "In Progress") = 0 Then

						$all_epics_done = False

						if ProcessExists($pid) = False Then

							_GUICtrlListView_SetItemText($listview, $index, "Done", 2)

						EndIf
					EndIf
				Next

				if $all_epics_done = True Then

					ExitLoop
				EndIf

				Sleep(1000)
			WEnd

			$html = 				"<!DOCTYPE html>" & @CRLF & _
									"<html>" & @CRLF & _
									"<head>" & @CRLF & _
									"<style>" & @CRLF & _
									"table, th, td {" & @CRLF & _
									"    border: 1px solid black;" & @CRLF & _
									"    border-collapse: collapse;" & @CRLF & _
									"}" & @CRLF & _
									"</style>" & @CRLF & _
									"</head>" & @CRLF & _
									"<body>" & @CRLF & _
									"<h1>Test Completion Report</h1>" & @CRLF & _
									"<h2>Test Cases Planned vs Actuals</h2>" & @CRLF

			SQLite_to_HTML_table("SELECT EpicKey AS ""Epic ID"",IssueKey AS ""Story ID"",IssueSummary AS ""Story Description"",IssueType,count(TestCaseID) AS Planned,count(TestCaseResult) AS Actual,(count(TestCaseID) - count(TestCaseResult)) AS Difference FROM report GROUP BY IssueKey ORDER BY EpicKey,IssueKey;", "center,center,left,center,center,center,center", "")

			$html = $html &			"<h2>Test Cases Status Summary</h2>" & @CRLF

			SQLite_to_HTML_table("SELECT DISTINCT EpicKey AS ""Epic ID"",IssueKey AS ""Story ID"",(SELECT count(b.TestCaseID) FROM report AS b WHERE a.IssueKey = b.IssueKey AND b.TestCaseResult = 'Passed') AS ""Passed"",(SELECT count(b.TestCaseID) FROM report AS b WHERE a.IssueKey = b.IssueKey AND b.TestCaseResult = 'Failed') AS ""Failed"",(SELECT count(b.TestCaseID) FROM report AS b WHERE a.IssueKey = b.IssueKey AND b.TestCaseResult = 'Blocked') AS ""Blocked"",(SELECT count(b.TestCaseID) FROM report AS b WHERE a.IssueKey = b.IssueKey AND b.TestCaseResult = 'Incomplete Automation') AS ""Incomplete Automation"",(SELECT count(b.TestCaseID) FROM report AS b WHERE a.IssueKey = b.IssueKey AND b.TestCaseResult = 'Untested') AS ""Untested"",(SELECT count(b.TestCaseID) FROM report AS b WHERE a.IssueKey = b.IssueKey AND b.TestCaseResult = 'Removed from Run') AS ""Removed from Run"",(SELECT count(b.TestCaseID) FROM report AS b WHERE a.IssueKey = b.IssueKey AND b.TestCaseResult = 'Retest') AS ""Retest"",(SELECT count(b.TestCaseID) FROM report AS b WHERE a.IssueKey = b.IssueKey AND b.TestCaseResult = 'Known Issue') AS ""Known Issue"",(SELECT count(b.TestCaseID) FROM report AS b WHERE a.IssueKey = b.IssueKey AND b.TestCaseResult = 'In Progress') AS ""In Progress"",(SELECT count(b.TestCaseID) FROM report AS b WHERE a.IssueKey = b.IssueKey) AS ""Total"" FROM report AS a ORDER BY a.EpicKey,a.IssueKey;", "center,center,center,center,center,center,center,center,center,center,center,center", "")

			$html = $html &			"<h2>Failed Test Cases</h2>" & @CRLF

			SQLite_to_HTML_table("SELECT EpicKey AS ""Epic ID"",IssueKey AS ""Story ID"",TestCaseID AS ""Test Case ID"",DefectID AS ""Defect ID"",DefectDetails AS ""Defect Details"",ImpactActionRequired AS ""Impact and Action Required"" FROM report WHERE TestCaseResult = 'Failed' ORDER BY EpicKey,IssueKey", "center,center,center,center,left,left", "There are no failed test cases reported for this release.")

			$html = $html &			"<h2>Test Cases Not Executed</h2>" & @CRLF

			SQLite_to_HTML_table("SELECT EpicKey AS ""Epic ID"",IssueKey AS ""Story ID"",TestCaseID AS ""Test Case ID"",Reason,ImpactActionRequired AS ""Impact and Action Required"" FROM report WHERE TestCaseResult IN ('Untested','Blocked','Incomplete Automation') ORDER BY EpicKey,IssueKey", "center,center,center,left,left", "There are no test cases that weren't executed against this release.")

			$html = $html &			"<h2>Removed Test Cases</h2>" & @CRLF

			SQLite_to_HTML_table("SELECT EpicKey AS ""Epic ID"",IssueKey AS ""Story ID"",TestCaseID AS ""Test Case ID"",Reason,ImpactActionRequired AS ""Impact and Action Required"" FROM report WHERE TestCaseResult = 'Removed from Run' ORDER BY EpicKey,IssueKey", "center,center,center,left,left", "There are no removed test cases reported for this release.")

			$html = $html &			"<h2>Outstanding Defects</h2>" & @CRLF
			$html = $html &			"<p><i>TODO</i></p>" & @CRLF
			$html = $html &			"</body>" & @CRLF & _
									"</html> " & @CRLF

			FileDelete(@ScriptDir & "\html_report.html")
			FileWrite(@ScriptDir & "\html_report.html", $html)

			Local $aResult, $iRows, $iColumns, $iRval

			$iRval = _SQLite_GetTable2d(-1, "SELECT * FROM report;", $aResult, $iRows, $iColumns)

			If $iRval = $SQLITE_OK Then

				Local $data = _SQLite_Display2DResult($aResult, 0, True)
				FileDelete(@ScriptDir & "\data.txt")
				FileWrite(@ScriptDir & "\data.txt", $data)
			Else
				MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
			EndIf

			GUICtrlSetData($progress, 0)
			GUICtrlSetData($status_input, "")
			GUICtrlSetState($epic_key_input, $GUI_ENABLE)
			GUICtrlSetState($start_button, $GUI_ENABLE)
			GUICtrlSetState($display_report_button, $GUI_ENABLE)
			GUICtrlSetState($display_data_button, $GUI_ENABLE)
			GUISetCursor(2, 0, $main_gui)

		Case $display_report_button

			ShellExecute(@ScriptDir & "\html_report.html")

		case $display_data_button

			ShellExecute("notepad", "data.txt", @ScriptDir)


	EndSwitch

WEnd

GUIDelete($main_gui)

; Shutdown TestRail

GUICtrlSetData($status_input, "Closing Jira ... ")
_JiraShutdown()


Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg
    Local $hWndFrom, $iIDFrom, $iCode
    $hWndFrom = $lParam
    $iIDFrom = BitAND($wParam, 0xFFFF) ; Low Word
    $iCode = BitShift($wParam, 16) ; Hi Word

	Switch $hWndFrom

        Case GUICtrlGetHandle($testrail_project_combo)

			Switch $iCode

                Case $CBN_SELCHANGE ; Sent when the user changes the current selection in the list box of a combo box

					query_testrail_plans()
            EndSwitch

        Case GUICtrlGetHandle($testrail_plan_combo)

			Switch $iCode

                Case $CBN_SELCHANGE ; Sent when the user changes the current selection in the list box of a combo box

					query_testrail_runs()

            EndSwitch
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND

Func SQLite_to_HTML_table($query, $td_alignments, $empty_message)

	Local $td_alignment = StringSplit($td_alignments, ",", 2)

	Local $aResult, $iRows, $iColumns, $iRval

	$iRval = _SQLite_GetTable2d(-1, $query, $aResult, $iRows, $iColumns)

	If $iRval = $SQLITE_OK Then

		_SQLite_Display2DResult($aResult)

		Local $num_rows = UBound($aResult, 1)
		Local $num_cols = UBound($aResult, 2)

		if $num_rows < 2 Then

			$html = $html &	"<p>" & $empty_message & "</p>" & @CRLF
		Else

			$html = $html &	"<table>" & @CRLF
			$html = $html & "<tr>"

			for $i = 0 to ($num_cols - 1)

				$html = $html & "<th>" & $aResult[0][$i] & "</th>" & @CRLF
			Next

			$html = $html & "</tr>" & @CRLF

			for $i = 1 to ($num_rows - 1)

				$html = $html & "<tr>"

				for $j = 0 to ($num_cols - 1)

					$html = $html & "<td align=""" & $td_alignment[$j] & """>" & $aResult[$i][$j] & "</td>" & @CRLF
				Next

				$html = $html & "</tr>" & @CRLF
			Next

			$html = $html &	"</table>" & @CRLF
		EndIf
	Else
		MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	EndIf
EndFunc


Func query_testrail_plans()

	Local $project_id_name = GUICtrlRead($testrail_project_combo)

	IniWrite($ini_filename, "main", "testrailproject", $project_id_name)

	Local $project_part = StringSplit($project_id_name, " - ", 3)

	GUICtrlSetData($status_input, "Querying TestRail Plans ... ")

	Local $plan_id_name = _TestRailGetPlansIDAndNameArray($project_part[0])
	Local $plan_id_str = ""

	for $i = 0 to (UBound($plan_id_name) - 1)

		if StringLen($plan_id_str) > 0 Then

			$plan_id_str = $plan_id_str & "|"
		EndIf

		$plan_id_str = $plan_id_str & $plan_id_name[$i][0] & " - " & $plan_id_name[$i][1]
	Next

	GUICtrlSetData($testrail_plan_combo, $plan_id_str)
	GUICtrlSetData($status_input, "")
	GUICtrlSetState($testrail_plan_combo, $GUI_ENABLE)

EndFunc

Func query_testrail_runs()

	Local $plan_id_name = GUICtrlRead($testrail_plan_combo)
	Local $plan_part = StringSplit($plan_id_name, " - ", 3)

	GUICtrlSetData($status_input, "Querying TestRail Runs ... ")
	Local $run_id = _TestRailGetPlanRunsID($plan_part[0])
	$run_ids = _ArrayToString($run_id)
	GUICtrlSetData($status_input, "")
	GUICtrlSetState($start_button, $GUI_ENABLE)

EndFunc
