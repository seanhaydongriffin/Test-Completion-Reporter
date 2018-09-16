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

Global $app_name = "JTR - Data Extractor"


;#cs
Global $testrail_username = $CmdLine[1]
Global $testrail_password = $CmdLine[2]
Global $testrail_run_ids = $CmdLine[3]
Global $jira_username = $CmdLine[4]
Global $jira_password = $CmdLine[5]
Global $jira_epic_key = $CmdLine[6]
;#ce

#cs
Global $testrail_username = $CmdLine[1]
Global $testrail_password = $CmdLine[2]
Global $testrail_plan = $CmdLine[4]
Global $jira_username = $CmdLine[5]
Global $jira_password = $CmdLine[6]
Global $jira_epic_key = $CmdLine[7]
#ce

Global $main_gui = GUICreate($app_name & " - Epic Key " & $jira_epic_key, 1200, 680)

Global $listview = GUICtrlCreateListView("Epic Key|Issue Key|Issue Summary|Issue Type|Test Project Name|Test Case ID|Test Case Name|Test Case Result", 10, 120, 1020, 480, $LVS_SHOWSELALWAYS)
_GUICtrlListView_SetColumnWidth(-1, 0, 80)
_GUICtrlListView_SetColumnWidth(-1, 1, 80)
_GUICtrlListView_SetColumnWidth(-1, 2, 200)
_GUICtrlListView_SetColumnWidth(-1, 3, 80)
_GUICtrlListView_SetColumnWidth(-1, 4, 140)
_GUICtrlListView_SetColumnWidth(-1, 5, 80)
_GUICtrlListView_SetColumnWidth(-1, 6, 500)
_GUICtrlListView_SetColumnWidth(-1, 7, 100)
_GUICtrlListView_SetExtendedListViewStyle($listview, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES))

Global $status_input = GUICtrlCreateInput("", 10, 680 - 25, 400, 20, $ES_READONLY, $WS_EX_STATICEDGE)
Global $progress = GUICtrlCreateProgress(420, 680 - 25, 610, 20)


GUISetState(@SW_SHOW, $main_gui)


; Startup SQLite

_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)
_SQLite_Open(@ScriptDir & "\Jira TestRail Reporter.sqlite")

; Startup TestRail

GUICtrlSetData($status_input, "Starting the TestRail connection ... ")
_TestRailDomainSet("https://janison.testrail.com")
_TestRailLogin($testrail_username, $testrail_password)

; Authentication

GUICtrlSetData($status_input, "Authenticating against TestRail ... ")
_TestRailAuth()


GUICtrlSetData($status_input, "Querying TestRail Projects ... ")

Local $project_id_name = _TestRailGetProjectsIDAndNameArray()
Local $project_id_str = ""

for $i = 0 to (UBound($project_id_name) - 1)

	if StringLen($project_id_str) > 0 Then

		$project_id_str = $project_id_str & "|"
	EndIf

	$project_id_str = $project_id_str & $project_id_name[$i][0] & " - " & $project_id_name[$i][1]
Next

GUICtrlSetData($progress, 0)
GUISetCursor(15, 1, $main_gui)
_GUICtrlListView_DeleteAllItems($listview)
GUICtrlSetData($status_input, "")

; Startup Jira

GUICtrlSetData($status_input, "Starting the Jira connection ... ")
_JiraSetup()
_JiraDomainSet("https://janisoncls.atlassian.net")
_JiraLogin($jira_username, $jira_password)

Local $testrail_test_case_id = ""
Local $testrail_test_case_name = ""

GUICtrlSetData($status_input, "Querying Epic " & $jira_epic_key & " ... ")
$issue = _JiraGetSearchResultKeysSummariesAndIssueTypeNames("summary,issuetype", """Epic Link"" = " & $jira_epic_key & " OR parent in (""" & $jira_epic_key & """)")
GUICtrlSetData($progress, 10)


; Loop until the user exits.
;While 1

	; GUI msg loop...
	$msg = GUIGetMsg()

;	Switch $msg

;		Case $GUI_EVENT_CLOSE

;			ExitLoop
;	EndSwitch

	for $i = 0 to (UBound($issue) - 1) Step 3

		GUICtrlSetData($status_input, "Querying Issue " & $issue[$i] & " ... ")

		_TestRailGetTestCases($issue[$i])

		$testrail_test_case_id = StringRegExp($testrail_html, '(?U)(?s)<div class="grid-column text-ppp".*>(.*)<', 3)
		$testrail_test_case_and_project_name = StringRegExp($testrail_html, '(?U)(?s)<a href=.*>(.*)<', 3)

		if IsArray($testrail_test_case_id) = True Then

			;_ArrayDisplay($tmp)

			Local $testrail_test_case_and_project_name_index = 0

			for $j = 0 to (UBound($testrail_test_case_id) - 1)

				$tmp_testrail_test_case_id = StringStripWS($testrail_test_case_id[$j], 3)
				$tmp_testrail_test_case_name = StringStripWS($testrail_test_case_and_project_name[$testrail_test_case_and_project_name_index], 3)
				$testrail_test_case_and_project_name_index = $testrail_test_case_and_project_name_index + 1
				$tmp_testrail_project_name = StringStripWS($testrail_test_case_and_project_name[$testrail_test_case_and_project_name_index], 3)
				$testrail_test_case_and_project_name_index = $testrail_test_case_and_project_name_index + 1
				$tmp_testrail_result = "Untested"
				Local $run_id = StringSplit($testrail_run_ids, "|", 2)

				for $k = 0 to (UBound($run_id) - 1)

					_TestRailGetResultsForCase($run_id[$k], StringReplace($tmp_testrail_test_case_id, "C", ""))

					if StringLen($testrail_json) > 9 and StringInStr($testrail_json, "No (active) test found for the run") < 1 Then

						$rr = StringRegExp($testrail_json, '"status_id":(\d+),', 1)
						$tmp_testrail_result = $rr[0]

						Switch $tmp_testrail_result

							case "1"

								$tmp_testrail_result = "Passed"

							case "2"

								$tmp_testrail_result = "Blocked"

							case "4"

								$tmp_testrail_result = "Retest"

							case "5"

								$tmp_testrail_result = "Failed"

							case "6"

								$tmp_testrail_result = "Known Issue"

							case "7"

								$tmp_testrail_result = "Incomplete Automation"

							case "8"

								$tmp_testrail_result = "Removed from Run"

							case "9"

								$tmp_testrail_result = "In Progress"
						EndSwitch

						ExitLoop
					EndIf
				Next

				GUICtrlCreateListViewItem($jira_epic_key & "|" & $issue[$i] & "|" & $issue[$i + 1] & "|" & $issue[$i + 2] & "|" & $tmp_testrail_project_name & "|" & $tmp_testrail_test_case_id & "|" & $tmp_testrail_test_case_name & "|" & $tmp_testrail_result, $listview)

;				if StringCompare($tmp_testrail_result, "Untested") = 0 Then

;					$tmp_testrail_result = "null"
;				Else

					$tmp_testrail_result = "'" & $tmp_testrail_result & "'"
;				EndIf

				_SQLite_Exec(-1, "INSERT INTO report(EpicKey,IssueKey,IssueSummary,IssueType,TestProjectName,TestCaseID,TestCaseName,TestCaseResult,DefectID,DefectDetails,Reason,ImpactActionRequired) VALUES ('" & $jira_epic_key & "','" & $issue[$i] & "','" & $issue[$i + 1] & "','" & $issue[$i + 2] & "','" & $tmp_testrail_project_name & "','" & $tmp_testrail_test_case_id & "','" & $tmp_testrail_test_case_name & "'," & $tmp_testrail_result & ",'<i>TODO</i>','<i>TODO</i>','<i>TODO</i>','<i>TODO</i>');") ; INSERT Data
			Next
		Else

			$tmp_testrail_project_name = "-"
			$tmp_testrail_test_case_id = "-"
			$tmp_testrail_test_case_name = "-"
			GUICtrlCreateListViewItem($jira_epic_key & "|" & $issue[$i] & "|" & $issue[$i + 1] & "|" & $issue[$i + 2] & "||||Untested", $listview)
			_SQLite_Exec(-1, "INSERT INTO report(EpicKey,IssueKey,IssueSummary,IssueType,TestProjectName,TestCaseID,TestCaseName,TestCaseResult,DefectID,DefectDetails,Reason,ImpactActionRequired) VALUES ('" & $jira_epic_key & "','" & $issue[$i] & "','" & $issue[$i + 1] & "','" & $issue[$i + 2] & "',null,null,null,null,null,null,null,null);") ; INSERT Data
		EndIf

		GUICtrlSetData($progress, 10 + ((($i + 1) / UBound($issue)) * 90))

	Next

	GUICtrlSetData($progress, 0)
	GUICtrlSetData($status_input, "")
	GUISetCursor(2, 0, $main_gui)



;WEnd

GUIDelete($main_gui)

; Shutdown TestRail

GUICtrlSetData($status_input, "Closing Jira ... ")
_JiraShutdown()



