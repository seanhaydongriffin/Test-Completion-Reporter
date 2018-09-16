#include-once
#Include <Array.au3>
#include <GuiEdit.au3>
#include "cURL.au3"
#Region Header
#cs
	Title:   		Janison Insights Automation UDF Library for AutoIt3
	Filename:  		JanisonInsights.au3
	Description: 	A collection of functions for creating, attaching to, reading from and manipulating Janison Insights
	Author:   		seangriffin
	Version:  		V0.1
	Last Update: 	25/02/18
	Requirements: 	AutoIt3 3.2 or higher,
					Janison Insights Release x.xx,
					cURL xxx
	Changelog:		---------24/12/08---------- v0.1
					Initial release.
#ce
#EndRegion Header
#Region Global Variables and Constants
Global $jira_domain = ""
Global $jira_username = ""
Global $jira_password = ""
Global $jira_json = ""
Global $jira_html = ""
#EndRegion Global Variables and Constants
#Region Core functions
; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsSetup()
; Description ...:	Setup activities including cURL initialization.
; Syntax.........:	_InsightsSetup()
; Parameters ....:
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _JiraSetup()

	; Initialise cURL
	cURL_initialise()


EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsShutdown()
; Description ...:	Setup activities including cURL initialization.
; Syntax.........:	_InsightsShutdown()
; Parameters ....:
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _JiraShutdown()

	; Clean up cURL
	cURL_cleanup()

EndFunc


; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsDomainSet()
; Description ...:	Sets the domain to use in all other functions.
; Syntax.........:	_InsightsDomainSet($domain)
; Parameters ....:	$win_title			- Optional: The title of the SAP window (within the session) to attach to.
;											The window "SAP Easy Access" is used if one isn't provided.
;											This may be a substring of the full window title.
;					$sap_transaction	- Optional: a SAP transaction to run after attaching to the session.
;											A "/n" will be inserted at the beginning of the transaction
;											if one isn't provided.
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _JiraDomainSet($domain)

	$jira_domain = $domain
EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsLogin()
; Description ...:	Login a user to Janison Insights.
; Syntax.........:	_InsightsLogin($username, $password)
; Parameters ....:	$win_title			- Optional: The title of the SAP window (within the session) to attach to.
;											The window "SAP Easy Access" is used if one isn't provided.
;											This may be a substring of the full window title.
;					$sap_transaction	- Optional: a SAP transaction to run after attaching to the session.
;											A "/n" will be inserted at the beginning of the transaction
;											if one isn't provided.
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _JiraLogin($username, $password)

	$jira_username = $username
	$jira_password = $password
EndFunc


; Search

Func _JiraSearchIssues($fields, $jql)

	Local $jql_encoded = EncodeUrl($jql)
	$response = cURL_easy($jira_domain & "/rest/api/2/search?fields=" & $fields & "&jql=" & $jql_encoded, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $jira_username & ":" & $jira_password)
;	$response = cURL_easy($jira_domain & "/rest/api/2/search?jql=" & $jql_encoded, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $jira_username & ":" & $jira_password)
	$jira_json = $response[2]
EndFunc


Func _JiraGetSearchResultKeysAndIssueTypeNames($fields, $jql)

	_JiraSearchIssues($fields, $jql)

	$rr = StringRegExp($jira_json, '(?U)"key":"(.*)".*"name":"(.*)"', 3)

	Return $rr

EndFunc

Func _JiraGetSearchResultKeysSummariesAndIssueTypeNames($fields, $jql)

	_JiraSearchIssues($fields, $jql)


	; "key":"SEAB-4681","fields":{"summary":"1.13 - Regression Requirements - upload organisation units to the tenant","issuetype":

	$rr = StringRegExp($jira_json, '(?U)"key":"(.*)".*"summary":"(.*)".*"name":"(.*)"', 3)

	Return $rr

EndFunc

Func _JiraBrowseIssue($key)

	$response = cURL_easy($jira_domain & "/browse/" & $key, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $jira_username & ":" & $jira_password)
	$jira_html = $response[2]
EndFunc



Func _JiraGetTestRailTestCasesFromIssue($key)

	;_JiraBrowseIssue($key)

;	Local $testrail_url = StringRegExp($jira_html, '(?U)"url":"(https://jira.testrail.net/issues/references.*)"', 3)

;	_TestRailGetTestCases($key) ;$testrail_url[0])


;	$rr = StringRegExp($jira_json, '(?U)"key":"(.*)"', 3)

;	Return $rr

EndFunc




Func EncodeUrl($src)
    Local $i
    Local $ch
    Local $NewChr
    Local $buff

    ;Init Counter
    $i = 1

    While ($i <= StringLen($src))
        ;Get byte code from string
        $ch = Asc(StringMid($src, $i, 1))

        ;Look for what bytes we have
        Switch $ch
            ;Looks ok here
            Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122, 126
                $buff &= Chr($ch)
                ;Space found
            Case 32
                $buff &= "+"
            Case Else
                ;Convert $ch to hexidecimal
                $buff &= "%" & Hex($ch, 2)
        EndSwitch
        ;INC Counter
        $i += 1
    WEnd

    Return $buff
EndFunc   ;==>EncodeUrl

Func cURL_easy_retry($url, $cookie_file = "", $cookie_action = 0, $output_type = 0, $output_file = "", $request_headers = "", $request_data = "", $ssl_verifypeer = 0, $noprogress = 1, $followlocation = 0, $num_of_retries = 10)

	Local $response

	for $i = 1 to $num_of_retries

		$response = cURL_easy($url, $cookie_file, $cookie_action, $output_type, $output_file, $request_headers, $request_data, $ssl_verifypeer, $noprogress, $followlocation)

		if $response[0] <> 500 Then

			Return $response
		EndIf

		ConsoleWrite("url failed with response code " & $response[0] & " - " & $url & @CRLF)
		Sleep(500)
	Next
EndFunc

