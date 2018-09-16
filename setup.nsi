;!include nsDialogs.nsh
;!include LogicLib.nsh


; example1.nsi
;
; This script is perhaps one of the simplest NSIs you can make. All of the
; optional settings are left to their default settings. The installer simply 
; prompts the user asking them where to install, and drops a copy of example1.nsi
; there. 

XPStyle on

;--------------------------------

; The name of the installer
Name "Test Completion Reporter"

; The file to write
OutFile "setup.exe"

; The default installation directory
InstallDir "C:\Test Completion Reporter"

; Request application privileges for Windows Vista
RequestExecutionLevel user

;--------------------------------


; Pages

Page directory
Page instfiles


;--------------------------------


; The stuff to install
Section "" ;No components page, name is not important

  ; Set output path to the installation directory.
  SetOutPath $INSTDIR
  
  ; Put file there
  File "Test Completion Reporter.exe"
  File "data_extractor.exe"
  File "curl.exe"
  File *.dll

  CreateDirectory "$SMPROGRAMS\Test Completion Reporter"
  CreateShortCut "$SMPROGRAMS\Test Completion Reporter\Test Completion Reporter.lnk" "$INSTDIR\Test Completion Reporter.exe"

SectionEnd ; end the section
