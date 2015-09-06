;Matthew1471 WinBlogX

!define MUI_PRODUCT "WinBlogX" ; Name Of The Software
!define MUI_NAME "Matthew1471 WinBlogX Client ${MUI_VERSION}" ;Installer name
!define MUI_VERSION "1.04.14" ; Software Version

!include "MUI.nsh"

;--------------------------------
;Configuration

  ;General
  OutFile "..\..\..\WinBlogX Setup.exe"

  ;Folder selection page
  InstallDir "$PROGRAMFILES\${MUI_PRODUCT}"
  
Function .onInit
  MessageBox MB_YESNO "This will install WinBlogX. Do you wish to continue?" IDYES gogogo
    Abort
  gogogo:
FunctionEnd

;--------------------------------
;Modern UI Configuration

  !define MUI_WELCOMEPAGE
  !define MUI_LICENSEPAGE
  !define MUI_COMPONENTSPAGE
    !define MUI_COMPONENTSPAGE_SMALLDESC
  !define MUI_DIRECTORYPAGE
  !define MUI_FINISHPAGE
    !define MUI_FINISHPAGE_RUN "$INSTDIR\WinBlogX.exe"
    !define MUI_FINISHPAGE_SHOWREADME "$INSTDIR\Docs\Readme.htm"
    !define MUI_FINISHPAGE_SHOWREADME_NOTCHECKED
  
  !define MUI_ABORTWARNING
  
  !define MUI_UNINSTALLER
  !define MUI_UNCONFIRMPAGE
  
  !define MUI_HEADERBITMAP "${NSISDIR}\Contrib\Icons\modern-header.bmp"

  !define MUI_ICON "${NSISDIR}\Contrib\Icons\new_nsis_3.ico"
  !define MUI_UNICON "${NSISDIR}\Contrib\Icons\new_nsis_3.ico"
  
;--------------------------------
;Languages

  !insertmacro MUI_LANGUAGE "Language"

;--------------------------------
;Data
  
  LicenseData "License.txt"

;--------------------------------
;Reserve Files
  
  ;Things that need to be extracted on first (keep these lines before any File command!)
  ;Only useful for BZIP2 compression
  
  ReserveFile "${NSISDIR}\Contrib\Icons\modern-header.bmp"

;--------------------------------
;Installer Sections

Section "!WinBlogX Application" SecCopyUI
  SectionIn RO
  SetOutPath "$INSTDIR"
  File "..\..\WinBlogX.exe"

  SetOutPath "$SYSDIR"
  File "..\Dependencies\*.ocx"
  File "..\Dependencies\*.dll"

  RegDLL "$SYSDIR\MSCOMCTL.ocx"
  RegDLL "$SYSDIR\MSINET.ocx"
  RegDLL "$SYSDIR\msstdfmt.dll"

SectionEnd

Section "Documentation" SecCopyDoc

  SetOutPath "$INSTDIR\Docs"
  File "..\..\Docs\Readme.htm"

  SetOutPath "$INSTDIR\Docs\Images"
  File "..\..\Docs\Images\*.jpg"

SectionEnd

Section "Source Code" SecCopySource
  SetOutPath "$INSTDIR\Source"
  File "..\*.bas"
  File "..\*.frm"
  File "..\*.frx"
  File "..\*.vbp"
  File "..\*.vbw"

  SetOutPath "$INSTDIR\Source\Installer\"
  File "*.nsh"
  File "*.txt"
  File "*.nsi"

  SetOutPath "$INSTDIR\Source\Dependencies\"
  File "..\Dependencies\*.ocx"
  File "..\Dependencies\*.dll"

  SetOutPath "$INSTDIR\Source\Images\"
  File "..\Images\*.gif"
  File "..\Images\*.ico"
  
SectionEnd

Section -post

  SetOutPath "$INSTDIR"
  File "..\..\License.txt"

  ;Create uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  ;Write Start Menu Shortcuts
  CreateDirectory "$SMPROGRAMS\Matthew1471\"
  CreateShortCut "$SMPROGRAMS\Matthew1471\-WinBlogX-.lnk" "$INSTDIR\WinBlogX.exe" "" "" "0" "" "" "Start WinBlogX"
  CreateShortCut "$SMPROGRAMS\Matthew1471\Uninstall.lnk" "$INSTDIR\Uninstall.exe"

  IfFileExists "$INSTDIR\Docs\Readme.htm" "" +2
    CreateShortCut "$SMPROGRAMS\Matthew1471\Readme.lnk" "$INSTDIR\Docs\ReadMe.htm" ""

  IfFileExists "$INSTDIR\Source\*.*" "" +2
    CreateShortCut "$SMPROGRAMS\Matthew1471\Source.lnk" "$INSTDIR\Source\" "" "" "0" "" "" "Files Used To Create The Program"

  ; write uninstall strings
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "Comments" "BlogX is a simple client program to upload to a Matthew1471 ASP BlogX server."
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "Contact" "Matthew1471"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "DisplayName" "Matthew1471 WinBlogX Client"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "DisplayIcon" "$INSTDIR\Uninstall.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "DisplayVersion" "${MUI_VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "HelpLink" "http://matthew1471.co.uk/Contact.asp"
  WriteRegDWord HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "NoModify" 1
  WriteRegDWord HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "NoRepair" 1
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "Publisher" "Matthew1471"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "Readme" "$INSTDIR\DOC\Readme.htm"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "InstallSource" "$EXEDIR\"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "URLInfoAbout" "http://matthew1471.co.uk"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "URLUpdateInfo" "http://matthew1471.co.uk/Blog/"

  IfFileExists "$PROGRAMFILES\Internet Explorer\iexplore.exe" 0 Skip
    
    MessageBox MB_YESNO|MB_ICONQUESTION "Would you like to add a shortcut in Internet Explorer to WinBlogX?" IDNO Skip ; Skipped if file doesn't exist

      WriteRegStr HKLM "Software\Microsoft\Internet Explorer\Extensions\{1FC3B46A-8899-45ED-AFD5-5CE83C381EC2}" "Exec" "$INSTDIR\WinBlogX.exe"
      WriteRegStr HKLM "Software\Microsoft\Internet Explorer\Extensions\{1FC3B46A-8899-45ED-AFD5-5CE83C381EC2}" "Icon" "$INSTDIR\WinBlogX.exe,001"
      WriteRegStr HKLM "Software\Microsoft\Internet Explorer\Extensions\{1FC3B46A-8899-45ED-AFD5-5CE83C381EC2}" "HotIcon" "$INSTDIR\WinBlogX.exe,001"
      WriteRegStr HKLM "Software\Microsoft\Internet Explorer\Extensions\{1FC3B46A-8899-45ED-AFD5-5CE83C381EC2}" "ButtonText" "Matthew1471 WinBlogX!"
      WriteRegStr HKLM "Software\Microsoft\Internet Explorer\Extensions\{1FC3B46A-8899-45ED-AFD5-5CE83C381EC2}" "CLSID" "{1FBA04EE-3024-11d2-8F1F-0000F87ABD16}"
      WriteRegStr HKLM "Software\Microsoft\Internet Explorer\Extensions\{1FC3B46A-8899-45ED-AFD5-5CE83C381EC2}" "Default Visible" "Yes"

    Skip: ; Skipped Section (Either IE not installed, or user said NO)

SectionEnd

;Display the Finish header
;Insert this macro after the sections if you are not using a finish page
!insertmacro MUI_SECTIONS_FINISHHEADER

;--------------------------------
;Descriptions

!insertmacro MUI_FUNCTIONS_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SecCopyUI} "Copy the WinBlogX application."
  !insertmacro MUI_DESCRIPTION_TEXT ${SecCopyDoc} "Copy the WinBlogX documentation."
  !insertmacro MUI_DESCRIPTION_TEXT ${SecCopySource} "Source code, should you want to Re-program WinBlogX."
!insertmacro MUI_FUNCTIONS_DESCRIPTION_END
 
;--------------------------------
;Uninstaller Section

Section "Uninstall"

  Delete "$INSTDIR\License.txt"
  Delete "$INSTDIR\ReadMe.htm"
  Delete "$INSTDIR\Uninstall.exe"
  Delete "$INSTDIR\WinBlog.ini"
  Delete "$INSTDIR\WinBlogX.exe"
  Delete "$INSTDIR\Update.asp"

  RMDir /r "$INSTDIR\Docs\"
  RMDir /r "$INSTDIR\Images\"
  RMDir /r "$INSTDIR\Source\"

  RMDir "$INSTDIR"

  RMDir /r "$SMPROGRAMS\Matthew1471\"

  DeleteRegKey /ifempty HKCU "Software\${MUI_PRODUCT}"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}"
  DeleteRegKey HKLM "Software\Microsoft\Internet Explorer\Extensions\{1FC3B46A-8899-45ED-AFD5-5CE83C381EC2}"

  ;Display the Finish header
  !insertmacro MUI_UNFINISHHEADER

SectionEnd