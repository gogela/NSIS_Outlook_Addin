              
############################################################################################
#      Setup for Outlook Addin by J.Gogela Oct 2013
#
# 	Base part created using NSIS Quick Setup Script Generator v1.09.18
#############################################################################################

!define APP_NAME "OutlookAddin"
!define COMP_NAME "samplecomp"
!define WEB_SITE "http://www.samplecomp.com"
!define VERSION "1.00.00.00"
!define COPYRIGHT ""
!define DESCRIPTION "outlook plugin"
!define INSTALLER_NAME "C:\nsis\Output\OutlookAddin\setup.exe"
!define INSTALL_TYPE "SetShellVarContext current"
!define REG_ROOT "HKCU"
!define UNINSTALL_PATH "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"

######################################################################

VIProductVersion  "${VERSION}"
VIAddVersionKey "ProductName"  "${APP_NAME}"
VIAddVersionKey "CompanyName"  "${COMP_NAME}"
VIAddVersionKey "LegalCopyright"  "${COPYRIGHT}"
VIAddVersionKey "FileDescription"  "${DESCRIPTION}"
VIAddVersionKey "FileVersion"  "${VERSION}"

######################################################################

SetCompressor ZLIB
Name "${APP_NAME}"
Caption "${APP_NAME}"
OutFile "${INSTALLER_NAME}"
BrandingText "${APP_NAME}"
XPStyle on

InstallDir "$PROGRAMFILES\OutlookAddin"

######################################################################



######################################################################

AutoCloseWindow true

Section -MainProgram
${INSTALL_TYPE}
SetOverwrite ifnewer
SetOutPath "$INSTDIR"

#######################################################################
# install files from Visual studio project release
#######################################################################
File "C:\nsis\infiles\OutlookAddIn.dll"
File "C:\infiles\OutlookAddIn.dll.config"
File "C:\infiles\OutlookAddIn.dll.manifest"
File "C:\nsis\infiles\OutlookAddIn.vsto"
File "C:\nsis\infiles\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"
File "C:\nsis\infiles\Microsoft.Office.Tools.dll"
File "C:\nsis\infiles\Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll"
File "C:\nsis\infiles\Microsoft.Office.Tools.v4.0.Framework.dll"
SectionEnd

Section -Reg
#######################################################################
# create registry keys - described here: 
# http://msdn.microsoft.com/en-us/library/office/bb206787 
# http://msdn.microsoft.com/en-us/library/vstudio/bb386106.aspx
#######################################################################

# form region - required only if you use some
WriteRegStr HKCU "Software\Microsoft\Office\Outlook\FormRegions\IPM.Note.myform" "OutlookAddIn.FormRegion2" "=OutlookAddIn"

# the add-in itself
WriteRegStr HKCU "Software\Microsoft\Office\Outlook\Addins\OutlookAddIn" "Description" "Adds additional functionality MS Outlook"
WriteRegStr HKCU "Software\Microsoft\Office\Outlook\Addins\OutlookAddIn" "FriendlyName" "OutlookAddIn" 
WriteRegDWORD HKCU "Software\Microsoft\Office\Outlook\Addins\OutlookAddIn" "LoadBehavior" 0x00000003
WriteRegStr HKCU "Software\Microsoft\Office\Outlook\Addins\OutlookAddIn" "Manifest" "file:///$INSTDIR\OutlookAddIn.vsto|vstolocal"


SectionEnd



######################################################################

Section -Icons_Reg
SetOutPath "$INSTDIR"
WriteUninstaller "$INSTDIR\uninstall.exe"

WriteRegStr ${REG_ROOT} "${UNINSTALL_PATH}"  "DisplayName" "${APP_NAME}"
WriteRegStr ${REG_ROOT} "${UNINSTALL_PATH}"  "UninstallString" "$INSTDIR\uninstall.exe"
WriteRegStr ${REG_ROOT} "${UNINSTALL_PATH}"  "DisplayVersion" "${VERSION}"
WriteRegStr ${REG_ROOT} "${UNINSTALL_PATH}"  "Publisher" "${COMP_NAME}"

!ifdef WEB_SITE
WriteRegStr ${REG_ROOT} "${UNINSTALL_PATH}"  "URLInfoAbout" "${WEB_SITE}"
!endif
SectionEnd

######################################################################

Section Uninstall
${INSTALL_TYPE}

#clean registry
DeleteRegKey HKCU "Software\Microsoft\Office\Outlook\FormRegions\IPM.Note.myform" 
DeleteRegKey HKCU "Software\Microsoft\Office\Outlook\Addins\OutlookAddIn"


Delete "$INSTDIR\OutlookAddIn.dll"
Delete "$INSTDIR\OutlookAddIn.dll.config"
Delete "$INSTDIR\OutlookAddIn.dll.manifest"
Delete "$INSTDIR\OutlookAddIn.vsto"
Delete "$INSTDIR\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"
Delete "$INSTDIR\Microsoft.Office.Tools.dll"
Delete "$INSTDIR\Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll"
Delete "$INSTDIR\Microsoft.Office.Tools.v4.0.Framework.dll"
Delete "$INSTDIR\uninstall.exe"
!ifdef WEB_SITE
Delete "$INSTDIR\${APP_NAME} website.url"
!endif

RmDir "$INSTDIR"

DeleteRegKey ${REG_ROOT} "${UNINSTALL_PATH}"
SectionEnd

######################################################################
