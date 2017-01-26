; 该脚本使用 HM VNISEdit 脚本编辑器向导产生

; 安装程序初始定义常量
!define PRODUCT_NAME "6000DDM Calitor"
!define PRODUCT_VERSION "3.0"
!define PRODUCT_PUBLISHER "BJHHFA, Inc."
!define PRODUCT_WEB_SITE "http://www.bjhhfa.com"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\Reboot.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"
!define PRODUCT_STARTMENU_REGVAL "NSIS:StartMenuDir"

SetCompressor /SOLID lzma
SetCompressorDictSize 32

; ------ MUI 现代界面定义 (1.67 版本以上兼容) ------
!include "MUI.nsh"

; MUI 预定义常量
!define MUI_ABORTWARNING
!define MUI_ICON "..\IntuiLinkMM.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

; 语言选择窗口常量设置
!define MUI_LANGDLL_REGISTRY_ROOT "${PRODUCT_UNINST_ROOT_KEY}"
!define MUI_LANGDLL_REGISTRY_KEY "${PRODUCT_UNINST_KEY}"
!define MUI_LANGDLL_REGISTRY_VALUENAME "NSIS:Language"

; 欢迎页面
!insertmacro MUI_PAGE_WELCOME
; 组件选择页面
!insertmacro MUI_PAGE_COMPONENTS
; 安装目录选择页面
!insertmacro MUI_PAGE_DIRECTORY
; 开始菜单设置页面
var ICONS_GROUP
!define MUI_STARTMENUPAGE_NODISABLE
!define MUI_STARTMENUPAGE_DEFAULTFOLDER "6000DDM Calitor"
!define MUI_STARTMENUPAGE_REGISTRY_ROOT "${PRODUCT_UNINST_ROOT_KEY}"
!define MUI_STARTMENUPAGE_REGISTRY_KEY "${PRODUCT_UNINST_KEY}"
!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "${PRODUCT_STARTMENU_REGVAL}"
!insertmacro MUI_PAGE_STARTMENU Application $ICONS_GROUP
; 安装过程页面
!insertmacro MUI_PAGE_INSTFILES
; 安装完成页面
!define MUI_FINISHPAGE_RUN "$INSTDIR\6000DDM Calitor.exe"
!insertmacro MUI_PAGE_FINISH

; 安装卸载过程页面
!insertmacro MUI_UNPAGE_INSTFILES

; 安装界面包含的语言设置
!insertmacro MUI_LANGUAGE "English"
!insertmacro MUI_LANGUAGE "SimpChinese"

; 安装预释放文件
!insertmacro MUI_RESERVEFILE_LANGDLL
!insertmacro MUI_RESERVEFILE_INSTALLOPTIONS
; ------ MUI 现代界面定义结束 ------

Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "..\6000DDM Calitor Setup.exe"
InstallDir "$PROGRAMFILES\6000DDM Calitor"
InstallDirRegKey HKLM "${PRODUCT_UNINST_KEY}" "UninstallString"
ShowInstDetails show
ShowUnInstDetails show
BrandingText "6000DDM Calitor"

Section "Main Program" SEC01
  SetOutPath "$INSTDIR"
  SetOverwrite ifnewer
  File "..\Reboot.exe"
  File "..\6000DDM Calitor.exe"
  File "..\IntuiLinkMM.ico"
  File "..\TIPOFDAY.TXT"

; 创建开始菜单快捷方式
  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
  CreateDirectory "$SMPROGRAMS\$ICONS_GROUP"
  CreateShortCut "$SMPROGRAMS\$ICONS_GROUP\6000DDM Calitor.lnk" "$INSTDIR\6000DDM Calitor.exe"
  CreateShortCut "$DESKTOP\6000DDM Calitor.lnk" "$INSTDIR\6000DDM Calitor.exe"
  !insertmacro MUI_STARTMENU_WRITE_END
SectionEnd

Section "Default Setting" SEC02
  SetOutPath "$INSTDIR\Configuration"
  SetOverwrite on
  File "..\Configuration\修正系数.txt"
  File "..\Configuration\软件通用设置.txt"
  File "..\Configuration\量程段编号设置.txt"
  File "..\Configuration\按键代码.txt"
  File "..\Configuration\RS232端口设置.txt"
  File "..\Configuration\LAN设置.txt"
  SetOutPath "$INSTDIR"
  File "..\Config.ini"

; 创建开始菜单快捷方式
  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
  !insertmacro MUI_STARTMENU_WRITE_END
SectionEnd

Section "Skin" SEC03
  File "..\SkinH_VB6.dll"
  SetOutPath "$INSTDIR\Themes"
  File "..\Themes\炫绿.she"
  File "..\Themes\积木.she"
  File "..\Themes\xmp.she"
  File "..\Themes\Xenes.she"
  File "..\Themes\wish.she"
  File "..\Themes\whitefire.she"
  File "..\Themes\vista.she"
  File "..\Themes\storm.she"
  File "..\Themes\royale.she"
  File "..\Themes\REAL.she"
  File "..\Themes\QQ影音.she"
  File "..\Themes\qqgame.she"
  File "..\Themes\QQ2009_窄_底边.she"
  File "..\Themes\QQ2009_宽_底边.she"
  File "..\Themes\QQ2009.she"
  File "..\Themes\qq2008.she"
  File "..\Themes\pixos.she"
  File "..\Themes\ouframe.she"
  File "..\Themes\office2007.she"
  File "..\Themes\MSN.she"
  File "..\Themes\longhorn.she"
  File "..\Themes\itunes.she"
  File "..\Themes\insomnia.she"
  File "..\Themes\homestead.she"
  File "..\Themes\hlong.she"
  File "..\Themes\gem.she"
  File "..\Themes\enjoy.she"
  File "..\Themes\elegance.she"
  File "..\Themes\dogmax.she"
  File "..\Themes\darkroyale.she"
  File "..\Themes\compact.she"
  File "..\Themes\china.she"
  File "..\Themes\black.she"
  File "..\Themes\asus.she"
  File "..\Themes\aero.she"
  File "..\Themes\adamant.she"

; 创建开始菜单快捷方式
  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
  !insertmacro MUI_STARTMENU_WRITE_END
SectionEnd

Section "Support Dlls" SEC04
  SetOutPath "$SYSDIR"
  SetOverwrite off
  File "C:\WINDOWS\system32\CMDLGCHS.DLL"
  File "C:\WINDOWS\system32\comdlg32.ocx"
  File "C:\WINDOWS\system32\MSCMCCHS.DLL"
  File "C:\WINDOWS\system32\MSCOMCHS.DLL"
  File "C:\WINDOWS\system32\MSCOMCTL.OCX"
  File "C:\WINDOWS\system32\MSCOMM32.OCX"
  File "C:\WINDOWS\system32\RCHTXCHS.DLL"
  File "C:\WINDOWS\system32\RICHTX32.OCX"
  File "C:\WINDOWS\system32\vb6chs.dll"
  File "C:\WINDOWS\system32\ws2_32.dll"
  File "C:\WINDOWS\system32\stdole2.tlb"

; 创建开始菜单快捷方式
  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
  !insertmacro MUI_STARTMENU_WRITE_END
SectionEnd

Section -AdditionalIcons
  SetOutPath $INSTDIR
  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
  WriteIniStr "$INSTDIR\${PRODUCT_NAME}.url" "InternetShortcut" "URL" "${PRODUCT_WEB_SITE}"
  CreateShortCut "$SMPROGRAMS\$ICONS_GROUP\Website.lnk" "$INSTDIR\${PRODUCT_NAME}.url"
  CreateShortCut "$SMPROGRAMS\$ICONS_GROUP\Uninstall.lnk" "$INSTDIR\uninst.exe"
  !insertmacro MUI_STARTMENU_WRITE_END
SectionEnd

Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR\6000DDM Calitor.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\6000DDM Calitor.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"
SectionEnd

#-- 根据 NSIS 脚本编辑规则，所有 Function 区段必须放置在 Section 区段之后编写，以避免安装程序出现未可预知的问题。--#

; 区段组件描述
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC01} "程序运行必须组件"
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC02} "默认设定，建议安装"
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC03} "皮肤组件"
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC04} "运行VB程序所必须的支持库文件，建议安装"
!insertmacro MUI_FUNCTION_DESCRIPTION_END

Function .onInit
  !insertmacro MUI_LANGDLL_DISPLAY
FunctionEnd

/******************************
 *  以下是安装程序的卸载部分  *
 ******************************/

Section Uninstall
  !insertmacro MUI_STARTMENU_GETFOLDER "Application" $ICONS_GROUP
  Delete "$INSTDIR\${PRODUCT_NAME}.url"
  Delete "$INSTDIR\uninst.exe"
  Delete "$INSTDIR\Themes\adamant.she"
  Delete "$INSTDIR\Themes\aero.she"
  Delete "$INSTDIR\Themes\asus.she"
  Delete "$INSTDIR\Themes\black.she"
  Delete "$INSTDIR\Themes\china.she"
  Delete "$INSTDIR\Themes\compact.she"
  Delete "$INSTDIR\Themes\darkroyale.she"
  Delete "$INSTDIR\Themes\dogmax.she"
  Delete "$INSTDIR\Themes\elegance.she"
  Delete "$INSTDIR\Themes\enjoy.she"
  Delete "$INSTDIR\Themes\gem.she"
  Delete "$INSTDIR\Themes\hlong.she"
  Delete "$INSTDIR\Themes\homestead.she"
  Delete "$INSTDIR\Themes\insomnia.she"
  Delete "$INSTDIR\Themes\itunes.she"
  Delete "$INSTDIR\Themes\longhorn.she"
  Delete "$INSTDIR\Themes\MSN.she"
  Delete "$INSTDIR\Themes\office2007.she"
  Delete "$INSTDIR\Themes\ouframe.she"
  Delete "$INSTDIR\Themes\pixos.she"
  Delete "$INSTDIR\Themes\qq2008.she"
  Delete "$INSTDIR\Themes\QQ2009.she"
  Delete "$INSTDIR\Themes\QQ2009_宽_底边.she"
  Delete "$INSTDIR\Themes\QQ2009_窄_底边.she"
  Delete "$INSTDIR\Themes\qqgame.she"
  Delete "$INSTDIR\Themes\QQ影音.she"
  Delete "$INSTDIR\Themes\REAL.she"
  Delete "$INSTDIR\Themes\royale.she"
  Delete "$INSTDIR\Themes\storm.she"
  Delete "$INSTDIR\Themes\vista.she"
  Delete "$INSTDIR\Themes\whitefire.she"
  Delete "$INSTDIR\Themes\wish.she"
  Delete "$INSTDIR\Themes\Xenes.she"
  Delete "$INSTDIR\Themes\xmp.she"
  Delete "$INSTDIR\Themes\积木.she"
  Delete "$INSTDIR\Themes\炫绿.she"
  Delete "$INSTDIR\SkinH_VB6.dll"
  Delete "$INSTDIR\Config.ini"
  Delete "$INSTDIR\Configuration\LAN设置.txt"
  Delete "$INSTDIR\Configuration\RS232端口设置.txt"
  Delete "$INSTDIR\Configuration\按键代码.txt"
  Delete "$INSTDIR\Configuration\量程段编号设置.txt"
  Delete "$INSTDIR\Configuration\软件通用设置.txt"
  Delete "$INSTDIR\Configuration\修正系数.txt"
  Delete "$INSTDIR\6000DDM Calitor.exe"
  Delete "$INSTDIR\Reboot.exe"
  Delete "$INSTDIR\IntuiLinkMM.ico"
  Delete "$INSTDIR\TIPOFDAY.TXT"

  Delete "$SMPROGRAMS\$ICONS_GROUP\Uninstall.lnk"
  Delete "$SMPROGRAMS\$ICONS_GROUP\Website.lnk"
  Delete "$DESKTOP\6000DDM Calitor.lnk"
  Delete "$SMPROGRAMS\$ICONS_GROUP\6000DDM Calitor.lnk"

  RMDir "$SMPROGRAMS\$ICONS_GROUP"
  RMDir "$INSTDIR\Themes"
  RMDir "$INSTDIR\Configuration"

  RMDir "$INSTDIR"

  DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
  DeleteRegKey HKLM "${PRODUCT_DIR_REGKEY}"
  SetAutoClose true
SectionEnd

#-- 根据 NSIS 脚本编辑规则，所有 Function 区段必须放置在 Section 区段之后编写，以避免安装程序出现未可预知的问题。--#

Function un.onInit
!insertmacro MUI_UNGETLANGUAGE
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "您确实要完全移除 $(^Name) ，及其所有的组件？" IDYES +2
  Abort
FunctionEnd

Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name) 已成功地从您的计算机移除。"
FunctionEnd
