; �ýű�ʹ�� HM VNISEdit �ű��༭���򵼲���

; ��װ�����ʼ���峣��
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

; ------ MUI �ִ����涨�� (1.67 �汾���ϼ���) ------
!include "MUI.nsh"

; MUI Ԥ���峣��
!define MUI_ABORTWARNING
!define MUI_ICON "..\IntuiLinkMM.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

; ����ѡ�񴰿ڳ�������
!define MUI_LANGDLL_REGISTRY_ROOT "${PRODUCT_UNINST_ROOT_KEY}"
!define MUI_LANGDLL_REGISTRY_KEY "${PRODUCT_UNINST_KEY}"
!define MUI_LANGDLL_REGISTRY_VALUENAME "NSIS:Language"

; ��ӭҳ��
!insertmacro MUI_PAGE_WELCOME
; ���ѡ��ҳ��
!insertmacro MUI_PAGE_COMPONENTS
; ��װĿ¼ѡ��ҳ��
!insertmacro MUI_PAGE_DIRECTORY
; ��ʼ�˵�����ҳ��
var ICONS_GROUP
!define MUI_STARTMENUPAGE_NODISABLE
!define MUI_STARTMENUPAGE_DEFAULTFOLDER "6000DDM Calitor"
!define MUI_STARTMENUPAGE_REGISTRY_ROOT "${PRODUCT_UNINST_ROOT_KEY}"
!define MUI_STARTMENUPAGE_REGISTRY_KEY "${PRODUCT_UNINST_KEY}"
!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "${PRODUCT_STARTMENU_REGVAL}"
!insertmacro MUI_PAGE_STARTMENU Application $ICONS_GROUP
; ��װ����ҳ��
!insertmacro MUI_PAGE_INSTFILES
; ��װ���ҳ��
!define MUI_FINISHPAGE_RUN "$INSTDIR\6000DDM Calitor.exe"
!insertmacro MUI_PAGE_FINISH

; ��װж�ع���ҳ��
!insertmacro MUI_UNPAGE_INSTFILES

; ��װ�����������������
!insertmacro MUI_LANGUAGE "English"
!insertmacro MUI_LANGUAGE "SimpChinese"

; ��װԤ�ͷ��ļ�
!insertmacro MUI_RESERVEFILE_LANGDLL
!insertmacro MUI_RESERVEFILE_INSTALLOPTIONS
; ------ MUI �ִ����涨����� ------

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

; ������ʼ�˵���ݷ�ʽ
  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
  CreateDirectory "$SMPROGRAMS\$ICONS_GROUP"
  CreateShortCut "$SMPROGRAMS\$ICONS_GROUP\6000DDM Calitor.lnk" "$INSTDIR\6000DDM Calitor.exe"
  CreateShortCut "$DESKTOP\6000DDM Calitor.lnk" "$INSTDIR\6000DDM Calitor.exe"
  !insertmacro MUI_STARTMENU_WRITE_END
SectionEnd

Section "Default Setting" SEC02
  SetOutPath "$INSTDIR\Configuration"
  SetOverwrite on
  File "..\Configuration\����ϵ��.txt"
  File "..\Configuration\���ͨ������.txt"
  File "..\Configuration\���̶α������.txt"
  File "..\Configuration\��������.txt"
  File "..\Configuration\RS232�˿�����.txt"
  File "..\Configuration\LAN����.txt"
  SetOutPath "$INSTDIR"
  File "..\Config.ini"

; ������ʼ�˵���ݷ�ʽ
  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
  !insertmacro MUI_STARTMENU_WRITE_END
SectionEnd

Section "Skin" SEC03
  File "..\SkinH_VB6.dll"
  SetOutPath "$INSTDIR\Themes"
  File "..\Themes\����.she"
  File "..\Themes\��ľ.she"
  File "..\Themes\xmp.she"
  File "..\Themes\Xenes.she"
  File "..\Themes\wish.she"
  File "..\Themes\whitefire.she"
  File "..\Themes\vista.she"
  File "..\Themes\storm.she"
  File "..\Themes\royale.she"
  File "..\Themes\REAL.she"
  File "..\Themes\QQӰ��.she"
  File "..\Themes\qqgame.she"
  File "..\Themes\QQ2009_խ_�ױ�.she"
  File "..\Themes\QQ2009_��_�ױ�.she"
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

; ������ʼ�˵���ݷ�ʽ
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

; ������ʼ�˵���ݷ�ʽ
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

#-- ���� NSIS �ű��༭�������� Function ���α�������� Section ����֮���д���Ա��ⰲװ�������δ��Ԥ֪�����⡣--#

; �����������
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC01} "�������б������"
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC02} "Ĭ���趨�����鰲װ"
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC03} "Ƥ�����"
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC04} "����VB�����������֧�ֿ��ļ������鰲װ"
!insertmacro MUI_FUNCTION_DESCRIPTION_END

Function .onInit
  !insertmacro MUI_LANGDLL_DISPLAY
FunctionEnd

/******************************
 *  �����ǰ�װ�����ж�ز���  *
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
  Delete "$INSTDIR\Themes\QQ2009_��_�ױ�.she"
  Delete "$INSTDIR\Themes\QQ2009_խ_�ױ�.she"
  Delete "$INSTDIR\Themes\qqgame.she"
  Delete "$INSTDIR\Themes\QQӰ��.she"
  Delete "$INSTDIR\Themes\REAL.she"
  Delete "$INSTDIR\Themes\royale.she"
  Delete "$INSTDIR\Themes\storm.she"
  Delete "$INSTDIR\Themes\vista.she"
  Delete "$INSTDIR\Themes\whitefire.she"
  Delete "$INSTDIR\Themes\wish.she"
  Delete "$INSTDIR\Themes\Xenes.she"
  Delete "$INSTDIR\Themes\xmp.she"
  Delete "$INSTDIR\Themes\��ľ.she"
  Delete "$INSTDIR\Themes\����.she"
  Delete "$INSTDIR\SkinH_VB6.dll"
  Delete "$INSTDIR\Config.ini"
  Delete "$INSTDIR\Configuration\LAN����.txt"
  Delete "$INSTDIR\Configuration\RS232�˿�����.txt"
  Delete "$INSTDIR\Configuration\��������.txt"
  Delete "$INSTDIR\Configuration\���̶α������.txt"
  Delete "$INSTDIR\Configuration\���ͨ������.txt"
  Delete "$INSTDIR\Configuration\����ϵ��.txt"
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

#-- ���� NSIS �ű��༭�������� Function ���α�������� Section ����֮���д���Ա��ⰲװ�������δ��Ԥ֪�����⡣--#

Function un.onInit
!insertmacro MUI_UNGETLANGUAGE
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "��ȷʵҪ��ȫ�Ƴ� $(^Name) ���������е������" IDYES +2
  Abort
FunctionEnd

Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name) �ѳɹ��ش����ļ�����Ƴ���"
FunctionEnd
