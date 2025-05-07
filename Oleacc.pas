unit Oleacc;

Interface

uses
  Windows, Variants, Classes;

Const
  ACCDLL = 'OLEACC.DLL';

Const
  DISPID_ACC_PARENT                   =-5000;
  DISPID_ACC_CHILDCOUNT               =-5001;
  DISPID_ACC_CHILD                    =-5002;

  DISPID_ACC_NAME                     =-5003;
  DISPID_ACC_VALUE                    =-5004;
  DISPID_ACC_DESCRIPTION              =-5005;
  DISPID_ACC_ROLE                     =-5006;
  DISPID_ACC_STATE                    =-5007;
  DISPID_ACC_HELP                     =-5008;
  DISPID_ACC_HELPTOPIC                =-5009;
  DISPID_ACC_KEYBOARDSHORTCUT         =-5010;
  DISPID_ACC_FOCUS                    =-5011;
  DISPID_ACC_SELECTION                =-5012;
  DISPID_ACC_DEFAULTACTION            =-5013;

  DISPID_ACC_SELECT                   =-5014;
  DISPID_ACC_LOCATION                 =-5015;
  DISPID_ACC_NAVIGATE                 =-5016;
  DISPID_ACC_HITTEST                  =-5017;
  DISPID_ACC_DODEFAULTACTION          =-5018;

Const
  NAVDIR_MIN                      =$00000000;
  NAVDIR_UP                       =$00000001;
  NAVDIR_DOWN                     =$00000002;
  NAVDIR_LEFT                     =$00000003;
  NAVDIR_RIGHT                    =$00000004;
  NAVDIR_NEXT                     =$00000005;
  NAVDIR_PREVIOUS                 =$00000006;
  NAVDIR_FIRSTCHILD               =$00000007;
  NAVDIR_LASTCHILD                =$00000008;
  NAVDIR_MAX                      =$00000009;

  SELFLAG_NONE                    =$00000000;
  SELFLAG_TAKEFOCUS               =$00000001;
  SELFLAG_TAKESELECTION           =$00000002;
  SELFLAG_EXTENDSELECTION         =$00000004;
  SELFLAG_ADDSELECTION            =$00000008;
  SELFLAG_REMOVESELECTION         =$00000016;
  SELFLAG_VALID                   =$00000032;

  STATE_SYSTEM_UNAVAILABLE        =$00000001;
  STATE_SYSTEM_SELECTED           =$00000002;
  STATE_SYSTEM_FOCUSED            =$00000004;
  STATE_SYSTEM_PRESSED            =$00000008;
  STATE_SYSTEM_CHECKED            =$00000010;
  STATE_SYSTEM_MIXED              =$00000020;
  STATE_SYSTEM_INDETERMINATE      =STATE_SYSTEM_MIXED;
  STATE_SYSTEM_READONLY           =$00000040;
  STATE_SYSTEM_HOTTRACKED         =$00000080;
  STATE_SYSTEM_DEFAULT            =$00000100;
  STATE_SYSTEM_EXPANDED           =$00000200;
  STATE_SYSTEM_COLLAPSED          =$00000400;
  STATE_SYSTEM_BUSY               =$00000800;
  STATE_SYSTEM_FLOATING           =$00001000;
  STATE_SYSTEM_MARQUEED           =$00002000;
  STATE_SYSTEM_ANIMATED           =$00004000;
  STATE_SYSTEM_INVISIBLE          =$00008000;
  STATE_SYSTEM_OFFSCREEN          =$00010000;
  STATE_SYSTEM_SIZEABLE           =$00020000;
  STATE_SYSTEM_MOVEABLE           =$00040000;
  STATE_SYSTEM_SELFVOICING        =$00080000;
  STATE_SYSTEM_FOCUSABLE          =$00100000;
  STATE_SYSTEM_SELECTABLE         =$00200000;
  STATE_SYSTEM_LINKED             =$00400000;
  STATE_SYSTEM_TRAVERSED          =$00800000;
  STATE_SYSTEM_MULTISELECTABLE    =$01000000;
  STATE_SYSTEM_EXTSELECTABLE      =$02000000;
  STATE_SYSTEM_ALERT_LOW          =$04000000;
  STATE_SYSTEM_ALERT_MEDIUM       =$08000000;
  STATE_SYSTEM_ALERT_HIGH         =$10000000;
  STATE_SYSTEM_PROTECTED          =$20000000;
  STATE_SYSTEM_HASPOPUP           =$40000000;
  STATE_SYSTEM_VALID              =$1FFFFFFF;

  ROLE_SYSTEM_TITLEBAR            =$00000001;
  ROLE_SYSTEM_MENUBAR             =$00000002;
  ROLE_SYSTEM_SCROLLBAR           =$00000003;
  ROLE_SYSTEM_GRIP                =$00000004;
  ROLE_SYSTEM_SOUND               =$00000005;
  ROLE_SYSTEM_CURSOR              =$00000006;
  ROLE_SYSTEM_CARET               =$00000007;
  ROLE_SYSTEM_ALERT               =$00000008;
  ROLE_SYSTEM_WINDOW              =$00000009;
  ROLE_SYSTEM_CLIENT              =$00000010;
  ROLE_SYSTEM_MENUPOPUP           =$00000011;
  ROLE_SYSTEM_MENUITEM            =$00000012;
  ROLE_SYSTEM_TOOLTIP             =$00000013;
  ROLE_SYSTEM_APPLICATION         =$00000014;
  ROLE_SYSTEM_DOCUMENT            =$00000015;
  ROLE_SYSTEM_PANE                =$00000016;
  ROLE_SYSTEM_CHART               =$00000017;
  ROLE_SYSTEM_DIALOG              =$00000018;
  ROLE_SYSTEM_BORDER              =$00000019;
  ROLE_SYSTEM_GROUPING            =$00000020;
  ROLE_SYSTEM_SEPARATOR           =$00000021;
  ROLE_SYSTEM_TOOLBAR             =$00000022;
  ROLE_SYSTEM_STATUSBAR           =$00000023;
  ROLE_SYSTEM_TABLE               =$00000024;
  ROLE_SYSTEM_COLUMNHEADER        =$00000025;
  ROLE_SYSTEM_ROWHEADER           =$00000026;
  ROLE_SYSTEM_COLUMN              =$00000027;
  ROLE_SYSTEM_ROW                 =$00000028;
  ROLE_SYSTEM_CELL                =$00000029;
  ROLE_SYSTEM_LINK                =$00000030;
  ROLE_SYSTEM_HELPBALLOON         =$00000031;
  ROLE_SYSTEM_CHARACTER           =$00000032;
  ROLE_SYSTEM_LIST                =$00000033;
  ROLE_SYSTEM_LISTITEM            =$00000034;
  ROLE_SYSTEM_OUTLINE             =$00000035;
  ROLE_SYSTEM_OUTLINEITEM         =$00000036;
  ROLE_SYSTEM_PAGETAB             =$00000037;
  ROLE_SYSTEM_PROPERTYPAGE        =$00000038;
  ROLE_SYSTEM_INDICATOR           =$00000039;
  ROLE_SYSTEM_GRAPHIC             =$00000040;
  ROLE_SYSTEM_STATICTEXT          =$00000041;
  ROLE_SYSTEM_TEXT                =$00000042;
  ROLE_SYSTEM_PUSHBUTTON          =$00000043;
  ROLE_SYSTEM_CHECKBUTTON         =$00000044;
  ROLE_SYSTEM_RADIOBUTTON         =$00000045;
  ROLE_SYSTEM_COMBOBOX            =$00000046;
  ROLE_SYSTEM_DROPLIST            =$00000047;
  ROLE_SYSTEM_PROGRESSBAR         =$00000048;
  ROLE_SYSTEM_DIAL                =$00000049;
  ROLE_SYSTEM_HOTKEYFIELD         =$00000050;
  ROLE_SYSTEM_SLIDER              =$00000051;
  ROLE_SYSTEM_SPINBUTTON          =$00000052;
  ROLE_SYSTEM_DIAGRAM             =$00000053;
  ROLE_SYSTEM_ANIMATION           =$00000054;
  ROLE_SYSTEM_EQUATION            =$00000055;
  ROLE_SYSTEM_BUTTONDROPDOWN      =$00000056;
  ROLE_SYSTEM_BUTTONMENU          =$00000057;
  ROLE_SYSTEM_BUTTONDROPDOWNGRID  =$00000058;
  ROLE_SYSTEM_WHITESPACE          =$00000059;
  ROLE_SYSTEM_PAGETABLIST         =$00000060;
  ROLE_SYSTEM_CLOCK               =$00000061;
  ROLE_SYSTEM_SPLITBUTTON         =$00000062;
  ROLE_SYSTEM_IPADDRESS           =$00000063;
  ROLE_SYSTEM_OUTLINEBUTTON       =$00000064;

Const
  LIBID_Accessibility: TGUID = '{1ea4dbf0-3c3b-11cf-810c-00aa00389b71}';
  IID_IAccessible:     TGUID = '{618736e0-3c3d-11cf-810c-00aa00389b71}';
  IID_IDispatch:       TGUID = '{a6ef9860-c720-11d0-9337-00a0c90dcaa9}';
  IID_IEnumVARIANT:    TGUID = '{00020404-0000-0000-c000-000000000046}';

{$EXTERNALSYM IEnumVariant}
type
  IEnumVariant = interface(IUnknown)
    ['{00020404-0000-0000-C000-000000000046}']
    function Next(celt: LongWord; var rgvar: OleVariant; out pceltFetched: LongWord): HResult; stdcall;
    function Skip(celt: LongWord): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out Enum: IEnumVariant): HResult; stdcall;
  end;

type
  IAccessible = interface;
  LPACCESSIBLE = ^IAccessible;
  LPDispatch = ^IDispatch;

  IAccessible=interface(IDispatch)
    ['{618736e0-3c3d-11cf-810c-00aa00389b71}']
    function Get_accParent(out ppdispParent: IDispatch): HResult; stdcall;
    function Get_accChildCount(out pcountChildren: Integer): HResult; stdcall;
    function Get_accChild(varChild: OleVariant; out ppdispChild: IDispatch): HResult; stdcall;
    function Get_accName(varChild: OleVariant; out pszName: WideString): HResult; stdcall;
    function Get_accValue(varChild: OleVariant; out pszValue: WideString): HResult; stdcall;
    function Get_accDescription(varChild: OleVariant; out pszDescription: WideString): HResult; stdcall;
    function Get_accRole(varChild: OleVariant; out pvarRole: OleVariant): HResult; stdcall;
    function Get_accState(varChild: OleVariant; out pvarState: OleVariant): HResult; stdcall;
    function Get_accHelp(varChild: OleVariant; out pszHelp: WideString): HResult; stdcall;
    function Get_accHelpTopic(out pszHelpFile: WideString; varChild: OleVariant; out pidTopic: Integer): HResult; stdcall;
    function Get_accKeyboardShortcut(varChild: OleVariant; out pszKeyboardShortcut: WideString): HResult; stdcall;
    function Get_accFocus(out pvarChild: OleVariant): HResult; stdcall;
    function Get_accSelection(out pvarChildren: OleVariant): HResult; stdcall;
    function Get_accDefaultAction(varChild: OleVariant; out pszDefaultAction: WideString): HResult; stdcall;
    function accSelect(flagsSelect: Integer; varChild: OleVariant): HResult; stdcall;
    function accLocation(out pxLeft: Integer; out pyTop: Integer; out pcxWidth: Integer; out pcyHeight: Integer; varChild: OleVariant): HResult; stdcall;
    function accNavigate(navDir: Integer; varStart: OleVariant; out pvarEndUpAt: OleVariant): HResult; stdcall;
    function accHitTest(xLeft: Integer; yTop: Integer; out pvarChild: OleVariant): HResult; stdcall;
    function accDoDefaultAction(varChild: OleVariant): HResult; stdcall;
    function Set_accName(varChild: OleVariant; const pszName: WideString): HResult; stdcall;
    function Set_accValue(varChild: OleVariant; const pszValue: WideString): HResult; stdcall;
  end;

type
  LPFNLresultFromObject = function (riid: TGUID; wParam: WPARAM; pAcc: LPACCESSIBLE): HResult; stdcall;
  LPFNObjectFromLresult = function (lResult: LRESULT; riid: TGUID; wParam: WPARAM; ppvObject: LPACCESSIBLE): HResult; stdcall;
  LPFNAccessibleObjectFromWindow = function (wnd: HWND; dwId: DWORD; riid: TGUID; ppvObject: pointer): HResult; stdcall;
  LPFNAccessibleObjectFromPoint = function (ptScreen: TPOINT; pAcc: LPACCESSIBLE; var pvarChild: olevariant): HResult; stdcall;
  LPFNCreateStdAccessibleObject = function (wnd: HWND; idObject: longint; riid: TGUID; ppvObject: LPACCESSIBLE): HResult; stdcall;
  LPFNAccessibleChildren = function (paccContainer: LPAccessible; iChildStart, cChildren: longint;
                                     rgvarChildren: pvariant; var pcObtained: longint): HResult; stdcall;

  function GetOleaccVersionInfo(pdwVer, pdwBuild: pdword): HResult; stdcall;
  function LresultFromObject(riid: TGUID; wParam: WPARAM; pAcc: LPACCESSIBLE): HResult; stdcall;
  function ObjectFromLresult(lResult: LRESULT; riid: TGUID; wParam: WPARAM; ppvObject: LPACCESSIBLE): HResult; stdcall;
  function WindowFromAccessibleObject(pAcc: IACCESSIBLE; var phwnd: HWND): HResult; stdcall;
  function AccessibleObjectFromWindow(hWnd: HWnd; dwId: DWord; const riid: TGUID; var ppvObject: pointer): HResult; stdcall;
  function AccessibleObjectFromEvent(wnd: HWND; dwId: DWORD; dwChildId: DWORD; pAcc: LPACCESSIBLE; var pvarChild: variant): HResult; stdcall;
  function AccessibleObjectFromPoint(ptScreen: TPOINT; pAcc: LPACCESSIBLE; var pvarChild: olevariant): HResult; stdcall;
  function CreateStdAccessibleObject(wnd: HWND; idObject: longint; riid: TGUID; ppvObject: LPACCESSIBLE): HResult; stdcall;
  function CreateStdAccessibleProxyA(wnd: HWND; pszClassName: pchar; idObject: longint; riid: TGUID; ppvObject: pointer): HResult; stdcall;
  function CreateStdAccessibleProxyW(wnd: HWND; pszClassName: pchar; idObject: longint; riid: TGUID; ppvObject: pointer): HResult; stdcall;
  function CreateStdAccessibleProxy(wnd: HWND; pszClassName: pchar; idObject: longint; riid: TGUID; ppvObject: pointer): HResult; stdcall;
  function AccessibleChildren (paccContainer: IAccessible; iChildStart, cChildren: longint; rgvarChildren: pvariant; var pcObtained: longint): HResult; stdcall;
  function GetRoleTextA(lRole: DWORD; lpszRole: pchar; cchRoleMax: byte): HResult; stdcall;
  function GetRoleTextW(lRole: DWORD; lpszRole: pchar; cchRoleMax: byte): HResult; stdcall;
  function GetRoleText (lRole: DWORD; lpszRole: pchar; cchRoleMax: byte): HResult; stdcall;
  function GetStateTextA(lStateBit: DWORD; lpszState: pwidechar; cchState: byte): HResult; stdcall;
  function GetStateTextW(lStateBit: DWORD; lpszState: pwidechar; cchState: byte): HResult; stdcall;
  function GetStateText (lStateBit: DWORD; lpszState: pwidechar; cchState: byte): HResult; stdcall;

implementation

  function GetOleaccVersionInfo(pdwVer, pdwBuild: pdword): HResult; stdcall; external ACCDLL;
  function LresultFromObject(riid: TGUID; wParam: WPARAM; pAcc: LPACCESSIBLE): HResult; external ACCDLL;
  function ObjectFromLresult(lResult: LRESULT; riid: TGUID; wParam: WPARAM; ppvObject: LPACCESSIBLE): HResult; external ACCDLL;
  function WindowFromAccessibleObject (pAcc: IACCESSIBLE; var phwnd: HWND): HResult; external ACCDLL;
  function AccessibleObjectFromWindow(hWnd: HWnd; dwId: DWord; const riid: TGUID; var ppvObject: pointer): HResult; external ACCDLL name'AccessibleObjectFromWindow';
  function AccessibleObjectFromEvent(wnd: HWND; dwId: DWORD; dwChildId: DWORD; pAcc: LPACCESSIBLE; var pvarChild: variant): HResult; external ACCDLL;
  function AccessibleObjectFromPoint(ptScreen: TPOINT; pAcc: LPACCESSIBLE; var pvarChild: olevariant): HResult; external ACCDLL name'AccessibleObjectFromPoint';
  function CreateStdAccessibleObject(wnd: HWND; idObject: longint; riid: TGUID; ppvObject: LPACCESSIBLE): HResult; external ACCDLL;
  function CreateStdAccessibleProxyA(wnd: HWND; pszClassName: pchar; idObject: longint; riid: TGUID; ppvObject: pointer): HResult; stdcall; external ACCDLL;
  function CreateStdAccessibleProxyW(wnd: HWND; pszClassName: pchar; idObject: longint; riid: TGUID; ppvObject: pointer): HResult; stdcall; external ACCDLL;
  function CreateStdAccessibleProxy(wnd: HWND; pszClassName: pchar; idObject: longint; riid: TGUID; ppvObject: pointer): HResult; stdcall; external ACCDLL name 'CreateStdAccessibleProxyA';
  function AccessibleChildren (paccContainer: IACCESSIBLE; iChildStart, cChildren: longint; rgvarChildren: Pvariant; var pcObtained: longint): HResult; external ACCDLL;
  function GetRoleTextA(lRole: DWORD; lpszRole: pchar; cchRoleMax: byte): HResult; external ACCDLL;
  function GetRoleTextW(lRole: DWORD; lpszRole: pchar; cchRoleMax: byte): HResult; external ACCDLL;
  function GetRoleText(lRole: DWORD; lpszRole: pchar; cchRoleMax: byte): HResult; external ACCDLL name 'GetRoleTextA';
  function GetStateTextA(lStateBit: DWORD; lpszState: pwidechar; cchState: byte): HResult; external ACCDLL;
  function GetStateTextW(lStateBit: DWORD; lpszState: pwidechar; cchState: byte): HResult; external ACCDLL;
  function GetStateText(lStateBit: DWORD; lpszState: pwidechar; cchState: byte): HResult; external ACCDLL name 'GetStateTextA';

begin
end.
