unit Unt_Principal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Oleacc;

type
  TFrm_Principal = class(TForm)
    Lbl_Acessibilidade: TLabel;
    MemoInfo: TMemo;
    Mmo_TxtLeitor: TMemo;
    Lbl_TextoLeitor: TLabel;
    Tmr_Info: TTimer;
    procedure Tmr_InfoTimer(Sender: TObject);
  private
    TextoVelho,
    TextoNovo: String;
    Pt: TPoint;

    procedure GetInfo(BuscaFilhos: Boolean);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Frm_Principal: TFrm_Principal;
  Acc: IAccessible;
  VarParent: OleVariant;
  Res: HResult;

implementation

{$R *.dfm}

function GetAccRole(Role: LongInt):string;
var
  R: Array[0..255] of Char;
begin
  GetRoleText(Role, @R, 255);
  Result := R;
end;

function Getpropriete(Prop: Integer; Acc: IAccessible; varID: OleVariant):string;
var
  Str: WideString;
  V: OleVariant;
begin
  case Prop of
    DISPID_ACC_NAME  :Res := Acc.Get_AccName(VarId, Str);
    DISPID_ACC_VALUE :Res := Acc.Get_AccValue(VarId, Str);
    DISPID_ACC_ROLE  :Res := Acc.Get_AccRole(VarId, V);
  end;

  Case Res of
    S_OK:
    Case Prop of
      DISPID_ACC_NAME,
      DISPID_ACC_VALUE: Result := Str;
      DISPID_ACC_ROLE : Result := GetAccRole(V);
    end;

    S_FALSE              : Result := 'Resultado False';
    E_INVALIDARG         : Result := 'Erro de Parametro';
    DISP_E_MEMBERNOTFOUND: Result := 'Esse objeto não suporta Acessibilidade'
    else
      Result := 'Erro Desconhecido';
  end;
end;

function GetName(Acc: IAccessible; varID: OleVariant):string;
begin
  Result := Getpropriete(DISPID_ACC_NAME, Acc, varID);
end;

function GetValue(Acc: IAccessible; varID: OleVariant):string;
begin
  Result := Getpropriete(DISPID_ACC_VALUE, Acc, varID);
end;

function GetRole(Acc: IAccessible; varID: OleVariant):string;
begin
  Result := Getpropriete(DISPID_ACC_ROLE, Acc, varID);
end;

procedure TFrm_Principal.GetInfo(BuscaFilhos: Boolean);
var
  R, Z: array[0..255] of char;
  Win: hwnd;
  Info,
  Nome,
  Valor,
  Role,
  vWinTexto,
  vClassNome: String;
begin
  Info := Info + 'Nome : '+ GetName(Acc, VarParent) +#13#10;
  Info := Info + 'Valor : '+ GetValue(Acc, VarParent) +#13#10;
  Info := Info + 'Role : '+ GetRole(Acc, VarParent) +#13#10;

  WindowFromAccessibleObject(Acc, Win);

  GetClassName(Win, Z, 255);
  Info := Info + 'Nome da Classe : '+ Z +#13#10;

  GetWindowText(Win, R, 255);
  Info := Info + 'TextoWindows : '+ R +#13#10;

  Info := Info + '-------------------' +#13#10;;

  MemoInfo.text := Info;

  TextoNovo := '';

  Nome  := GetName(Acc, Varparent);
  Valor := GetValue(Acc, Varparent);
  Role  := GetRole(Acc, Varparent);

  vWinTexto  := R;
  vClassNome := Z;

  if ((vClassNome <> '') and (vClassNome <> 'Resultado False') and (vClassNome = 'TrayClockWClass')) then
    TextoNovo := 'Relógio '+ vWinTexto +' '+ FormatDateTime('dddd", "dd "de" mmmm "de" yyyy', Now)
  else
  if ((Nome <> '') and (Nome <> 'Resultado False') and (Nome = 'Modo de Exibição de Itens')) then
    TextoNovo := 'Visualizando Pasta'
  else
  if ((Nome <> '') and (Nome <> 'Resultado False') and (Nome <> 'Nome')) then
    TextoNovo := Nome
  else
  if ((Valor <> '') and (Valor <> 'Resultado False') and (Valor <> '0')) then
    TextoNovo := Valor
  else
  if ((vWinTexto <> '') and (vWinTexto <> 'Resultado False')) then
  begin
    if (vWinTexto = 'FolderView') then
      TextoNovo := 'Visualizando Pasta'
    else
      TextoNovo := vWinTexto;
  end
  else
  if ((Role <> '') and (Role <> 'Resultado False') and (Role <> 'cliente')) then
    TextoNovo := Role;

  if ((TextoNovo <> '') and (TextoNovo <> TextoVelho)) then
  begin
    TextoVelho := TextoNovo;
    Mmo_TxtLeitor.Lines.Add(TextoNovo);
  end;
end;

procedure TFrm_Principal.Tmr_InfoTimer(Sender: TObject);
var
  Tpt: TPoint;
  Res: HResult;
begin
  GetCursorPos(Tpt);

  if (Tpt.X = Pt.x) and (Tpt.y = Pt.y) then
    Exit;

  Pt := Tpt;

  Res := AccessibleObjectFrompoint(Pt, @Acc, VarParent);

  if Res <> S_OK then
    Exit;

  GetInfo(False);
end;

end.
