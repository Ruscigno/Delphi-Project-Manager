unit udpmMainForm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBClient, Grids, DBGrids, ExtCtrls, ComCtrls, StdCtrls, Buttons, registry,
  Mask, JvBaseDlg, JvSelectDirectory;

type
  TPositionType = (tpQualquerLugar, tpInicio, tpFim);

  TfdpmMainForm = class(TForm)
    cdsProjects: TClientDataSet;
    cdsBPL: TClientDataSet;
    cdsLibrary: TClientDataSet;
    cdsPath: TClientDataSet;
    cdsVersions: TClientDataSet;
    dsProjetos: TDataSource;
    dsBPL: TDataSource;
    dsLibrary: TDataSource;
    cdsProjectsnmProjeto: TStringField;
    cdsProjectsdePath: TStringField;
    cdsProjectscdVersao: TIntegerField;
    cdsProjectsCC_cdVersao: TStringField;
    Panel1: TPanel;
    pbSalvar: TBitBtn;
    pbAplicar: TBitBtn;
    pbLerDelphi: TBitBtn;
    pbFechar: TBitBtn;
    cdsProjectscdProjeto: TIntegerField;
    dsPath: TDataSource;
    Panel5: TPanel;
    grProjetos: TDBGrid;
    Panel6: TPanel;
    pbNovo: TBitBtn;
    pbExcluir: TBitBtn;
    cdsProjectsdeUnidade: TStringField;
    OpenBPL: TOpenDialog;
    stGerenciador: TSplitter;
    Panel2: TPanel;
    pcGerenciador: TPageControl;
    tsPathBPL: TTabSheet;
    Panel3: TPanel;
    pbUp: TBitBtn;
    pbDown: TBitBtn;
    pbIncluirPath: TBitBtn;
    pbExcluirPath: TBitBtn;
    grPathBPL: TDBGrid;
    tsBPL: TTabSheet;
    grBPL: TDBGrid;
    Panel4: TPanel;
    pbNovoBPL: TBitBtn;
    pbExcluirBPL: TBitBtn;
    tsLibrary: TTabSheet;
    grLibrary: TDBGrid;
    Panel7: TPanel;
    pbNovoLibrary: TBitBtn;
    pbExcluirLibrary: TBitBtn;
    tsConfiguracao: TTabSheet;
    Label1: TLabel;
    edPathDAT: TEdit;
    cbAutoStart: TCheckBox;
    pbDiretorio: TBitBtn;
    pbAplicarAlteracoes: TButton;
    pbSetarDrive: TBitBtn;
    Label2: TLabel;
    tbColunas: TEdit;
    fdPathDAT: TJvSelectDirectory;

    procedure FormCreate(Sender: TObject);
    procedure pbFecharClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure pbSalvarClick(Sender: TObject);
    procedure pbNovoClick(Sender: TObject);
    procedure pbExcluirClick(Sender: TObject);
    procedure dsProjetosStateChange(Sender: TObject);
    procedure cdsProjectsBeforePost(DataSet: TDataSet);
    procedure cdsProjectsNewRecord(DataSet: TDataSet);
    procedure cdsBPLNewRecord(DataSet: TDataSet);
    procedure cdsLibraryNewRecord(DataSet: TDataSet);
    procedure cdsBPLBeforePost(DataSet: TDataSet);
    procedure cdsLibraryBeforePost(DataSet: TDataSet);
    procedure pbLerDelphiClick(Sender: TObject);
    procedure cdsPathNewRecord(DataSet: TDataSet);
    procedure pbMoveUpClick(Sender: TObject);
    procedure cdsPathBeforePost(DataSet: TDataSet);
    procedure pbDownClick(Sender: TObject);
    procedure cdsProjectsAfterScroll(DataSet: TDataSet);
    procedure pbIncluirPathClick(Sender: TObject);
    procedure pbExcluirPathClick(Sender: TObject);
    procedure pbNovoBPLClick(Sender: TObject);
    procedure pbExcluirBPLClick(Sender: TObject);
    procedure pbNovoLibraryClick(Sender: TObject);
    procedure pbExcluirLibraryClick(Sender: TObject);
    procedure pbAplicarClick(Sender: TObject);
    procedure grProjetosEditButtonClick(Sender: TObject);
    procedure grBPLEditButtonClick(Sender: TObject);
    procedure grPathBPLEditButtonClick(Sender: TObject);
    procedure grLibraryEditButtonClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure pbDiretorioClick(Sender: TObject);
    procedure pbAplicarAlteracoesClick(Sender: TObject);
    procedure edPathDATEnter(Sender: TObject);
    procedure edPathDATExit(Sender: TObject);
    procedure cbAutoStartClick(Sender: TObject);
    procedure pbSetarDriveClick(Sender: TObject);
    procedure tbColunasChange(Sender: TObject);
  private
    { Private declarations }

    nMaiorCdProjeto : integer;
    nUltimoDelphi   : integer;
    sUltimoDrive    : string;
    sUltimoDirProj  : string;
    sUltimoDirPath  : string;
    sUltimoDirBPL   : string;
    sUltimoDirLib   : string;
    sOldPath        : string;
    bOldAutoStart   : boolean;

    procedure EnableButtons;
    procedure atualizaRegistro;
    function descobreMaiorCdPath: integer;
    function descobreMenorCdPath: integer;
    function descobreKnownPackages: string;
    function descobreLibrary: string;
    function descobrePath: string;
    function getVersaoDelphi (cdVersao: integer = 0): string;
    function AbreRegistroDelphi(rRegistroDelphi: TRegistry;
      sChave: string): boolean;
    function AbreRegistroPathDelphi(rRegistroDelphi: TRegistry): boolean;
    procedure DesinstalaPacotes;
    procedure InsereLibraryPath(var sErros: string);
    procedure InserePath(var sErros: string);
    procedure InstalaPacotes(var sErros: string);
    function JaTemPacote(slPacotesInstalados: TStringList; sPacote: string;
      var nPacoteInstalado: integer): boolean;
    function MontaArquivo(sDiretorio, sArquivo: string): string;
    function PCharOrNil(const S: AnsiString): PAnsiChar;
    procedure RetiraLibraryPath;
    procedure RetiraPath;
    function SetaDrive(pLetra, pDiretorio: string): boolean;
    function ShellExecAndWait(const FileName, Parameters, Verb: string;
      CmdShow: Integer): Boolean;
    procedure strDelete(var sOriginal: string; sExcluir: string;
      tpPosicao: TPositionType; sSubstituiPor: string);
    function Subst(pDiretorio: string): boolean;
    procedure ConfiguraDelphi;
    procedure abreDataSets;
    procedure SaveData;
    function getVersaoAplicativo: string;
  public
    { Public declarations }
  end;

var
  fdpmMainForm: TfdpmMainForm;

implementation

{$R *.DFM}

uses
  FileCtrl, ShellApi;

const
  sKey = 'SetaDrive';

function TfdpmMainForm.getVersaoAplicativo: string;
var
  iBufferSize, iDummy : DWORD;
  pBuffer, pFileInfo  : pointer;
  iVer                : array[1..4] of word;

begin
  result := '';

  iBufferSize := getFileVersionInfoSize (PChar (paramStr(0)), iDummy);
  if (iBufferSize > 0) then
  begin
    getMem(pBuffer, iBufferSize);

    try
      getFileVersionInfo (PChar (ParamStr(0)), 0, iBufferSize, pBuffer);
      verQueryValue(pBuffer, '\', pFileInfo, iDummy);

      iVer[1] := hiWord (PVSFixedFileInfo (pFileInfo)^.dwFileVersionMS);
      iVer[2] := loWord (PVSFixedFileInfo (pFileInfo)^.dwFileVersionMS);
      iVer[3] := hiWord (PVSFixedFileInfo (pFileInfo)^.dwFileVersionLS);
      iVer[4] := loWord (PVSFixedFileInfo (pFileInfo)^.dwFileVersionLS);
    finally
      FreeMem(pBuffer);
    end;

    result := format('%d.%d.%d-%d', [iVer[1], iVer[2], iVer[3], iVer[4]]);
  end;
end;

procedure TfdpmMainForm.FormCreate(Sender: TObject);

  function getPathDelphi (cdVersao: integer): string;
  var
    oRegVer : TRegIniFile;

  begin
    result  := '';
    oRegVer := TRegIniFile.Create;
    try
      oRegVer.RootKey := HKEY_LOCAL_MACHINE;

      if oRegVer.OpenKey(getVersaoDelphi (cdVersao), True) then
      begin
        result := TRegistry (oRegVer).readString ('RootDir');
        oRegVer.CloseKey;
      end;
    finally
      oRegVer.free;
    end;
  end;

  procedure insereRegistro (cdVersao: integer; deVersao: string);
  begin
    cdsVersions.append;
    cdsVersions.fieldByName ('cdVersao').asInteger := cdVersao;
    cdsVersions.fieldByName ('deVersao').asString  := deVersao;
    cdsVersions.fieldByName ('dePath'  ).asString  := getPathDelphi (cdVersao);
    cdsVersions.post;
  end;

var
  oReg : TRegIniFile;

begin
  caption := caption + ' - Versão: ' + getVersaoAplicativo;

  if not cdsVersions.active then
    cdsVersions.createDataSet;

  nUltimoDelphi  := 7;
  sUltimoDrive   := 'S';
  sUltimoDirProj := extractFileDir (paramStr (0));
  sUltimoDirPath := sUltimoDirProj;
  sUltimoDirBPL  := sUltimoDirProj;
  sUltimoDirLib  := sUltimoDirProj;

  oReg := TRegIniFile.Create;
  try
    oReg.RootKey := HKEY_LOCAL_MACHINE;

    if oReg.OpenKey('\Software\Gerenciador de Projetos', True) then
    begin
      nUltimoDelphi  := oReg.ReadInteger ('Configuração', 'Versão' , nUltimoDelphi);
      sUltimoDrive   := oReg.ReadString  ('Configuração', 'Drive'  , sUltimoDrive);
      sUltimoDirProj := oReg.ReadString  ('Configuração', 'DirProj', sUltimoDirProj);
      sUltimoDirPath := oReg.ReadString  ('Configuração', 'DirPath', sUltimoDirPath);
      sUltimoDirBPL  := oReg.ReadString  ('Configuração', 'DirBPL' , sUltimoDirBPL);
      sUltimoDirLib  := oReg.ReadString  ('Configuração', 'DirLib' , sUltimoDirLib);
      edPathDAT.text := oReg.ReadString  ('Configuração', 'PathDAT', extractFilePath (paramStr (0)));
      cbAutoStart.checked := oReg.Readbool ('Configuração', 'AutoStart', true);
      tbColunas.Text := oreg.ReadString('Configuração', 'Colunas', '100');
      sOldPath            := trim (edPathDAT.text);
      fdPathDAT.InitialDir:= trim (edPathDat.text);
      bOldAutoStart       := cbAutoStart.checked;
      oReg.CloseKey;
    end;
  finally
    oReg.free;
  end;

  insereRegistro (03, '3.0');
  insereRegistro (04, '4.0');
  insereRegistro (05, '5.0');
  insereRegistro (06, '6.0');
  insereRegistro (07, '7.0');
  insereRegistro (08, '8.0');
  insereRegistro (09, '2005');
  insereRegistro (10, '2006');
  insereRegistro (11, '2007');

  abreDataSets;
end;

procedure TfdpmMainForm.pbFecharClick(Sender: TObject);
begin
  close;
end;

procedure TfdpmMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  fdpmMainForm := nil;
  action := caFree;
  self.release;
end;

procedure TfdpmMainForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
  nRes : Word;

begin
  canClose := true;
  nRes     := mrNo;

  if (cdsProjects.ChangeCount > 0) or
     (cdsBPL.ChangeCount > 0) or
     (cdsLibrary.ChangeCount > 0) then
    nRes := messageDLG ('Existem alterações que não foram salvas. Deseja salvar?',
       mtConfirmation, [mbYes, mbNo, mbCancel], 0);

  if nRes = mrCancel then
    CanClose := false
  else
    if nRes = mrYes then
      pbSalvarClick (sender);
end;

procedure TfdpmMainForm.pbSalvarClick(Sender: TObject);
begin
  SaveData;
end;

procedure TfdpmMainForm.pbNovoClick(Sender: TObject);
begin
  if cdsBPL.state in [dsEdit, dsInsert] then
    cdsBPL.post;

  if cdsLibrary.state in [dsEdit, dsInsert] then
    cdsLibrary.post;

  cdsProjects.append;
  activeControl := grProjetos;
  grProjetos.SelectedField := cdsProjects.fieldByName ('nmProjeto');

  pcGerenciador.activePage := tsPathBPL;
end;

procedure TfdpmMainForm.pbExcluirClick(Sender: TObject);
var
  cdProjeto : integer;

begin
  if messageDLG ('Confirma a exclusão do projeto ' +
       cdsProjects.fieldByName ('nmProjeto').asString + '?',
       mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    cdProjeto := cdsProjects.fieldByName ('cdProjeto').asInteger;

    while cdsBPL.Locate ('cdProjeto', cdProjeto, []) do
      cdsBPL.delete;

    while cdsLibrary.Locate ('cdProjeto', cdProjeto, []) do
      cdsLibrary.delete;

    cdsProjects.Delete;
    EnableButtons;
  end;
end;

procedure TfdpmMainForm.EnableButtons;
begin
  pbSalvar.enabled    := True{cdsProjects.active and
                         ((cdsProjects.ChangeCount > 0) or
                          (cdsBPL.ChangeCount > 0) or
                          (cdsLibrary.ChangeCount > 0))};
  pbNovo.enabled      := cdsProjects.active and (cdsProjects.State = dsBrowse);
  pbExcluir.enabled   := cdsProjects.active and
                         (cdsProjects.recordCount > 0) and
                         (cdsProjects.State = dsBrowse);
  pbSetarDrive.enabled := pbExcluir.enabled;
  pbAplicar.enabled   := pbExcluir.enabled;
  pbLerDelphi.enabled := cdsProjects.active and (cdsProjects.recordCount > 0) and
                         (cdsProjects.state = dsBrowse);

  pbIncluirPath.enabled := cdsPath.active and pbAplicar.enabled;
  pbExcluirPath.enabled := pbIncluirPath.enabled and cdsPath.active and (cdsPath.recordCount > 0);
  pbUp.enabled          := pbIncluirPath.enabled and cdsPath.active and (cdsPath.recordCount > 1);
  pbDown.enabled        := pbUp.enabled;

  pbNovoBPL.enabled     := cdsBPL.active and pbAplicar.enabled;
  pbExcluirBPL.enabled  := pbNovoBPL.enabled and cdsBPL.active and (cdsBPL.recordCount > 0);

  pbNovoLibrary.enabled    := cdsLibrary.active and pbAplicar.enabled;
  pbExcluirLibrary.enabled := pbNovoLibrary.enabled and cdsLibrary.active and (cdsLibrary.recordCount > 0);
end;

procedure TfdpmMainForm.dsProjetosStateChange(Sender: TObject);
begin
  EnableButtons;
end;

procedure TfdpmMainForm.cdsProjectsBeforePost(DataSet: TDataSet);
begin
  if trim (cdsProjects.fieldByName ('nmProjeto').asString) = '' then
  begin
    messageDLG ('O campo Nome é obrigatório', mtError, [mbok], 0);
    activeControl := grProjetos;
    grProjetos.SelectedField := cdsProjects.fieldByName ('nmProjeto');
    abort;
  end;

  if trim (cdsProjects.fieldByName ('dePath').asString) = '' then
  begin
    messageDLG ('O campo Path é obrigatório', mtError, [mbok], 0);
    activeControl := grProjetos;
    grProjetos.SelectedField := cdsProjects.fieldByName ('dePath');
    abort;
  end;

  if trim (cdsProjects.fieldByName ('deUnidade').asString) = '' then
  begin
    messageDLG ('O campo Drive é obrigatório', mtError, [mbok], 0);
    activeControl := grProjetos;
    grProjetos.SelectedField := cdsProjects.fieldByName ('deUnidade');
    abort;
  end;

  if trim (cdsProjects.fieldByName ('cdVersao').asString) = '' then
  begin
    messageDLG ('O campo Delphi é obrigatório', mtError, [mbok], 0);
    activeControl := grProjetos;
    grProjetos.SelectedField := cdsProjects.fieldByName ('cdVersao');
    abort;
  end;

  nUltimoDelphi  := cdsProjects.fieldByName ('cdVersao').asInteger;
  sUltimoDrive   := cdsProjects.fieldByName ('deUnidade').asString;
  sUltimoDirProj := extractFileDir (cdsProjects.fieldByName ('dePath').asString);

  atualizaRegistro;
end;

procedure TfdpmMainForm.cdsProjectsNewRecord(DataSet: TDataSet);
begin
  inc (nMaiorCdProjeto);

  cdsProjects.fieldByName ('cdProjeto').asInteger := nMaiorCdProjeto;
  cdsProjects.fieldByName ('cdVersao' ).asInteger := nUltimoDelphi;
  cdsProjects.fieldByName ('deUnidade').asString  := sUltimoDrive;
end;

procedure TfdpmMainForm.cdsBPLNewRecord(DataSet: TDataSet);
begin
  cdsBPL.fieldByName ('cdProjeto').asInteger :=
    cdsProjects.fieldByName ('cdProjeto').asInteger;
end;

function TfdpmMainForm.descobreMaiorCdPath: integer;
var
  oClone : TClientDataSet;

begin
  result := 1;
  oClone := TClientDataSet.create (nil);
  try
    oClone.CloneCursor (cdsPath, True);
    oClone.filtered := false;
    oClone.filter   := 'cdProjeto = ' + cdsProjects.fieldByName ('cdProjeto').asString;
    oClone.filtered := true;
    oClone.first;

    while not oClone.EOF do
    begin
      if oClone.fieldByName ('cdPath').asInteger > result then
        result := oClone.fieldByName ('cdPath').asInteger;

        oClone.next;
    end;
  finally
    oClone.free;
  end;
end;

function TfdpmMainForm.descobreMenorCdPath: integer;
var
  oClone : TClientDataSet;

begin
  result := 1;
  oClone := TClientDataSet.create (nil);
  try
    oClone.CloneCursor (cdsPath, True);
    oClone.filtered := false;
    oClone.filter   := 'cdProjeto = ' + cdsProjects.fieldByName ('cdProjeto').asString;
    oClone.filtered := true;
    oClone.first;

    while not oClone.EOF do
    begin
      if oClone.fieldByName ('cdPath').asInteger < result then
        result := oClone.fieldByName ('cdPath').asInteger;

        oClone.next;
    end;
  finally
    oClone.free;
  end;
end;

procedure TfdpmMainForm.cdsLibraryNewRecord(DataSet: TDataSet);
begin
  cdsLibrary.fieldByName ('cdProjeto').asInteger :=
    cdsProjects.fieldByName ('cdProjeto').asInteger;
end;

procedure TfdpmMainForm.cdsBPLBeforePost(DataSet: TDataSet);
begin
  if trim (cdsBPL.fieldByName ('deBPL').asString) = '' then
  begin
    messageDLG ('O campo BPL é obrigatório', mtError, [mbok], 0);
    pcGerenciador.activePage := tsBPL;
    activeControl := grBPL;
    grBPL.SelectedField := cdsBPL.fieldByName ('deBPL');
    abort;
  end;
end;

procedure TfdpmMainForm.cdsLibraryBeforePost(DataSet: TDataSet);
begin
  if trim (cdsLibrary.fieldByName ('dePath').asString) = '' then
  begin
    messageDLG ('O campo Path é obrigatório', mtError, [mbok], 0);
    pcGerenciador.activePage := tsLibrary;
    activeControl := grLibrary;
    grLibrary.SelectedField := cdsLibrary.fieldByName ('dePath');
    abort;
  end;
end;

procedure TfdpmMainForm.pbLerDelphiClick(Sender: TObject);
var
  oReg     : TRegIniFile;
  oLista   : TStringList;
  i        : integer;
  a, b, bm : string;

begin
  if messageDLG ('A configuração do projeto ' +
       cdsProjects.fieldByName ('nmProjeto').asString +
       ' será sobreescrita com o conteúdo do Registro do Windows. Confirma a operação?',
       mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    oReg   := TRegIniFile.Create;
    oLista := TStringList.create;
    try
      oReg.RootKey := HKEY_CURRENT_USER;

      if oReg.OpenKey(descobreKnownPackages, True) then
      begin
        oReg.GetValueNames (oLista);

        if cdsBPL.state in [dsEdit, dsInsert] then
          cdsBPL.cancel;

        while cdsBPL.locate ('cdProjeto',
                cdsProjects.fieldByName ('cdProjeto').asInteger, []) do
          cdsBPL.delete;

        for i := 0 to oLista.count -1 do
        begin
          if trim (oLista.Strings [i]) <> '' then
          begin
            cdsBPL.append;
            cdsBPL.fieldbyname ('deBPL').asString := oLista.Strings [i];
            cdsBPL.post;
          end;
        end;

        oReg.CloseKey;
      end;

      // Lê as bibliotecas (library path)
      oLista.clear;
      oReg.RootKey := HKEY_CURRENT_USER;

      if oReg.OpenKey(descobreLibrary, False) then
      begin
        a := TRegistry (oReg).ReadString ('Search Path');

        oReg.CloseKey;

        if cdsLibrary.state in [dsEdit, dsInsert] then
          cdsLibrary.cancel;

        while cdsLibrary.locate ('cdProjeto',
                cdsProjects.fieldByName ('cdProjeto').asInteger, []) do
          cdsLibrary.delete;

        b := '';
        for i := 1 to length (a) do
        begin
          if a[i] <> ';' then
            b := b + a[i]
          else
          begin
            if b <> '' then
            begin
              cdsLibrary.append;
              cdsLibrary.fieldByName ('dePath').asString := b;
              cdsLibrary.post;
            end;

            b := '';
          end;
        end;

        if b <> '' then
        begin
          cdsLibrary.append;
          cdsLibrary.fieldByName ('dePath').asString := b;
          cdsLibrary.post;
        end;
      end;

      // Lê o Path do delphi32
      oLista.clear;
      if (cdsProjects.fieldByName ('cdVersao').asInteger < 11) then
        oReg.RootKey := HKEY_LOCAL_MACHINE
      else
        oReg.RootKey := HKEY_CURRENT_USER;

      if oReg.OpenKey(descobrePath, False) then
      begin
        if (cdsProjects.fieldByName ('cdVersao').asInteger < 11) then
          a := TRegistry (oReg).ReadString ('Path')
        else
          a := TRegistry (oReg).ReadString ('App');

        oReg.CloseKey;

        b := '';
        for i := 1 to length (a) do
        begin
          if a[i] <> ';' then
            b := b + a[i]
          else
          begin
            oLista.add (b);
            b := '';
          end;
        end;

        if b <> '' then
          oLista.add (b);
      end;

      // Adiciona Paths que foram encontrados nos pacotes
      cdsBPL.DisableControls;
      try
        bm := cdsBPL.bookMark;
        cdsBPL.first;
        while not cdsBPL.EOF do
        begin
          if oLista.indexOf (ansiLowerCase (
              extractFileDir (cdsBPL.fieldByName ('deBPL').asString))) = -1 then
            oLista.add (extractFileDir (cdsBPL.fieldByName ('deBPL').asString));

          cdsBPL.next;
        end;
      finally
        cdsBPL.bookMark := bm;
        cdsBPL.EnableControls;
      end;

      if cdsPath.state in [dsEdit, dsInsert] then
        cdsPath.cancel;

      while cdsPath.locate ('cdProjeto',
              cdsProjects.fieldByName ('cdProjeto').asInteger, []) do
        cdsPath.delete;

      for i := 0 to oLista.count -1 do
      begin
        if trim (oLista.Strings [i]) <> '' then
        begin
          cdsPath.append;
          cdsPath.fieldByName ('dePath').asString := oLista.Strings [i];
          cdsPath.post;
        end;
      end;
    finally
      oReg.free;

      oLista.clear;
      oLista.free;
    end;
  end;
end;

function TfdpmMainForm.getVersaoDelphi (cdVersao: integer): string;
var
  nVersao : integer;

begin
  if cdVersao > 0 then
    nVersao := cdVersao
  else
    nVersao := cdsProjects.fieldByName ('cdVersao').asInteger;

  case nVersao of
     3 : result := '\Software\Borland\Delphi\3.0';
     4 : result := '\Software\Borland\Delphi\4.0';
     5 : result := '\Software\Borland\Delphi\5.0';
     6 : result := '\Software\Borland\Delphi\6.0';
     7 : result := '\Software\Borland\Delphi\7.0';
     8 : result := '\Software\Borland\BDS\2.0'; //testar
     9 : result := '\Software\Borland\BDS\3.0'; //testar
    10 : result := '\Software\Borland\BDS\4.0'; //testar
    11 : result := '\Software\Borland\BDS\5.0'; //testar
  else
    result := '';
  end;
end;

function TfdpmMainForm.descobreKnownPackages: string;
begin
  result := GetVersaoDelphi + '\Known Packages';
end;

function TfdpmMainForm.descobreLibrary: string;
begin
  result := GetVersaoDelphi + '\Library';
end;

function TfdpmMainForm.descobrePath: string;
begin
  // testar no delphi 8, 2005 e 2006
  if (cdsProjects.fieldByName ('cdVersao').asInteger < 11) then
    result := '\Software\Microsoft\Windows\CurrentVersion\App Paths\Delphi32.exe'
  else
    result := '\Software\Borland\BDS\5.0'
end;

procedure TfdpmMainForm.cdsPathNewRecord(DataSet: TDataSet);
begin
  cdsPath.fieldByName ('cdProjeto').asInteger :=
    cdsProjects.fieldByName ('cdProjeto').asInteger;
  cdsPath.fieldByName ('cdPath').asInteger := descobreMaiorCdPath + 1;
end;

procedure TfdpmMainForm.pbMoveUpClick(Sender: TObject);
var
  cdPath1, cdPath2 : integer;
  dePath1, dePath2 : string;

begin
  if cdspath.fieldByName ('cdPath').asInteger > descobreMenorCdPath then
  begin
    cdsPath.disableControls;
    try
      cdPath1 := cdsPath.fieldByName ('cdPath').asInteger;
      dePath1 := cdsPath.fieldByName ('dePath').asString;

      cdsPath.Prior;

      cdPath2 := cdsPath.fieldByName ('cdPath').asInteger;
      dePath2 := cdsPath.fieldByName ('dePath').asString;

      cdsPath.edit;
      cdsPath.fieldByName ('dePath').asString := dePath1;
      cdsPath.post;

      cdsPath.locate ('cdPath', cdPath1, []);
      cdsPath.edit;
      cdsPath.fieldByName ('dePath').asString := dePath2;
      cdsPath.post;

      cdsPath.locate ('cdPath', cdPath2, []);
    finally
      cdsPath.enableControls;
    end;
  end;
end;

procedure TfdpmMainForm.cdsPathBeforePost(DataSet: TDataSet);
begin
  if trim (cdsPath.fieldByName ('dePath').asString) = '' then
  begin
    messageDLG ('O campo Path é obrigatório', mtError, [mbok], 0);
    pcGerenciador.activePage := tsPathBPL;
    activeControl := grPathBPL;
    grPathBPL.SelectedField := cdsPath.fieldByName ('dePath');
    abort;
  end;
end;

procedure TfdpmMainForm.pbDownClick(Sender: TObject);
var
  cdPath1, cdPath2 : integer;
  dePath1, dePath2 : string;

begin
  if cdspath.fieldByName ('cdPath').asInteger < descobreMaiorCdPath then
  begin
    cdsPath.disableControls;
    try
      cdPath1 := cdsPath.fieldByName ('cdPath').asInteger;
      dePath1 := cdsPath.fieldByName ('dePath').asString;

      cdsPath.Next;

      cdPath2 := cdsPath.fieldByName ('cdPath').asInteger;
      dePath2 := cdsPath.fieldByName ('dePath').asString;

      cdsPath.edit;
      cdsPath.fieldByName ('dePath').asString := dePath1;
      cdsPath.post;

      cdsPath.locate ('cdPath', cdPath1, []);
      cdsPath.edit;
      cdsPath.fieldByName ('dePath').asString := dePath2;
      cdsPath.post;

      cdsPath.locate ('cdPath', cdPath2, []);
    finally
      cdsPath.enableControls;
    end;
  end;
end;

procedure TfdpmMainForm.cdsProjectsAfterScroll(DataSet: TDataSet);
begin
  EnableButtons;
end;

procedure TfdpmMainForm.pbIncluirPathClick(Sender: TObject);
begin
  cdsPath.insert;
end;

procedure TfdpmMainForm.pbExcluirPathClick(Sender: TObject);
begin
  if messageDLG ('Confirma a exclusão do Path: ' +
       cdsPath.fieldByName ('dePath').asString + ' ?',
       mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    cdsPath.delete;
    EnableButtons;
  end;
end;

procedure TfdpmMainForm.pbNovoBPLClick(Sender: TObject);
begin
  cdsBPL.insert;
end;

procedure TfdpmMainForm.pbExcluirBPLClick(Sender: TObject);
begin
  if messageDLG ('Confirma a exclusão da BPL: ' +
       cdsBPL.fieldByName ('deBPL').asString + ' ?',
       mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    cdsBPL.delete;
    EnableButtons;
  end;
end;

procedure TfdpmMainForm.pbNovoLibraryClick(Sender: TObject);
begin
  cdsLibrary.insert;
end;

procedure TfdpmMainForm.pbExcluirLibraryClick(Sender: TObject);
begin
  if messageDLG ('Confirma a exclusão do Library Path: ' +
       cdsLibrary.fieldByName ('dePath').asString + ' ?',
       mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    cdsLibrary.delete;
    EnableButtons;
  end;
end;

procedure TfdpmMainForm.pbAplicarClick(Sender: TObject);
var
  sErros: string;

begin
  if messageDLG ('Confirma a seleção do projeto: ' +
       cdsProjects.fieldByName ('nmProjeto').asString + ' ?',
       mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    if trim (cdsProjects.fieldByName ('cdProjeto').asString) <> '' then
    begin
      if not SetaDrive(cdsProjects.FieldByName('deUnidade').asString,
                       cdsProjects.FieldByName('dePath').asString) then
        messageDLG ('Não foi possível criar o drive ' +
          cdsProjects.fieldByName ('deUnidade').asString + '.',
          mtError, [mbOk], 0)
      else
      begin
        // Limpa o registro atual do Delphi
        DesinstalaPacotes;
        RetiraLibraryPath;
        RetiraPath;

        // Registra a configuração do sistema corrente em cdSistema
        InstalaPacotes (sErros);
        InsereLibraryPath (sErros);
        InserePath (sErros);
        ConfiguraDelphi;
        messageDLG ('O projeto: ' +
          cdsProjects.fieldByName ('nmProjeto').asString + ' foi registrado.',
          mtInformation, [mbOk], 0);
      end;
    end;

    if sErros <> '' then
      messageDLG ('Erro: ' + sErros,
        mtError, [mbOk], 0);
  end;
end;

procedure TfdpmMainForm.grBPLEditButtonClick(Sender: TObject);
begin
  if trim (cdsBPL.fieldByName ('deBPL').asString) <> '' then
  begin
    OpenBPL.initialDir := extractFileDir (cdsBPL.fieldByName ('deBPL').asString);
    OpenBPL.fileName   := extractFileName (cdsBPL.fieldByName ('deBPL').asString);
  end else
    OpenBPL.initialDir := sUltimoDirBPL;

  if OpenBPL.Execute and
     (trim (OpenBPL.fileName) <> '') and
     (ansiLowerCase (extractFileExt (OpenBPL.fileName)) = '.bpl') then
  begin
    if not (cdsBPL.state in [dsEdit, dsInsert]) then
      cdsBPL.edit;

    cdsBPL.fieldByName ('deBPL').asString := OpenBPL.fileName;
    sUltimoDirBPL := extractFileDir (OpenBPL.fileName);
    atualizaRegistro;
  end;
end;

procedure TfdpmMainForm.grPathBPLEditButtonClick(Sender: TObject);
var
  sDir : string;

begin
  sDir := '';

  if trim (cdsPath.fieldByName ('dePath').asString) <> '' then
    sDir := cdsPath.fieldByName ('dePath').asString
  else
    sDir := sUltimoDirPath;

  if SelectDirectory (sDir, [sdAllowCreate, sdPerformCreate, sdPrompt], 0) and
     directoryExists (sDir) then
  begin
    if not (cdsPath.state in [dsEdit, dsInsert]) then
      cdsPath.edit;

    cdsPath.fieldByName ('dePath').asString := sDir;
    sUltimoDirPath := sDir;
    atualizaRegistro;
  end;
end;

procedure TfdpmMainForm.grLibraryEditButtonClick(Sender: TObject);
var
  sDir : string;

begin
  sDir := '';

  if trim (cdsLibrary.fieldByName ('dePath').asString) <> '' then
    sDir := cdsLibrary.fieldByName ('dePath').asString
  else
    sDir := sUltimoDirLib;

  if SelectDirectory (sDir, [sdAllowCreate, sdPerformCreate, sdPrompt], 0) and
     directoryExists (sDir) then
  begin
    if not (cdsLibrary.state in [dsEdit, dsInsert]) then
      cdsLibrary.edit;

    cdsLibrary.fieldByName ('dePath').asString := sDir;
    sUltimoDirLib := sDir;
    atualizaRegistro;
  end;
end;

procedure TfdpmMainForm.grProjetosEditButtonClick(Sender: TObject);
var
  sDir : string;

begin
  sDir := '';

  if trim (cdsProjects.fieldByName ('dePath').asString) <> '' then
    sDir := cdsProjects.fieldByName ('dePath').asString
  else
    sDir := sUltimoDirProj;

  if SelectDirectory (sDir, [sdAllowCreate, sdPerformCreate, sdPrompt], 0) and
     directoryExists (sDir) then
  begin
    if not (cdsProjects.state in [dsEdit, dsInsert]) then
      cdsProjects.edit;

    cdsProjects.fieldByName ('dePath').asString := sDir;
    sUltimoDirProj := sDir;
    atualizaRegistro;
  end;
end;

procedure TfdpmMainForm.atualizaRegistro;
var
  oReg : TRegIniFile;

begin
  oReg := TRegIniFile.Create;
  try
    oReg.RootKey := HKEY_LOCAL_MACHINE;

    if oReg.OpenKey('\Software\Gerenciador de Projetos', True) then
    begin
      oReg.WriteInteger ('Configuração', 'Versão' , nUltimoDelphi);
      oReg.WriteString  ('Configuração', 'Drive'  , sUltimoDrive);
      oReg.WriteString  ('Configuração', 'DirProj', sUltimoDirProj);
      oReg.WriteString  ('Configuração', 'DirPath', sUltimoDirPath);
      oReg.WriteString  ('Configuração', 'DirBPL' , sUltimoDirBPL);
      oReg.WriteString  ('Configuração', 'DirLib' , sUltimoDirLib);
      oReg.CloseKey;
    end;
  finally
    oReg.free;
  end;
end;

procedure TfdpmMainForm.FormResize(Sender: TObject);
begin
  grPathBPL.columns  [0].width := grPathBPL.width - 34;
  grBPL.columns      [0].width := grBPL.width     - 34;
  grLibrary.columns  [0].width := grLibrary.width - 34;
  grProjetos.columns [1].width := grProjetos.width - 37 -
                                  grProjetos.columns [0].width -
                                  grProjetos.columns [2].width -
                                  grProjetos.columns [3].width;
end;

function TfdpmMainForm.AbreRegistroPathDelphi (rRegistroDelphi: TRegistry): boolean;
begin
  if (cdsProjects.fieldByName ('cdVersao').asInteger < 11) then
    rRegistroDelphi.RootKey := HKEY_LOCAL_MACHINE
  else
    rRegistroDelphi.RootKey := HKEY_CURRENT_USER;

  result := rRegistroDelphi.OpenKey (descobrePath, false);
end;

function TfdpmMainForm.AbreRegistroDelphi(
  rRegistroDelphi: TRegistry; sChave: string): boolean;
begin
  rRegistroDelphi.RootKey := HKEY_CURRENT_USER;
  result := rRegistroDelphi.OpenKey (GetVersaoDelphi + '\' + sChave, false);
end;

procedure TfdpmMainForm.DesinstalaPacotes;
var
  DelphiPackage : TRegistry;
  slPacotes: TStringList;
  nCont: integer;

begin
  DelphiPackage := TRegistry.Create;
  slPacotes     := TStringList.Create;
  try
    if AbreRegistroDelphi (DelphiPackage, 'Known Packages') then
    begin
      DelphiPackage.GetValueNames (slPacotes);
      for nCont := 0 to slPacotes.count - 1 do
        DelphiPackage.DeleteValue (slPacotes[nCont]);
    end;
  finally
    DelphiPackage.Free;
    slPacotes.Free;
  end;
end;

procedure TfdpmMainForm.RetiraLibraryPath;
var
  DelphiPackage : TRegistry;

begin
  DelphiPackage := TRegistry.create;
  try
    if AbreRegistroDelphi (DelphiPackage, 'Library') then
      DelphiPackage.WriteString ('Search Path', '')
  finally
    DelphiPackage.free;
  end;
end;

procedure TfdpmMainForm.strDelete (var sOriginal: string;
  sExcluir: string; tpPosicao: TPositionType; sSubstituiPor: string);
var
  nPos: integer;
begin
  if tpPosicao = tpFim then
    nPos := Length (sOriginal) - Length (sExcluir) + 1
  else
    nPos := Pos (AnsiUpperCase (sExcluir), AnsiUpperCase (sOriginal));
  if ((tpPosicao = tpQualquerLugar) and (nPos > 0)) or
    ((tpPosicao = tpInicio) and (nPos = 1)) or
    ((tpPosicao = tpFim) and (AnsiUpperCase (Copy (sOriginal, nPos,
      Length (sExcluir))) = AnsiUpperCase (sExcluir))) then
  begin
    Delete (sOriginal, nPos, Length (sExcluir));
    Insert (sSubstituiPor, sOriginal, nPos);
  end;
end;

procedure TfdpmMainForm.RetiraPath;
var
  regPathDelphi : TRegistry;
  sPathDelphi   : string;

begin
  regPathDelphi := TRegistry.create;
  try
    if AbreRegistroPathDelphi (regPathDelphi) then
    begin
      if (cdsProjects.fieldByName ('cdVersao').asInteger < 11) then
        sPathDelphi := ExtractFilePath (regPathDelphi.ReadString (''))
      else
        sPathDelphi := ExtractFilePath (regPathDelphi.ReadString ('App'));
      strDelete (sPathDelphi, '\', tpFim, '');
      regPathDelphi.WriteString ('Path', sPathDelphi)
    end;
  finally
    regPathDelphi.free;
  end;
end;

function TfdpmMainForm.MontaArquivo(sDiretorio, sArquivo: string): string;
begin
  if (Length (sArquivo) >= 2) and ((trim (sArquivo)[2] = ':') or
    (trim (sArquivo)[1] = '$')) then
    result := sArquivo
  else
  begin
    if (sDiretorio[Length (sDiretorio)] <> '\') and (sArquivo[1] <> '\') then
      sDiretorio := sDiretorio + '\';
    if (sArquivo[1] = '\') and (sDiretorio[Length (sDiretorio)] = '\') then
      Delete (sArquivo, 1, 1);

    result := trim (sDiretorio) + trim (sArquivo);
  end;
end;

function TfdpmMainForm.JaTemPacote(slPacotesInstalados: TStringList;
  sPacote: string; var nPacoteInstalado: integer): boolean;
var
  nCont: integer;
begin
  result := false;
  sPacote := AnsiUpperCase (ExtractFileName (sPacote));
  for nCont := 0 to slPacotesInstalados.Count - 1 do
  begin
    if sPacote = AnsiUpperCase (ExtractFileName (
      slPacotesInstalados[nCont])) then
    begin
      result := true;
      nPacoteInstalado := nCont;
      break;
    end;
  end;
end;

procedure TfdpmMainForm.InstalaPacotes (var sErros: string);
var
  DelphiPackage      : TRegistry;
  slPacotesInstalados: TStringList;
  sPacote, sDescricao, sErroPacote, sDirInstalacao: string;
  sPacoteDisco     : string;
  nPacoteExistente : integer;

begin
  DelphiPackage       := TRegistry.create;
  slPacotesInstalados := TStringList.create;
  try
    DelphiPackage.RootKey := HKEY_LOCAL_MACHINE;

    if DelphiPackage.OpenKey (getVersaoDelphi + '\', false) then
    begin
      sDirInstalacao := DelphiPackage.ReadString ('RootDir');
      DelphiPackage.CloseKey;

      if AbreRegistroDelphi (DelphiPackage, 'Known Packages') then
      begin
        slPacotesInstalados.clear;
        DelphiPackage.GetValueNames (slPacotesInstalados);
        cdsBPL.first;

        while not cdsBPL.eof do
        begin
          sPacote := MontaArquivo (cdsProjects.fieldByName ('dePath').asString,
            cdsBPL.fieldByName ('deBPL').asString);

          sPacoteDisco := sPacote;
          strDelete (sPacoteDisco, '$(DELPHI)', tpInicio, sDirInstalacao);
          if FileExists (sPacoteDisco) then
          begin
            sDescricao := GetPackageDescription (PChar (sPacoteDisco));
            if JaTemPacote (slPacotesInstalados, sPacote,
              nPacoteExistente) then
            begin
              DelphiPackage.DeleteValue (slPacotesInstalados [nPacoteExistente]);
              slPacotesInstalados.Delete (nPacoteExistente);
            end;
            DelphiPackage.WriteString (sPacote, sDescricao);
          end
          else
            sErroPacote := sErroPacote + sPacote + ', ';

          cdsBPL.Next;
        end;
      end;
    end;
  finally
    DelphiPackage.free;
    slPacotesInstalados.free;
  end;

  if sErroPacote <> '' then
    sErros := sErros + '   Os pacotes ' + sErroPacote +
      'não existem no disco e não podem ser instalados.';
end;

procedure TfdpmMainForm.InsereLibraryPath(var sErros: string);
var
  DelphiPackage : TRegistry;
  sPath, sLibraryPath: string;
begin
  DelphiPackage := TRegistry.create;
  try
    if AbreRegistroDelphi (DelphiPackage, 'Library') and
      (cdsLibrary.RecordCount > 0) then
    begin
      cdsLibrary.first;
      sLibraryPath := DelphiPackage.ReadString ('Search Path');
      while not cdsLibrary.eof do
      begin
        sPath := MontaArquivo (cdsProjects.fieldByName ('dePath').asString,
          cdsLibrary.fieldByName ('dePath').asString);
        sLibraryPath := sLibraryPath + ';' + sPath;
        cdsLibrary.Next;
      end;

      DelphiPackage.WriteString ('Search Path', sLibraryPath);
    end;
  finally
    DelphiPackage.free;
  end;
end;

procedure TfdpmMainForm.InserePath(var sErros: string);
var
  regPathDelphi : TRegistry;
  sPath, sPathDelphi: string;

begin
  regPathDelphi := TRegistry.create;
  try
    if AbreRegistroPathDelphi (regPathDelphi) and
      (cdsPath.RecordCount > 0) then
    begin
      cdsPath.first;
      sPathDelphi := regPathDelphi.ReadString ('Path');
      while not cdsPath.eof do
      begin
        sPath := MontaArquivo (cdsProjects.fieldByName ('dePath').asString,
          cdsPath.fieldByName ('dePath').asString);
        sPathDelphi := sPathDelphi + ';' + sPath;
        cdsPath.Next;
      end;
      regPathDelphi.WriteString ('Path', sPathDelphi)
    end;
  finally
    regPathDelphi.free;
  end;
end;

function TfdpmMainForm.PCharOrNil(const S: AnsiString): PAnsiChar;
begin
  if Length(S) = 0 then
    Result := nil
  else
    Result := PAnsiChar(S);
end;

function TfdpmMainForm.ShellExecAndWait(const FileName: string; const Parameters: string;
  const Verb: string; CmdShow: Integer): Boolean;
var
  Sei: TShellExecuteInfo;

begin
  FillChar(Sei, SizeOf(Sei), #0);
  Sei.cbSize := SizeOf(Sei);
  Sei.fMask := SEE_MASK_DOENVSUBST or SEE_MASK_FLAG_NO_UI or SEE_MASK_NOCLOSEPROCESS;
  Sei.lpFile := PChar(FileName);
  Sei.lpParameters := PCharOrNil(Parameters);
  Sei.lpVerb := PCharOrNil(Verb);
  Sei.nShow := CmdShow;
  Result := ShellExecuteEx(@Sei);
  if Result then
  begin
    WaitForInputIdle(Sei.hProcess, INFINITE);
    WaitForSingleObject(Sei.hProcess, INFINITE);
    CloseHandle(Sei.hProcess);
  end;
end;

function TfdpmMainForm.Subst(pDiretorio: string): boolean;
begin
  result := ShellExecAndWait('subst', pDiretorio, '', SW_HIDE);
end;

function TfdpmMainForm.SetaDrive(pLetra, pDiretorio: string): boolean;
var
  bDirExists: boolean;
  oReg: TRegistry;

begin
  if not(cbAutoStart.Checked) then
  begin
    result := True;
    Exit;
  end;
  
  bDirExists := DirectoryExists(pDiretorio);
  pDiretorio := '"' + pDiretorio + '"';
  //--
  result := bDirExists;

  Subst(pLetra + ': /d');
  if result then
    result := Subst(pLetra + ': ' + pDiretorio);

  if cbAutoStart.checked then
  begin
    oReg := TRegistry.Create;

    try
      oReg.RootKey := HKEY_LOCAL_MACHINE;
      if oReg.OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Run', True) then
      begin
        try
          if (not result)then
          begin
            if oReg.ValueExists(sKey + pLetra) then
              oReg.DeleteValue(sKey + pLetra);
          end else
            oReg.WriteString(sKey + pLetra, 'subst ' + pLetra + ': ' + pDiretorio);
        finally
          oReg.CloseKey;
        end;
      end;
    finally
      oReg.Free;
    end;
  end;

  if (not bDirExists) then
  begin
    MessageDlg('Não é possível criar o drive ' + pLetra + ': pois o ' +
      ' diretório "' + pDiretorio + '" não existe.', mtError, [mbOK], 0);
  end;
end;

procedure TfdpmMainForm.pbDiretorioClick(Sender: TObject);
begin
  if fdPathDAT.Execute then
    edPathDAT.text := trim (fdPathDAT.Directory);
end;

procedure TfdpmMainForm.cbAutoStartClick(Sender: TObject);
begin
  pbAplicarAlteracoes.enabled :=
    (trim (sOldPath) <> trim (edPathDAT.text)) or
    (bOldAutoStart <> cbAutoStart.checked);
end;

procedure TfdpmMainForm.edPathDATEnter(Sender: TObject);
begin
  sOldPath := trim (edPathDAT.text);
end;

procedure TfdpmMainForm.edPathDATExit(Sender: TObject);
begin
  if not DirectoryExists (edPathDAT.text) then
  begin
    showMessage ('O diretório informado não existe.');
    edPathDAT.text := trim (sOldPath);
  end else
    pbAplicarAlteracoes.enabled :=
      (trim (sOldPath) <> trim (edPathDAT.text)) or
      (bOldAutoStart <> cbAutoStart.checked);
end;

procedure TfdpmMainForm.pbAplicarAlteracoesClick(Sender: TObject);
var
  oReg : TRegIniFile;
  resp : word;
  sAUX : string;

begin
  if messageDLG ('As alterações serão aplicadas. Confirma Operação?', mtConfirmation,
     [mbYes, mbNo], 0) = mrYes then
  begin
    if pbSalvar.enabled then
    begin
      resp := messageDLG ('Existem alterações pendentes de salvamento. Deseja salvar?',
               mtConfirmation, [mbYes, mbNo, mbCancel], 0);

      if resp = mrCancel then
        abort;

      if resp = mrYes then
      begin
        sAUX := edPathDAT.text;
        try
          edPathDAT.text := sOldPath;
          SaveData;
        finally
          edPathDAT.text := sAUX;
        end;
      end;
    end;

    oReg := TRegIniFile.Create;
    try
      oReg.RootKey := HKEY_LOCAL_MACHINE;

      if oReg.OpenKey('\Software\Gerenciador de Projetos', True) then
      begin
        oReg.WriteString ('Configuração', 'PathDAT'  , edPathDAT.text);
        oReg.WriteBool   ('Configuração', 'AutoStart', cbAutoStart.checked);
        oReg.WriteString('Configuração', 'Colunas', tbColunas.Text);
        oReg.CloseKey;

        if edPathDat.text <> '' then
          fdPathDAT.InitialDir := edPathDat.text
        else
          fdPathDAT.InitialDir := extractFilePath (paramStr (0));
      end;

      if (not cbAutoStart.checked) and
         oReg.OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Run', True) then
      try
        if oReg.ValueExists(sKey) then
          oReg.DeleteValue(sKey);
      finally
        oReg.CloseKey;
      end;
    finally
      oReg.free;
    end;

    if (trim (sOldPath) <> trim (edPathDAT.text)) and
       (messageDLG ('Deseja salvar os dados na nova pasta?', mtConfirmation,
        [mbYes, mbNo], 0) = mrYes) then
      SaveData;

    sOldPath      := trim (edPathDAT.text);
    bOldAutoStart := cbAutoStart.checked;

    cdsProjects.close;
    cdsBPL.close;
    cdsLibrary.close;
    cdsPath.close;

    abreDataSets;

    pbAplicarAlteracoes.enabled := false;
  end;
end;

procedure TfdpmMainForm.abreDataSets;
begin
  if not cdsProjects.active then
    cdsProjects.createDataSet;

  if not cdsBPL.active then
    cdsBPL.createDataSet;

  if not cdsLibrary.active then
    cdsLibrary.createDataSet;

  if not cdsPath.active then
    cdsPath.createDataSet;

  if fileExists (IncludeTrailingPathDelimiter(edPathDAT.text) + 'cdsProjects.cds') then
    cdsProjects.LoadFromFile (IncludeTrailingPathDelimiter (edPathDAT.text) + 'cdsProjects.cds');

  cdsProjects.disableControls;
  try
    cdsProjects.first;
    while not cdsProjects.EOF do
    begin
      if cdsProjects.fieldByName ('cdProjeto').asInteger > nMaiorCdProjeto then
        nMaiorCdProjeto := cdsProjects.fieldByName ('cdProjeto').asInteger;

      cdsProjects.next;
    end;
  finally
    cdsProjects.first;
    cdsProjects.enableControls;
  end;

  if fileExists (IncludeTrailingPathDelimiter (edPathDAT.text) + 'cdsBPL.cds') then
    cdsBPL.LoadFromFile (IncludeTrailingPathDelimiter (edPathDAT.text) + 'cdsBPL.cds');

  if fileExists (IncludeTrailingPathDelimiter (edPathDAT.text) + 'cdsLibrary.cds') then
    cdsLibrary.LoadFromFile (IncludeTrailingPathDelimiter (edPathDAT.text) + 'cdsLibrary.cds');

  if fileExists (IncludeTrailingPathDelimiter (edPathDAT.text) + 'cdsPath.cds') then
    cdsPath.loadFromFile (IncludeTrailingPathDelimiter (edPathDAT.text) + 'cdsPath.cds');

  pcGerenciador.activePage := tsPathBPL;
  activeControl := grProjetos;

  cdsProjects.MergeChangeLog;
  cdsProjects.CancelUpdates;
  cdsBPL.MergeChangeLog;
  cdsBPL.CancelUpdates;
  cdsLibrary.MergeChangeLog;
  cdsLibrary.CancelUpdates;
  cdsPath.MergeChangeLog;
  cdsPath.CancelUpdates;

  EnableButtons;
end;

procedure TfdpmMainForm.SaveData;
begin
  cdsProjects.MergeChangeLog;
  cdsProjects.CancelUpdates;
  cdsBPL.MergeChangeLog;
  cdsBPL.CancelUpdates;
  cdsLibrary.MergeChangeLog;
  cdsLibrary.CancelUpdates;
  cdsPath.MergeChangeLog;
  cdsPath.CancelUpdates;

  if cdsProjects.recordCount > 0 then
    cdsProjects.SaveToFile (IncludeTrailingPathDelimiter (edPathDAT.text) + 'cdsProjects.cds', dfXML);

  if cdsBPL.recordCount > 0 then
    cdsBPL.SaveToFile (IncludeTrailingPathDelimiter (edPathDAT.text) + 'cdsBPL.cds', dfXML);

  if cdsLibrary.recordCount > 0 then
    cdsLibrary.SaveToFile (IncludeTrailingPathDelimiter (edPathDAT.text) + 'cdsLibrary.cds', dfXML);

  if cdsPath.recordCount > 0 then
    cdsPath.SaveToFile (IncludeTrailingPathDelimiter (edPathDAT.text) + 'cdsPath.cds', dfXML);

  EnableButtons;
end;

procedure TfdpmMainForm.pbSetarDriveClick(Sender: TObject);
begin
  if DirectoryExists(cdsProjects.fieldByName ('depath').asString) then
  begin
    Subst(cdsProjects.fieldByName ('deUnidade').asString + ': /d');
    Subst(cdsProjects.fieldByName ('deUnidade').asString + ': ' +
          cdsProjects.fieldByName ('depath'   ).asString);
  end else
    showMessage ('O diretório informado no projeto selecionado não existe.');
end;

procedure TfdpmMainForm.tbColunasChange(Sender: TObject);
begin
  pbAplicarAlteracoes.enabled := true;
end;

procedure TfdpmMainForm.ConfiguraDelphi;
var
  DelphiPackage : TRegistry;
  nErro, nVal: integer;
begin
  // autocreateforms, colunas, stop on delphi exceptions, showcompilerprogress, tab, grid}
  DelphiPackage := TRegistry.create;
  try
    if AbreRegistroDelphi (DelphiPackage, 'Editor\Options') and
      (cdsLibrary.RecordCount > 0) then
    begin
      cdsLibrary.first;
      DelphiPackage.WriteString ('Tab Character', 'False');
      DelphiPackage.WriteString ('Tab Stops', '2');
      Val(tbColunas.Text, nVal, nErro);
      if nErro <> 0 then
        nVal := 100;
      DelphiPackage.WriteInteger('Right Margin', nVal);
    end;
    if AbreRegistroDelphi (DelphiPackage, 'Debugging') and
      (cdsLibrary.RecordCount > 0) then
    begin
      cdsLibrary.first;
      DelphiPackage.WriteString ('Break On Delphi Exceptions', '1');
    end;
    if AbreRegistroDelphi (DelphiPackage, 'Compiling') and
      (cdsLibrary.RecordCount > 0) then
    begin
      cdsLibrary.first;
      DelphiPackage.WriteString ('Show Compiler Progress', 'True');
    end;
    if AbreRegistroDelphi (DelphiPackage, 'Form Design') and
      (cdsLibrary.RecordCount > 0) then
    begin
      cdsLibrary.first;
      DelphiPackage.WriteString ('Auto Create Forms', 'False');
      DelphiPackage.WriteInteger('Grid Size X', 2);
      DelphiPackage.WriteInteger('Grid Size Y', 2);
      DelphiPackage.WriteString('Snap to Grid', '-1');
      DelphiPackage.WriteString('Display Grid', '-1');
    end;
  finally
    DelphiPackage.free;
  end;
end;

end.

