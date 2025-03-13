unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics, WinSvc, Vcl.Controls, Vcl.Forms,
  Vcl.Dialogs, TLHelp32, Vcl.StdCtrls, MaskUtils, Vcl.ExtCtrls, Shellapi, Vcl.ComCtrls, Vcl.Mask, Vcl.AppEvnts, Vcl.Imaging.jpeg,
  IniFiles;

type
  TFrm_Main = class(TForm)
    Tm_Hora: TTimer;
    Btn_Iniciar: TButton;
    Btn_Sair: TButton;
    Lbl_Hora_Agora: TLabel;
    Msk_Hora_Informada: TMaskEdit;
    Lbl_Hora_Execução: TLabel;
    Lbl_Status_Execução: TLabel;
    lbl_Status: TLabel;
    Btn_Parar: TButton;
    Tray_Esconder_Form: TTrayIcon;
    App_Event_Fechar_Form: TApplicationEvents;
    Img_Fundo: TImage;
    btnIniciarServicos: TButton;
    btnPararServicos: TButton;
    procedure Tm_HoraTimer(Sender: TObject);
    procedure Btn_SairClick(Sender: TObject);
    procedure Btn_IniciarClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Btn_PararClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure Tray_Esconder_FormDblClick(Sender: TObject);
    procedure btnPararServicosClick(Sender: TObject);
    procedure btnIniciarServicosClick(Sender: TObject);
  private
    procedure Split(Delimitador: Char; Texto: String; Lista: TStringList);
  public
    { Public declarations }
  end;

var
  Frm_Main: TFrm_Main;
  continua: Boolean;
  Hora_Atual: String;

implementation

{$R *.dfm}

function KillTask(ExeFileName: String): Integer; // Função para matar processo informado
const
  PROCESS_TERMINATE = $0001;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
begin
  Result := 0;

  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);

  FProcessEntry32.dwSize := SizeOf(FProcessEntry32);

  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);

  while Integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = UpperCase(ExeFileName)) or (UpperCase(FProcessEntry32.szExeFile) = UpperCase(ExeFileName))) then
      Result := Integer(TerminateProcess(OpenProcess(PROCESS_TERMINATE, BOOL(0), FProcessEntry32.th32ProcessID), 0));

    ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
  end;

  CloseHandle(FSnapshotHandle);
end;

function DeleteFile(const FileName: String): Boolean;
begin
  Result := Winapi.Windows.DeleteFile(PChar(FileName));
end;

function ServiceStart(sMachine, sService: String): Boolean;
var
  schSCManager, schService: SC_HANDLE;
  ssStatus: TServiceStatus;
  dwWaitTime: Integer;
  lpServiceArgVectors: LPCWSTR;
begin
  schSCManager := OpenSCManager(PChar(sMachine), nil, SC_MANAGER_CONNECT);

  if (schSCManager = 0) then
    RaiseLastOSError;

  try
    schService := OpenService(schSCManager, PChar(sService),SERVICE_START or SERVICE_QUERY_STATUS);

    if (schService = 0) then
      RaiseLastOSError;

    try
      // checa o status caso o serviço não pare.
      if not QueryServiceStatus(schService, ssStatus) then
      begin
        if (ERROR_SERVICE_NOT_ACTIVE <> GetLastError()) then
          RaiseLastOSError;

        ssStatus.dwCurrentState := SERVICE_STOPPED;
      end;

      // Checa se o serviço continua rodando
      if (ssStatus.dwCurrentState <> SERVICE_STOPPED) and (ssStatus.dwCurrentState <> SERVICE_STOP_PENDING) then
      begin
        Result := True;

        Exit;
      end;

      // Espera o serviço parar antes de iniciar novamente.

      while (ssStatus.dwCurrentState = SERVICE_STOP_PENDING) do
      begin
        // Faz um intervalo para que não demore mais do que 10 seguundo.
        dwWaitTime := ssStatus.dwWaitHint div 10;

        if (dwWaitTime < 1000) then
          dwWaitTime := 1000
        else if (dwWaitTime > 10000) then
          dwWaitTime := 10000;

        Sleep(dwWaitTime);

        // Verifique o status até que o serviço não pare mais de aguardar.
        if not QueryServiceStatus(schService, ssStatus) then
        begin
          if (ERROR_SERVICE_NOT_ACTIVE <> GetLastError()) then
            RaiseLastOSError;

          Break;
        end;
      end;

      // Tentar iniciar o serviço.
      lpServiceArgVectors := nil;

      if not StartService(schService, 0, lpServiceArgVectors) then
          RaiseLastOSError;

      // Determina qulaquer serviço que esteja rodando.
      Result := (ssStatus.dwCurrentState = SERVICE_RUNNING);
    finally
      CloseServiceHandle(schService);
    end;
  finally
    CloseServiceHandle(schSCManager);
  end;
end;

// Função que para o serviço selcionado
function ServiceStop(aMachineName, aServiceName: string): boolean;
var
  schm,schs: SC_Handle;
  ss: TServiceStatus;
  dwChkP: DWord;
begin
  schm := OpenSCManager(PChar(aMachineName), nil, SC_MANAGER_CONNECT);

  if (schm > 0) then begin
    schs := OpenService(schm,  PChar(aServiceName), SERVICE_STOP or SERVICE_QUERY_STATUS);

    if (schs > 0) then  begin
      if (ControlService(schs, SERVICE_CONTROL_STOP, ss)) then begin
        if (QueryServiceStatus(schs,ss)) then begin
          while (SERVICE_STOPPED<> ss.dwCurrentState) do begin
            dwChkP := ss.dwCheckPoint;

            Sleep(ss.dwWaitHint);

            if (not QueryServiceStatus(schs,ss)) then break;

            if (ss.dwCheckPoint < dwChkP) then break;
          end;
        end;
      end;

      CloseServiceHandle(schs);
    end;

    CloseServiceHandle(schm);
  end;

  Result := SERVICE_STOPPED = ss.dwCurrentState;
end;

procedure TFrm_Main.Btn_IniciarClick(Sender: TObject);
begin
  continua           := True;
  Tm_Hora.Enabled    := True; // Habilita o inicio do relogio
  lbl_Status.Caption := 'EM EXECUÇÃO'; // Altera o caption do status
end;

procedure TFrm_Main.Btn_PararClick(Sender: TObject);
begin
  continua           := False; // Altera a propriedade do Timer
  lbl_Status.Caption := 'PARADO'; // Altera o caption do status
end;

procedure TFrm_Main.Btn_SairClick(Sender: TObject);
begin
  Application.Terminate; // Finaliza o form
end;

procedure TFrm_Main.btnIniciarServicosClick(Sender: TObject);
var
  ArqINI: TIniFile;
  meuteste, Servicos: String;
  Lista: TStringList;
  I: Integer;
begin
  meuteste := ExtractFilePath(Application.ExeName) + 'Config.ini';
  ArqINI := TIniFile.Create(meuteste);

  // Iniciar todos os serviços informados
  Servicos := ArqINI.ReadString('Servicos', 'Servicos', '');

  Lista := TStringList.Create;
  try
    Split('/', Servicos, Lista) ;

    for I := 0 to Lista.Count-1 do
    begin
      Sleep(1000);
      try
        ServiceStart('', Lista[I]);
      except
        //
      end;
    end;
    ShowMessage('Todos os serviços iniciados!');
  finally
    Lista.Free;
  end;
end;

procedure TFrm_Main.btnPararServicosClick(Sender: TObject);
var
  ArqINI: TIniFile;
  meuteste, Servicos: String;
  Lista: TStringList;
  I: Integer;
begin
  meuteste := ExtractFilePath(Application.ExeName) + 'Config.ini';
  ArqINI := TIniFile.Create(meuteste);

  // Para todos os serviços informados
  Servicos := ArqINI.ReadString('Servicos', 'Servicos', '');

  Lista := TStringList.Create;
  try
    Split('/', Servicos, Lista) ;

    for I := 0 to Lista.Count-1 do
      ServiceStop('', Lista[I]);
    ShowMessage('Todos os serviços parados!');
  finally
    Lista.Free;
  end;
end;

procedure TFrm_Main.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := False; // Impedir o fechamento do form
  Self.Hide(); // esconder o fomr
  Self.WindowState := wsMinimized; // deixar o staus da janela minimizado
  Tray_Esconder_Form.Visible := true; // deixar o icone do form visivel na barra
end;

procedure TFrm_Main.FormShow(Sender: TObject);
begin
  continua           := True; // Altera a propriedade do Timer
  lbl_Status.Caption := 'AGUARDANDO'; // Altera o caption do status
  Tm_Hora.Enabled    := True; // Habilita o inicio do relogio
end;

procedure TFrm_Main.Split(Delimitador: Char; Texto: String; Lista: TStringList);
begin
  Lista.Clear;
  Lista.Delimiter       := Delimitador;
  Lista.StrictDelimiter := True; // Requires D2006 or newer.
  Lista.DelimitedText   := Texto;
end;

procedure TFrm_Main.Tm_HoraTimer(Sender: TObject);
var
  ArqINI: TIniFile;
  meuteste, Servicos: String;
  Lista: TStringList;
  I: Integer;
begin
  Lbl_Hora_Agora.Caption := TimeToStr(Now); // Mostra hora atual

  if continua then
  begin
    Hora_Atual := FormatDateTime('hh:MM:ss', Now); // Pega a hora em tempo real da maquina

    if (Msk_Hora_Informada.Text = Hora_Atual) then // Verifica se hora atual é igual a hora informada
    begin
      try
        try
          meuteste := ExtractFilePath(Application.ExeName) + 'Config.ini';

          ArqINI := TIniFile.Create(meuteste);

          // Para todos os serviços informados
          Servicos := ArqINI.ReadString('Servicos', 'Servicos', '');

          Lista := TStringList.Create;
          try
            Split('/', Servicos, Lista) ;

            for I := 0 to Lista.Count-1 do
              ServiceStop('', Lista[I]);
          finally
            Lista.Free;
          end;
          (*
          ServiceStop('', 'ServicoAgendamentos');
          ServiceStop('', 'srvSite_Clientes_Cadastros');
          ServiceStop('', 'srvSite_Clientes');
          ServiceStop('', 'srvSite_Cobranca');
          ServiceStop('', 'srvSite_Estoque');
          ServiceStop('', 'srvSite_Impostos');
          ServiceStop('', 'srvSitePedidos');
          ServiceStop('', 'srvSite_Precos');
          ServiceStop('', 'srvSite_Produtos');
          ServiceStop('', 'srvSite_RecebePedido');
          ServiceStop('', 'IntegracaoWSMServicos');
          ServiceStop('', 'srvAuditorias');
          ServiceStop('', 'srvSite_Promocionais');
          ServiceStop('', 'srvIntegracao_Magis_Estoque');
          ServiceStop('', 'srvIntegracao_Magis_V3');
          *)



          KillTask('AppServer.exe');
          KillTask('Impacto.VCL.API.exe');



          PostMessage(FindWindow('AppClient', nil), WM_CLOSE, 0, 0);

          if FileExists(ArqINI.ReadString('Configurations', 'Original', '')) and FileExists(ArqINI.ReadString('Configurations', 'Atualizacao', '')) then // verifica o caminho e a existencia do arquivo
          begin
            // Server Padrão
            RenameFile(ArqINI.ReadString('Configurations', 'Original', ''), ArqINI.ReadString('Configurations', 'Renomeado', '')+ FormatDateTime('-dd-MM-yyyy', Now) + '.bpl'); // Renomeia o arquivo de origem

            CopyFile(PWideChar(ArqINI.ReadString('Configurations', 'Atualizacao', '')), PWideChar(ArqINI.ReadString('Configurations', 'Original', '')), True); // Copia o arquivo das pastas informadas

            // Serve Vendas
            RenameFile(ArqINI.ReadString('Configurations', 'OriginalVendas', ''), ArqINI.ReadString('Configurations', 'RenomeadoVendas', '')+ FormatDateTime('-dd-MM-yyyy', Now) + '.bpl'); // Renomeia o arquivo de origem

            CopyFile(PWideChar(ArqINI.ReadString('Configurations', 'Atualizacao', '')), PWideChar(ArqINI.ReadString('Configurations', 'OriginalVendas', '')), True); // Copia o arquivo das pastas informadas

            // Server Estoque
            RenameFile(ArqINI.ReadString('Configurations', 'OriginalEstoque', ''), ArqINI.ReadString('Configurations', 'RenomeadoEstoque', '')+ FormatDateTime('-dd-MM-yyyy', Now) + '.bpl'); // Renomeia o arquivo de origem

            CopyFile(PWideChar(ArqINI.ReadString('Configurations', 'Atualizacao', '')), PWideChar(ArqINI.ReadString('Configurations', 'OriginalEstoque', '')), True); // Copia o arquivo das pastas informadas


            DeleteFile('C:\Users\Administrator\Desktop\Arquivos_Atualizar_AppServer\ASCustoms.bpl');
          end;

          // O sleep server para impedir que o programa fique executando em loop e aguarda o serviço parar
          Sleep(25000);
          ShellExecute(Handle, 'open', PChar(ArqINI.ReadString('Configurations', 'Server', '')), nil, nil, sw_show); // Server Padrão do Impacto

          Sleep(2500);
          ShellExecute(Handle, 'open', PChar(ArqINI.ReadString('Configurations', 'ServerFabrica', '')), nil, nil, sw_show); // Server da Fábrica

          Sleep(2500);
          ShellExecute(Handle, 'open', PChar(ArqINI.ReadString('Configurations', 'ServerWR', '')), nil, nil, sw_show); // Server da WR

          Sleep(2500);
          ShellExecute(Handle, 'open', PChar(ArqINI.ReadString('Configurations', 'API_Impacto', '')), nil, nil, sw_show); // API do Impacto

          Sleep(2500);
          ShellExecute(Handle, 'open', PChar(ArqINI.ReadString('Configurations', 'ServerLicencas', '')), nil, nil, sw_show); // Server de Licenças do SFA

          Sleep(2500);
          ShellExecute(Handle, 'open', PChar(ArqINI.ReadString('Configurations', 'ServerTeste', '')), nil, nil, sw_show); // Server de Teste

          Sleep(2500);
          ShellExecute(Handle, 'open', PChar(ArqINI.ReadString('Configurations', 'ServerVendas', '')), nil, nil, sw_show); // Server de Vendas

          Sleep(2500);
          ShellExecute(Handle, 'open', PChar(ArqINI.ReadString('Configurations', 'ServerEstoque', '')), nil, nil, sw_show); // Server do Estoque


          // Iniciar todos os serviços informados
          Servicos := ArqINI.ReadString('Servicos', 'Servicos', '');

          Lista := TStringList.Create;
          try
            Split('/', Servicos, Lista) ;

            for I := 0 to Lista.Count-1 do
            begin
              Sleep(1000);
              try
                ServiceStart('', Lista[I]);
              except
                //
              end;
            end;
          finally
            Lista.Free;
          end;
          (*
          ServiceStart('', 'ServicoAgendamentos');
          ServiceStart('', 'srvSite_Clientes_Cadastros');
          ServiceStart('', 'srvSite_Clientes');
          ServiceStart('', 'srvSite_Cobranca');
          ServiceStart('', 'srvSite_Estoque');
          ServiceStart('', 'srvSite_Impostos');
          ServiceStart('', 'srvSitePedidos');
          ServiceStart('', 'srvSite_Precos');
          ServiceStart('', 'srvSite_Produtos');
          ServiceStart('', 'srvSite_RecebePedido');
          ServiceStart('', 'IntegracaoWSMServicos');
          ServiceStart('', 'srvAuditorias');
          ServiceStart('', 'srvSite_Promocionais');
          ServiceStart('', 'srvIntegracao_Magis_Estoque');
          ServiceStart('', 'srvIntegracao_Magis_V3');
          *)
        finally
          ArqINI.Free;
        end;
      except
        on E: Exception do
        begin
          ShowMessage('ERRO: ' + E.Message);

          ArqINI.Free;
        end;
      end;
    end;
  end;
end;

procedure TFrm_Main.Tray_Esconder_FormDblClick(Sender: TObject);
begin
  Tray_Esconder_Form.Visible := False; // deixar o icone do form oculto na barra
  Show(); // mostrar form
  WindowState := wsNormal; // deixar o staus da janela normal
  Application.BringToFront(); // Trazer aplicação para a tela principal
end;

end.
