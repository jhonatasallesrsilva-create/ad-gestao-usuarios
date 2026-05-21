# ============================================================
# EMPRESA TI - GESTAO DE USUARIOS
# Criacao + Desligamento de Colaborador
# Interface Unificada com Abas
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class Win32UI {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, string lParam);
}
"@

# Forcar TLS 1.2 (exigido pela Microsoft Graph API)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ============================================================
# REMOVER ACENTOS (corrige nomes como Joao, Angela, etc.)
# ============================================================
function Remove-Acentos {
    param([string]$Texto)
    $n = $Texto.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = [System.Text.StringBuilder]::new()
    foreach ($c in $n.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne
            [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    return $sb.ToString().Normalize([System.Text.NormalizationForm]::FormC)
}

# ============================================================
# DEFAULTS - sobrescritos pelo config.xml na inicializacao
# ============================================================
# ============================================================
# CONFIGURACOES FIXAS - Edite aqui para nao precisar configurar
# manualmente a cada execucao. As credenciais Graph API ficam
# salvas criptografadas (DPAPI) no config.xml apos o primeiro uso.
# ============================================================

# --- Microsoft Graph API (Azure AD App Registration) ---
# Preencha APENAS uma vez. Apos salvar, nao precisa mais.
$CFG_TenantId     = "SEU_TENANT_ID_AQUI"
$CFG_ClientId     = "SEU_CLIENT_ID_AQUI"
$CFG_ClientSecret = "SEU_CLIENT_SECRET_AQUI"

# --- Configuracoes Gerais ---
$CFG_Dominio      = "suaempresa.com.br"
$CFG_SenhaInicial = "@SenhaInicial2025"
$CFG_Telefone     = "(XX)XXXX-XXXX"
$CFG_PaginaWeb    = "www.suaempresa.com.br"
$CFG_UsageLocation= "BR"
$CFG_TargetOU     = "OU=Usuarios Desabilitados,DC=suaempresa,DC=com,DC=br"
$CFG_Servidor     = "IP_DO_SERVIDOR_AQUI"

# --- Credencial do servidor AD / AAD Connect ---
$CFG_SyncUser     = "DOMINIO\administrador"   # Use DOMINIO\usuario ou .\usuario para conta local
$CFG_SyncSenha    = 'SENHA_DO_SERVIDOR_AQUI'

# ============================================================
# DEFAULTS internos (nao altere - usam os valores acima)
# ============================================================
$Dominio          = $CFG_Dominio
$SenhaInicial     = $CFG_SenhaInicial
$Telefone         = $CFG_Telefone
$PaginaWeb        = $CFG_PaginaWeb
$UsageLocation    = $CFG_UsageLocation
$TargetOU         = $CFG_TargetOU
$AADConnectServer = $CFG_Servidor
$AADSyncUserDefault  = $CFG_SyncUser
$AADSyncSenhaDefault = $CFG_SyncSenha
$LogDir           = "$PSScriptRoot\logs"
$HistCriacaoCsv   = "$PSScriptRoot\logs\historico_criacao.csv"
$HistDesligCsv    = "$PSScriptRoot\logs\historico_desligamentos.csv"
$ConfigFile       = "$PSScriptRoot\config.xml"

# ============================================================
# SESSAO AD REMOTA - funciona em qualquer maquina sem RSAT
# ============================================================
$script:ADSession = $null
$script:ADRemote  = $false

# ============================================================
# CREDENCIAIS + CONFIG GERAL (criptografados via DPAPI)
# ============================================================
function Get-StoredCredentials {
    if (Test-Path $ConfigFile) {
        try {
            $cfg = Import-Clixml -Path $ConfigFile
            if (-not $cfg -or -not $cfg.ClientSecret) { throw "Config invalido" }
            return @{
                TenantId         = $cfg.TenantId
                ClientId         = $cfg.ClientId
                ClientSecret     = $cfg.ClientSecret | ConvertTo-SecureString
                Dominio          = if ($cfg.Dominio)          { $cfg.Dominio }          else { $script:Dominio }
                AADConnectServer = if ($cfg.AADConnectServer) { $cfg.AADConnectServer } else { $script:AADConnectServer }
                TargetOU         = if ($cfg.TargetOU)         { $cfg.TargetOU }         else { $script:TargetOU }
                Telefone         = if ($cfg.Telefone)         { $cfg.Telefone }         else { $script:Telefone }
                PaginaWeb        = if ($cfg.PaginaWeb)        { $cfg.PaginaWeb }        else { $script:PaginaWeb }
                SenhaInicial     = if ($cfg.SenhaInicial)     { $cfg.SenhaInicial | ConvertTo-SecureString } `
                                   else { ConvertTo-SecureString $script:SenhaInicial -AsPlainText -Force }
                AADSyncUser      = if ($cfg.AADSyncUser)      { $cfg.AADSyncUser }      else { "" }
                AADSyncSenha     = if ($cfg.AADSyncSenha)     { $cfg.AADSyncSenha | ConvertTo-SecureString } else { $null }
            }
        } catch {
            Remove-Item -Path $ConfigFile -Force -ErrorAction SilentlyContinue
            return $null
        }
    }
    return $null
}

function Save-Credentials {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [System.Security.SecureString]$ClientSecret,
        [string]$Dominio,
        [string]$AADConnectServer,
        [string]$TargetOU,
        [string]$Telefone,
        [string]$PaginaWeb,
        [System.Security.SecureString]$SenhaInicial,
        [string]$AADSyncUser = "",
        [System.Security.SecureString]$AADSyncSenha = $null
    )
    @{
        TenantId         = $TenantId
        ClientId         = $ClientId
        ClientSecret     = $ClientSecret | ConvertFrom-SecureString
        Dominio          = $Dominio
        AADConnectServer = $AADConnectServer
        TargetOU         = $TargetOU
        Telefone         = $Telefone
        PaginaWeb        = $PaginaWeb
        SenhaInicial     = $SenhaInicial | ConvertFrom-SecureString
        AADSyncUser      = $AADSyncUser
        AADSyncSenha     = if ($AADSyncSenha) { $AADSyncSenha | ConvertFrom-SecureString } else { $null }
    } | Export-Clixml -Path $ConfigFile -Force
}

function Resolve-SecureStringPlain {
    param([System.Security.SecureString]$SecStr)
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecStr)
    $plain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
    return $plain
}

# ============================================================
# ENSURE-ADMODULE - Carrega AD local ou via PSSession remota.
# Funciona em qualquer maquina Windows no dominio.
# ============================================================
$script:ADModuleError = ""

function Ensure-ADModule {
    $script:ADModuleError = ""

    # 1. Ja carregado?
    if (Get-Module -Name ActiveDirectory) { return $true }

    # 2. Disponivel localmente (RSAT)?
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        try { Import-Module ActiveDirectory -ErrorAction Stop; $script:ADRemote = $false; return $true } catch {}
    }

    $srv = $script:AADConnectServer
    if ([string]::IsNullOrWhiteSpace($srv)) {
        $script:ADModuleError = "Servidor nao configurado. Clique na engrenagem (canto superior direito) e preencha o campo 'Servidor principal'."
        return $false
    }

    # Reutilizar sessao existente se ainda aberta
    if ($script:ADSession -and $script:ADSession.State -eq "Opened") {
        if (Get-Command Get-ADUser -ErrorAction SilentlyContinue) { return $true }
    }

    # 3. PSSession remota para o servidor
    # Preparar TrustedHosts (necessario para NTLM via IP)
    try {
        $thCurrent = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction Stop).Value
        if ($thCurrent -ne "*" -and $thCurrent -notlike "*$srv*") {
            $thNew = if ([string]::IsNullOrWhiteSpace($thCurrent)) { $srv } else { "$thCurrent,$srv" }
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value $thNew -Force -ErrorAction Stop
        }
    } catch { }

    if ($script:ADSession) {
        Remove-PSSession $script:ADSession -ErrorAction SilentlyContinue
        $script:ADSession = $null
    }

    # Preparar credencial com varios formatos de usuario
    $credList = @()
    if ($script:AADSyncUser -and $script:AADSyncSenha) {
        $u = $script:AADSyncUser
        # Gerar variantes do nome de usuario
        $variants = @($u)
        if ($u -notmatch '\|@') {
            # sem dominio: tentar conta local e com dominio extraido do proprio servidor
            $variants = @(".\$u", $u)
        }
        foreach ($v in $variants) {
            $credList += New-Object System.Management.Automation.PSCredential($v, $script:AADSyncSenha)
        }
    }

    # Resolver hostname a partir do IP (Kerberos funciona com hostname, nao com IP)
    $srvHost = $srv
    try {
        $dns = [System.Net.Dns]::GetHostEntry($srv)
        if ($dns.HostName -and $dns.HostName -ne $srv) {
            $srvHost = $dns.HostName
        }
    } catch { }

    $allErrs = [System.Collections.ArrayList]::new()

    function Try-PSSession($computer, $auth, $cred) {
        try {
            $sp = @{ ComputerName = $computer; ErrorAction = "Stop" }
            if ($auth)  { $sp["Authentication"] = $auth }
            if ($cred)  { $sp["Credential"]     = $cred }
            $sess = New-PSSession @sp
            Import-PSSession -Session $sess -Module ActiveDirectory `
                -AllowClobber -DisableNameChecking -ErrorAction Stop | Out-Null
            return $sess
        } catch {
            $who  = if ($cred) { $cred.UserName } else { "(usuario logado)" }
            $host2 = if ($computer -eq $srv) { $computer } else { "$computer (hostname)" }
            [void]$allErrs.Add("Auth=$(if($auth){$auth}else{'Default'}) Host=$host2 User=$who => $($_.Exception.Message)")
            return $null
        }
    }

    # TENTATIVA 1: hostname + sem credencial (Kerberos com usuario logado - melhor opcao)
    if ($srvHost -ne $srv) {
        $sess = Try-PSSession $srvHost $null $null
        if ($sess) { $script:ADSession = $sess; $script:ADRemote = $true; return $true }
    }

    # TENTATIVA 2: hostname + credenciais explicitas (Kerberos + credencial)
    if ($srvHost -ne $srv -and $credList.Count -gt 0) {
        foreach ($cred in $credList) {
            $sess = Try-PSSession $srvHost "Negotiate" $cred
            if ($sess) { $script:ADSession = $sess; $script:ADRemote = $true; return $true }
        }
    }

    # TENTATIVA 3: IP + sem credencial (NegotiateWithImplicitCredential - usuario logado via NTLM)
    $sess = Try-PSSession $srv "NegotiateWithImplicitCredential" $null
    if ($sess) { $script:ADSession = $sess; $script:ADRemote = $true; return $true }

    # TENTATIVA 4: IP + credenciais explicitas + Negotiate (NTLM)
    foreach ($cred in $credList) {
        $sess = Try-PSSession $srv "Negotiate" $cred
        if ($sess) { $script:ADSession = $sess; $script:ADRemote = $true; return $true }
    }

    # TENTATIVA 5: IP + credenciais + Default
    foreach ($cred in $credList) {
        $sess = Try-PSSession $srv $null $cred
        if ($sess) { $script:ADSession = $sess; $script:ADRemote = $true; return $true }
    }

    $script:ADModuleError = ($allErrs | ForEach-Object { "  - $_" }) -join "`n"

    $script:ADSession     = $null
    $script:ADRemote      = $false
    $script:ADModuleError = $lastErr
    return $false
}

function Show-CredentialForm {
    $cf = New-Object System.Windows.Forms.Form
    $cf.Text        = "Configuracoes - Empresa TI Gestao de Usuarios"
    $cf.Size        = New-Object System.Drawing.Size(520, 730)
    $cf.StartPosition   = "CenterScreen"
    $cf.FormBorderStyle = "FixedDialog"
    $cf.MaximizeBox = $false; $cf.MinimizeBox = $false
    $cf.BackColor   = [System.Drawing.Color]::FromArgb(0, 53, 92)
    $cf.ForeColor   = [System.Drawing.Color]::White
    $cf.Font        = New-Object System.Drawing.Font("Segoe UI", 10)
    $ci = [System.Drawing.Color]::FromArgb(0, 68, 110)

    function _Lbl($text, $y, $bold=$false, $color=$null) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text
        $l.Font = if($bold){ New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold) } `
                  else      { New-Object System.Drawing.Font("Segoe UI",9.5) }
        $l.ForeColor = if($color){ $color } else { [System.Drawing.Color]::FromArgb(180,180,200) }
        $l.AutoSize  = $true
        $l.Location  = New-Object System.Drawing.Point(20,$y)
        $cf.Controls.Add($l)
    }
    function _Inp($name, $y, $pw=$false) {
        $t = New-Object System.Windows.Forms.TextBox
        $t.Name        = $name
        $t.Size        = New-Object System.Drawing.Size(460,28)
        $t.Location    = New-Object System.Drawing.Point(20,$y)
        $t.BackColor   = $ci
        $t.ForeColor   = [System.Drawing.Color]::White
        $t.BorderStyle = "FixedSingle"
        if ($pw) { $t.UseSystemPasswordChar = $true }
        $cf.Controls.Add($t)
        return $t
    }

    # --- SECAO: GRAPH API ---
    _Lbl "  Credenciais Microsoft Graph API" 15 $true ([System.Drawing.Color]::FromArgb(0,120,212))
    _Lbl "App Registration do Azure AD - salvos criptografados (DPAPI)." 38

    _Lbl "Tenant ID:"     65;  _Inp "cred_T" 83  | Out-Null
    _Lbl "Client ID:"     118; _Inp "cred_C" 136 | Out-Null
    _Lbl "Client Secret:" 171; _Inp "cred_S" 189 $true | Out-Null

    # Separador
    $sep1 = New-Object System.Windows.Forms.Label
    $sep1.Size = New-Object System.Drawing.Size(460,1); $sep1.Location = New-Object System.Drawing.Point(20,228)
    $sep1.BackColor = [System.Drawing.Color]::FromArgb(80,84,110); $cf.Controls.Add($sep1)

    # --- SECAO: CONFIG GERAL ---
    _Lbl "  Configuracoes Gerais" 237 $true ([System.Drawing.Color]::FromArgb(46,204,113))

    _Lbl "Dominio AD (ex: empresa.com.br):"                    262; _Inp "cfg_Dom" 280 | Out-Null
    _Lbl "Senha Padrao de Criacao:"                            315; _Inp "cfg_Sen" 333 $true | Out-Null
    _Lbl "(em branco = manter senha ja salva)" 365
    _Lbl "Servidor principal (DC/AADConnect - usado para AD remoting):"             383; _Inp "cfg_AAD" 401 | Out-Null
    _Lbl "OU de Usuarios Desabilitados (DN completo):"        436; _Inp "cfg_OU"  454 | Out-Null
    _Lbl "Telefone padrao:"                                    489; _Inp "cfg_Tel" 507 | Out-Null

    # Separador ADSync
    $sep2 = New-Object System.Windows.Forms.Label
    $sep2.Size = New-Object System.Drawing.Size(460,1); $sep2.Location = New-Object System.Drawing.Point(20,540)
    $sep2.BackColor = [System.Drawing.Color]::FromArgb(80,84,110); $cf.Controls.Add($sep2)

    _Lbl "  Credencial do Servidor (AD Remoting + AAD Connect Sync)" 548 $true ([System.Drawing.Color]::FromArgb(241,196,15))
    _Lbl "Usuario admin do servidor (ex: DOMINIO\Administrador) - para remoting AD e AAD Sync:" 572
    _Inp "cfg_SyncUsr" 590 | Out-Null
    _Lbl "Senha do servidor AAD Connect:" 623; _Inp "cfg_SyncPwd" 641 $true | Out-Null
    _Lbl "(deixe em branco para manter senha ja salva)" 673

    # Botao salvar
    $btnS = New-Object System.Windows.Forms.Button
    $btnS.Text     = "SALVAR E CONTINUAR"
    $btnS.Font     = New-Object System.Drawing.Font("Segoe UI",11,[System.Drawing.FontStyle]::Bold)
    $btnS.Size     = New-Object System.Drawing.Size(460,40)
    $btnS.Location = New-Object System.Drawing.Point(20,690)
    $btnS.FlatStyle = "Flat"; $btnS.FlatAppearance.BorderSize = 0
    $btnS.BackColor = [System.Drawing.Color]::FromArgb(0,120,212)
    $btnS.ForeColor = [System.Drawing.Color]::White
    $btnS.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnS.Add_Click({
        $tT       = $cf.Controls["cred_T"].Text.Trim()
        $tC       = $cf.Controls["cred_C"].Text.Trim()
        $tS       = $cf.Controls["cred_S"].Text.Trim()
        $tDom     = $cf.Controls["cfg_Dom"].Text.Trim()
        $tSen     = $cf.Controls["cfg_Sen"].Text.Trim()
        $tAAD     = $cf.Controls["cfg_AAD"].Text.Trim()
        $tOU      = $cf.Controls["cfg_OU"].Text.Trim()
        $tTel     = $cf.Controls["cfg_Tel"].Text.Trim()
        $tSyncUsr = $cf.Controls["cfg_SyncUsr"].Text.Trim()
        $tSyncPwd = $cf.Controls["cfg_SyncPwd"].Text.Trim()

        if ([string]::IsNullOrWhiteSpace($tT) -or [string]::IsNullOrWhiteSpace($tC)) {
            [System.Windows.Forms.MessageBox]::Show("Preencha Tenant ID e Client ID.","Erro","OK","Warning"); return }
        if ([string]::IsNullOrWhiteSpace($tDom)) {
            [System.Windows.Forms.MessageBox]::Show("Preencha o Dominio AD.","Erro","OK","Warning"); return }

        $ex = Get-StoredCredentials

        $secSecret = if ([string]::IsNullOrWhiteSpace($tS)) {
            if ($ex) { $ex.ClientSecret }
            else { [System.Windows.Forms.MessageBox]::Show("Informe o Client Secret.","Erro","OK","Warning"); return }
        } else { ConvertTo-SecureString $tS -AsPlainText -Force }

        $secSenha = if ([string]::IsNullOrWhiteSpace($tSen)) {
            if ($ex) { $ex.SenhaInicial }
            else { ConvertTo-SecureString $script:SenhaInicial -AsPlainText -Force }
        } else { ConvertTo-SecureString $tSen -AsPlainText -Force }

        # Credencial ADSync
        $secSyncPwd = if ([string]::IsNullOrWhiteSpace($tSyncPwd)) {
            if ($ex -and $ex.AADSyncSenha) { $ex.AADSyncSenha } else { $null }
        } else { ConvertTo-SecureString $tSyncPwd -AsPlainText -Force }

        $syncUsrFinal = if ($tSyncUsr) { $tSyncUsr } elseif ($ex -and $ex.AADSyncUser) { $ex.AADSyncUser } else { "" }

        $tAAD = if($tAAD){ $tAAD } else { $script:AADConnectServer }
        $tOU  = if($tOU ) { $tOU  } else { $script:TargetOU }
        $tTel = if($tTel) { $tTel } else { $script:Telefone }

        Save-Credentials -TenantId $tT -ClientId $tC -ClientSecret $secSecret `
            -Dominio $tDom -AADConnectServer $tAAD -TargetOU $tOU `
            -Telefone $tTel -PaginaWeb $script:PaginaWeb -SenhaInicial $secSenha `
            -AADSyncUser $syncUsrFinal -AADSyncSenha $secSyncPwd

        $cf.Tag = @{
            TenantId=$tT; ClientId=$tC; ClientSecret=$secSecret
            Dominio=$tDom; AADConnectServer=$tAAD; TargetOU=$tOU
            Telefone=$tTel; PaginaWeb=$script:PaginaWeb; SenhaInicial=$secSenha
            AADSyncUser=$syncUsrFinal; AADSyncSenha=$secSyncPwd
        }
        $cf.DialogResult = "OK"; $cf.Close()
    })
    $cf.Controls.Add($btnS)

    # Pre-preencher com valores existentes
    $ex = Get-StoredCredentials
    if ($ex) {
        $cf.Controls["cred_T"].Text    = $ex.TenantId
        $cf.Controls["cred_C"].Text    = $ex.ClientId
        $cf.Controls["cfg_Dom"].Text   = $ex.Dominio
        $cf.Controls["cfg_AAD"].Text   = $ex.AADConnectServer
        $cf.Controls["cfg_OU"].Text    = $ex.TargetOU
        $cf.Controls["cfg_Tel"].Text   = $ex.Telefone
        $cf.Controls["cfg_SyncUsr"].Text = if ($ex.AADSyncUser) { $ex.AADSyncUser } else { "" }
    } else {
        $cf.Controls["cfg_Dom"].Text   = $script:Dominio
        $cf.Controls["cfg_AAD"].Text   = $script:AADConnectServer
        $cf.Controls["cfg_OU"].Text    = $script:TargetOU
        $cf.Controls["cfg_Tel"].Text   = $script:Telefone
    }

    if ($cf.ShowDialog() -eq "OK") { return $cf.Tag }
    return $null
}

# ============================================================
# CARREGAR CREDENCIAIS E CONFIGS
# Prioridade: 1) valores CFG_ no script  2) config.xml  3) formulario
# ============================================================
$script:CFG_Preenchido = (
    $CFG_TenantId     -ne "COLE_SEU_TENANT_ID_AQUI" -and
    $CFG_ClientId     -ne "COLE_SEU_CLIENT_ID_AQUI"  -and
    $CFG_ClientSecret -ne "COLE_SEU_CLIENT_SECRET_AQUI"
)

if ($script:CFG_Preenchido) {
    # Valores fixos no script -> sempre usa, recria config.xml (evita problemas de DPAPI entre maquinas)
    $secSecret = ConvertTo-SecureString $CFG_ClientSecret -AsPlainText -Force
    $secSenha  = ConvertTo-SecureString $CFG_SenhaInicial -AsPlainText -Force
    $secSync   = ConvertTo-SecureString $CFG_SyncSenha   -AsPlainText -Force

    # Remover config.xml antigo se existir (pode ser de outra maquina)
    if (Test-Path $ConfigFile) { Remove-Item $ConfigFile -Force -ErrorAction SilentlyContinue }

    Save-Credentials `
        -TenantId         $CFG_TenantId `
        -ClientId         $CFG_ClientId `
        -ClientSecret     $secSecret `
        -Dominio          $CFG_Dominio `
        -AADConnectServer $CFG_Servidor `
        -TargetOU         $CFG_TargetOU `
        -Telefone         $CFG_Telefone `
        -PaginaWeb        $CFG_PaginaWeb `
        -SenhaInicial     $secSenha `
        -AADSyncUser      $CFG_SyncUser `
        -AADSyncSenha     $secSync

    $creds = Get-StoredCredentials
    if (-not $creds) {
        [System.Windows.Forms.MessageBox]::Show(
            "Erro ao salvar configuracoes fixas.`nTente executar como Administrador.",
            "Erro","OK","Error"); exit
    }
} else {
    # CFG_ nao preenchidos -> tenta config.xml salvo
    $creds = Get-StoredCredentials

    if (-not $creds) {
        # Nenhum config disponivel -> mostrar formulario (apenas primeira vez)
        [System.Windows.Forms.MessageBox]::Show(
            "Configuracao necessaria!`n`n" +
            "As credenciais do Azure AD ainda nao foram preenchidas.`n`n" +
            "OPCAO 1 (recomendado): Edite o arquivo .ps1 e preencha:`n" +
            "  CFG_TenantId, CFG_ClientId e CFG_ClientSecret`n`n" +
            "OPCAO 2: Preencha agora na janela de configuracao.`n" +
            "Ficara salvo e nao aparecera mais.",
            "Configuracao Inicial", "OK", "Information")
        $creds = Show-CredentialForm
        if (-not $creds) {
            [System.Windows.Forms.MessageBox]::Show(
                "Credenciais nao configuradas. Encerrando.",
                "Erro","OK","Error"); exit
        }
    }
}

$script:TenantId     = $creds.TenantId
$script:ClientId     = $creds.ClientId
$script:SecureSecret = $creds.ClientSecret
$Dominio          = $creds.Dominio
$AADConnectServer = $creds.AADConnectServer
$TargetOU         = $creds.TargetOU
$Telefone         = $creds.Telefone
$PaginaWeb        = $creds.PaginaWeb
$SenhaInicial     = Resolve-SecureStringPlain $creds.SenhaInicial
$script:AADSyncUser  = if ($creds.AADSyncUser)  { $creds.AADSyncUser }  else { $AADSyncUserDefault }
$script:AADSyncSenha = if ($creds.AADSyncSenha) { $creds.AADSyncSenha } else { ConvertTo-SecureString $AADSyncSenhaDefault -AsPlainText -Force }

# ============================================================
# TOKEN GRAPH API (cache + renovacao automatica)
# ============================================================
$script:GraphTokenCache  = $null
$script:GraphTokenExpiry = [datetime]::MinValue

function Get-GraphToken {
    if ($script:GraphTokenCache -and [datetime]::UtcNow -lt $script:GraphTokenExpiry) {
        return $script:GraphTokenCache
    }
    $plain = Resolve-SecureStringPlain $script:SecureSecret
    $body  = @{
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $script:ClientId
        client_secret = $plain
    }
    $plain = $null
    $resp  = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$($script:TenantId)/oauth2/v2.0/token" `
                               -Method Post -Body $body -ErrorAction Stop
    $script:GraphTokenCache  = @{ Authorization = "Bearer $($resp.access_token)"; "Content-Type" = "application/json" }
    $script:GraphTokenExpiry = [datetime]::UtcNow.AddMinutes(50)
    return $script:GraphTokenCache
}

# ============================================================
# FUNCOES DE LOG
# ============================================================
function Write-LogFile {
    param([string]$Mensagem, [string]$Tipo = "INFO", [string]$Prefixo = "geral")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    Add-Content -Path (Join-Path $LogDir "${Prefixo}_$(Get-Date -Format 'yyyy-MM').log") `
                -Value "[$ts] [$Tipo] $Mensagem" -Encoding UTF8
}

function Add-Log {
    param([System.Windows.Forms.RichTextBox]$Box, [string]$Mensagem, [string]$Tipo = "INFO", [string]$LogPrefixo = "geral")
    $ts   = Get-Date -Format "HH:mm:ss"
    $icon = switch ($Tipo) { "OK"{[char]0x2714};"ERRO"{[char]0x2718};"AVISO"{[char]0x26A0};"INFO"{[char]0x27A4};"ETAPA"{[char]0x25B6};default{" "} }
    $cor  = switch ($Tipo) {
        "OK"   {[System.Drawing.Color]::FromArgb(46,204,113)}
        "ERRO" {[System.Drawing.Color]::FromArgb(231,76,60)}
        "AVISO"{[System.Drawing.Color]::FromArgb(241,196,15)}
        "INFO" {[System.Drawing.Color]::FromArgb(52,152,219)}
        "ETAPA"{[System.Drawing.Color]::FromArgb(255,255,255)}
        default{[System.Drawing.Color]::FromArgb(180,180,200)}
    }
    $Box.SelectionStart = $Box.TextLength; $Box.SelectionLength = 0; $Box.SelectionColor = $cor
    $Box.AppendText("[$ts] $icon $Mensagem`r`n"); $Box.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
    Write-LogFile $Mensagem $Tipo $LogPrefixo
}

function DoEvents-Sleep {
    param([int]$Seconds)
    for ($i = 0; $i -lt ($Seconds * 4); $i++) {
        Start-Sleep -Milliseconds 250
        [System.Windows.Forms.Application]::DoEvents()
    }
}

# ============================================================
# CORES E FONTES
# ============================================================
$corFundo       = [System.Drawing.Color]::FromArgb(0, 53, 92)     # OBER Blue (Logo #00355C)
$corPainel      = [System.Drawing.Color]::FromArgb(0, 68, 110)
$corPainelClaro = [System.Drawing.Color]::FromArgb(0, 83, 130)
$corInput       = [System.Drawing.Color]::FromArgb(0, 68, 110)
$corBorda       = [System.Drawing.Color]::FromArgb(0, 103, 150)
$corAbaAtiva    = [System.Drawing.Color]::FromArgb(0, 68, 110)
$corAbaInativa  = [System.Drawing.Color]::FromArgb(0, 53, 92)
$corAutocomp    = [System.Drawing.Color]::FromArgb(0, 68, 110)
$corAutocompSel = [System.Drawing.Color]::FromArgb(21, 193, 242)  # Logo Sky Blue (#15C1F2)

$corVerde       = [System.Drawing.Color]::FromArgb(16, 185, 129)  # Emerald 500
$corAzul        = [System.Drawing.Color]::FromArgb(21, 193, 242)  # Logo Sky Blue (#15C1F2)
$corAzulHover   = [System.Drawing.Color]::FromArgb(56, 209, 250)  # Lighter Sky Blue
$corAzulEsc     = [System.Drawing.Color]::FromArgb(2, 132, 199)   # Sky 600
$corVermelho    = [System.Drawing.Color]::FromArgb(239, 68, 68)   # Red 500
$corVermelhoHov = [System.Drawing.Color]::FromArgb(248, 113, 113) # Red 400
$corVermelhoEsc = [System.Drawing.Color]::FromArgb(220, 38, 38)   # Red 600

$corTexto        = [System.Drawing.Color]::FromArgb(248, 250, 252) # Slate 50
$corTextoCinza   = [System.Drawing.Color]::FromArgb(148, 163, 184) # Slate 400

$fonteTitulo    = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Bold)
$fonteSubtitulo = New-Object System.Drawing.Font("Segoe UI",8.5)
$fonteLabel     = New-Object System.Drawing.Font("Segoe UI",9.5)
$fonteLabelBold = New-Object System.Drawing.Font("Segoe UI",9.5,[System.Drawing.FontStyle]::Bold)
$fonteInput     = New-Object System.Drawing.Font("Segoe UI",10)
$fonteBotao     = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$fonteLog       = New-Object System.Drawing.Font("Consolas",9)
$fonteAba       = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)

# ============================================================
# FORMULARIO PRINCIPAL
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text            = "Empresa TI - Gestao de Usuarios | AD + Microsoft 365"
$form.Size            = New-Object System.Drawing.Size(920,620)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox     = $true
$form.MinimumSize     = New-Object System.Drawing.Size(860,548)
$form.ShowInTaskbar   = $true
$form.TopMost         = $false
$form.BackColor       = $corFundo
$form.ForeColor       = $corTexto
$form.Font            = $fonteLabel

# ============================================================
# CABECALHO
# ============================================================
$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Height    = 55
$pnlHeader.Dock      = "Top"
$pnlHeader.BackColor = $corPainel

$lblTitulo = New-Object System.Windows.Forms.Label
$lblTitulo.Text     = "  Empresa TI - Gestao de Usuarios"
$lblTitulo.Font     = $fonteTitulo; $lblTitulo.ForeColor = [System.Drawing.Color]::White
$lblTitulo.Size     = New-Object System.Drawing.Size(370,26)
$lblTitulo.Location = New-Object System.Drawing.Point(150,5)
$pnlHeader.Controls.Add($lblTitulo)

$lblSubtitulo = New-Object System.Windows.Forms.Label
$lblSubtitulo.Text     = "  Active Directory + Microsoft 365 | Sua Empresa"
$lblSubtitulo.Font     = $fonteSubtitulo
$lblSubtitulo.ForeColor= [System.Drawing.Color]::FromArgb(200,220,255)
$lblSubtitulo.Size     = New-Object System.Drawing.Size(420,18)
$lblSubtitulo.Location = New-Object System.Drawing.Point(150,32)
$pnlHeader.Controls.Add($lblSubtitulo)

    $logoBase64 = "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"
$btnConfig = New-Object System.Windows.Forms.Button
$btnConfig.Text     = [char]0x2699
$btnConfig.Font     = New-Object System.Drawing.Font("Segoe UI",14)
$btnConfig.Size     = New-Object System.Drawing.Size(38,38)
$btnConfig.Anchor   = "Top, Right"
$btnConfig.Location = New-Object System.Drawing.Point($($pnlHeader.Width - 55), 9)
$btnConfig.FlatStyle = "Flat"; $btnConfig.FlatAppearance.BorderSize = 0
$btnConfig.BackColor = $corAzulEsc; $btnConfig.ForeColor = [System.Drawing.Color]::White
$btnConfig.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnConfig.Add_Click({
    $r = Show-CredentialForm
    if ($r) {
        $script:TenantId     = $r.TenantId
        $script:ClientId     = $r.ClientId
        $script:SecureSecret = $r.ClientSecret
        $script:GraphTokenCache = $null  # forcar renovacao do token
        $script:Dominio          = $r.Dominio
        $script:AADConnectServer = $r.AADConnectServer
        $script:TargetOU         = $r.TargetOU
        $script:Telefone         = $r.Telefone
        $SenhaInicial = Resolve-SecureStringPlain $r.SenhaInicial
        $cSenha.Text  = $SenhaInicial
        $script:AADSyncUser  = if ($r.AADSyncUser)  { $r.AADSyncUser }  else { "" }
        $script:AADSyncSenha = if ($r.AADSyncSenha) { $r.AADSyncSenha } else { $null }
        # Atualizar variaveis de sessao com os novos valores
        $Dominio          = $r.Dominio
        $AADConnectServer = $r.AADConnectServer
        $TargetOU         = $r.TargetOU
        $Telefone         = $r.Telefone
        [System.Windows.Forms.MessageBox]::Show("Configuracoes atualizadas!`n`nAs alteracoes ficam salvas no config.xml e sao carregadas automaticamente na proxima execucao.","Configuracao","OK","Information")
    }
})
$pnlHeader.Controls.Add($btnConfig)
$form.Controls.Add($pnlHeader)

# ============================================================
# ABAS CUSTOMIZADAS
# ============================================================
$pnlAbas = New-Object System.Windows.Forms.Panel
$pnlAbas.Width     = 240
$pnlAbas.Dock      = "Left"
$pnlAbas.BackColor = $corAbaInativa
$form.Controls.Add($pnlAbas)

    $imgBytes = [System.Convert]::FromBase64String($logoBase64)
    $ms = New-Object System.IO.MemoryStream(,$imgBytes)
    $logoImage = [System.Drawing.Image]::FromStream($ms)

    $picLogo = New-Object System.Windows.Forms.PictureBox
    $picLogo.SizeMode  = "Zoom"
    $picLogo.Size      = New-Object System.Drawing.Size(140,46)
    $picLogo.Anchor    = "Bottom, Left"
    $picLogo.Location  = New-Object System.Drawing.Point(40, 490)
    $picLogo.BackColor = [System.Drawing.Color]::Transparent
    $picLogo.Image     = $logoImage
    $pnlAbas.Controls.Add($picLogo)
    try { $bmp = New-Object System.Drawing.Bitmap($logoImage,32,32); $form.Icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon()) } catch {}


function New-TabButton($text, $y, $cor, $corHov) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text; $btn.Font = $fonteAba
    $btn.Size = New-Object System.Drawing.Size(240, 50)
    $btn.Location = New-Object System.Drawing.Point(0, $y)
    $btn.FlatStyle = "Flat"; $btn.FlatAppearance.BorderSize = 0
    $btn.BackColor = $cor; $btn.ForeColor = [System.Drawing.Color]::White
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.TextAlign = "MiddleLeft"
    $btn.Padding = New-Object System.Windows.Forms.Padding(20,0,0,0)
    $btn.Tag = $corHov
    return $btn
}

$btnAbaCriar    = New-TabButton "Criar Usuario"            20   $corAzul      $corAzulHover
$btnAbaDesligar = New-TabButton ([char]0x26A0 + " Desligar Colaborador") 70 $corAbaInativa $corVermelho
$pnlAbas.Controls.Add($btnAbaCriar)
$pnlAbas.Controls.Add($btnAbaDesligar)

# Indicador da aba ativa (Sidebar)
$indAba = New-Object System.Windows.Forms.Panel
$indAba.Size      = New-Object System.Drawing.Size(4, 50)
$indAba.Location  = New-Object System.Drawing.Point(0, 20)
$indAba.BackColor = $corTexto
$pnlAbas.Controls.Add($indAba)

# ============================================================
# PAINEL CONTEUDO
# ============================================================
$pnlConteudo = New-Object System.Windows.Forms.Panel
$pnlConteudo.Dock      = "Fill"
$pnlConteudo.BackColor = $corFundo
$form.Controls.Add($pnlConteudo)

# Fix WinForms Docking Z-Order
# Highest z-index (front) = Innermost. Lowest z-index (back) = Outermost.
$pnlHeader.SendToBack()     # Header will take the entire top
$pnlAbas.BringToFront()     # Sidebar will be below header
$pnlConteudo.BringToFront() # Content will fill the rest

# ============================================================
# HELPERS DE UI
# ============================================================
function New-Lbl($text, $x, $y, $bold = $false, $parent = $null) {
    $l = New-Object System.Windows.Forms.Label; $l.Text = $text
    $l.Font     = if ($bold) { $fonteLabelBold } else { $fonteLabel }
    $l.ForeColor = $corTextoCinza; $l.AutoSize = $true
    $l.Location = New-Object System.Drawing.Point($x,$y)
    if ($parent) { $parent.Controls.Add($l) }
    return $l
}
function New-Inp($x, $y, $w, $parent = $null, $Watermark = "") {
    $t = New-Object System.Windows.Forms.TextBox; $t.Font = $fonteInput
    $t.Size        = New-Object System.Drawing.Size($w,32)
    $t.Location    = New-Object System.Drawing.Point($x,$y)
    $t.BackColor   = $corInput; $t.ForeColor = $corTexto; $t.BorderStyle = "FixedSingle"
    if ($Watermark) {
        $t.Add_HandleCreated({ [Win32UI]::SendMessage([System.IntPtr]$this.Handle, 0x1501, [IntPtr]1, $Watermark) }.GetNewClosure())
    }
    if ($parent) { $parent.Controls.Add($t) }
    return $t
}
function New-Sep($x, $y, $w, $parent) {
    $s = New-Object System.Windows.Forms.Label
    $s.Size = New-Object System.Drawing.Size($w,1)
    $s.Location = New-Object System.Drawing.Point($x,$y)
    $s.BackColor = $corBorda; $parent.Controls.Add($s)
}
function New-LogBox($parent) {
    $p = New-Object System.Windows.Forms.Panel
    $p.Size     = New-Object System.Drawing.Size(610,210)
    $p.Location = New-Object System.Drawing.Point(35,268)
    $p.BackColor= $corPainel; $parent.Controls.Add($p)
    $lh = New-Object System.Windows.Forms.Label
    $lh.Text     = " Log de Execucao"
    $lh.Font     = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $lh.ForeColor= $corTextoCinza
    $lh.Size     = New-Object System.Drawing.Size(610,20)
    $lh.Location = New-Object System.Drawing.Point(0,0)
    $lh.BackColor= [System.Drawing.Color]::FromArgb(35,37,52)
    $p.Controls.Add($lh)
    $rb = New-Object System.Windows.Forms.RichTextBox
    $rb.Size      = New-Object System.Drawing.Size(606,185)
    $rb.Location  = New-Object System.Drawing.Point(2,22)
    $rb.BackColor = [System.Drawing.Color]::FromArgb(22,22,34)
    $rb.ForeColor = $corTextoCinza; $rb.Font = $fonteLog
    $rb.ReadOnly  = $true; $rb.BorderStyle = "None"
    $rb.ScrollBars= "Vertical"; $rb.WordWrap = $false
    $p.Controls.Add($rb); return $rb
}
function New-ProgBar($parent, $y) {
    $pb = New-Object System.Windows.Forms.ProgressBar
    $pb.Size = New-Object System.Drawing.Size(610,10); $pb.Location = New-Object System.Drawing.Point(35,$y)
    $pb.Minimum = 0; $pb.Maximum = 100; $pb.Value = 0; $pb.Style = "Continuous"
    $parent.Controls.Add($pb); return $pb
}
function New-StatusLabel($parent, $y) {
    $l = New-Object System.Windows.Forms.Label
    $l.Font     = New-Object System.Drawing.Font("Segoe UI",9)
    $l.ForeColor= $corTextoCinza
    $l.Size     = New-Object System.Drawing.Size(610,18)
    $l.Location = New-Object System.Drawing.Point(35,$y)
    $parent.Controls.Add($l); return $l
}

# ============================================================
# AUTOCOMPLETE HELPER - ListBox flutuante para busca no AD
# Retorna o objeto { Timer, ListBox }
# onSelect: scriptblock chamado com o objeto ADUser selecionado
# ============================================================
function New-ADAutoComplete {
    param(
        [System.Windows.Forms.TextBox]$Campo,
        [System.Windows.Forms.Control]$PainelHost,  # painel pai do listbox (para posicionamento)
        [System.Windows.Forms.Control]$PainelRef,   # painel de referencia do campo
        [int]$ListX, [int]$ListY, [int]$ListW,
        [scriptblock]$OnSelect,
        [string]$FiltroExtra = ""                   # ex: "-and Enabled -eq 'True'"
    )

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 450

    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Font        = New-Object System.Drawing.Font("Segoe UI",9.5)
    $lst.BackColor   = $corAutocomp
    $lst.ForeColor   = $corTexto
    $lst.BorderStyle = "FixedSingle"
    $lst.Visible     = $false
    $lst.Size        = New-Object System.Drawing.Size($ListW, 10)
    $lst.Location    = New-Object System.Drawing.Point($ListX, $ListY)
    $lst.Tag         = [System.Collections.ArrayList]::new()   # armazena os ADUsers
    $PainelHost.Controls.Add($lst)
    $lst.BringToFront()

    # Capturar variaveis locais em script-scope para os scriptblocks
    # GetNewClosure() garante que $timer, $lst, $Campo, $OnSelect sao capturados corretamente
    $applySelect = {
        if ($lst.SelectedIndex -ge 0 -and $lst.Tag.Count -gt $lst.SelectedIndex) {
            $adUser = $lst.Tag[$lst.SelectedIndex]
            $lst.Visible = $false
            $timer.Stop()
            & $OnSelect $adUser
        }
    }.GetNewClosure()

    $lst.Add_Click($applySelect)

    $lst.Add_KeyDown({
        if ($_.KeyCode -eq "Return") { & $applySelect }
        if ($_.KeyCode -eq "Escape") { $lst.Visible = $false; $timer.Stop() }
    }.GetNewClosure())

    # Seta para baixo no campo move foco para lista
    $Campo.Add_KeyDown({
        if ($_.KeyCode -eq "Down" -and $lst.Visible -and $lst.Items.Count -gt 0) {
            $lst.Focus(); $lst.SelectedIndex = 0
        }
        if ($_.KeyCode -eq "Escape") { $lst.Visible = $false; $timer.Stop() }
    }.GetNewClosure())

    # Timer faz a busca assíncrona
    $checkTimer = New-Object System.Windows.Forms.Timer
    $checkTimer.Interval = 100
    $state = @{ PS = $null; Handle = $null; Runspace = $null }

    $checkTimer.Add_Tick({
        if ($state.Handle -and $state.Handle.IsCompleted) {
            $checkTimer.Stop()
            try {
                $users = $state.PS.EndInvoke($state.Handle)
                $state.PS.Dispose()
                $state.Runspace.Close()
                $state.Runspace.Dispose()
                
                $lst.Items.Clear()
                $lst.Tag.Clear()
                if ($users.Count -eq 0) {
                    [void]$lst.Items.Add("  Nenhum usuario encontrado")
                    $lst.ForeColor = [System.Drawing.Color]::FromArgb(180,180,200)
                } else {
                    $lst.ForeColor = $corTexto
                    foreach ($u in $users) {
                        $setor = if ($u.Department) { $u.Department } else { "Sem setor" }
                        [void]$lst.Tag.Add($u)
                        [void]$lst.Items.Add("  $($u.SamAccountName)   $($u.DisplayName) - $setor")
                    }
                }
                $itemH = $lst.ItemHeight
                $qtd   = [Math]::Min($lst.Items.Count, 7)
                $lst.Height  = ($itemH * $qtd) + 4
            } catch {
                $lst.Items.Clear()
                [void]$lst.Items.Add("  Erro na busca: $($_.Exception.Message)")
                $lst.ForeColor = [System.Drawing.Color]::FromArgb(231,76,60)
                $lst.Height  = $lst.ItemHeight + 4
            }
        }
    }.GetNewClosure())

    $timer.Add_Tick({
        $timer.Stop()
        $busca = $Campo.Text.Trim()
        if ($busca.Length -lt 2) { $lst.Visible = $false; return }

        $lst.Items.Clear(); $lst.Tag.Clear()
        [void]$lst.Items.Add("  Buscando...")
        $lst.ForeColor = [System.Drawing.Color]::FromArgb(180,180,200)
        $lst.Height  = $lst.ItemHeight + 4
        $lst.Visible = $true; $lst.BringToFront()

        if ($state.PS) { try { $state.PS.Stop(); $state.PS.Dispose(); $state.Runspace.Close(); $state.Runspace.Dispose() } catch {} }

        $rsScript = {
            param($buscaArg)
            Add-Type -AssemblyName System.DirectoryServices
            $searcher = New-Object System.DirectoryServices.DirectorySearcher
            $searcher.Filter = "(&(objectCategory=person)(objectClass=user)(|(displayName=*$buscaArg*)(sAMAccountName=*$buscaArg*)(givenName=*$buscaArg*)(sn=*$buscaArg*)))"
            $searcher.PropertiesToLoad.AddRange([string[]]@("sAMAccountName","displayName","department","userAccountControl","memberOf","givenName","sn"))
            $searcher.SizeLimit = 12
            $rawResults = $searcher.FindAll()

            $users = [System.Collections.ArrayList]::new()
            foreach ($r in $rawResults) {
                $uac = if ($r.Properties["useraccountcontrol"].Count -gt 0) { [int]$r.Properties["useraccountcontrol"][0] } else { 0 }
                if ($uac -band 2) { continue }
                $sam  = if ($r.Properties["samaccountname"].Count -gt 0) { $r.Properties["samaccountname"][0] } else { "" }
                $disp = if ($r.Properties["displayname"].Count -gt 0)    { $r.Properties["displayname"][0]    } else { $sam }
                $dept = if ($r.Properties["department"].Count -gt 0)     { $r.Properties["department"][0]     } else { "" }
                $gn   = if ($r.Properties["givenname"].Count -gt 0)      { $r.Properties["givenname"][0]      } else { "" }
                $sn2  = if ($r.Properties["sn"].Count -gt 0)             { $r.Properties["sn"][0]             } else { "" }
                $mof  = if ($r.Properties["memberof"].Count -gt 0)       { $r.Properties["memberof"]          } else { @() }
                [void]$users.Add([PSCustomObject]@{
                    SamAccountName = $sam; DisplayName = $disp
                    Department = $dept; GivenName = $gn; Surname = $sn2; MemberOf = $mof
                })
            }
            $rawResults.Dispose()
            return ($users | Sort-Object DisplayName)
        }

        $state.Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $state.Runspace.ThreadOptions = "ReuseThread"
        $state.Runspace.Open()

        $state.PS = [System.Management.Automation.PowerShell]::Create()
        $state.PS.Runspace = $state.Runspace
        [void]$state.PS.AddScript($rsScript).AddArgument($busca)
        $state.Handle = $state.PS.BeginInvoke()
        $checkTimer.Start()
    }.GetNewClosure())

    # Dispara busca ao digitar
    $Campo.Add_TextChanged({
        $timer.Stop()
        if ($Campo.Text.Trim().Length -ge 2) { $timer.Start() }
        else { $lst.Visible = $false }
    }.GetNewClosure())

    return @{ Timer = $timer; ListBox = $lst }
}

# ============================================================
# ====  ABA 1 - CRIAR USUARIO  ====
# ============================================================
$pnlCriar = New-Object System.Windows.Forms.Panel
$pnlCriar.Dock      = "Fill"
$pnlCriar.BackColor = $corFundo
$pnlConteudo.Controls.Add($pnlCriar)

$pnlCForm = New-Object System.Windows.Forms.Panel
$pnlCForm.Size      = New-Object System.Drawing.Size(610,195)
$pnlCForm.Location  = New-Object System.Drawing.Point(35,15)
$pnlCForm.BackColor = $corPainel
$pnlCriar.Controls.Add($pnlCForm)

New-Lbl "Primeiro Nome *"    15  10 $false $pnlCForm | Out-Null
$cNome      = New-Inp 15  28 290 $pnlCForm "Ex: Joao"

New-Lbl "Sobrenome *"        320 10 $false $pnlCForm | Out-Null
$cSobrenome = New-Inp 315 28 290 $pnlCForm "Ex: da Silva"

New-Lbl "Usuario Template *" 15  58 $false $pnlCForm | Out-Null
$cTemplate  = New-Inp 15  76 290 $pnlCForm "Busque por nome, e-mail ou setor..."

New-Lbl "Senha Inicial"      320 58 $false $pnlCForm | Out-Null
$cSenha     = New-Inp 315 76 290 $pnlCForm "Senha inicial"
$cSenha.Text = $SenhaInicial

New-Lbl "Licenca Microsoft 365" 15 108 $false $pnlCForm | Out-Null
$cmbLicenca = New-Object System.Windows.Forms.ComboBox
$cmbLicenca.Font = $fonteInput; $cmbLicenca.Size = New-Object System.Drawing.Size(290,28)
$cmbLicenca.Location = New-Object System.Drawing.Point(15,126)
$cmbLicenca.BackColor = $corInput; $cmbLicenca.ForeColor = $corTexto
$cmbLicenca.FlatStyle = "Flat"; $cmbLicenca.DropDownStyle = "DropDownList"
$cmbLicenca.Items.AddRange(@("Automatico (recomendado)","Business Standard (SPB)","Apps Enterprise + E1","Nao atribuir licenca"))
$cmbLicenca.SelectedIndex = 0; $pnlCForm.Controls.Add($cmbLicenca)

New-Lbl "DADOS GERADOS" 320 108 $true $pnlCForm | Out-Null
(New-Lbl "DADOS GERADOS" 320 108 $true $pnlCForm).ForeColor = $corAzul

New-Lbl "Login:" 320 132 $false $pnlCForm | Out-Null
$cUsername = New-Inp 370 130 235 $pnlCForm "sAMAccountName gerado"
$cUsername.ReadOnly = $true; $cUsername.ForeColor = $corVerde

New-Lbl "Email:" 320 160 $false $pnlCForm | Out-Null
$cEmail = New-Inp 370 158 235 $pnlCForm "E-mail principal"
$cEmail.ReadOnly = $true; $cEmail.ForeColor = $corVerde

# --- Gerar username com suporte a acentos ---
function Update-GeneratedUsername {
    $n = Remove-Acentos $cNome.Text.Trim()
    $s = Remove-Acentos $cSobrenome.Text.Trim()
    if ($n -ne "" -and $s -ne "") {
        $u = ("$n.$s").ToLower() -replace '\s+','' -replace '[^a-z0-9.]',''
        $cUsername.Text = $u
        $cEmail.Text    = "$u@$Dominio"
    } else {
        $cUsername.Text = ""; $cEmail.Text = ""
    }
}

# --- Timer debounce para verificar username duplicado no AD ---
$script:TimerUsername = New-Object System.Windows.Forms.Timer
$script:TimerUsername.Interval = 800
$script:TimerUsername.Add_Tick({
    $script:TimerUsername.Stop()
    $u = $cUsername.Text.Trim()
    if ($u -eq "") { return }
    try {
        # Verificar duplicidade via LDAP nativo
        Add-Type -AssemblyName System.DirectoryServices
        $s = New-Object System.DirectoryServices.DirectorySearcher
        $s.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$u))"
        $s.PropertiesToLoad.Add("sAMAccountName") | Out-Null
        $found = $s.FindOne()
        if ($found) {
            $cUsername.BackColor = [System.Drawing.Color]::FromArgb(70,20,20)
            $cUsername.ForeColor = [System.Drawing.Color]::FromArgb(231,76,60)
            $lblStCriar.Text     = "$([char]0x26A0) Username '$u' ja existe no AD!"
            $lblStCriar.ForeColor = $corVermelho
        } else {
            $cUsername.BackColor = $corInput
            $cUsername.ForeColor = $corVerde
            $lblStCriar.Text     = "$([char]0x2714) Username '$u' disponivel."
            $lblStCriar.ForeColor = $corVerde
        }
    } catch {
        $cUsername.BackColor = $corInput
        $cUsername.ForeColor = $corVerde
    }
})

$cNome.Add_TextChanged({
    Update-GeneratedUsername
    $script:TimerUsername.Stop(); $script:TimerUsername.Start()
})
$cSobrenome.Add_TextChanged({
    Update-GeneratedUsername
    $script:TimerUsername.Stop(); $script:TimerUsername.Start()
})

# --- Botoes Criar ---
$btnCriar = New-Object System.Windows.Forms.Button
$btnCriar.Text = "CRIAR USUARIO"; $btnCriar.Font = $fonteBotao
$btnCriar.Size = New-Object System.Drawing.Size(300,36); $btnCriar.Location = New-Object System.Drawing.Point(35,210)
$btnCriar.FlatStyle = "Flat"; $btnCriar.FlatAppearance.BorderSize = 0
$btnCriar.BackColor = $corAzul; $btnCriar.ForeColor = [System.Drawing.Color]::White
$btnCriar.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnCriar.Add_MouseEnter({ $btnCriar.BackColor = $corAzulHover })
$btnCriar.Add_MouseLeave({ $btnCriar.BackColor = $corAzul })
$pnlCriar.Controls.Add($btnCriar)

$btnLimpar = New-Object System.Windows.Forms.Button
$btnLimpar.Text = "LIMPAR CAMPOS"; $btnLimpar.Font = $fonteBotao
$btnLimpar.Size = New-Object System.Drawing.Size(300,36); $btnLimpar.Location = New-Object System.Drawing.Point(345,210)
$btnLimpar.FlatStyle = "Flat"; $btnLimpar.FlatAppearance.BorderSize = 0
$btnLimpar.BackColor = $corPainelClaro; $btnLimpar.ForeColor = $corTexto
$btnLimpar.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnLimpar.Add_MouseEnter({ $btnLimpar.BackColor = $corBorda })
$btnLimpar.Add_MouseLeave({ $btnLimpar.BackColor = $corPainelClaro })
$pnlCriar.Controls.Add($btnLimpar)

$pbCriar    = New-ProgBar $pnlCriar 248
$lblStCriar = New-StatusLabel $pnlCriar 256
$lblStCriar.Text = "Pronto para criar usuario."
$logCriar   = New-LogBox $pnlCriar

$btnLimpar.Add_Click({
    $cNome.Text = ""; $cSobrenome.Text = ""; $cTemplate.Text = ""
    $cUsername.Text = ""; $cEmail.Text = ""
    $cUsername.BackColor = $corInput; $cUsername.ForeColor = $corVerde
    $cmbLicenca.SelectedIndex = 0
    $logCriar.Clear(); $pbCriar.Value = 0
    $lblStCriar.Text = "Pronto para criar usuario."; $lblStCriar.ForeColor = $corTextoCinza
    $cNome.Focus()
})

# --- AUTOCOMPLETE DO TEMPLATE (busca por nome no AD) ---
# Posicao relativa a pnlCriar: pnlCForm(12,8) + cTemplate(15,76+32)
$acTemplate = New-ADAutoComplete `
    -Campo      $cTemplate `
    -PainelHost $pnlCriar `
    -PainelRef  $pnlCForm `
    -ListX      27 -ListY 118 -ListW 310 `
    -OnSelect   {
        param($adUser)
        if ($adUser -and $adUser.SamAccountName) {
            $cTemplate.Text = $adUser.SamAccountName
        }
    }

# Fechar lista do template ao clicar fora
$pnlCriar.Add_Click({ $acTemplate.ListBox.Visible = $false })
$pnlCForm.Add_Click({ $acTemplate.ListBox.Visible = $false })

# ============================================================
# ====  ABA 2 - DESLIGAR COLABORADOR  ====
# ============================================================
$pnlDesligar = New-Object System.Windows.Forms.Panel
$pnlDesligar.Dock      = "Fill"
$pnlDesligar.BackColor = $corFundo; $pnlDesligar.Visible = $false
$pnlConteudo.Controls.Add($pnlDesligar)

$pnlDForm = New-Object System.Windows.Forms.Panel
$pnlDForm.Size      = New-Object System.Drawing.Size(610,195)
$pnlDForm.Location  = New-Object System.Drawing.Point(35,15)
$pnlDForm.BackColor = $corPainel; $pnlDesligar.Controls.Add($pnlDForm)

New-Lbl "Nome ou Login do Colaborador *" 15 12 $false $pnlDForm | Out-Null

$dUsername = New-Inp 15 32 340 $pnlDForm "Busque pelo nome ou login..."

# Flag para evitar que o TextChanged limpe as labels quando o autocomplete preenche
$script:DesligarAutoFill = $false

$dUsername.Add_TextChanged({
    if (-not $script:DesligarAutoFill) {
        $u = $dUsername.Text.Trim()
        $dLblEmail.Text  = if ($u) { "Email: $u@$Dominio" } else { "Email: ---" }
        $dLblNome.Text   = "Nome no AD: ---"; $dLblNome.ForeColor = $corTextoCinza
        $dLblGrupos.Text = "Grupos AD: ---"
    }
})

$btnVerificar = New-Object System.Windows.Forms.Button
$btnVerificar.Text = "Verificar"
$btnVerificar.Font = New-Object System.Drawing.Font("Segoe UI",9.5,[System.Drawing.FontStyle]::Bold)
$btnVerificar.Size = New-Object System.Drawing.Size(100,28); $btnVerificar.Location = New-Object System.Drawing.Point(362,32)
$btnVerificar.FlatStyle = "Flat"; $btnVerificar.FlatAppearance.BorderSize = 0
$btnVerificar.BackColor = $corAzul; $btnVerificar.ForeColor = [System.Drawing.Color]::White
$btnVerificar.Cursor = [System.Windows.Forms.Cursors]::Hand; $pnlDForm.Controls.Add($btnVerificar)

$dLblEmail  = New-Lbl "Email: ---"         15  68  $false $pnlDForm
$dLblNome   = New-Lbl "Nome no AD: ---"    15  86  $false $pnlDForm
$dLblGrupos = New-Lbl "Grupos AD: ---"     15  104 $false $pnlDForm

New-Sep 15 126 580 $pnlDForm
New-Lbl "Opcoes de desligamento:" 15 134 $false $pnlDForm | Out-Null

$dChkExchange = New-Object System.Windows.Forms.CheckBox
$dChkExchange.Text    = "Converter mailbox para compartilhada (Exchange Online via Graph API)"
$dChkExchange.Font    = $fonteLabel; $dChkExchange.ForeColor = $corTexto; $dChkExchange.Checked = $true
$dChkExchange.AutoSize= $true; $dChkExchange.Location = New-Object System.Drawing.Point(15,152)
$pnlDForm.Controls.Add($dChkExchange)

$dChkLicencas = New-Object System.Windows.Forms.CheckBox
$dChkLicencas.Text    = "Remover licencas Microsoft 365 (via Graph API)"
$dChkLicencas.Font    = $fonteLabel; $dChkLicencas.ForeColor = $corTexto; $dChkLicencas.Checked = $true
$dChkLicencas.AutoSize= $true; $dChkLicencas.Location = New-Object System.Drawing.Point(15,172)
$pnlDForm.Controls.Add($dChkLicencas)

$btnDesligar = New-Object System.Windows.Forms.Button
$btnDesligar.Text = ([char]0x26A0) + "  DESLIGAR COLABORADOR"; $btnDesligar.Font = $fonteBotao
$btnDesligar.Size = New-Object System.Drawing.Size(610,36); $btnDesligar.Location = New-Object System.Drawing.Point(35,210)
$btnDesligar.FlatStyle = "Flat"; $btnDesligar.FlatAppearance.BorderSize = 0
$btnDesligar.BackColor = $corVermelho; $btnDesligar.ForeColor = [System.Drawing.Color]::White
$btnDesligar.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnDesligar.Add_MouseEnter({ if ($btnDesligar.Enabled) { $btnDesligar.BackColor = $corVermelhoHov } })
$btnDesligar.Add_MouseLeave({ if ($btnDesligar.Enabled) { $btnDesligar.BackColor = $corVermelho } })
$pnlDesligar.Controls.Add($btnDesligar)

$pbDesligar    = New-ProgBar $pnlDesligar 248
$lblStDesligar = New-StatusLabel $pnlDesligar 256
$lblStDesligar.Text = "Pronto. Informe o nome ou login e clique em DESLIGAR COLABORADOR."
$logDesligar   = New-LogBox $pnlDesligar

# --- AUTOCOMPLETE DO DESLIGAR (busca por nome, preenche o username) ---
# Posicao relativa a pnlDesligar: pnlDForm(12,8) + dUsername(15, 32+32)
$acDesligar = New-ADAutoComplete `
    -Campo      $dUsername `
    -PainelHost $pnlDesligar `
    -PainelRef  $pnlDForm `
    -ListX      27 -ListY 74 -ListW 460 `
    -OnSelect   {
        param($adUser)
        if ($adUser -and $adUser.SamAccountName) {
            $script:DesligarAutoFill = $true
            $dUsername.Text  = $adUser.SamAccountName
            $dLblEmail.Text  = "Email: $($adUser.SamAccountName)@$Dominio"
            $nome = if ($adUser.DisplayName) { $adUser.DisplayName } else { "$($adUser.GivenName) $($adUser.Surname)" }
            $dLblNome.Text   = "Nome no AD: $nome"
            $dLblNome.ForeColor = $corVerde
            $qtd = if ($adUser.MemberOf) { $adUser.MemberOf.Count } else { 0 }
            $dLblGrupos.Text = "Grupos AD: $qtd grupo(s)"
            $script:DesligarAutoFill = $false
        }
    }

# Fechar lista ao clicar fora
$pnlDesligar.Add_Click({ $acDesligar.ListBox.Visible = $false })
$pnlDForm.Add_Click({ $acDesligar.ListBox.Visible = $false })

# ============================================================
# TROCA DE ABAS
# ============================================================
function Switch-Tab($aba) {
    if ($aba -eq "criar") {
        $pnlCriar.Visible = $true; $pnlDesligar.Visible = $false
        $btnAbaCriar.BackColor = $corAzul; $btnAbaDesligar.BackColor = $corAbaInativa
        $indAba.Location = New-Object System.Drawing.Point(0,20); $indAba.BackColor = $corAzul
    } else {
        $pnlCriar.Visible = $false; $pnlDesligar.Visible = $true
        $btnAbaCriar.BackColor = $corAbaInativa; $btnAbaDesligar.BackColor = $corVermelho
        $indAba.Location = New-Object System.Drawing.Point(0,70); $indAba.BackColor = $corVermelho
    }
}

$btnAbaCriar.Add_Click({ Switch-Tab "criar" })
$btnAbaDesligar.Add_Click({ Switch-Tab "desligar" })

# ============================================================
# ACAO: VERIFICAR NO AD
# ============================================================
$btnVerificar.Add_Click({
    $u = $dUsername.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($u)) {
        [System.Windows.Forms.MessageBox]::Show("Informe o nome ou username para verificar.","Aviso","OK","Warning"); return
    }
    $dLblNome.Text = "Nome no AD: verificando..."; $dLblGrupos.Text = "Grupos AD: verificando..."
    $dLblNome.ForeColor = $corTextoCinza; [System.Windows.Forms.Application]::DoEvents()
    try {
        # Verificar usuario via LDAP nativo
        Add-Type -AssemblyName System.DirectoryServices
        $s = New-Object System.DirectoryServices.DirectorySearcher
        $s.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$u))"
        $s.PropertiesToLoad.AddRange([string[]]@("displayName","department","memberOf","givenName","sn"))
        $found = $s.FindOne()
        if ($found) {
            $gn   = if ($found.Properties["givenname"].Count -gt 0)   { $found.Properties["givenname"][0]   } else { "" }
            $sn2  = if ($found.Properties["sn"].Count -gt 0)          { $found.Properties["sn"][0]          } else { "" }
            $disp = if ($found.Properties["displayname"].Count -gt 0) { $found.Properties["displayname"][0] } else { "$gn $sn2".Trim() }
            $dept = if ($found.Properties["department"].Count -gt 0)  { $found.Properties["department"][0]  } else { "N/A" }
            $qtd  = $found.Properties["memberof"].Count
            $dLblNome.Text   = "Nome no AD: $disp"; $dLblNome.ForeColor = $corVerde
            $dLblGrupos.Text = "Grupos AD: $qtd grupo(s)"
            $dLblEmail.Text  = "Email: $u@$Dominio"
        } else {
            $dLblNome.Text      = "Nome no AD: USUARIO NAO ENCONTRADO"
            $dLblNome.ForeColor = [System.Drawing.Color]::FromArgb(231,76,60)
            $dLblGrupos.Text    = "Grupos AD: ---"
        }
    } catch {
        $dLblNome.Text      = "Nome no AD: Erro ao consultar ($($_.Exception.Message))"
        $dLblNome.ForeColor = [System.Drawing.Color]::FromArgb(231,76,60)
        $dLblGrupos.Text    = "Grupos AD: ---"
    }
})

# ============================================================
# ACAO: CRIAR USUARIO
# ============================================================
$btnCriar.Add_Click({
    if ([string]::IsNullOrWhiteSpace($cNome.Text))      { [System.Windows.Forms.MessageBox]::Show("Preencha o PRIMEIRO NOME.","Campo obrigatorio","OK","Warning");   $cNome.Focus(); return }
    if ([string]::IsNullOrWhiteSpace($cSobrenome.Text)) { [System.Windows.Forms.MessageBox]::Show("Preencha o SOBRENOME.","Campo obrigatorio","OK","Warning");       $cSobrenome.Focus(); return }
    if ([string]::IsNullOrWhiteSpace($cTemplate.Text))  { [System.Windows.Forms.MessageBox]::Show("Preencha o USUARIO TEMPLATE.","Campo obrigatorio","OK","Warning");$cTemplate.Focus(); return }
    if ([string]::IsNullOrWhiteSpace($cUsername.Text))  { [System.Windows.Forms.MessageBox]::Show("Username nao gerado. Verifique nome e sobrenome.","Erro","OK","Error"); return }

    $usrCheck = $cUsername.Text.Trim()
    if ($usrCheck -notmatch '^[a-z][a-z0-9.]{1,62}$') {
        [System.Windows.Forms.MessageBox]::Show("Username '$usrCheck' contem caracteres invalidos.`nUse apenas letras minusculas, numeros e ponto.","Username invalido","OK","Warning"); return
    }

    $ti     = (Get-Culture).TextInfo
    $PNome  = $ti.ToTitleCase($cNome.Text.ToLower().Trim())
    $PSnome = $ti.ToTitleCase($cSobrenome.Text.ToLower().Trim())
    $PUser  = $cUsername.Text.Trim()
    $PTmpl  = $cTemplate.Text.Trim()
    $PEmail = $cEmail.Text.Trim()
    $PLic   = $cmbLicenca.SelectedItem.ToString()
    $PSenha = $cSenha.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($PSenha)) { $PSenha = $SenhaInicial }

    if (([System.Windows.Forms.MessageBox]::Show(
        "Confirma a criacao?`n`nNome    : $PNome $PSnome`nLogin   : $PUser`nEmail   : $PEmail`nTemplate: $PTmpl`nLicenca : $PLic`nSenha   : $PSenha",
        "Confirmacao","YesNo","Question")) -ne "Yes") { return }

    $btnCriar.Enabled = $false; $btnLimpar.Enabled = $false; $btnCriar.Text = "PROCESSANDO..."
    $logCriar.Clear(); $pbCriar.Value = 0
    $Senha = ConvertTo-SecureString $PSenha -AsPlainText -Force

    function RestoreCriar { $btnCriar.Enabled = $true; $btnLimpar.Enabled = $true; $btnCriar.Text = "CRIAR USUARIO" }

    # ETAPA 1 - Pre-requisitos
    $lblStCriar.Text = "Etapa 1/6: Verificando pre-requisitos..."; $lblStCriar.ForeColor = $corAzul
    Add-Log $logCriar "ETAPA 1: Verificando pre-requisitos" "ETAPA" "criacao"
    $adMsg = if ($script:ADRemote) { " (remoting -> $($script:AADConnectServer))" } else { " (local)." }
    if (Ensure-ADModule) { Add-Log $logCriar "Modulo ActiveDirectory pronto$adMsg" "OK" "criacao" }
    else {
        Add-Log $logCriar "Falha ao conectar ao AD (todas as tentativas):" "ERRO" "criacao"
        foreach ($errLine in ($script:ADModuleError -split "`n")) {
            if ($errLine.Trim()) { Add-Log $logCriar $errLine.Trim() "ERRO" "criacao" }
        }
        Add-Log $logCriar "DICA: Execute no servidor como Admin:" "AVISO" "criacao"
        Add-Log $logCriar "  Enable-PSRemoting -Force" "AVISO" "criacao"
        Add-Log $logCriar "  Set-Item WSMan:\localhost\Client\TrustedHosts -Value '*' -Force" "AVISO" "criacao"
        RestoreCriar; return
    }

    try {
        Get-ADUser -Identity $PUser -ErrorAction Stop | Out-Null
        Add-Log $logCriar "Usuario '$PUser' ja existe no AD!" "ERRO" "criacao"
        $lblStCriar.Text = "ERRO: Usuario ja existe."; $lblStCriar.ForeColor = $corVermelho
        RestoreCriar; return
    } catch {
        if ($_.CategoryInfo.Category -eq "ObjectNotFound" -or
            $_.Exception.GetType().Name -like "*NotFound*" -or
            $_.Exception.Message -like "*not found*" -or
            $_.Exception.Message -like "*nao foi encontrado*") {
            Add-Log $logCriar "Username '$PUser' disponivel." "OK" "criacao"
        } else {
            Add-Log $logCriar "Aviso ao verificar duplicidade: $_" "AVISO" "criacao"
        }
    }
    $pbCriar.Value = 10

    # ETAPA 2 - Template
    $lblStCriar.Text = "Etapa 2/6: Buscando usuario template..."
    Add-Log $logCriar "ETAPA 2: Buscando template '$PTmpl'" "ETAPA" "criacao"
    try { $tmpl = Get-ADUser -Identity $PTmpl -Properties MemberOf,Department,Title,Company,Office,Description,Manager }
    catch { Add-Log $logCriar "Template '$PTmpl' nao encontrado!" "ERRO" "criacao"; RestoreCriar; return }

    $OU         = $tmpl.DistinguishedName -replace '^CN=[^,]+,',''
    $Setor      = $tmpl.Department
    $NomeExibido= "$PNome $PSnome - $Setor"
    Add-Log $logCriar "Template OK | Setor: $Setor" "OK" "criacao"
    Add-Log $logCriar "OU destino: $OU" "INFO" "criacao"
    $pbCriar.Value = 20

    # ETAPA 3 - Criar no AD
    $lblStCriar.Text = "Etapa 3/6: Criando usuario no Active Directory..."
    Add-Log $logCriar "ETAPA 3: Criando usuario no AD" "ETAPA" "criacao"
    try {
        $params = @{
            GivenName            = $PNome
            Surname              = $PSnome
            Name                 = $NomeExibido
            DisplayName          = $NomeExibido
            SamAccountName       = $PUser
            UserPrincipalName    = $PEmail
            EmailAddress         = $PEmail
            OfficePhone          = $Telefone
            HomePage             = $PaginaWeb
            AccountPassword      = $Senha
            Enabled              = $true
            ChangePasswordAtLogon= $true
            Path                 = $OU
        }
        if ($tmpl.Department)  { $params["Department"]  = $tmpl.Department }
        if ($tmpl.Title)       { $params["Title"]       = $tmpl.Title }
        if ($tmpl.Company)     { $params["Company"]     = $tmpl.Company }
        if ($tmpl.Office)      { $params["Office"]      = $tmpl.Office }
        if ($tmpl.Description) { $params["Description"] = $tmpl.Description }
        New-ADUser @params
        if ($tmpl.Manager) { Set-ADUser -Identity $PUser -Manager $tmpl.Manager; Add-Log $logCriar "Gestor copiado do template." "OK" "criacao" }
        Add-Log $logCriar "Usuario '$PUser' criado no AD!" "OK" "criacao"
    } catch {
        Add-Log $logCriar "Falha ao criar usuario: $_" "ERRO" "criacao"
        $lblStCriar.Text = "ERRO ao criar usuario no AD."; $lblStCriar.ForeColor = $corVermelho
        RestoreCriar; return
    }
    $pbCriar.Value = 40

    # ETAPA 4 - Grupos
    $lblStCriar.Text = "Etapa 4/6: Copiando grupos do template..."
    Add-Log $logCriar "ETAPA 4: Copiando grupos" "ETAPA" "criacao"
    $gOk = 0
    foreach ($g in $tmpl.MemberOf) {
        try {
            Add-ADGroupMember -Identity $g -Members $PUser -ErrorAction Stop
            $nG = ($g -split ',')[0] -replace 'CN=',''
            Add-Log $logCriar "Grupo: $nG" "OK" "criacao"; $gOk++
        } catch {
            $nG = ($g -split ',')[0] -replace 'CN=',''
            Add-Log $logCriar "Falha: $nG" "AVISO" "criacao"
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    Add-Log $logCriar "Total: $gOk grupo(s) copiado(s)" "INFO" "criacao"
    $pbCriar.Value = 55

    # ETAPA 5 - Sync Azure AD
    $lblStCriar.Text = "Etapa 5/6: Sincronizando com Microsoft 365..."
    Add-Log $logCriar "ETAPA 5: Sincronizacao Azure AD Connect ($AADConnectServer)" "ETAPA" "criacao"
    $syncOk = $false

    # Montar credencial do servidor ADSync se configurada
    $aadSyncCred = $null
    if ($script:AADSyncUser -and $script:AADSyncSenha) {
        try { $aadSyncCred = New-Object System.Management.Automation.PSCredential($script:AADSyncUser, $script:AADSyncSenha) } catch {}
    }

    # Verificar se estamos no proprio servidor ADSync (rodar localmente sem remoting)
    $localIP = ""
    try { $localIP = [System.Net.Dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -First 1 -ExpandProperty IPAddressToString } catch {}
    $isLocal = ($AADConnectServer -in @($env:COMPUTERNAME, "localhost", "127.0.0.1", "::1")) -or ($AADConnectServer -eq $localIP)

    if ($isLocal) {
        try {
            Import-Module ADSync -ErrorAction Stop
            Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
            Add-Log $logCriar "Sync Delta iniciado localmente (servidor local)." "OK" "criacao"
            $syncOk = $true
        } catch { Add-Log $logCriar "ADSync nao disponivel localmente. Sync pode ser necessario manualmente." "AVISO" "criacao" }
    } else {
        # Tentar remoting com credencial (obrigatorio ao usar IP)
        try {
            $icParams = @{
                ComputerName = $AADConnectServer
                ScriptBlock  = { Import-Module ADSync -ErrorAction Stop; Start-ADSyncSyncCycle -PolicyType Delta | Out-Null }
                ErrorAction  = "Stop"
            }
            if ($aadSyncCred) { $icParams["Credential"] = $aadSyncCred }
            Invoke-Command @icParams
            Add-Log $logCriar "Sync Delta iniciado remotamente em $AADConnectServer." "OK" "criacao"
            $syncOk = $true
        } catch {
            $errMsg = $_.Exception.Message
            if ($errMsg -match "CannotUseIPAddress" -or $errMsg -match "credencial" -or $errMsg -match "credential") {
                Add-Log $logCriar "Remoting falhou: credencial do servidor ADSync nao configurada." "AVISO" "criacao"
                Add-Log $logCriar "Acesse Configuracoes (icone engrenagem) e informe usuario/senha do servidor $AADConnectServer." "AVISO" "criacao"
            } else {
                Add-Log $logCriar "Remoting para $AADConnectServer falhou: $errMsg" "AVISO" "criacao"
            }
            # Fallback: tentar localmente
            try {
                Import-Module ADSync -ErrorAction Stop
                Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
                Add-Log $logCriar "Sync Delta iniciado localmente (fallback)." "OK" "criacao"
                $syncOk = $true
            } catch {
                Add-Log $logCriar "ADSync nao disponivel. O sync ocorrera automaticamente no proximo ciclo (~30 min)." "AVISO" "criacao"
            }
        }
    }

    # Polling M365 - aguarda o usuario aparecer (max 5 min)
    $lblStCriar.Text = "Etapa 5/6: Aguardando usuario aparecer no M365..."
    Add-Log $logCriar "Aguardando propagacao no M365 (polling a cada 15s, max 5 min)..." "INFO" "criacao"
    $m365Ok = $false
    $tentativasSync = if ($syncOk) { 16 } else { 20 }  # ~4 ou 5 min
    for ($s = 1; $s -le $tentativasSync; $s++) {
        $lblStCriar.Text = "Etapa 5/6: Aguardando M365... tentativa $s/$tentativasSync"
        $pbCriar.Value   = [math]::Min(70, 55 + (($s / $tentativasSync) * 15))
        DoEvents-Sleep -Seconds 15
        try {
            $hdrs = Get-GraphToken
            Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$PEmail" -Headers $hdrs -ErrorAction Stop | Out-Null
            Add-Log $logCriar "Usuario encontrado no M365 apos $($s * 15)s." "OK" "criacao"
            $m365Ok = $true; break
        } catch { }
    }
    if (-not $m365Ok) { Add-Log $logCriar "Usuario nao apareceu no M365 em tempo esperado. Verifique o AAD Connect." "AVISO" "criacao" }
    $pbCriar.Value = 70

    # ETAPA 6 - Licenca M365
    $nomeLic = ""
    if ($PLic -eq "Nao atribuir licenca") {
        Add-Log $logCriar "ETAPA 6: Licenca ignorada pelo usuario." "AVISO" "criacao"
        $pbCriar.Value = 100
    } else {
        $lblStCriar.Text = "Etapa 6/6: Atribuindo licenca Microsoft 365..."
        Add-Log $logCriar "ETAPA 6: Atribuindo licenca M365" "ETAPA" "criacao"
        try {
            $hdrs = Get-GraphToken
            Add-Log $logCriar "Token Graph API obtido." "OK" "criacao"

            $u365 = $null
            for ($i = 1; $i -le 10; $i++) {
                $lblStCriar.Text = "Etapa 6/6: Procurando usuario no M365... tentativa $i/10"
                try { $u365 = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$PEmail" -Headers $hdrs -ErrorAction Stop; break }
                catch { if ($i -lt 10) { DoEvents-Sleep -Seconds 30 } }
                $pbCriar.Value = [math]::Min(90, 70 + ($i * 1.5))
            }

            if (-not $u365) {
                Add-Log $logCriar "Usuario nao encontrado no M365 apos tentativas. Atribua manualmente." "AVISO" "criacao"
            } else {
                Add-Log $logCriar "Usuario encontrado no M365!" "OK" "criacao"
                $bodyLoc = @{ usageLocation = $UsageLocation } | ConvertTo-Json
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$PEmail" -Method Patch -Headers $hdrs -Body $bodyLoc | Out-Null
                Add-Log $logCriar "UsageLocation: $UsageLocation" "OK" "criacao"
                DoEvents-Sleep -Seconds 10

                $lics    = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -Headers $hdrs).value
                $skuIds  = @()
                if ($PLic -eq "Business Standard (SPB)") {
                    $l = $lics | Where-Object { $_.skuPartNumber -eq "SPB" -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
                    if ($l) { $skuIds = @($l.skuId); $nomeLic = "Business Standard" }
                    else    { Add-Log $logCriar "Licenca SPB nao disponivel!" "ERRO" "criacao" }
                } elseif ($PLic -eq "Apps Enterprise + E1") {
                    $la = $lics | Where-Object { $_.skuPartNumber -eq "OFFICESUBSCRIPTION" -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
                    $le = $lics | Where-Object { $_.skuPartNumber -eq "STANDARDPACK"       -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
                    if ($la -and $le) { $skuIds = @($la.skuId, $le.skuId); $nomeLic = "Apps Enterprise + E1" }
                    else { Add-Log $logCriar "Licencas Apps+E1 nao disponiveis!" "ERRO" "criacao" }
                } else {
                    $std  = $lics | Where-Object { $_.skuPartNumber -eq "SPB"              -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
                    $apps = $lics | Where-Object { $_.skuPartNumber -eq "OFFICESUBSCRIPTION"-and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
                    $e1   = $lics | Where-Object { $_.skuPartNumber -eq "STANDARDPACK"      -and ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 }
                    if ($std)          { $skuIds = @($std.skuId); $nomeLic = "Business Standard" }
                    elseif ($apps -and $e1) { $skuIds = @($apps.skuId, $e1.skuId); $nomeLic = "Apps Enterprise + E1" }
                    else { Add-Log $logCriar "Nenhuma licenca disponivel!" "AVISO" "criacao" }
                }

                $jaTemLics = @()
                try { $jaTemLics = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($u365.id)/licenseDetails" -Headers $hdrs).value.skuId } catch {}

                $licAtribuida = $false
                foreach ($id in $skuIds) {
                    if ($jaTemLics -contains $id) {
                        Add-Log $logCriar "Licenca ja existente (ignorando conflito): $id" "AVISO" "criacao"
                        $licAtribuida = $true; continue
                    }
                    for ($tentativa = 1; $tentativa -le 3; $tentativa++) {
                        try {
                            $hdrsUtf8 = @{ Authorization = $hdrs.Authorization; "Content-Type" = "application/json; charset=utf-8" }
                            $bL = @{ addLicenses = @(@{ skuId = $id }); removeLicenses = @() } | ConvertTo-Json -Depth 5
                            Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($u365.id)/assignLicense" `
                                -Method Post -Headers $hdrsUtf8 -Body ([System.Text.Encoding]::UTF8.GetBytes($bL)) -ErrorAction Stop | Out-Null
                            $licAtribuida = $true; break
                        } catch {
                            if ($tentativa -lt 3) { Add-Log $logCriar "Tentativa $tentativa/3 falhou para $id. Aguardando 10s..." "AVISO" "criacao"; DoEvents-Sleep -Seconds 10 }
                            else { Add-Log $logCriar "Erro ao atribuir licenca $id apos 3 tentativas: $_" "ERRO" "criacao" }
                        }
                    }
                }
                if ($nomeLic -and $licAtribuida) { Add-Log $logCriar "Licenca configurada: $nomeLic" "OK" "criacao" }
            }
        } catch { Add-Log $logCriar "Erro na licenca M365: $_" "AVISO" "criacao"; Add-Log $logCriar "Atribua manualmente em admin.microsoft.com" "AVISO" "criacao" }
        $pbCriar.Value = 100
    }

    # Resumo
    Add-Log $logCriar "" ""
    Add-Log $logCriar "======================================" "ETAPA" "criacao"
    Add-Log $logCriar "  USUARIO CRIADO COM SUCESSO!"        "OK" "criacao"
    Add-Log $logCriar "  Nome    : $NomeExibido"             "OK" "criacao"
    Add-Log $logCriar "  Login   : $PUser"                   "OK" "criacao"
    Add-Log $logCriar "  Email   : $PEmail"                  "OK" "criacao"
    Add-Log $logCriar "  Grupos  : $gOk copiado(s)"          "OK" "criacao"
    Add-Log $logCriar "  Licenca : $(if($nomeLic){$nomeLic}else{'N/A'})" "OK" "criacao"
    Add-Log $logCriar "======================================" "ETAPA" "criacao"

    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    [PSCustomObject]@{
        DataHora  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Nome      = $NomeExibido; Username = $PUser; Email = $PEmail
        Setor     = $Setor; Template = $PTmpl
        Licenca   = if ($nomeLic) { $nomeLic } else { "N/A" }
        Grupos    = $gOk; CriadoPor = $env:USERNAME
    } | Export-Csv -Path $HistCriacaoCsv -Append -NoTypeInformation -Encoding UTF8

    $lblStCriar.Text = "CONCLUIDO! Usuario '$PUser' criado com sucesso."; $lblStCriar.ForeColor = $corVerde
    [System.Windows.Forms.MessageBox]::Show(
        "Usuario criado com sucesso!`n`nNome    : $NomeExibido`nLogin   : $PUser`nEmail   : $PEmail`nSenha   : $PSenha$(if($nomeLic){"`nLicenca : $nomeLic"})",
        "Sucesso","OK","Information")
    RestoreCriar
})

# ============================================================
# ACAO: DESLIGAR COLABORADOR
# ============================================================
$btnDesligar.Add_Click({
    $DUser = $dUsername.Text.Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($DUser)) {
        [System.Windows.Forms.MessageBox]::Show("Informe o nome ou username.","Campo obrigatorio","OK","Warning"); return
    }

    $msg  = "ATENCAO! Esta acao e IRREVERSIVEL.`n`nColaborador : $DUser`nEmail       : $DUser@$Dominio`n`n"
    $msg += "Acoes que serao executadas:`n"
    $msg += "  - Redefinir senha com senha aleatoria`n  - Desabilitar e expirar conta no AD`n"
    $msg += "  - Mover para OU de desabilitados`n  - Remover gestor e subordinados`n"
    $msg += "  - Ocultar do catalogo de enderecos`n  - Zerar horarios de logon`n"
    $msg += "  - Remover de todos os grupos AD`n"
    if ($dChkExchange.Checked) { $msg += "  - Converter mailbox para compartilhada`n" }
    if ($dChkLicencas.Checked) { $msg += "  - Remover licencas Microsoft 365`n" }
    $msg += "`nConfirma o desligamento?"

    if (([System.Windows.Forms.MessageBox]::Show($msg,"Confirmar Desligamento","YesNo","Warning")) -ne "Yes") { return }

    $btnDesligar.Enabled = $false; $btnVerificar.Enabled = $false
    $dUsername.ReadOnly = $true; $dChkExchange.Enabled = $false; $dChkLicencas.Enabled = $false
    $btnDesligar.BackColor = [System.Drawing.Color]::FromArgb(80,84,110); $btnDesligar.Text = "PROCESSANDO..."

    function RestoreDesligar {
        $btnDesligar.Enabled = $true; $btnVerificar.Enabled = $true
        $dUsername.ReadOnly = $false; $dChkExchange.Enabled = $true; $dChkLicencas.Enabled = $true
        $btnDesligar.BackColor = $corVermelho; $btnDesligar.Text = ([char]0x26A0) + "  DESLIGAR COLABORADOR"
    }

    $logDesligar.Clear(); $pbDesligar.Value = 0
    $DMailbox = "$DUser@$Dominio"; $DDate = Get-Date -Format "yyyy-MM-dd HH:mm"
    $PassWD   = [System.Web.Security.Membership]::GeneratePassword(14,4)

    Add-Log $logDesligar "======================================" "ETAPA" "desligamento"
    Add-Log $logDesligar "  INICIO DO PROCESSO DE DESLIGAMENTO" "ETAPA" "desligamento"
    Add-Log $logDesligar "  Usuario : $DUser | Email: $DMailbox" "ETAPA" "desligamento"
    Add-Log $logDesligar "  Operador: $env:USERNAME"              "ETAPA" "desligamento"
    Add-Log $logDesligar "======================================" "ETAPA" "desligamento"

    # ETAPA 1 - Verificar AD
    $lblStDesligar.Text = "Etapa 1/5: Verificando usuario no Active Directory..."; $pbDesligar.Value = 5
    Add-Log $logDesligar "ETAPA 1: Verificando usuario no Active Directory" "ETAPA" "desligamento"
    $adMsg = if ($script:ADRemote) { " (remoting -> $($script:AADConnectServer))" } else { " (local)." }
    if (Ensure-ADModule) { Add-Log $logDesligar "Modulo ActiveDirectory pronto$adMsg" "OK" "desligamento" }
    else {
        Add-Log $logDesligar "Falha ao conectar ao AD (todas as tentativas):" "ERRO" "desligamento"
        foreach ($errLine in ($script:ADModuleError -split "`n")) {
            if ($errLine.Trim()) { Add-Log $logDesligar $errLine.Trim() "ERRO" "desligamento" }
        }
        Add-Log $logDesligar "DICA: Execute no servidor como Admin:" "AVISO" "desligamento"
        Add-Log $logDesligar "  Enable-PSRemoting -Force" "AVISO" "desligamento"
        [System.Windows.Forms.MessageBox]::Show(
            "Falha ao conectar ao servidor AD.`n`n$($script:ADModuleError)`n`nVerifique se o WinRM esta ativo no servidor:`n  Enable-PSRemoting -Force",
            "Erro de Conexao","OK","Error")
        RestoreDesligar; return
    }
    try {
        $DAD   = Get-ADUser -Identity $DUser -Properties MemberOf,DirectReports,Manager,DisplayName,Department -ErrorAction Stop
        $DNome = if ($DAD.DisplayName) { $DAD.DisplayName } else { "$($DAD.GivenName) $($DAD.Surname)" }
        $DSetor= if ($DAD.Department)  { $DAD.Department  } else { "N/A" }
        Add-Log $logDesligar "Usuario encontrado: $DNome | Setor: $DSetor" "OK" "desligamento"
    } catch {
        Add-Log $logDesligar "Usuario '$DUser' nao encontrado no AD!" "ERRO" "desligamento"
        [System.Windows.Forms.MessageBox]::Show("Usuario '$DUser' nao encontrado no Active Directory.","Nao encontrado","OK","Error")
        RestoreDesligar; return
    }
    $pbDesligar.Value = 10

    # ETAPA 2 - Acoes AD
    $lblStDesligar.Text = "Etapa 2/5: Executando acoes no Active Directory..."
    Add-Log $logDesligar "ETAPA 2: Acoes no Active Directory" "ETAPA" "desligamento"

    try { Set-ADAccountPassword -Identity $DUser -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $PassWD -Force) -ErrorAction Stop; Add-Log $logDesligar "Senha redefinida com senha aleatoria segura." "OK" "desligamento" }
    catch { Add-Log $logDesligar "Erro ao redefinir senha: $_" "ERRO" "desligamento" }

    try { Disable-ADAccount -Identity $DUser -ErrorAction Stop; Add-Log $logDesligar "Conta desabilitada." "OK" "desligamento" }
    catch { Add-Log $logDesligar "Erro ao desabilitar: $_" "ERRO" "desligamento" }

    try { Set-ADAccountExpiration -Identity $DUser -DateTime "2000-01-01" -ErrorAction Stop; Add-Log $logDesligar "Conta expirada." "OK" "desligamento" }
    catch { Add-Log $logDesligar "Erro ao expirar: $_" "ERRO" "desligamento" }
    $pbDesligar.Value = 18

    try { Move-ADObject -Identity $DAD.DistinguishedName -TargetPath $TargetOU -ErrorAction Stop; Add-Log $logDesligar "Movido para OU de desabilitados." "OK" "desligamento" }
    catch { Add-Log $logDesligar "Erro ao mover OU: $_" "AVISO" "desligamento" }

    try { Set-ADUser -Identity $DUser -Description "Desligado em $DDate" -ErrorAction Stop; Add-Log $logDesligar "Descricao: 'Desligado em $DDate'" "OK" "desligamento" }
    catch { Add-Log $logDesligar "Erro descricao: $_" "AVISO" "desligamento" }

    try { Set-ADUser -Identity $DUser -Clear Manager -ErrorAction Stop; Add-Log $logDesligar "Gestor removido." "OK" "desligamento" }
    catch { Add-Log $logDesligar "Erro gestor: $_" "AVISO" "desligamento" }

    $subOk = 0
    foreach ($r in $DAD.DirectReports) {
        try { Set-ADUser -Identity $r -Clear Manager -ErrorAction Stop; $subOk++ } catch {}
        [System.Windows.Forms.Application]::DoEvents()
    }
    if ($subOk -gt 0) { Add-Log $logDesligar "$subOk subordinado(s) atualizados." "OK" "desligamento" }

    try { Set-ADUser -Identity $DUser -Replace @{ msExchHideFromAddressLists = $true } -ErrorAction Stop; Add-Log $logDesligar "Ocultado do catalogo de enderecos." "OK" "desligamento" }
    catch { Add-Log $logDesligar "Erro catalogo: $_" "AVISO" "desligamento" }

    try { Set-ADUser -Identity $DUser -Replace @{ LogonHours = ([byte[]](0..20 | ForEach-Object { 0 })) } -ErrorAction Stop; Add-Log $logDesligar "Horarios de logon zerados." "OK" "desligamento" }
    catch { Add-Log $logDesligar "Erro logon hours: $_" "AVISO" "desligamento" }
    $pbDesligar.Value = 35

    # ETAPA 3 - Grupos
    $lblStDesligar.Text = "Etapa 3/5: Removendo grupos do Active Directory..."
    Add-Log $logDesligar "ETAPA 3: Remocao dos grupos AD" "ETAPA" "desligamento"
    $dGOk = 0; $dGErr = 0
    foreach ($gDN in $DAD.MemberOf) {
        try {
            $grp = Get-ADGroup -Identity $gDN
            Remove-ADGroupMember -Identity $grp.DistinguishedName -Members $DUser -Confirm:$false -ErrorAction Stop
            Add-Log $logDesligar "Removido do grupo: $($grp.Name)" "OK" "desligamento"; $dGOk++
        } catch {
            $nG = ($gDN -split ',')[0] -replace 'CN=',''
            Add-Log $logDesligar "Erro grupo '$nG': $_" "ERRO" "desligamento"; $dGErr++
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    Add-Log $logDesligar "Grupos removidos: $dGOk | Erros: $dGErr" "INFO" "desligamento"
    $pbDesligar.Value = 55

    # ETAPA 4 - Exchange (converter para shared mailbox)
    # Roda em Runspace separado para nao travar a interface grafica
    if ($dChkExchange.Checked) {
        $lblStDesligar.Text = "Etapa 4/5: Convertendo mailbox para compartilhada..."
        Add-Log $logDesligar "ETAPA 4: Convertendo mailbox para compartilhada" "ETAPA" "desligamento"
        Add-Log $logDesligar "Abrindo sessao Exchange Online (aguarde a janela de login)..." "INFO" "desligamento"

        # Arquivo temporario para comunicar resultado do runspace
        $tmpResult = [System.IO.Path]::GetTempFileName()
        $tmpResult = [System.IO.Path]::ChangeExtension($tmpResult, ".txt")

        $rsScript = [scriptblock]::Create(@"
param(`$mailbox, `$resultFile)
try {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction SilentlyContinue | Out-Null
        Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop | Out-Null
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Connect-ExchangeOnline -ShowBanner:`$false -ErrorAction Stop
    Set-Mailbox -Identity `$mailbox -Type Shared -ErrorAction Stop
    Disconnect-ExchangeOnline -Confirm:`$false -ErrorAction SilentlyContinue
    Set-Content -Path `$resultFile -Value "OK" -Encoding UTF8
} catch {
    Set-Content -Path `$resultFile -Value "ERRO:`$(`$_.Exception.Message)" -Encoding UTF8
}
"@)

        $rs  = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.Open()
        $ps  = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript($rsScript).AddArgument($DMailbox).AddArgument($tmpResult)
        $handle = $ps.BeginInvoke()

        # Aguarda conclusao mantendo a UI responsiva
        $elapsed = 0
        while (-not $handle.IsCompleted) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 300
            $elapsed += 300
            if ($elapsed % 5000 -eq 0) {
                $lblStDesligar.Text = "Etapa 4/5: Aguardando Exchange Online... ($([int]($elapsed/1000))s)"
            }
            # Timeout de 3 minutos
            if ($elapsed -ge 180000) {
                Add-Log $logDesligar "Timeout: Exchange Online demorou mais de 3 minutos." "AVISO" "desligamento"
                break
            }
        }
        [void]$ps.EndInvoke($handle)
        $ps.Dispose(); $rs.Close(); $rs.Dispose()

        # Ler resultado do arquivo temporario
        $convertOk = $false
        if (Test-Path $tmpResult) {
            $exoResult = Get-Content $tmpResult -Raw -Encoding UTF8
            Remove-Item $tmpResult -Force -ErrorAction SilentlyContinue

            if ($exoResult -match '^OK') {
                Add-Log $logDesligar "Mailbox convertida para compartilhada com sucesso." "OK" "desligamento"
                $convertOk = $true
            } else {
                $exoErr = $exoResult -replace '^ERRO:',''
                Add-Log $logDesligar "Exchange Online falhou: $exoErr" "AVISO" "desligamento"
                Add-Log $logDesligar "Tentando via Graph API como fallback..." "INFO" "desligamento"

                # Fallback: Graph API beta
                try {
                    $hdrs     = Get-GraphToken
                    $mu       = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$DMailbox" -Headers $hdrs -ErrorAction Stop
                    $hdrsUtf8 = @{ Authorization = $hdrs.Authorization; "Content-Type" = "application/json; charset=utf-8" }
                    Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/users/$($mu.id)/mailboxSettings" `
                        -Method Patch -Headers $hdrsUtf8 `
                        -Body ([System.Text.Encoding]::UTF8.GetBytes('{"userPurpose":"shared"}')) -ErrorAction Stop
                    Add-Log $logDesligar "Mailbox convertida via Graph API." "OK" "desligamento"
                    $convertOk = $true
                } catch {
                    $graphErr = $_.Exception.Message
                    Add-Log $logDesligar "Graph API tambem falhou: $graphErr" "ERRO" "desligamento"
                    Add-Log $logDesligar "ACAO MANUAL: Exchange Admin Center > Caixas de correio > Converter para compartilhada" "AVISO" "desligamento"
                }
            }
        } else {
            Add-Log $logDesligar "Nao foi possivel obter resultado do Exchange Online." "ERRO" "desligamento"
            Add-Log $logDesligar "ACAO MANUAL: Exchange Admin Center > Caixas de correio > Converter para compartilhada" "AVISO" "desligamento"
        }
    } else { Add-Log $logDesligar "ETAPA 4: Exchange ignorado pelo usuario." "AVISO" "desligamento" }
    $pbDesligar.Value = 75

    # ETAPA 5 - Licencas
    if ($dChkLicencas.Checked) {
        $lblStDesligar.Text = "Etapa 5/5: Removendo licencas Microsoft 365..."
        Add-Log $logDesligar "ETAPA 5: Remocao de licencas via Graph API" "ETAPA" "desligamento"
        try {
            $hdrs = Get-GraphToken
            Add-Log $logDesligar "Token Graph API obtido." "OK" "desligamento"
            $mu   = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$DMailbox" -Headers $hdrs -ErrorAction Stop
            $lics = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($mu.id)/licenseDetails" -Headers $hdrs).value
            $lOk  = 0
            if ($lics.Count -eq 0) {
                Add-Log $logDesligar "Nenhuma licenca encontrada." "AVISO" "desligamento"
            } else {
                $hdrsUtf8 = @{ Authorization = $hdrs.Authorization; "Content-Type" = "application/json; charset=utf-8" }
                foreach ($l in $lics) {
                    for ($tentativa = 1; $tentativa -le 3; $tentativa++) {
                        try {
                            $bR = @{ addLicenses = @(); removeLicenses = @($l.skuId) } | ConvertTo-Json -Depth 5
                            Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($mu.id)/assignLicense" `
                                -Method Post -Headers $hdrsUtf8 -Body ([System.Text.Encoding]::UTF8.GetBytes($bR)) -ErrorAction Stop
                            Add-Log $logDesligar "Licenca removida: $($l.skuPartNumber)" "OK" "desligamento"; $lOk++; break
                        } catch {
                            if ($tentativa -lt 3) { Add-Log $logDesligar "Tentativa $tentativa/3 falhou para $($l.skuPartNumber). Aguardando 10s..." "AVISO" "desligamento"; DoEvents-Sleep -Seconds 10 }
                            else { Add-Log $logDesligar "Erro ao remover $($l.skuPartNumber) apos 3 tentativas: $_" "ERRO" "desligamento" }
                        }
                    }
                    [System.Windows.Forms.Application]::DoEvents()
                }
                Add-Log $logDesligar "$lOk licenca(s) removida(s)." "INFO" "desligamento"
            }
        } catch {
            Add-Log $logDesligar "Erro ao remover licencas: $_" "ERRO" "desligamento"
            Add-Log $logDesligar "Remova manualmente em admin.microsoft.com" "AVISO" "desligamento"
        }
    } else { Add-Log $logDesligar "ETAPA 5: Licencas ignoradas pelo usuario." "AVISO" "desligamento" }
    $pbDesligar.Value = 100

    # Resumo
    Add-Log $logDesligar "" ""
    Add-Log $logDesligar "======================================" "ETAPA" "desligamento"
    Add-Log $logDesligar "  DESLIGAMENTO CONCLUIDO!"             "OK" "desligamento"
    Add-Log $logDesligar "  Username : $DUser"                   "OK" "desligamento"
    Add-Log $logDesligar "  Nome     : $DNome"                   "OK" "desligamento"
    Add-Log $logDesligar "  Setor    : $DSetor"                  "OK" "desligamento"
    Add-Log $logDesligar "  Grupos   : $dGOk removido(s)"        "OK" "desligamento"
    Add-Log $logDesligar "  Operador : $env:USERNAME"             "OK" "desligamento"
    Add-Log $logDesligar "======================================" "ETAPA" "desligamento"

    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    [PSCustomObject]@{
        DataHora      = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Username      = $DUser; NomeCompleto = $DNome; Email = $DMailbox
        Setor         = $DSetor; GruposRemov = $dGOk; Operador = $env:USERNAME
    } | Export-Csv -Path $HistDesligCsv -Append -NoTypeInformation -Encoding UTF8

    $lblStDesligar.Text = "CONCLUIDO! Colaborador '$DUser' desligado com sucesso."; $lblStDesligar.ForeColor = $corVerde
    [System.Windows.Forms.MessageBox]::Show(
        "Colaborador desligado com sucesso!`n`nUsername : $DUser`nNome     : $DNome`nSetor    : $DSetor`nGrupos   : $dGOk removido(s)`n`nLog salvo em:`n$LogDir",
        "Desligamento Concluido","OK","Information")

    RestoreDesligar
    $dUsername.Clear(); $pbDesligar.Value = 0
    $dLblNome.Text = "Nome no AD: ---"; $dLblNome.ForeColor = $corTextoCinza
    $dLblGrupos.Text = "Grupos AD: ---"; $dLblEmail.Text = "Email: ---"
    $lblStDesligar.Text = "Pronto. Informe o nome ou login e clique em DESLIGAR COLABORADOR."
    $lblStDesligar.ForeColor = $corTextoCinza
})

# ============================================================
# INICIAR
# ============================================================
Switch-Tab "criar"
$logCriar.SelectionColor = $corTextoCinza;    $logCriar.AppendText("Pronto. Preencha os campos e clique em CRIAR USUARIO.`r`n")
$logDesligar.SelectionColor = $corTextoCinza; $logDesligar.AppendText("Pronto. Informe o nome ou login e clique em DESLIGAR COLABORADOR.`r`n")
$cNome.Focus()

$form.Add_FormClosed({
    if ($script:ADSession) { Remove-PSSession $script:ADSession -ErrorAction SilentlyContinue; $script:ADSession = $null }
    [System.Windows.Forms.Application]::ExitThread()
})
[System.Windows.Forms.Application]::Run($form)
