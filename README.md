# 🔷 Ober TI — Gestão de Usuários
**Active Directory + Microsoft 365 | Onboarding & Offboarding Automatizado**

> Interface gráfica em PowerShell + WinForms para criação e desligamento de colaboradores com integração completa ao Active Directory e Microsoft 365.

---

## 📸 Interface

> *Aba Criação — Aba Desligamento*
> *(adicione screenshots aqui)*

---

## ✨ Funcionalidades

### 🟢 Aba Criação — Onboarding (6 etapas)

| Etapa | Descrição |
|-------|-----------|
| 1 | Verificação de pré-requisitos e conectividade com o AD |
| 2 | Leitura do usuário template (OU, cargo, grupos, gestor, departamento) |
| 3 | Criação do usuário no Active Directory com todos os atributos |
| 4 | Cópia automática de grupos de segurança do template |
| 5 | Sincronização com Azure AD via ADConnect + polling de propagação no M365 |
| 6 | Atribuição automática de licença Microsoft 365 via Graph API |

### 🔴 Aba Desligamento — Offboarding (5 etapas)

| Etapa | Descrição |
|-------|-----------|
| 1 | Verificação do usuário no Active Directory |
| 2 | Redefinição de senha aleatória, desabilitação, expiração, movimentação para OU de desabilitados, remoção de gestor/subordinados, ocultação do catálogo de endereços e bloqueio de logon |
| 3 | Remoção de todos os grupos AD |
| 4 | Conversão da mailbox para compartilhada via Exchange Online (fallback: Graph API beta) |
| 5 | Remoção de licenças Microsoft 365 via Graph API (com 3 tentativas por licença) |

---

## 🔐 Segurança

- Credenciais do Azure AD App Registration e servidor ADSync protegidas com **criptografia DPAPI** (vinculada ao usuário/máquina)
- Senha de desligamento gerada aleatoriamente (14 caracteres, mínimo 4 especiais)
- Suporte a múltiplos métodos de autenticação remota: **Kerberos**, **NTLM** e credencial explícita

---

## ⚙️ Requisitos

- Windows 10/11 ou Windows Server 2016+
- PowerShell 5.1+
- Acesso de rede ao servidor AD / AAD Connect
- **Azure AD App Registration** com permissões:
  - `User.ReadWrite.All`
  - `Directory.ReadWrite.All`
  - `Organization.Read.All`
- **Não é necessário RSAT instalado** — o módulo ActiveDirectory é carregado via PSSession remota automaticamente

---

## 🚀 Como usar

### 1. Configuração inicial

Edite o arquivo `ober_gestao_usuarios.ps1` e preencha as variáveis no topo do arquivo:

```powershell
$CFG_TenantId     = "SEU_TENANT_ID"
$CFG_ClientId     = "SEU_CLIENT_ID"
$CFG_ClientSecret = "SEU_CLIENT_SECRET"

$CFG_Dominio      = "suaempresa.com.br"
$CFG_SenhaInicial = "@SenhaInicial2025"
$CFG_Servidor     = "10.0.0.1"          # IP ou hostname do servidor ADConnect
$CFG_SyncUser     = "DOMINIO\admin"
$CFG_SyncSenha    = "SenhaDo Servidor"
$CFG_TargetOU     = "OU=Usuarios Desabilitados,DC=suaempresa,DC=com,DC=br"
```

As credenciais são salvas criptografadas em `config.xml` após o primeiro uso. Não é necessário configurar novamente em execuções futuras **na mesma máquina**.

### 2. Executar

Dê um duplo clique no arquivo `gestao_usuarios.bat` (solicita elevação UAC automaticamente).

Ou execute diretamente via PowerShell como Administrador:

```powershell
powershell -ExecutionPolicy Bypass -File ".\ober_gestao_usuarios.ps1"
```

### 3. Criar usuário

1. Vá para a aba **Criação**
2. Preencha: Primeiro Nome, Sobrenome e Usuário Template
3. Login e e-mail são gerados automaticamente
4. Selecione a licença Microsoft 365 desejada
5. Clique em **CRIAR USUÁRIO**

### 4. Desligar colaborador

1. Vá para a aba **Desligamento**
2. Digite o username ou nome do colaborador
3. Clique em **Verificar** para confirmar os dados no AD
4. Marque as opções desejadas (Exchange Online, Remoção de Licenças)
5. Clique em **⚠️ DESLIGAR COLABORADOR**

---

## 📁 Estrutura de arquivos

```
📦 gestao-usuarios/
├── gestao_usuarios.bat          # Launcher (eleva UAC automaticamente)
├── ober_gestao_usuarios.ps1     # Script principal
├── config.xml                   # Credenciais criptografadas (gerado automaticamente)
└── logs/
    ├── criacao_YYYY-MM.log      # Log mensal de criações
    ├── desligamento_YYYY-MM.log # Log mensal de desligamentos
    ├── historico_criacao.csv    # Histórico completo de criações
    └── historico_desligamentos.csv # Histórico completo de desligamentos
```

---

## 🛠️ Detalhes técnicos

- **PSSession remota**: carrega o módulo ActiveDirectory via WinRM em qualquer máquina da rede, sem necessidade de RSAT. Suporta autenticação por Kerberos (hostname) e NTLM (IP).
- **Graph API token cache**: token OAuth 2.0 com cache de 50 minutos e renovação automática.
- **Polling de propagação M365**: após o sync ADConnect, verifica a cada 15 segundos (até 5 minutos) se o usuário apareceu no Microsoft 365 antes de tentar atribuir a licença.
- **Exchange Online em Runspace separado**: a conversão de mailbox roda em thread paralela para não travar a interface gráfica, com timeout de 3 minutos.
- **Fallback Exchange → Graph API**: se o Exchange Online falhar, tenta converter via endpoint beta da Graph API (`mailboxSettings`).

---

## 📊 Logs e rastreabilidade

Todas as operações geram:
- Log em tempo real na interface (com ícones coloridos por tipo: ✔ OK | ✘ ERRO | ⚠ AVISO | ➤ INFO | ▶ ETAPA)
- Arquivo de log mensal em `logs/`
- Registro em CSV com data/hora, operador, usuário e resultado

---

## 🤝 Contribuição

Pull requests são bem-vindos! Para mudanças maiores, abra uma issue primeiro para discutirmos o que você gostaria de mudar.

---

## 📄 Licença

MIT License — veja o arquivo [LICENSE](LICENSE) para detalhes.

---

*Desenvolvido por [Seu Nome] — Ober Tecnologia da Informação*
