# 📊 Indicadores RH — AliseoSA

> Sistema web para acompanhamento do efetivo e comparação automática das planilhas semanais de colaboradores, com autenticação via Active Directory e relatórios em Excel.

![Python](https://img.shields.io/badge/Python-3.12-blue?logo=python)
![Flask](https://img.shields.io/badge/Flask-3.1-lightgrey?logo=flask)
![Docker](https://img.shields.io/badge/Docker-Compose-2496ED?logo=docker)
![Bootstrap](https://img.shields.io/badge/Bootstrap-5.3-purple?logo=bootstrap)

---

## 🎯 O que faz

A cada semana o RH envia uma planilha `.xlsx` com todos os colaboradores ativos. Este sistema:

- **Compara** a nova planilha com a anterior automaticamente
- **Detecta** admissões, desligamentos e alterações de cargo/departamento/gestor
- **Exibe** as diferenças em uma interface web com abas e filtro por nome
- **Lista** o efetivo completo com filtros por tipo, cargo, departamento e gestor
- **Gera** relatório em Excel formatado com cores para download
- **Controla** acesso via Active Directory (um ou múltiplos grupos)

---

## ✨ Funcionalidades

| Funcionalidade | Descrição |
|---|---|
| 🔐 Login AD | Autenticação com usuário e senha de rede (LDAP/LDAPS) |
| 👥 Múltiplos grupos | Libere acesso para TI, Diretoria, RH — quantos grupos quiser |
| 🌙 Dark mode | Alternância claro/escuro salva na preferência do usuário |
| 📊 Dashboard | Cards clicáveis com resumo da última comparação |
| 📤 Upload de planilha | Envio do `.xlsx` via drag‑and‑drop ou seleção |
| 🆕 Admissões | Colaboradores que entraram entre as duas planilhas |
| ❌ Desligamentos | Colaboradores que saíram entre as duas planilhas |
| 🔄 Alterações | Mudanças de CARGO, DEPARTAMENTO ou GESTOR |
| 👤 Efetivo completo | Lista todos os colaboradores com filtros avançados |
| 📜 Histórico | Dashboard com todas as comparações anteriores |
| 📥 Export Excel | Relatório com células coloridas por tipo de evento |

---

## 🧱 Stack

- **Backend:** Python 3.12 + Flask 3.1 + Gunicorn
- **Planilhas:** pandas + openpyxl
- **Autenticação:** ldap3 (Active Directory / LDAP)
- **Frontend:** Bootstrap 5.3 + Bootstrap Icons
- **Infraestrutura:** Docker + Docker Compose

---

## 📋 Pré-requisitos

### No servidor de destino
- Ubuntu 22.04 / 24.04 (ou qualquer Linux com Docker)
- Docker Engine 24+
- Docker Compose v2+
- Acesso de rede ao servidor AD na porta **389** (LDAP) ou **636** (LDAPS)
- Porta **5000** liberada no firewall (ou outra porta que você configurar)

### Instalar Docker no Ubuntu (se ainda não tiver)
```bash
curl -fsSL https://get.docker.com | sh
sudo usermod -aG docker $USER && newgrp docker
docker --version
```

---

## 🚀 Deploy — Passo a Passo

### 1. Clonar o repositório no servidor

```bash
cd /opt
sudo git clone https://github.com/tuliodutra27/Indicadores-RH.git
sudo chown -R $USER:$USER /opt/App-RH
cd /opt/App-RH/web
```

### 2. Criar e configurar o `.env`

```bash
cp .env.example .env
nano .env
```

Preencha com os dados reais do seu ambiente:

```env
# Gere com: python3 -c "import secrets; print(secrets.token_hex(32))"
SECRET_KEY=cole-aqui-a-chave-gerada

APP_PORT=5000

# Active Directory
AD_SERVER=ldap://192.168.1.10
AD_PORT=389
AD_USE_SSL=false
AD_DOMAIN=aliseo.local
AD_BASE_DN=DC=aliseo,DC=local

# Grupos com acesso — separe múltiplos grupos com ";"
AD_TI_GROUP_CN=GRP_SA_OPE_TI_RW;GRP_SA_OPE_DIRETORIA_RW

DEV_MODE=false
DEV_USERS={}
```

> 💡 **Gerar SECRET_KEY segura:**
> ```bash
> python3 -c "import secrets; print(secrets.token_hex(32))"
> ```

### 3. Subir o container

```bash
docker compose up -d --build
```

A imagem será construída (~2 min na primeira vez) e o container iniciará automaticamente.

### 4. Verificar que está rodando

```bash
docker compose ps          # deve mostrar status "healthy"
docker compose logs -f     # acompanhar logs em tempo real (Ctrl+C para sair)
```

### 5. Acessar no browser

```
http://IP-DO-SERVIDOR:5000
```

---

## ⚙️ Variáveis de ambiente

| Variável | Obrigatória | Padrão | Descrição |
|---|---|---|---|
| `SECRET_KEY` | ✅ | — | Chave secreta Flask (gere aleatoriamente) |
| `APP_PORT` | ❌ | `5000` | Porta exposta pelo Docker |
| `AD_SERVER` | ✅ | — | Endereço LDAP do AD (`ldap://IP`) |
| `AD_PORT` | ❌ | `389` | Porta LDAP (636 para LDAPS) |
| `AD_USE_SSL` | ❌ | `false` | Usar LDAPS (true/false) |
| `AD_DOMAIN` | ✅ | — | Domínio AD (ex: `aliseo.local`) |
| `AD_BASE_DN` | ✅ | — | Base DN (ex: `DC=aliseo,DC=local`) |
| `AD_TI_GROUP_CN` | ✅ | — | Grupos AD com acesso — separados por `;` |
| `DEV_MODE` | ❌ | `false` | Login sem AD para testes |
| `DEV_USERS` | ❌ | `{}` | Usuários de teste em JSON |

### Formatos aceitos em `AD_TI_GROUP_CN`

```env
# Um grupo (só o nome)
AD_TI_GROUP_CN=GRP_SA_OPE_TI_RW

# Um grupo com prefixo CN=
AD_TI_GROUP_CN=CN=GRP_SA_OPE_TI_RW

# Múltiplos grupos (separados por ";")
AD_TI_GROUP_CN=GRP_SA_OPE_TI_RW;GRP_SA_OPE_DIRETORIA_RW

# DN completo também funciona
AD_TI_GROUP_CN=CN=GRP_SA_OPE_TI_RW,OU=Seguranca,DC=aliseo,DC=local
```

O acesso é liberado se o usuário pertencer a **qualquer um** dos grupos listados.

---

## 📁 Estrutura de dados persistida

Os dados ficam em `web/data/` (volume Docker montado em `/data` no container):

```
data/
├── historico.json            ← índice de todas as comparações
├── snapshots/                ← cópias das planilhas enviadas (base de comparação)
│   └── snapshot_20250527_143000_26-05.xlsx
├── relatorios/               ← resultados em JSON por comparação
│   └── resultado_20250527_143000.json
└── uploads/                  ← arquivos enviados pelos usuários
```

> ⚠️ **Faça backup periódico da pasta `data/`**, especialmente `snapshots/` e `historico.json`.

---

## 📌 Regras de negócio

### Chave de identificação dos colaboradores

| Tipo | Campo-chave | Chave interna |
|---|---|---|
| **CLT** | MATRÍCULA numérica | `CLT::610` |
| **PJ** | NOME (sem matrícula numérica) | `PJ::FULANO DA SILVA` |

### Campos monitorados para "alteração"
- **CARGO**
- **DEPARTAMENTO**
- **GESTOR**

### Fluxo de comparação
1. **Primeira planilha** → criada como baseline (nenhuma diferença exibida)
2. **Planilhas seguintes** → comparadas contra o snapshot anterior

---

## 🔄 Como atualizar o sistema

```bash
cd /opt/App-RH
git pull origin main
cd web
docker compose up -d --build
```

> ⚠️ **Nota:** após a atualização que renomeou o projeto para "Indicadores RH",
> o nome do container mudou de `analisador-rh` para `indicadores-rh`.
> Na primeira vez após o update, remova o container antigo:
> ```bash
> docker compose down   # para e remove o container antigo
> docker compose up -d --build   # sobe com o novo nome
> ```

---

## 🛑 Comandos úteis

```bash
# Reiniciar (ex: após editar .env sem rebuild)
docker compose restart

# Parar tudo
docker compose down

# Ver logs em tempo real
docker compose logs -f

# Ver logs das últimas 100 linhas
docker compose logs --tail=100

# Reconstruir imagem (após update do código)
docker compose up -d --build

# Verificar saúde do container
docker inspect --format='{{.State.Health.Status}}' indicadores-rh
```

---

## 🧪 Modo desenvolvimento (sem AD)

Para testar sem precisar de Active Directory:

```env
# no .env
DEV_MODE=true
DEV_USERS={"admin":"admin123","rh":"rh123"}
```

```bash
docker compose restart
```

> ⚠️ **NUNCA deixe `DEV_MODE=true` em produção!**

---

## 🔒 Boas práticas de segurança

- Use uma `SECRET_KEY` longa e aleatória (mínimo 32 bytes)
- Nunca versione o arquivo `.env` (já está no `.gitignore`)
- Considere usar LDAPS (`AD_USE_SSL=true`, porta 636)
- Coloque um **Nginx como proxy reverso** na frente se expor externamente (adiciona HTTPS)
- Faça backup regular da pasta `data/`

---

## ❓ Troubleshooting

### Container não sobe / fica reiniciando
```bash
docker compose logs
```

### Erro de autenticação AD
- Verifique conectividade: `ping 192.168.1.10` e `nc -zv 192.168.1.10 389`
- Confirme `AD_DOMAIN` e `AD_BASE_DN`
- O usuário deve pertencer a um dos grupos em `AD_TI_GROUP_CN`

### "Acesso negado" mesmo com usuário no grupo correto
- Verifique se o nome do grupo em `AD_TI_GROUP_CN` está correto
- Tente colocar o CN completo: `CN=NOME_DO_GRUPO`
- Qualquer formato funciona: `NOME`, `CN=NOME` ou DN completo

### Planilha não processa / erro de colunas
- Verifique se não há linhas em branco **no topo** (antes do cabeçalho)
- Colunas esperadas: `MATRÍCULA, NOME, CARGO, DEPARTAMENTO, GESTOR, ADMISSÃO`

### Porta 5000 já em uso
```env
APP_PORT=5001
```
```bash
docker compose up -d
```

---

## 👨‍💻 Desenvolvido por

**Setor de TI — AliseoSA**
Tulio Pereira · `tuliodutra27@gmail.com`

---

## 📄 Licença

Uso interno — AliseoSA. Todos os direitos reservados.
