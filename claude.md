# Contexto do Projeto: Comparador de Planilhas do RH — AliseoSA

## Objetivo
Automatizar a detecção de mudanças nas planilhas semanais de colaboradores
que o RH envia ao setor de TI. A cada nova planilha, o sistema compara com o
snapshot anterior e exibe **adições (admissões), remoções (desligamentos)
e alterações** em uma interface web com controle de acesso via AD.

**Usuário:** assistente de TI na AliseoSA, automatiza tarefas via código.

---

## Estado atual
**SISTEMA WEB FUNCIONANDO EM DOCKER.**
- Script CLI original: `comparador_rh_v2.py` (mantido como referência — continua no modelo
  ANTIGO da planilha, com Matrícula/Cargo/Gestor; não foi adaptado ao modelo novo)
- O módulo de organograma foi removido do projeto (decisão de nov/2026); fica no histórico do git
- Aplicação web: pasta `web/` (Flask + Docker)

---

## Arquivos do projeto

### Script CLI original (referência)
- `comparador_rh_v2.py` — script original que funcionou e validou a ideia

### Aplicação Web (`web/`)
```
web/
├── Dockerfile
├── docker-compose.yml
├── requirements.txt
├── .env.example          ← copiar para .env e configurar
├── README.md
└── app/
    ├── main.py           ← Flask: rotas e orquestração
    ├── auth.py           ← autenticação LDAP/AD
    ├── comparador.py     ← lógica de comparação (portada do CLI)
    ├── templates/
    │   ├── base.html
    │   ├── login.html
    │   ├── dashboard.html
    │   └── resultado.html
    └── static/
        └── style.css
```

### Dados persistidos (volume Docker → `web/data/`)
```
data/
├── historico.json        ← índice de todas as comparações
├── snapshots/            ← cópias das planilhas (base de comparação)
├── relatorios/           ← resultados em JSON por ID
└── uploads/              ← arquivos enviados pelos usuários
```

---

## Decisões e regras de negócio

1. **Chave de identificação**:
   - CLT → coluna CÓDIGO (antes MATRÍCULA) → chave `CLT::<código>`
   - PJ → CÓDIGO vazio ou "PJ" (qualquer caixa) → identificado pelo NOME → chave `PJ::<nome>`

2. **Campos monitorados para "alteração"**: FUNÇÃO (antes CARGO), DEPARTAMENTO
   (o campo GESTOR deixou de existir na planilha do RH e foi removido do sistema)

3. **Snapshots**: cada planilha enviada é salva como snapshot para a próxima comparação.
   Primeira vez = baseline (sem diferenças). Segunda vez em diante = mostra diff.

4. **Autenticação**: LDAP/AD. Só usuários do grupo configurado em `AD_TI_GROUP_CN` (padrão: `TI`) têm acesso.

5. **DEV_MODE**: variável de ambiente para testar sem AD. **NUNCA usar em produção.**

---

## Colunas da planilha — modelo atual do RH (nov/2026)
Arquivo `.xls`/`.xlsx` com duas abas: **`Ativos `** (efetivo — a única lida; os nomes das
abas vêm com espaço no final) e `Demitidos ` (ignorada). Na aba Ativos a linha 1 é um
título e o cabeçalho fica na linha 2 (localizado automaticamente, em qualquer linha).
Colunas: `Código, Nome, Função, Departamento, Data de Admissão`.
Nomes antigos (Matrícula, Cargo, Admissão) ainda são aceitos, para comparar com snapshots
antigos. Chaves do JSON dos relatórios continuam `matricula`/`cargo`/`admissao` de propósito
(relatórios gravados antes da mudança precisam continuar abrindo).

## Colunas do modelo ANTIGO (referência; após normalização para MAIÚSCULAS)
`MATRÍCULA, NOME, CARGO, DEPARTAMENTO, GESTOR, ADMISSÃO, MOI/MOD, ADM/TURNO, LETRA`

---

## Como rodar (produção)

```bash
cd web/
cp .env.example .env
# editar .env com dados reais do AD
docker compose up -d
# acesse http://IP:5000
```

## Como testar sem AD (desenvolvimento)

```bash
# no .env:
DEV_MODE=true
DEV_USERS={"admin":"admin123"}
docker compose up -d
```

---

## Bugs do script CLI (NÃO reintroduzir no web)

1. **Matrícula "610.0"**: pandas lê número como float. Script remove o `.0`.
2. **Linha em branco no rodapé** → `PJ::NAN`: descarta linhas sem matrícula E sem nome.
3. **Linha em branco no TOPO**: deslocava o cabeçalho.
   → Resolvido: o cabeçalho agora é detectado automaticamente (web). O CLI antigo ainda sofre disso.

## Problemas de dados conhecidos (na fonte, não no código)

1. **Matrícula 1241 duplicada** com duas pessoas diferentes — erro do RH.
   Script mantém última ocorrência e avisa no terminal.
2. **Alex (1129)** duplicata simples — descartável.

---

## Stack

- Python 3.12 + Flask 3.1 + gunicorn
- pandas + openpyxl + xlrd (processamento de planilhas `.xlsx` e `.xls`)
- ldap3 (autenticação AD)
- Bootstrap 5 (UI)
- Docker + Docker Compose

---

## Pendências / melhorias futuras

- [x] Auto-detectar header mesmo com linhas em branco no topo (feito na versão web)
- [ ] Tratar matrículas duplicadas com nomes diferentes de forma mais visível
- [ ] Notificação por e-mail ao processar uma planilha
- [ ] Filtro por data no histórico do dashboard
- [ ] Exportação PDF do relatório

---

## Como fazer o push para o GitHub (pendente — fazer da outra máquina)

Repositório destino: **https://github.com/tuliodutra27/App-RH**

Passos a executar na máquina nova:

```bash
# 1. Copie os arquivos do projeto para a nova máquina (ou transfira via pendrive/rede)
#    Pasta a copiar: "Analisador de Cargos/web/"

# 2. Inicializa git e faz o primeiro push
cd "Analisador de Cargos/web"

git init
git config user.name "Tulio Dutra"
git config user.email "tuliodutra27@gmail.com"

# Cria .gitignore antes do commit
cat > .gitignore << 'EOF'
.env
data/
__pycache__/
*.pyc
*.pyo
.DS_Store
EOF

git add .
git commit -m "feat: sistema web RH com Docker e autenticação AD

- Interface web Flask para comparação de planilhas do RH
- Autenticação via Active Directory (grupo TI)
- Upload de planilha com comparação automática
- Exibe admissões, desligamentos e alterações
- Download do relatório em Excel formatado
- Containerizado com Docker + gunicorn

Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"

git branch -M main
git remote add origin https://github.com/tuliodutra27/App-RH.git
git push -u origin main
```

> Se o repositório já tiver conteúdo: `git push -u origin main --force`
