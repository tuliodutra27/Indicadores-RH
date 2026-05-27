"""
Autenticação via Active Directory (LDAP).
Permite login de usuários pertencentes ao grupo TI.
"""

import os
import json
from ldap3 import Server, Connection, ALL, SUBTREE
from ldap3.core.exceptions import (
    LDAPBindError,
    LDAPSocketOpenError,
    LDAPException,
)

# ── Configuração via variáveis de ambiente ─────────────────────────────────
AD_SERVER      = os.environ.get("AD_SERVER",       "ldap://seu-servidor-ad")
AD_PORT        = int(os.environ.get("AD_PORT",     "389"))
AD_USE_SSL     = os.environ.get("AD_USE_SSL",      "false").lower() == "true"
AD_DOMAIN      = os.environ.get("AD_DOMAIN",       "aliseo.local")
AD_BASE_DN     = os.environ.get("AD_BASE_DN",      "DC=aliseo,DC=local")
AD_TI_GROUP_CN = os.environ.get("AD_TI_GROUP_CN",  "TI")

# DEV_MODE: NUNCA habilitar em produção. Permite login sem AD para testes.
# Defina DEV_USERS como JSON: '{"admin":"senha123"}'
DEV_MODE       = os.environ.get("DEV_MODE", "false").lower() == "true"
DEV_USERS      = json.loads(os.environ.get("DEV_USERS", "{}"))


def autenticar_ad(usuario: str, senha: str) -> tuple[bool, str]:
    """
    Autentica o usuário no AD e verifica se pertence ao grupo TI.

    Retorna:
        (True, display_name)         → autenticado com sucesso
        (False, mensagem_de_erro)    → falha na autenticação
    """

    # ── Modo desenvolvimento (sem AD) ──────────────────────────────────────
    if DEV_MODE:
        if DEV_USERS:
            if usuario in DEV_USERS and DEV_USERS[usuario] == senha:
                return True, f"{usuario} (DEV)"
            return False, "Usuário ou senha incorretos."
        # Sem usuários definidos em DEV_USERS → aceita qualquer credencial
        return True, f"{usuario} (DEV)"

    # ── Autenticação real via LDAP/AD ──────────────────────────────────────
    try:
        server = Server(
            AD_SERVER,
            port=AD_PORT,
            use_ssl=AD_USE_SSL,
            get_info=ALL,
            connect_timeout=8,
        )

        # Bind com credenciais do usuário (UPN: usuario@dominio)
        user_upn = f"{usuario}@{AD_DOMAIN}"
        conn = Connection(
            server,
            user=user_upn,
            password=senha,
            auto_bind=True,
            receive_timeout=10,
        )

        # Busca dados do usuário: nome e grupos
        conn.search(
            AD_BASE_DN,
            f"(&(objectClass=user)(sAMAccountName={usuario}))",
            search_scope=SUBTREE,
            attributes=["displayName", "memberOf", "sAMAccountName"],
        )

        if not conn.entries:
            conn.unbind()
            return False, "Usuário não encontrado no Active Directory."

        entry = conn.entries[0]
        display_name = (
            str(entry.displayName)
            if entry.displayName
            else usuario
        )

        # Verifica pertencimento a pelo menos um dos grupos permitidos
        # AD_TI_GROUP_CN aceita múltiplos grupos separados por ";"
        # Ex: "GRP_SA_OPE_TI_RW;GRP_SA_OPE_DIRETORIA_RW"
        member_of = []
        if entry.memberOf:
            member_of = [str(g).upper() for g in entry.memberOf.values]

        def _normalizar_grupo(raw: str) -> str:
            """Extrai só o CN independente do formato fornecido."""
            raw = raw.strip()
            # DN completo (tem vírgula) → pega só o primeiro componente
            if "," in raw:
                raw = raw.split(",")[0].strip()
            # Remove prefixo CN= se presente
            if raw.upper().startswith("CN="):
                raw = raw[3:]
            return raw.upper()

        grupos_permitidos = [
            _normalizar_grupo(g)
            for g in AD_TI_GROUP_CN.split(";")
            if g.strip()
        ]

        in_group = any(
            f"CN={grupo}," in membro
            for grupo in grupos_permitidos
            for membro in member_of
        )

        conn.unbind()

        if not in_group:
            nomes = ", ".join(grupos_permitidos)
            return (
                False,
                f"Acesso negado. Apenas usuários dos grupos [{nomes}] "
                "têm permissão para acessar este sistema.",
            )

        return True, display_name

    except LDAPBindError:
        return False, "Usuário ou senha incorretos."
    except LDAPSocketOpenError:
        return (
            False,
            f"Não foi possível conectar ao servidor AD ({AD_SERVER}). "
            "Verifique a configuração de rede.",
        )
    except LDAPException as exc:
        return False, f"Erro de autenticação LDAP: {exc}"
    except Exception as exc:  # noqa: BLE001
        return False, f"Erro inesperado ao autenticar: {exc}"
