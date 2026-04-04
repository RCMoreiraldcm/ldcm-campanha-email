"""Processa e-mails de inscricao/desinscricao na caixa do remetente."""

import logging
import os
import sys

import requests
from dotenv import load_dotenv

import auth
import email_comum

load_dotenv()

logger = logging.getLogger("biblioteca_ldcm")
logger.setLevel(logging.INFO)
fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
sh = logging.StreamHandler(sys.stdout)
sh.setFormatter(fmt)
logger.addHandler(sh)

GRAPH_URL = "https://graph.microsoft.com/v1.0"
REMETENTE = os.getenv("EMAIL_REMETENTE", "rodrigo.moreira@ldcm.com.br")

SUBJECT_INSCREVER = "INSCREVER BIBLIOTECA DIARIO"
SUBJECT_DESINSCREVER = "DESINSCREVER BIBLIOTECA DIARIO"


def _headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}


def processar(dry_run: bool = True):
    token = auth.get_token()
    logger.info("Buscando e-mails de inscricao/desinscricao...")

    # Buscar mensagens com os assuntos relevantes
    url = (
        f"{GRAPH_URL}/users/{REMETENTE}/messages"
        f"?$filter=contains(subject,'{SUBJECT_INSCREVER}') or contains(subject,'{SUBJECT_DESINSCREVER}')"
        f"&$select=id,subject,from&$top=100"
    )

    resp = requests.get(url, headers=_headers(token), timeout=60)
    if not resp.ok:
        logger.error("Erro ao ler inbox: %s - %s", resp.status_code, resp.text[:300])
        return

    messages = resp.json().get("value", [])
    logger.info("Encontradas %d mensagens para processar", len(messages))

    if not messages:
        return

    subscribers = set(email_comum.ler_subscribers())
    inscritos = 0
    desinscritos = 0
    msg_ids_processados = []

    for msg in messages:
        subject = msg.get("subject", "")
        sender = msg.get("from", {}).get("emailAddress", {}).get("address", "")
        if not sender:
            continue

        sender = sender.lower().strip()
        msg_ids_processados.append(msg["id"])

        if SUBJECT_INSCREVER in subject.upper():
            if sender not in subscribers:
                subscribers.add(sender)
                inscritos += 1
                logger.info("  + INSCRITO: %s", sender)
            else:
                logger.info("  = JA INSCRITO: %s", sender)

        elif SUBJECT_DESINSCREVER in subject.upper():
            if sender in subscribers:
                subscribers.discard(sender)
                desinscritos += 1
                logger.info("  - DESINSCRITO: %s", sender)
            else:
                logger.info("  = NAO ESTAVA INSCRITO: %s", sender)

    logger.info("Resultado: +%d inscritos, -%d desinscritos, total: %d",
                inscritos, desinscritos, len(subscribers))

    if dry_run:
        logger.info("DRY RUN -- subscribers.txt NAO atualizado, e-mails NAO deletados.")
        return

    # Salvar lista atualizada
    email_comum.salvar_subscribers(list(subscribers))
    logger.info("subscribers.txt atualizado")

    # Deletar mensagens processadas
    for msg_id in msg_ids_processados:
        del_url = f"{GRAPH_URL}/users/{REMETENTE}/messages/{msg_id}"
        del_resp = requests.delete(del_url, headers=_headers(token), timeout=30)
        if del_resp.ok:
            logger.info("  Mensagem %s deletada", msg_id[:12])
        else:
            logger.warning("  Falha ao deletar %s: %s", msg_id[:12], del_resp.status_code)


if __name__ == "__main__":
    processar(dry_run="--executar" not in sys.argv)
