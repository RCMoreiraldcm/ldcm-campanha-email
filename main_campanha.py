"""Ponto de entrada — campanha de e-mail gradual da Biblioteca LDCM.

Uso:
    python main_campanha.py              # dry_run (padrao)
    python main_campanha.py --enviar     # modo producao
"""

import logging
import os
import sys

from dotenv import load_dotenv

import auth
import email_gradual

load_dotenv()

# ── Logging ─────────────────────────────────────────────────────────

logger = logging.getLogger("biblioteca_ldcm")
logger.setLevel(logging.INFO)

fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

os.makedirs("logs", exist_ok=True)

fh = logging.FileHandler("logs/campanha.log", encoding="utf-8")
fh.setFormatter(fmt)
logger.addHandler(fh)

sh = logging.StreamHandler(sys.stdout)
sh.setFormatter(fmt)
logger.addHandler(sh)


# ── Execucao ────────────────────────────────────────────────────────

def executar(dry_run: bool = True) -> None:
    modo = "DRY RUN" if dry_run else "PRODUCAO"
    logger.info("=" * 60)
    logger.info("Campanha gradual -- modo %s", modo)
    logger.info("=" * 60)

    token = auth.get_token()
    logger.info("Autenticacao OK")

    # Buscar todos os livros da campanha
    livros = email_gradual.buscar_livros_campanha(token)
    total = len(livros)
    logger.info("Total de livros desde %s: %d", email_gradual.DATA_CORTE, total)

    # Ler offset atual
    offset = email_gradual.ler_offset()
    logger.info("Offset atual: %d", offset)

    if offset >= total:
        logger.info("=" * 60)
        logger.info("CAMPANHA CONCLUIDA -- todos os %d titulos ja foram enviados.", total)
        logger.info("Desative o workflow quando quiser encerrar.")
        logger.info("=" * 60)
        return

    # Selecionar lote: offset -> offset+100
    lote = livros[offset:offset + email_gradual.LOTE_TAMANHO]
    novo_offset = offset + len(lote)

    logger.info("Lote: %d titulos (posicoes %d a %d de %d)", len(lote), offset + 1, novo_offset, total)

    for i, livro in enumerate(lote, 1):
        logger.info("  [%d/%d] #%05d -- %s", i, len(lote), livro["NumeroLista"], livro["Title"])

    # Montar HTML
    html = email_gradual.montar_html(lote)
    logger.info("HTML gerado: %d caracteres", len(html))

    # Salvar preview
    preview_path = "logs/ultimo_email_dry.html"
    with open(preview_path, "w", encoding="utf-8") as f:
        f.write(html)
    logger.info("Preview salvo em %s", preview_path)

    if dry_run:
        logger.info("Modo DRY RUN -- e-mail NAO enviado.")
        return

    # Enviar
    logger.info("Enviando e-mail para: %s", email_gradual.EMAIL_DESTINATARIOS)
    status = email_gradual.enviar_email(token, html, len(lote))
    logger.info("E-mail enviado com sucesso (HTTP %s)", status)

    # Salvar novo offset
    email_gradual.salvar_offset(novo_offset)
    logger.info("Offset atualizado: %d -> %d", offset, novo_offset)

    logger.info("-" * 60)
    logger.info("Resumo: %d enviados | offset %d de %d", len(lote), novo_offset, total)
    logger.info("-" * 60)


if __name__ == "__main__":
    enviar = "--enviar" in sys.argv
    executar(dry_run=not enviar)
