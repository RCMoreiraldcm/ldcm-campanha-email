"""Ponto de entrada — campanha de e-mail gradual da Biblioteca LDCM.

Uso:
    python main_campanha.py              # dry_run (padrão — não envia nada)
    python main_campanha.py --enviar     # modo produção — envia e marca
"""

import logging
import os
import sys
from datetime import datetime

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


# ── Execução ────────────────────────────────────────────────────────

def executar(dry_run: bool = True) -> None:
    """Executa um ciclo da campanha gradual.

    Args:
        dry_run: True = gera HTML e loga, mas não envia nem marca.
                 False = envia e-mail e marca itens no SharePoint.
    """
    modo = "DRY RUN" if dry_run else "PRODUÇÃO"
    logger.info("=" * 60)
    logger.info("Campanha gradual -- modo %s", modo)
    logger.info("=" * 60)

    # 1. Autenticação
    token = auth.get_token()
    logger.info("Autenticação OK")

    # 2. Buscar livros da campanha
    pendentes, total_campanha = email_gradual.buscar_livros_campanha(token)
    logger.info(
        "Campanha: %d total desde %s | %d pendentes",
        total_campanha, email_gradual.DATA_CORTE, len(pendentes),
    )

    # 3. Verificar se há itens pendentes
    if not pendentes:
        logger.info("=" * 60)
        logger.info(
            "CAMPANHA CONCLUÍDA — Todos os títulos desde %s foram enviados.",
            email_gradual.DATA_CORTE,
        )
        logger.info(
            "AÇÃO NECESSÁRIA: reativar os fluxos no Power Automate:"
        )
        logger.info("  → Fluxo 'Aviso semanal de novos livros'")
        logger.info("  → Fluxo 'Aviso diário de lote grande'")
        logger.info("=" * 60)
        return

    # 4. Calcular lote
    tamanho_lote = email_gradual.calcular_lote(total_campanha)
    lote = pendentes[:tamanho_lote]
    restantes = len(pendentes) - len(lote)

    logger.info(
        "Lote calculado: %d títulos (5%% de %d = %d, teto 100)",
        len(lote), total_campanha,
        email_gradual.calcular_lote(total_campanha),
    )

    # 5. Listar títulos do lote
    for i, livro in enumerate(lote, 1):
        logger.info(
            "  [%d/%d] #%05d — %s",
            i, len(lote), livro["NumeroLista"], livro["Title"],
        )

    # 6. Montar HTML
    html = email_gradual.montar_html(lote, restantes, total_campanha)
    logger.info("HTML gerado: %d caracteres", len(html))

    # 7. Salvar preview
    preview_path = "logs/ultimo_email_dry.html"
    with open(preview_path, "w", encoding="utf-8") as f:
        f.write(html)
    logger.info("Preview salvo em %s", preview_path)

    # 8. Enviar ou parar
    if dry_run:
        logger.info("Modo DRY RUN — e-mail NÃO enviado, itens NÃO marcados.")
        logger.info(
            "Para enviar de verdade: python main_campanha.py --enviar"
        )
        return

    # --- Modo produção ---
    logger.info("Enviando e-mail para: %s", email_gradual.EMAIL_DESTINATARIOS)
    status = email_gradual.enviar_email(token, html, len(lote))
    logger.info("✅ E-mail enviado com sucesso (HTTP %s)", status)

    # 9. Marcar itens (só se envio OK)
    logger.info("Marcando %d itens como CampanhaGradual = Sim...", len(lote))
    marcados, falhas = email_gradual.marcar_lote(token, lote)
    logger.info("Marcacao: %d OK, %d falhas", marcados, falhas)

    # 10. Resumo
    logger.info("-" * 60)
    logger.info("Resumo: %d enviados | %d restantes | %d total campanha",
                len(lote), restantes, total_campanha)
    logger.info("Proxima execucao: rodar novamente amanha.")
    if restantes == 0:
        logger.info(
            "ATENCAO: Este foi o ultimo lote! Reativar os fluxos do Power Automate."
        )
    logger.info("-" * 60)


if __name__ == "__main__":
    enviar = "--enviar" in sys.argv
    executar(dry_run=not enviar)
