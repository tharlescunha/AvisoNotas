import os
import re
import time
import sqlite3
import smtplib
from pathlib import Path
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr

import pandas as pd


DB_PATH = "envios.db"
EXCEL_PATH = "Lista de email.xlsx"
PDF_PATH = "Guia de Utilização da Plataforma.pdf"

SMTP_HOST = "smtp.office365.com"
SMTP_PORT = 587
SMTP_USER = "no-reply@alzgraos.com.br"
SMTP_PASS = "bwntblsntfxkdscl"

FROM_NAME = "ALZ Grãos"
ASSUNTO = "Guia de Utilização da Plataforma de Agendamento de Descarga do CIF – Tegram"

# Ajuste conforme sua necessidade
PAUSA_ENTRE_ENVIOS_SEGUNDOS = 2
MAX_TENTATIVAS = 2


def agora_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def validar_email(email: str) -> bool:
    """
    Validação simples de e-mail.
    """
    if not email:
        return False

    email = email.strip()
    padrao = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    return bool(re.match(padrao, email))


def conectar_db(db_path: str = DB_PATH) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def inicializar_banco(db_path: str = DB_PATH) -> None:
    """
    Cria a base SQLite para controle dos envios.
    """
    with conectar_db(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS envios_email (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT NOT NULL UNIQUE,
                status TEXT NOT NULL DEFAULT 'PENDENTE',
                tentativas INTEGER NOT NULL DEFAULT 0,
                enviado_em TEXT,
                ultimo_erro TEXT,
                criado_em TEXT NOT NULL,
                atualizado_em TEXT NOT NULL
            )
            """
        )
        conn.commit()


def importar_emails_do_excel(excel_path: str = EXCEL_PATH, db_path: str = DB_PATH) -> None:
    """
    Lê a planilha Excel e importa os e-mails da coluna 'Emails'
    para o SQLite, sem duplicar registros.
    """
    caminho = Path(excel_path)
    if not caminho.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {excel_path}")

    df = pd.read_excel(caminho)

    if "Emails" not in df.columns:
        raise ValueError("A planilha precisa ter a coluna 'Emails'.")

    df["Emails"] = df["Emails"].astype(str).str.strip()
    df = df[df["Emails"].notna()]
    df = df[df["Emails"] != ""]
    df = df.drop_duplicates(subset=["Emails"])

    inseridos = 0
    ignorados = 0
    invalidos = 0

    with conectar_db(db_path) as conn:
        for email in df["Emails"].tolist():
            email = email.strip()

            if not validar_email(email):
                invalidos += 1
                continue

            try:
                conn.execute(
                    """
                    INSERT INTO envios_email (
                        email, status, tentativas, enviado_em, ultimo_erro, criado_em, atualizado_em
                    )
                    VALUES (?, 'PENDENTE', 0, NULL, NULL, ?, ?)
                    """,
                    (email, agora_str(), agora_str()),
                )
                inseridos += 1
            except sqlite3.IntegrityError:
                ignorados += 1

        conn.commit()

    print(f"Importação concluída | inseridos={inseridos} | ignorados={ignorados} | invalidos={invalidos}")


def buscar_pendentes(db_path: str = DB_PATH, limite: int = 50) -> list[sqlite3.Row]:
    """
    Busca registros pendentes ou com erro e com tentativas abaixo do limite.
    """
    with conectar_db(db_path) as conn:
        rows = conn.execute(
            """
            SELECT *
            FROM envios_email
            WHERE (status = 'PENDENTE' OR status = 'ERRO')
              AND tentativas < ?
            ORDER BY id
            LIMIT ?
            """,
            (MAX_TENTATIVAS, limite),
        ).fetchall()

    return rows


def atualizar_status(
    email_id: int,
    status: str,
    tentativas: int | None = None,
    enviado_em: str | None = None,
    ultimo_erro: str | None = None,
    db_path: str = DB_PATH,
) -> None:
    """
    Atualiza o status do registro no banco.
    """
    with conectar_db(db_path) as conn:
        if tentativas is None:
            conn.execute(
                """
                UPDATE envios_email
                SET status = ?,
                    enviado_em = ?,
                    ultimo_erro = ?,
                    atualizado_em = ?
                WHERE id = ?
                """,
                (status, enviado_em, ultimo_erro, agora_str(), email_id),
            )
        else:
            conn.execute(
                """
                UPDATE envios_email
                SET status = ?,
                    tentativas = ?,
                    enviado_em = ?,
                    ultimo_erro = ?,
                    atualizado_em = ?
                WHERE id = ?
                """,
                (status, tentativas, enviado_em, ultimo_erro, agora_str(), email_id),
            )

        conn.commit()


def montar_corpo_html() -> str:
    """
    Monta um corpo HTML profissional, discreto e organizado.
    """
    return """
    <html>
      <body style="margin:0; padding:0; background-color:#f5f6f8; font-family:Arial, Helvetica, sans-serif; color:#222;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f5f6f8; padding:24px 0;">
          <tr>
            <td align="center">
              <table width="680" cellpadding="0" cellspacing="0" border="0"
                     style="background:#ffffff; border:1px solid #e5e7eb; border-radius:10px; overflow:hidden;">
                
                <tr>
                  <td style="background:#1f3b5b; padding:20px 28px;">
                    <div style="font-size:22px; font-weight:bold; color:#ffffff;">ALZ Grãos</div>
                    <div style="font-size:13px; color:#dbe5f0; margin-top:6px;">
                      Plataforma de Agendamento de Descarga do CIF – Tegram
                    </div>
                  </td>
                </tr>

                <tr>
                  <td style="padding:30px 28px 10px 28px;">
                    <div style="font-size:16px; line-height:1.7; color:#222;">
                      <p style="margin:0 0 18px 0;">Prezados,</p>

                      <p style="margin:0 0 18px 0;">
                        Encaminhamos em anexo o <strong>manual de utilização da plataforma de agendamento de descarga do CIF – Tegram</strong>,
                        com o objetivo de auxiliá-los na navegação e no uso das funcionalidades disponíveis.
                      </p>

                      <p style="margin:0 0 18px 0;">
                        Em caso de dúvidas, nossa equipe está à disposição para suporte por meio do canal abaixo:
                      </p>

                      <div style="margin:0 0 18px 0; padding:14px 16px; background:#f8fafc; border-left:4px solid #1f3b5b; border-radius:6px;">
                        <div style="font-size:14px; color:#333;">
                          <strong>E-mail de suporte:</strong> performance@alzgraos.com.br
                        </div>
                      </div>

                      <p style="margin:0 0 18px 0;">
                        Se preferirem, também podem entrar em contato diretamente com o comercial responsável pelo atendimento da sua empresa.
                      </p>

                      <p style="margin:0 0 18px 0;">
                        Agradecemos pela utilização da plataforma e pela parceria.
                      </p>
                    </div>
                  </td>
                </tr>

                <tr>
                  <td style="padding:8px 28px 28px 28px;">
                    <div style="font-size:12px; color:#6b7280; line-height:1.6; border-top:1px solid #e5e7eb; padding-top:16px;">
                      Este é um e-mail automático. Por favor, não responda a esta mensagem.
                      <br><br>
                      Atenciosamente,<br>
                      <strong>Equipe ALZ Grãos</strong>
                    </div>
                  </td>
                </tr>

              </table>
            </td>
          </tr>
        </table>
      </body>
    </html>
    """


def montar_corpo_texto() -> str:
    """
    Corpo texto simples para clientes que não renderizam HTML.
    """
    return """Prezados,

Encaminhamos em anexo o manual de utilização da plataforma de agendamento de descarga do CIF – Tegram, com o objetivo de auxiliá-los na navegação e no uso das funcionalidades disponíveis.

Em caso de dúvidas, nossa equipe está à disposição para suporte por meio do canal abaixo:
E-mail de suporte: performance@alzgraos.com.br

Se preferirem, também podem entrar em contato diretamente com o comercial responsável pelo atendimento da sua empresa.

Agradecemos pela utilização da plataforma e pela parceria.

Este é um e-mail automático. Por favor, não responda a esta mensagem.

Atenciosamente,
Equipe ALZ Grãos
"""


def criar_mensagem_email(destinatario: str, pdf_path: str = PDF_PATH) -> EmailMessage:
    """
    Monta a mensagem de e-mail com HTML + texto plano + anexo PDF.
    """
    caminho_pdf = Path(pdf_path)
    if not caminho_pdf.exists():
        raise FileNotFoundError(f"Arquivo PDF não encontrado: {pdf_path}")

    msg = EmailMessage()
    msg["From"] = formataddr((FROM_NAME, SMTP_USER))
    msg["To"] = destinatario
    msg["Subject"] = ASSUNTO

    # Corpo texto plano
    msg.set_content(montar_corpo_texto())

    # Corpo HTML
    msg.add_alternative(montar_corpo_html(), subtype="html")

    # Anexo PDF
    with open(caminho_pdf, "rb") as f:
        arquivo_bytes = f.read()

    msg.add_attachment(
        arquivo_bytes,
        maintype="application",
        subtype="pdf",
        filename=caminho_pdf.name,
    )

    return msg


def enviar_email(destinatario: str, pdf_path: str = PDF_PATH) -> None:
    """
    Envia um e-mail individual.
    """
    if not SMTP_PASS:
        raise RuntimeError("SMTP_PASS não definido nas variáveis de ambiente.")

    msg = criar_mensagem_email(destinatario, pdf_path)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)


def processar_envios(
    db_path: str = DB_PATH,
    pdf_path: str = PDF_PATH,
    pausa_segundos: int = PAUSA_ENTRE_ENVIOS_SEGUNDOS,
    limite_por_execucao: int = 50,
) -> None:
    """
    Processa a fila de envios:
    - busca pendentes/erro
    - envia um por um
    - registra sucesso/erro
    - aplica pausa entre envios
    """
    pendentes = buscar_pendentes(db_path=db_path, limite=limite_por_execucao)

    if not pendentes:
        print("Nenhum e-mail pendente para envio.")
        return

    print(f"Iniciando processamento | total={len(pendentes)}")

    for item in pendentes:
        email_id = item["id"]
        email = item["email"]
        tentativas_atuais = item["tentativas"] or 0
        nova_tentativa = tentativas_atuais + 1

        print(f"[{email_id}] Enviando para: {email} | tentativa={nova_tentativa}")

        try:
            atualizar_status(
                email_id=email_id,
                status="PROCESSANDO",
                tentativas=nova_tentativa,
                enviado_em=None,
                ultimo_erro=None,
                db_path=db_path,
            )

            enviar_email(destinatario=email, pdf_path=pdf_path)

            atualizar_status(
                email_id=email_id,
                status="ENVIADO",
                tentativas=nova_tentativa,
                enviado_em=agora_str(),
                ultimo_erro=None,
                db_path=db_path,
            )

            print(f"[{email_id}] SUCESSO -> {email}")

        except Exception as e:
            erro = str(e)[:4000]

            atualizar_status(
                email_id=email_id,
                status="ERRO",
                tentativas=nova_tentativa,
                enviado_em=None,
                ultimo_erro=erro,
                db_path=db_path,
            )

            print(f"[{email_id}] ERRO -> {email} | detalhe={erro}")

        print(f"Aguardando {pausa_segundos}s antes do próximo envio...")
        time.sleep(pausa_segundos)

    print("Processamento concluído.")


def resumo_envios(db_path: str = DB_PATH) -> None:
    """
    Exibe um resumo geral dos envios.
    """
    with conectar_db(db_path) as conn:
        total = conn.execute("SELECT COUNT(*) FROM envios_email").fetchone()[0]
        pendente = conn.execute("SELECT COUNT(*) FROM envios_email WHERE status = 'PENDENTE'").fetchone()[0]
        processando = conn.execute("SELECT COUNT(*) FROM envios_email WHERE status = 'PROCESSANDO'").fetchone()[0]
        enviado = conn.execute("SELECT COUNT(*) FROM envios_email WHERE status = 'ENVIADO'").fetchone()[0]
        erro = conn.execute("SELECT COUNT(*) FROM envios_email WHERE status = 'ERRO'").fetchone()[0]

    print("===== RESUMO =====")
    print(f"TOTAL       : {total}")
    print(f"PENDENTE    : {pendente}")
    print(f"PROCESSANDO : {processando}")
    print(f"ENVIADO     : {enviado}")
    print(f"ERRO        : {erro}")


def main():
    inicializar_banco()
    importar_emails_do_excel()
    resumo_envios()

    # Processa em lotes. Ajuste o limite se quiser.
    processar_envios(limite_por_execucao=100)

    resumo_envios()


if __name__ == "__main__":
    main()
