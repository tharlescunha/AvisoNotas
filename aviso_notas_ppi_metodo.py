import re
import smtplib
import time
from pathlib import Path
from email.message import EmailMessage
from email.utils import formataddr

import pandas as pd

EXCEL_PATH = "Lista de email_2.xlsx"  # planilha com a coluna: Emails
COLUNA_EMAIL = "Emails"

SMTP_HOST = "smtp.office365.com"
SMTP_PORT = 587
FROM_NAME = "ALZ Grãos"
ASSUNTO = "Aviso de notas fiscais sob responsabilidade da PPI"
PAUSA_ENTRE_ENVIOS_SEGUNDOS = 2

# =========================
# DADOS DO AVISO
# =========================
NOTAS_PENDENTES_LANCAMENTO = [
    {"nf": "00045871", "fornecedor": "Fornecedor Alpha", "vencimento": "08/04/2026", "valor": "R$ 12.450,00"},
    {"nf": "00045889", "fornecedor": "Fornecedor Beta", "vencimento": "09/04/2026", "valor": "R$ 7.980,00"},
    {"nf": "00045903", "fornecedor": "Fornecedor Gama", "vencimento": "10/04/2026", "valor": "R$ 4.320,00"},
]

NOTAS_PENDENTES_PAGAMENTO = [
    {"nf": "00045720", "fornecedor": "Fornecedor Delta", "vencimento": "06/04/2026", "valor": "R$ 15.700,00"},
    {"nf": "00045744", "fornecedor": "Fornecedor Épsilon", "vencimento": "07/04/2026", "valor": "R$ 9.150,00"},
]

NOTAS_CONCLUIDAS_MES_04 = [
    {"nf": "00045611", "fornecedor": "Fornecedor Zeta", "data_conclusao": "02/04/2026", "valor": "R$ 6.890,00"},
    {"nf": "00045635", "fornecedor": "Fornecedor Eta", "data_conclusao": "04/04/2026", "valor": "R$ 11.240,00"},
]


def validar_email(email: str) -> bool:
    if not email:
        return False
    email = str(email).strip()
    padrao = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    return bool(re.match(padrao, email))


def carregar_destinatarios(excel_path: str = EXCEL_PATH, coluna_email: str = COLUNA_EMAIL) -> list[str]:
    caminho = Path(excel_path)
    if not caminho.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {excel_path}")

    df = pd.read_excel(caminho)

    if coluna_email not in df.columns:
        raise ValueError(f"A planilha precisa ter a coluna '{coluna_email}'.")

    emails = df[coluna_email].dropna().astype(str).str.strip().unique().tolist()
    emails_validos = [email for email in emails if validar_email(email)]
    return emails_validos


def montar_linhas_html_notas(notas: list[dict], tipo: str) -> str:
    linhas = []

    for item in notas:
        if tipo in {"lancamento", "pagamento"}:
            data_ref = item["vencimento"]
            titulo_data = data_ref
        else:
            data_ref = item["data_conclusao"]
            titulo_data = data_ref

        linhas.append(
            f"""
            <tr>
                <td style="padding:10px; border:1px solid #dcdfe4;">{item['nf']}</td>
                <td style="padding:10px; border:1px solid #dcdfe4;">{item['fornecedor']}</td>
                <td style="padding:10px; border:1px solid #dcdfe4;">{titulo_data}</td>
                <td style="padding:10px; border:1px solid #dcdfe4; text-align:right;">{item['valor']}</td>
            </tr>
            """
        )

    return "".join(linhas)


def montar_corpo_html() -> str:
    total_lancamento = len(NOTAS_PENDENTES_LANCAMENTO)
    total_pagamento = len(NOTAS_PENDENTES_PAGAMENTO)
    total_concluidas = len(NOTAS_CONCLUIDAS_MES_04)

    linhas_lancamento = montar_linhas_html_notas(NOTAS_PENDENTES_LANCAMENTO, "lancamento")
    linhas_pagamento = montar_linhas_html_notas(NOTAS_PENDENTES_PAGAMENTO, "pagamento")
    linhas_concluidas = montar_linhas_html_notas(NOTAS_CONCLUIDAS_MES_04, "concluida")

    return f"""
    <html>
      <body style="margin:0; padding:0; background-color:#f4f6f8; font-family:Arial, Helvetica, sans-serif; color:#1f2937;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="padding:24px 0; background-color:#f4f6f8;">
          <tr>
            <td align="center">
              <table width="760" cellpadding="0" cellspacing="0" border="0" style="background:#ffffff; border:1px solid #e5e7eb; border-radius:12px; overflow:hidden;">
                <tr>
                  <td style="background:#16324f; padding:22px 28px;">
                    <div style="font-size:22px; font-weight:bold; color:#ffffff;">ALZ Grãos</div>
                    <div style="font-size:13px; color:#d8e3ef; margin-top:6px;">Aviso de notas fiscais sob responsabilidade da PPI</div>
                  </td>
                </tr>

                <tr>
                  <td style="padding:28px;">
                    <p style="margin:0 0 16px 0; font-size:15px;">Prezados,</p>

                    <p style="margin:0 0 16px 0; font-size:14px; line-height:1.7;">
                      Segue abaixo um <strong>aviso de acompanhamento</strong> referente às notas fiscais sob responsabilidade da <strong>PPI</strong>.
                    </p>

                    <div style="margin:0 0 20px 0; padding:14px 16px; background:#f8fafc; border-left:4px solid #16324f; border-radius:6px; font-size:14px; line-height:1.6;">
                      <div><strong>Notas pendentes de lançamento:</strong> {total_lancamento}</div>
                      <div><strong>Notas pendentes de pagamento:</strong> {total_pagamento}</div>
                      <div><strong>Notas concluídas no mês 04:</strong> {total_concluidas}</div>
                    </div>

                    <h3 style="font-size:16px; color:#16324f; margin:0 0 10px 0; background:#fdeaea; padding:10px 12px; border-radius:8px;">1. Notas pendentes de lançamento</h3>
                    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; margin:0 0 24px 0; font-size:13px;">
                      <tr style="background:#fdeaea;">
                        <th align="left" style="padding:10px; border:1px solid #dcdfe4;">NF</th>
                        <th align="left" style="padding:10px; border:1px solid #dcdfe4;">Fornecedor</th>
                        <th align="left" style="padding:10px; border:1px solid #dcdfe4;">Vencimento</th>
                        <th align="right" style="padding:10px; border:1px solid #dcdfe4;">Valor</th>
                      </tr>
                      {linhas_lancamento}
                    </table>

                    <h3 style="font-size:16px; color:#16324f; margin:0 0 10px 0; background:#fff8db; padding:10px 12px; border-radius:8px;">2. Notas pendentes de pagamento</h3>
                    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; margin:0 0 24px 0; font-size:13px;">
                      <tr style="background:#fff8db;">
                        <th align="left" style="padding:10px; border:1px solid #dcdfe4;">NF</th>
                        <th align="left" style="padding:10px; border:1px solid #dcdfe4;">Fornecedor</th>
                        <th align="left" style="padding:10px; border:1px solid #dcdfe4;">Vencimento</th>
                        <th align="right" style="padding:10px; border:1px solid #dcdfe4;">Valor</th>
                      </tr>
                      {linhas_pagamento}
                    </table>

                    <h3 style="font-size:16px; color:#16324f; margin:0 0 10px 0; background:#eaf7ea; padding:10px 12px; border-radius:8px;">3. Notas concluídas no mês 04</h3>
                    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; margin:0 0 24px 0; font-size:13px;">
                      <tr style="background:#eaf7ea;">
                        <th align="left" style="padding:10px; border:1px solid #dcdfe4;">NF</th>
                        <th align="left" style="padding:10px; border:1px solid #dcdfe4;">Fornecedor</th>
                        <th align="left" style="padding:10px; border:1px solid #dcdfe4;">Data de conclusão</th>
                        <th align="right" style="padding:10px; border:1px solid #dcdfe4;">Valor</th>
                      </tr>
                      {linhas_concluidas}
                    </table>

                    <p style="margin:0; font-size:14px; line-height:1.7;">
                      Favor verificar os itens pendentes e seguir com as devidas tratativas.
                    </p>
                  </td>
                </tr>

                <tr>
                  <td style="padding:18px 28px 28px 28px;">
                    <div style="font-size:12px; color:#6b7280; line-height:1.6; border-top:1px solid #e5e7eb; padding-top:16px;">
                      Este é um e-mail automático de aviso. Por favor, não responder.
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
    linhas_lancamento = "\n".join(
        [f"- NF {item['nf']} | {item['fornecedor']} | Venc.: {item['vencimento']} | {item['valor']}" for item in NOTAS_PENDENTES_LANCAMENTO]
    )
    linhas_pagamento = "\n".join(
        [f"- NF {item['nf']} | {item['fornecedor']} | Venc.: {item['vencimento']} | {item['valor']}" for item in NOTAS_PENDENTES_PAGAMENTO]
    )
    linhas_concluidas = "\n".join(
        [f"- NF {item['nf']} | {item['fornecedor']} | Concluída em: {item['data_conclusao']} | {item['valor']}" for item in NOTAS_CONCLUIDAS_MES_04]
    )

    return f"""Prezados,

Segue abaixo um aviso de acompanhamento referente às notas fiscais sob responsabilidade da PPI.

Notas pendentes de lançamento: {len(NOTAS_PENDENTES_LANCAMENTO)}
{linhas_lancamento}

Notas pendentes de pagamento: {len(NOTAS_PENDENTES_PAGAMENTO)}
{linhas_pagamento}

Notas concluídas no mês 04: {len(NOTAS_CONCLUIDAS_MES_04)}
{linhas_concluidas}

Favor verificar os itens pendentes e seguir com as devidas tratativas.

Este é um e-mail automático de aviso. Por favor, não responder.

Atenciosamente,
Equipe ALZ Grãos
"""


def criar_mensagem_email(destinatario: str, remetente_email: str) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = formataddr((FROM_NAME, remetente_email))
    msg["To"] = destinatario
    msg["Subject"] = ASSUNTO
    msg.set_content(montar_corpo_texto())
    msg.add_alternative(montar_corpo_html(), subtype="html")
    return msg


def enviar_email(destinatario: str, remetente_email: str, senha_email: str) -> bool:
    try:
        if not validar_email(destinatario):
            raise ValueError(f"Destinatário inválido: {destinatario}")

        if not validar_email(remetente_email):
            raise ValueError(f"E-mail do remetente inválido: {remetente_email}")

        if not senha_email or not str(senha_email).strip():
            raise ValueError("A senha do e-mail não foi informada.")

        msg = criar_mensagem_email(destinatario, remetente_email)

        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(remetente_email, senha_email)
            server.send_message(msg)

        return True
    except Exception as exc:
        raise RuntimeError(f"Erro ao enviar e-mail para {destinatario}: {exc}") from exc


def executar_envio(remetente_email: str, senha_email: str, excel_path: str = EXCEL_PATH) -> dict:
    try:
        destinatarios = carregar_destinatarios(excel_path=excel_path, coluna_email=COLUNA_EMAIL)

        if not destinatarios:
            raise ValueError("Nenhum e-mail válido encontrado na planilha.")

        resultados = []
        enviados_com_sucesso = 0
        enviados_com_erro = 0

        for i, destinatario in enumerate(destinatarios, start=1):
            try:
                sucesso = enviar_email(destinatario, remetente_email, senha_email)
                resultados.append(
                    {
                        "destinatario": destinatario,
                        "sucesso": sucesso,
                        "mensagem": "E-mail enviado com sucesso.",
                    }
                )
                enviados_com_sucesso += 1
            except Exception as exc:
                resultados.append(
                    {
                        "destinatario": destinatario,
                        "sucesso": False,
                        "mensagem": str(exc),
                    }
                )
                enviados_com_erro += 1

            if i < len(destinatarios):
                time.sleep(PAUSA_ENTRE_ENVIOS_SEGUNDOS)

        return {
            "sucesso": enviados_com_erro == 0,
            "mensagem": "Processo concluído.",
            "total_destinatarios": len(destinatarios),
            "enviados_com_sucesso": enviados_com_sucesso,
            "enviados_com_erro": enviados_com_erro,
            "resultados": resultados,
        }
    except Exception as exc:
        raise RuntimeError(f"Erro ao executar envio dos e-mails: {exc}") from exc


if __name__ == "__main__":
    # Exemplo de uso
    # resultado = executar_envio(
    #     remetente_email="seu-email@dominio.com",
    #     senha_email="sua_senha_ou_senha_de_app",
    #     excel_path="Lista de email_2.xlsx",
    # )
    # print(resultado)
    pass
