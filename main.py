import json
import sys
from pathlib import Path
from aviso_notas_ppi_metodo import executar_envio

def ler_payload_task():
    """
    Lê o arquivo de payload enviado pelo orquestrador.
    """

    task_file = None

    # Tenta pegar via argumento (ex: python main.py caminho.json)
    if len(sys.argv) > 1:
        task_file = sys.argv[1]

    # Se não veio nada
    if not task_file:
        raise ValueError("Nenhum arquivo de payload foi informado.")

    path = Path(task_file)

    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {task_file}")

    # Lê e retorna o JSON
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)

def main():
    """
    Função principal do bot.

    Responsável por:
    - Receber argumentos
    - Ler usuario e senha
    - Enviar e-mail
    - Retornar sucesso ou erro
    """

    print("=" * 80)
    print("INICIANDO BOT E-MAIL INICIADO")
    print("=" * 80)

    # 🔥 Lê o payload recebido
    payload = ler_payload_task()

    # pega o primeiro parâmetro
    param = payload["parameters"][0]

    # converte a string para dict
    params_json = json.loads(param["parameter_value"])

    # acessa os valores
    email = params_json["dados_acesso"]["email"]
    senha = params_json["dados_acesso"]["senha"]

    resultado = executar_envio(
        remetente_email=email,
        senha_email=senha,
        excel_path="Lista de email_2.xlsx",
    )

    # 🔥 Aqui seria o processamento real do bot
    # (neste exemplo apenas simulação)

    resultado = {
        "status": "success",
        "mensagem": "Payload recebido e processado com sucesso."
    }

    print(resultado)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        erro = {
            "status": "error",
            "mensagem": str(e)
        }

        print("\n[ERRO]")
        print(json.dumps(erro, indent=4, ensure_ascii=False))

        # Retorna código de erro pro orquestrador
        sys.exit(1)
        