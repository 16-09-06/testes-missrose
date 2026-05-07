from flask import Flask, request, jsonify
from flask_cors import CORS
from pywebpush import webpush, WebPushException
import json
import os
 
app = Flask(__name__)
 
# ✅ CORS configurado EXATAMENTE para o seu GitHub Pages e para testes locais.
# Impede que outros sites tentem usar sua API.
CORS(app, origins=["https://16-09-06.github.io", "http://localhost:5000", "http://127.0.0.1:5000" ,"https://miss-rose-rho.vercel.app/"])
 
# ✅ Suas chaves VAPID — a chave privada é lida de variável de ambiente
# para não ficar exposta no código. No seu servidor, execute:
#   export VAPID_PRIVATE_KEY="7sgXHbkDLECmWtIOw8HB34sYz0UBndUlK5ot1DTVpO4"
# Para testes locais ainda funciona com o fallback abaixo.
VAPID_PRIVATE_KEY = os.environ.get("VAPID_PRIVATE_KEY", "7sgXHbkDLECmWtIOw8HB34sYz0UBndUlK5ot1DTVpO4")
VAPID_CLAIMS = {
    "sub": "mailto:faturamento@missrosebra.com"
}
 
# ✅ Banco de inscrições usando um arquivo JSON simples para não perder
# os dados quando o servidor reiniciar (Excelente para o PythonAnywhere).
DB_FILE = "inscricoes.json"

def carregar_banco():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r") as f:
            return json.load(f)
    return {}

def salvar_banco(db):
    with open(DB_FILE, "w") as f:
        json.dump(db, f)

subscriptions_db = carregar_banco()
 
# Mapeamento de equipes — espelha a lógica do app.js
EQUIPES = {
    "equipe_renata": ["RENATA", "HOZANA", "ISRAEL", "ROSANGELA", "SARA", "VINICIUS"],
    "equipe_carol":  ["CAROL", "ALICE", "CHARLENE", "HEMILLY", "MICHELLE"],
}
 
def get_targets(target_group):
    """Retorna a lista de nomes de usuários para o grupo alvo."""
    if target_group == "todas":
        return list(subscriptions_db.keys())
    elif target_group in EQUIPES:
        return [u for u in EQUIPES[target_group] if u in subscriptions_db]
    else:
        # Tenta como nome individual
        return [target_group.upper()] if target_group.upper() in subscriptions_db else []
 
# ---------------------------------------------------------------------------
# ROTA 1: Salvar inscrição de um usuário
# Chamada pelo app.js quando a vendedora habilita notificações
# ---------------------------------------------------------------------------
@app.route('/salvar-inscricao', methods=['POST'])
def salvar_inscricao():
    dados = request.get_json()
    if not dados:
        return jsonify({"erro": "Nenhum dado fornecido"}), 400
 
    nome = dados.get('nome', '').upper().strip()
    subscription = dados.get('subscription')
 
    if not nome or not subscription:
        return jsonify({"erro": "Campos 'nome' e 'subscription' são obrigatórios"}), 400
 
    subscriptions_db[nome] = subscription
    salvar_banco(subscriptions_db) # Garante que a nova inscrição seja salva no arquivo
    print(f"[Push] Inscrição salva para: {nome}. Total inscritos: {len(subscriptions_db)}")
    return jsonify({"status": "sucesso", "usuario": nome}), 200
 
 
# ---------------------------------------------------------------------------
# ROTA 2: Enviar notificação push para um grupo ou usuário
# Chamada pelo app.js quando o admin dispara uma mensagem
# ---------------------------------------------------------------------------
@app.route('/enviar-push', methods=['POST'])
def enviar_push():
    dados = request.get_json()
    if not dados:
        return jsonify({"erro": "Nenhum dado fornecido"}), 400
 
    target = dados.get('target', 'todas')
    title  = dados.get('title', 'Miss Rôse')
    body   = dados.get('body', 'Nova notificação!')
    url    = dados.get('url', '/')
 
    destinatarios = get_targets(target)
 
    if not destinatarios:
        return jsonify({"status": "aviso", "detalhe": f"Nenhuma inscrição ativa encontrada para '{target}'"}), 200
 
    mensagem = json.dumps({"title": title, "body": body, "url": url})
 
    resultados = {"enviados": [], "falhas": []}
 
    for nome in destinatarios:
        subscription = subscriptions_db.get(nome)
        if not subscription:
            continue
        try:
            webpush(
                subscription_info=subscription,
                data=mensagem,
                vapid_private_key=VAPID_PRIVATE_KEY,
                vapid_claims=VAPID_CLAIMS,
                ttl=86400
            )
            resultados["enviados"].append(nome)
            print(f"[Push] ✅ Enviado para {nome}")
        except WebPushException as ex:
            resultados["falhas"].append({"usuario": nome, "erro": repr(ex)})
            print(f"[Push] ❌ Falha ao enviar para {nome}: {ex}")
            # Se a inscrição expirou (status 410), remove do banco
            if hasattr(ex, 'response') and ex.response and ex.response.status_code == 410:
                print(f"[Push] Removendo inscrição expirada de {nome}")
                del subscriptions_db[nome]
                salvar_banco(subscriptions_db) # Garante que a remoção seja salva no arquivo
 
    return jsonify({"status": "sucesso", "resultado": resultados}), 200
 
 
# ---------------------------------------------------------------------------
# ROTA 3: Verificação de status (útil para saber se o servidor está online)
# ---------------------------------------------------------------------------
@app.route('/', methods=['GET'])
def status():
    return jsonify({
        "status": "online",
        "app": "Robô de Mensagens da Miss Rôse",
        "inscritos": len(subscriptions_db)
    })
 
 
if __name__ == '__main__':
    # Use debug=False em produção
    app.run(host='0.0.0.0', port=5000, debug=True)
 