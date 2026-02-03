import os
import time
import threading
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from playwright.sync_api import sync_playwright

app = Flask(__name__)

# --- CONFIGURAÇÕES ---
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- VARIÁVEL GLOBAL (MEMÓRIA DO STATUS) ---
ESTADO_ROBO = {
    "rodando": False,
    "mensagem": "Aguardando...",
    "progresso": 0,
    "total": 0,
    "concluido": False,
    "sucesso": False
}

# --- FUNÇÕES DO ROBÔ (Inalteradas) ---

def preencher_blindado(page, placeholder, valor):
    if pd.isna(valor) or str(valor).strip() == "": return
    try:
        campo = page.get_by_placeholder(placeholder, exact=False)
        if campo.is_visible():
            campo.click()
            campo.fill("") 
            campo.type(str(valor), delay=100)
            time.sleep(1)
            page.keyboard.press("Tab")
    except: pass

def preencher_por_name(page, nome, valor):
    if pd.isna(valor) or str(valor).strip() == "": return
    try:
        campo = page.locator(f'input[name="{nome}"]')
        if campo.is_visible():
            campo.click()
            campo.fill("") 
            campo.type(str(valor), delay=100)
            time.sleep(1)
            page.keyboard.press("Tab")
    except: pass

def selecionar_credito(page, valor):
    try:
        if pd.isna(valor): return
        texto = str(valor).upper().strip()
        val = "0" if "ESPECIAL" in texto else "1"
        sel = page.locator('select[name="TipoCredito"]')
        if sel.is_visible():
            sel.select_option(value=val)
            time.sleep(1)
            page.keyboard.press("Tab")
    except: pass

def capturar_mensagem(page):
    try:
        # Seletor genérico para alertas
        el = page.locator(".alert, .toast-message, .alert-success, div[ng-message]")
        el.first.wait_for(state="visible", timeout=4000)
        return el.first.text_content().strip()
    except:
        return "Sem mensagem capturada"

# --- O MOTOR DO ROBÔ (Executa em Segundo Plano) ---

def worker_robo(caminho_arquivo):
    global ESTADO_ROBO
    playwright = None
    browser = None
    
    # Reset Status
    ESTADO_ROBO["rodando"] = True
    ESTADO_ROBO["concluido"] = False
    ESTADO_ROBO["mensagem"] = "🚀 Conectando ao Chrome..."
    
    try:
        playwright = sync_playwright().start()
        try:
            browser = playwright.chromium.connect_over_cdp("http://localhost:9222")
        except:
            ESTADO_ROBO["mensagem"] = "❌ Erro: Chrome Robô fechado. Abra o atalho!"
            ESTADO_ROBO["rodando"] = False
            return

        context = browser.contexts[0]
        page = None
        for aba in context.pages:
            if "aberturacredito" in aba.url:
                page = aba
                page.bring_to_front()
                break
        
        if not page:
            if len(context.pages) > 0: page = context.pages[0]
            else: 
                ESTADO_ROBO["mensagem"] = "❌ Nenhuma aba aberta."
                ESTADO_ROBO["rodando"] = False
                return

        # Lendo Excel
        ESTADO_ROBO["mensagem"] = "📂 Lendo Excel..."
        df = pd.read_excel(caminho_arquivo)
        df['Retorno Sistema'] = ""
        
        total = len(df)
        ESTADO_ROBO["total"] = total
        
        for index, row in df.iterrows():
            # Atualiza para o usuário ver
            ESTADO_ROBO["progresso"] = index + 1
            ESTADO_ROBO["mensagem"] = f"⚙️ Processando linha {index + 1} de {total}..."
            
            try:
                # Preenchimento
                preencher_blindado(page, "Unidade Executora", row.get('unidade'))
                preencher_blindado(page, "Função e Subfunção", row.get('funcao'))
                preencher_blindado(page, "Programa", row.get('programa'))
                preencher_blindado(page, "Ação", row.get('acao'))
                preencher_blindado(page, "Natureza da Despesa", row.get('natureza'))
                preencher_blindado(page, "Descrição da Despesa", row.get('descricao'))
                preencher_blindado(page, "Vínculo", row.get('vinculo'))
                selecionar_credito(page, row.get('credito'))
                preencher_blindado(page, "Data", row.get('data'))
                preencher_blindado(page, "Finalidade", row.get('finalidade'))
                preencher_por_name(page, "NumeroAto", row.get('numero_ato'))
                preencher_por_name(page, "NumeroLeiAutorizativa", row.get('lei_autorizativa'))

                # Botão Confirmar e Captura
                try:
                    btn = page.locator("button:has-text('Confirmar')")
                    if btn.is_visible():
                        btn.click()
                        ESTADO_ROBO["mensagem"] = f"📝 Linha {index + 1}: Salvando..."
                        msg = capturar_mensagem(page)
                        df.at[index, 'Retorno Sistema'] = msg
                        time.sleep(2)
                    else:
                        df.at[index, 'Retorno Sistema'] = "Botão não encontrado"
                except Exception as e:
                    df.at[index, 'Retorno Sistema'] = f"Erro botão: {str(e)}"

            except Exception as e:
                df.at[index, 'Retorno Sistema'] = f"Erro Linha: {str(e)}"
                continue

        # Finalização
        ESTADO_ROBO["mensagem"] = "💾 Gerando relatório..."
        caminho_relatorio = os.path.join(app.config['UPLOAD_FOLDER'], "relatorio_final.xlsx")
        df.to_excel(caminho_relatorio, index=False)
        
        ESTADO_ROBO["mensagem"] = "✨ Concluído com Sucesso!"
        ESTADO_ROBO["concluido"] = True
        ESTADO_ROBO["sucesso"] = True

    except Exception as e:
        ESTADO_ROBO["mensagem"] = f"❌ Erro Crítico: {str(e)}"
        ESTADO_ROBO["sucesso"] = False
    
    finally:
        ESTADO_ROBO["rodando"] = False
        if browser: 
            try: browser.close()
            except: pass
        if playwright: 
            try: playwright.stop()
            except: pass

# --- ROTAS (ENDPOINTS) ---

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/iniciar', methods=['POST'])
def iniciar():
    """Recebe o arquivo e inicia a Thread"""
    if 'arquivo' not in request.files: return jsonify({"erro": "Sem arquivo"}), 400
    
    arquivo = request.files['arquivo']
    if arquivo.filename == '': return jsonify({"erro": "Sem nome"}), 400
        
    caminho = os.path.join(app.config['UPLOAD_FOLDER'], "importacao.xlsx")
    arquivo.save(caminho)
    
    # Reseta Status
    global ESTADO_ROBO
    ESTADO_ROBO["rodando"] = True
    ESTADO_ROBO["mensagem"] = "Iniciando..."
    ESTADO_ROBO["progresso"] = 0
    ESTADO_ROBO["concluido"] = False

    # INICIA A THREAD (Isso impede o travamento)
    thread = threading.Thread(target=worker_robo, args=(caminho,))
    thread.daemon = True
    thread.start()
    
    return jsonify({"status": "ok"})

@app.route('/status')
def status():
    """O JavaScript chama isso a cada segundo"""
    return jsonify(ESTADO_ROBO)

@app.route('/download')
def download():
    path = os.path.join(app.config['UPLOAD_FOLDER'], "relatorio_final.xlsx")
    return send_file(path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, threaded=True)