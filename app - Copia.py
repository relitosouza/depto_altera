import os
import time
import pandas as pd
from flask import Flask, render_template, request
from playwright.sync_api import sync_playwright

app = Flask(__name__)

# Configurações
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- FUNÇÕES DE PREENCHIMENTO ---

def preencher_blindado(page, placeholder, valor):
    """Preenche campos buscando pelo TEXTO CINZA (Placeholder)"""
    if pd.isna(valor) or str(valor).strip() == "":
        return
    valor = str(valor)
    
    try:
        # Tenta achar o campo
        campo = page.get_by_placeholder(placeholder, exact=False)
        if campo.is_visible():
            campo.click()
            campo.fill("") 
            campo.type(valor, delay=100)
            time.sleep(1) 
            page.keyboard.press("Tab")
        else:
            print(f"   [Aviso] Placeholder '{placeholder}' não visível.")
    except Exception:
        # Ignora erros menores de campo para não travar o robô
        pass

def preencher_por_name(page, nome_tecnico, valor):
    """Preenche campos buscando pelo atributo NAME exato."""
    if pd.isna(valor) or str(valor).strip() == "":
        return
    valor = str(valor)

    try:
        campo = page.locator(f'input[name="{nome_tecnico}"]')
        if campo.is_visible():
            campo.click()
            campo.fill("") 
            campo.type(valor, delay=100)
            time.sleep(1)
            page.keyboard.press("Tab")
        else:
            print(f"   [Aviso] Campo name='{nome_tecnico}' não encontrado.")
    except Exception:
        pass

def selecionar_credito_blindado(page, valor_excel):
    """Seleciona o Menu Suspenso de Tipo de Crédito"""
    try:
        if pd.isna(valor_excel): return
        texto = str(valor_excel).upper().strip()
        valor_para_selecionar = "0" if "ESPECIAL" in texto else "1"
        
        select = page.locator('select[name="TipoCredito"]')
        if select.is_visible():
            select.select_option(value=valor_para_selecionar)
            time.sleep(1)
            page.keyboard.press("Tab")
    except Exception:
        pass

# --- ROBÔ PRINCIPAL ---

def executar_robo(caminho_arquivo):
    playwright = None
    browser = None
    
    try:
        playwright = sync_playwright().start()
        print("🔌 Conectando ao Chrome Robô...")
        
        # Conecta no navegador aberto
        try:
            browser = playwright.chromium.connect_over_cdp("http://localhost:9222")
        except Exception:
            return False, "Erro: Não consegui conectar. O Chrome Robô (tela preta) está aberto?"

        context = browser.contexts[0]
        
        # Busca a aba correta
        page = None
        for aba in context.pages:
            if "aberturacredito" in aba.url:
                page = aba
                page.bring_to_front()
                break
        
        if not page:
            # Se não achar a aba exata, pega a primeira visível
            if len(context.pages) > 0:
                page = context.pages[0]
            else:
                return False, "Nenhuma aba aberta no Chrome."

        # Lê Excel
        df = pd.read_excel(caminho_arquivo)
        registros = 0
        total = len(df)
        
        print(f"🚀 Iniciando processamento de {total} itens.")

        for index, row in df.iterrows():
            print(f"--> Linha {index + 1}/{total}")
            
            try:
                # 1. Campos Comuns
                preencher_blindado(page, "Unidade Executora", row.get('unidade'))
                preencher_blindado(page, "Função e Subfunção", row.get('funcao'))
                preencher_blindado(page, "Programa", row.get('programa'))
                preencher_blindado(page, "Ação", row.get('acao'))
                preencher_blindado(page, "Natureza da Despesa", row.get('natureza'))
                preencher_blindado(page, "Descrição da Despesa", row.get('descricao'))
                preencher_blindado(page, "Vínculo", row.get('vinculo'))
                
                # 2. Crédito
                selecionar_credito_blindado(page, row.get('credito'))
                
                # 3. Datas e Finalidade
                preencher_blindado(page, "Data", row.get('data'))
                preencher_blindado(page, "Finalidade", row.get('finalidade'))

                # 4. Novos Campos
                preencher_por_name(page, "NumeroAto", row.get('numero_ato'))
                preencher_por_name(page, "NumeroLeiAutorizativa", row.get('lei_autorizativa'))

                # 5. Botão Confirmar
                try:
                    botao = page.locator("button:has-text('Confirmar')")
                    if botao.is_visible():
                        print("   [Click] Confirmar")
                        botao.click()
                        time.sleep(3) # Espera salvar
                    else:
                        print("   [Aviso] Botão Confirmar não visto.")
                except:
                    pass

                registros += 1
                
            except Exception as e:
                print(f"   ❌ Erro na linha {index+1}: {e}")
                # Não para o robô, vai para a próxima linha
                continue 

        msg_sucesso = (
            f"✨ Processo finalizado!\n"
            f"✅ {registros} registros processados com sucesso."
        )
        return True, msg_sucesso

    except Exception as e:
        return False, f"Ocorreu um erro técnico: {e}"
        
    finally:
        # --- CORREÇÃO DO ERRO DE DESCONEXÃO ---
        # Aqui tentamos fechar suavemente. Se der erro, ignoramos silenciosamente.
        print("🏁 Finalizando conexão...")
        if browser:
            try:
                browser.close()
            except:
                pass # Ignora erro se já estiver desconectado
        
        if playwright:
            try:
                playwright.stop()
            except:
                pass

# --- FLASK ---

@app.route('/', methods=['GET', 'POST'])
def index():
    mensagem = None
    cor = "info"
    if request.method == 'POST':
        if 'arquivo' not in request.files:
            mensagem = "Envie um arquivo."
            cor = "warning"
        else:
            arquivo = request.files['arquivo']
            if arquivo.filename == '':
                mensagem = "Selecione um arquivo válido."
                cor = "warning"
            else:
                caminho = os.path.join(app.config['UPLOAD_FOLDER'], "importacao.xlsx")
                arquivo.save(caminho)
                sucesso, msg = executar_robo(caminho)
                mensagem = msg
                cor = "success" if sucesso else "danger"

    return render_template('index.html', mensagem=mensagem, cor=cor)

if __name__ == '__main__':
    # threaded=True é essencial para não travar a interface
    app.run(debug=True, threaded=True)