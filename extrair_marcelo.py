import pandas as pd
import os
import re

# ==============================================================================
# 1. LISTA CORRIGIDA DOS SEUS ARQUIVOS ALVO
# ==============================================================================
arquivos_alvo = [
    r"C:\Users\Aluno\Desktop\Aluno\aula3\pedro.xlsx",
    r"C:\Users\Aluno\Desktop\Aluno\aula3\ivanildo.xlsx",
    r"C:\Users\Aluno\Desktop\Aluno\aula3\neide.xlsx",
    r"C:\Users\Aluno\Desktop\Aluno\aula3\ildo.xlsx",
    r"C:\Users\Aluno\Desktop\Aluno\aula3\fora_estado.xlsx",
    r"C:\Users\Aluno\Desktop\Aluno\aula3\guarulhos.xlsx",
    r"C:\Users\Aluno\Desktop\Aluno\aula3\cliente.xlsx",
]

# Pasta onde você deseja salvar o resultado consolidado final
pasta_destino = r"C:\Users\Aluno\Downloads\original e editado"
# ==============================================================================

dados_cadastrais = []

print(f"Iniciando a extração cirúrgica de dados para os {len(arquivos_alvo)} arquivos definidos...")

for caminho_arquivo in arquivos_alvo:
    # Verifica se o arquivo realmente existe antes de tentar abrir
    if not os.path.exists(caminho_arquivo):
        print(f"⚠️ Atenção: O arquivo não foi encontrado e será pulado: {caminho_arquivo}")
        continue
        
    nome_arquivo = os.path.basename(caminho_arquivo)
    print(f"Processando arquivo: {nome_arquivo}")
    
    df = pd.read_excel(caminho_arquivo, header=None)
    total_linhas, total_colunas = df.shape
    
    # Localizar todas as linhas onde um bloco de cliente inicia
    linhas_inicio_cliente = []
    for r in range(total_linhas):
        if pd.notna(df.iloc[r, 0]) and "Dt Pedido" in str(df.iloc[r, 0]):
            linhas_inicio_cliente.append(r)
            
    # Varrer cada bloco horizontalmente de 20 em 20 colunas
    for r_inicio in lines_inicio_cliente if 'lines_inicio_cliente' in locals() else linhas_inicio_cliente:
        for c in range(0, total_colunas, 20):
            if c + 3 >= total_colunas:
                break
                
            nome_cliente = df.iloc[r_inicio + 1, c + 3] if pd.notna(df.iloc[r_inicio + 1, c + 3]) else None
            if not nome_cliente or str(nome_cliente).strip().lower() in ["nan", "0", ""]:
                continue
            
            cliente = {
                "Codigo_Cliente": None,
                "Nome_Cliente": str(nome_cliente).strip(),
                "CNPJ": None,
                "Inscricao_Estadual": None,
                "Data_Ultimo_Pedido": None,
                "Inicio_Relacionamento": None,
                "Sintegra": None,
                "Tabela_Preco": None,
                "Condicao_Pagamento": None,
                "Endereco": None,
                "Cidade": None,
                "Estado": None,
                "CEP": None,
                "Email": None,
                "Telefone": None,
                "Link_Google_Maps": None,
                "Contato_Responsavel": None,
                "Observacoes": None,
                "Arquivo_Origem": nome_arquivo
            }
            
            if pd.notna(df.iloc[r_inicio, c + 4]):
                cliente["Data_Ultimo_Pedido"] = df.iloc[r_inicio, c + 4]
            if pd.notna(df.iloc[r_inicio + 2, c + 2]):
                cliente["Endereco"] = str(df.iloc[r_inicio + 2, c + 2]).strip()
            if pd.notna(df.iloc[r_inicio + 4, c + 2]):
                cliente["CNPJ"] = str(df.iloc[r_inicio + 4, c + 2]).strip()
                
            if pd.notna(df.iloc[r_inicio + 5, c + 3]):
                cliente["Contato_Responsavel"] = str(df.iloc[r_inicio + 5, c + 3]).strip()
            elif pd.notna(df.iloc[r_inicio + 5, c + 4]):
                cliente["Contato_Responsavel"] = str(df.iloc[r_inicio + 5, c + 4]).strip()

            # Varredura interna do bloco do cliente
            for r_offset in range(7):
                r_atual = r_inicio + r_offset
                if r_atual >= total_linhas: 
                    break
                    
                for c_atual in range(c, c + 20):
                    if c_atual >= total_colunas: 
                        break
                        
                    celula = df.iloc[r_atual, c_atual]
                    if pd.isna(celula): 
                        continue
                    
                    texto = str(celula).strip()
                    texto_upper = texto.upper()
                    texto_rotulo = texto_upper.replace(" ", "").replace(":", "").replace(".", "")
                    
                    # Extração de Inscrição Estadual
                    if texto_rotulo == "IE" or texto_rotulo == "INSCRICAOESTADUAL":
                        numeros_ie = re.findall(r'\d[\d\.\-\/]*', texto)
                        if numeros_ie:
                            cliente["Inscricao_Estadual"] = numeros_ie[0]
                        elif c_atual + 1 < total_colunas and pd.notna(df.iloc[r_atual, c_atual + 1]):
                            val_dir = str(df.iloc[r_atual, c_atual + 1]).strip()
                            if val_dir.lower() != "nan" and not any(x in val_dir.upper() for x in ["MARCELO", "OBS:"]):
                                cliente["Inscricao_Estadual"] = val_dir

                    # Extração de Cidade
                    elif texto_rotulo == "CIDADE":
                        if c_atual + 1 < total_colunas and pd.notna(df.iloc[r_atual, c_atual + 1]):
                            val_dir = str(df.iloc[r_atual, c_atual + 1]).strip()
                            if val_dir.lower() != "nan" and val_dir != "":
                                cliente["Cidade"] = val_dir
                        if not cliente["Cidade"]:
                            for offset_c in range(1, 5):
                                if c_atual + offset_c < total_colunas:
                                    val_cidade = df.iloc[r_atual, c_atual + offset_c]
                                    if pd.notna(val_cidade) and str(val_cidade).strip().lower() not in ["nan", ""]:
                                        if not any(x in str(val_cidade).upper() for x in ["ESTADO", "CEP", "EMAIL"]):
                                            cliente["Cidade"] = str(val_cidade).strip()
                                            break

                    # Demais campos por proximidade direta à direita
                    elif texto_rotulo == "NºCLIENTE" or texto_rotulo == "NUMEROCLIENTE":
                        if c_atual + 1 < total_colunas and pd.notna(df.iloc[r_atual, c_atual + 1]):
                            cliente["Codigo_Cliente"] = str(df.iloc[r_atual, c_atual + 1]).strip()
                    elif texto_rotulo == "LOC":
                        if c_atual + 1 < total_colunas and pd.notna(df.iloc[r_atual, c_atual + 1]):
                            cliente["Link_Google_Maps"] = str(df.iloc[r_atual, c_atual + 1]).strip()
                    elif texto_rotulo == "TEL":
                        if c_atual + 1 < total_colunas and pd.notna(df.iloc[r_atual, c_atual + 1]):
                            cliente["Telefone"] = str(df.iloc[r_atual, c_atual + 1]).strip()
                    elif texto_rotulo == "CEP":
                        if c_atual + 1 < total_colunas and pd.notna(df.iloc[r_atual, c_atual + 1]):
                            cliente["CEP"] = str(df.iloc[r_atual, c_atual + 1]).strip()
                    elif texto_rotulo == "EMAIL":
                        if c_atual + 1 < total_colunas and pd.notna(df.iloc[r_atual, c_atual + 1]):
                            cliente["Email"] = str(df.iloc[r_atual, c_atual + 1]).strip()
                    elif texto_rotulo == "OBS":
                        if c_atual + 1 < total_colunas and pd.notna(df.iloc[r_atual, c_atual + 1]):
                            cliente["Observacoes"] = str(df.iloc[r_atual, c_atual + 1]).strip()
                    elif texto_rotulo == "ESTADO":
                        if c_atual + 1 < total_colunas and pd.notna(df.iloc[r_atual, c_atual + 1]):
                            cliente["Estado"] = str(df.iloc[r_atual, c_atual + 1]).strip()

                    # Rótulos acoplados com texto interno
                    if "INICIO:" in texto_upper:
                        cliente["Inicio_Relacionamento"] = texto.replace("INICIO:", "").strip()
                    elif "SINTEGRA:" in texto_upper:
                        cliente["Sintegra"] = texto.replace("SINTEGRA:", "").strip()
                    elif "TABELA:" in texto_upper:
                        cliente["Tabela_Preco"] = texto.replace("TABELA:", "").strip()
                    elif "COND.PGT:" in texto_upper:
                        cliente["Condicao_Pagamento"] = texto.replace("COND.PGT:", "").strip()

            dados_cadastrais.append(cliente)

# Consolidar tudo em um DataFrame final
df_final = pd.DataFrame(dados_cadastrais)

if not df_final.empty:
    df_final = df_final.drop_duplicates(subset=["Nome_Cliente", "CNPJ"], keep="last")

# Garante que a pasta destino existe e salva o arquivo consolidado
os.makedirs(pasta_destino, exist_ok=True)
caminho_salvar = os.path.join(pasta_destino, "banco_dados_clientes_cadastros_consolidado.xlsx")
df_final.to_excel(caminho_salvar, index=False)

print(f"\n✅ Concluído com sucesso! Arquivo consolidado gerado em:")
print(f"👉 {caminho_salvar}")
print(f"Total de clientes cadastrados unificados: {len(df_final)}")