import pandas as pd
import random
from collections import defaultdict

# Define a semente para reprodutibilidade
random.seed(42)

# Leitura dos dados dos orientadores
df_orientadores = pd.read_excel("planilhas/Resultado Final Exatas.xlsx")
orientadores = df_orientadores[["Nome do Orientador:", "Endereço de e-mail", "Departamento do Orientador:"]]
orientadores.columns = ["Nome", "Email", "Departamento"]

# Leitura dos avaliadores Ad Hoc
df_adhoc = pd.read_excel("planilhas/Ad Hoc - CIÊNCIAS EXATAS.xlsx", sheet_name="Pesquisadores")
avaliadores = df_adhoc[["Nome:", "E-mail Principal (id.uff):", "Área alocada no sistema"]]
avaliadores.columns = ["Nome", "Email", "Departamento"]

# Unifica avaliadores e orientadores (todos podem ser avaliadores)
avaliadores_total = pd.concat([orientadores, avaliadores], ignore_index=True).drop_duplicates(subset="Email")
avaliadores_total.rename(columns={"Departamento": "Área"}, inplace=True)
avaliadores_total.reset_index(drop=True, inplace=True)
Avaliadores = avaliadores_total.copy()

# Leitura dos projetos
df_projetos = pd.read_excel("planilhas/Ciências Exatas e da Terra-Engenharia-Agrárias.xlsx")
projetos = df_projetos.copy()

# Adiciona ID único na primeira coluna
projetos.reset_index(drop=True, inplace=True)
projetos.insert(0, "ID do Projeto", ["P{:03d}".format(i + 1) for i in projetos.index])

# Parâmetros
n_avaliadores_por_projeto = 4
avaliadores_em_ordem = list(Avaliadores.index) * ((len(projetos) * n_avaliadores_por_projeto) // len(Avaliadores) + 1)
random.shuffle(avaliadores_em_ordem)

# Atribuições com verificação: avaliador ≠ orientador do projeto
atribuições = defaultdict(list)
idx = 0
for i in projetos.index:
    atribuídos = set()
    orientador_nome = str(projetos.loc[i, "Nome do Orientador:"]).strip().lower()
    orientador_email = str(projetos.loc[i, "Endereço de e-mail"]).strip().lower()

    tentativas = 0
    while len(atribuídos) < n_avaliadores_por_projeto:
        if idx >= len(avaliadores_em_ordem):
            raise RuntimeError("Lista de avaliadores esgotada.")

        av_idx = avaliadores_em_ordem[idx]
        avaliador_nome = str(Avaliadores.loc[av_idx, "Nome"]).strip().lower()
        avaliador_email = str(Avaliadores.loc[av_idx, "Email"]).strip().lower()

        if (
            av_idx not in atribuídos and
            avaliador_nome != orientador_nome and
            avaliador_email != orientador_email
        ):
            atribuídos.add(av_idx)

        idx += 1
        tentativas += 1
        if tentativas > 1000:
            raise RuntimeError(f"Não foi possível encontrar avaliadores válidos para o projeto {i}")

    atribuições[i] = list(atribuídos)

# DataFrame com os projetos e avaliadores alocados
projetos_com_avaliadores = projetos.copy()
for k in range(1, n_avaliadores_por_projeto + 1):
    projetos_com_avaliadores[f"Avaliador {k}"] = projetos_com_avaliadores.index.map(lambda i: Avaliadores.loc[atribuições[i][k - 1], "Nome"])
    projetos_com_avaliadores[f"Email {k}"] = projetos_com_avaliadores.index.map(lambda i: Avaliadores.loc[atribuições[i][k - 1], "Email"])
    projetos_com_avaliadores[f"Área {k}"] = projetos_com_avaliadores.index.map(lambda i: Avaliadores.loc[atribuições[i][k - 1], "Área"])

# Mapeamento inverso: avaliador → infos dos projetos
projetos_por_avaliador = defaultdict(list)
for proj_id, avaliadores_ids in atribuições.items():
    projeto_info = {
        "ID": projetos.loc[proj_id, "ID do Projeto"],
        "Orientador": projetos.loc[proj_id, "Nome do Orientador:"],
        "Departamento": projetos.loc[proj_id, "Departamento do Orientador:"]
    }
    for av_id in avaliadores_ids:
        projetos_por_avaliador[av_id].append(projeto_info)

# Aba AvaliadoresPorProjeto com até 4 projetos
avaliador_linhas = []
for av_id, projetos_info in projetos_por_avaliador.items():
    linha = {
        "Nome do Avaliador": Avaliadores.loc[av_id, "Nome"],
        "Email": Avaliadores.loc[av_id, "Email"],
        "Área": Avaliadores.loc[av_id, "Área"]
    }
    for i in range(min(4, len(projetos_info))):
        linha[f"Projeto {i+1}"] = projetos_info[i]["ID"]
        linha[f"Orientador {i+1}"] = projetos_info[i]["Orientador"]
        linha[f"Departamento {i+1}"] = projetos_info[i]["Departamento"]
    avaliador_linhas.append(linha)

avaliadores_por_projeto_df = pd.DataFrame(avaliador_linhas)

# Exporta planilha com duas abas
with pd.ExcelWriter("saida/Projetos_com_Avaliadores.xlsx", engine="xlsxwriter") as writer:
    projetos_com_avaliadores.to_excel(writer, sheet_name="Projetos", index=False)
    avaliadores_por_projeto_df.to_excel(writer, sheet_name="AvaliadoresPorProjeto", index=False)

print("Arquivo 'Projetos_com_Avaliadores.xlsx' com duas abas gerado com sucesso.")
