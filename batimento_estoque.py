# -*- coding: utf-8 -*-
"""
Batimento de estoque SAP x WMS
==============================
Lê as 6 bases de origem, monta o df_analise, classifica cada série (coluna "Analise") e
gera um Excel com duas abas: "Analise" (df_analise inteiro) e "Resumo" (contagem por status).

Como usar
---------
1) Deixe este arquivo na mesma pasta dos arquivos de origem (ou ajuste PASTA_ENTRADA abaixo):
     IQ09.xlsx | GerenciadorInformacoesEspecificas.csv | EstoqueLocalPorLote.csv |
     Mapa de Locais.xlsx | 00 - Avarias.xlsx | 01 - Devoluções.xlsx
2) Rode:  python batimento_estoque.py
3) O Excel sai na mesma pasta, com data e hora no nome (não sobrescreve o anterior).

No VS Code, os marcadores "# %%" permitem rodar célula por célula, como no notebook.
"""

# %% Imports e configuração
import os
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

try:
    PASTA_SCRIPT = Path(__file__).resolve().parent
except NameError:                      # rodando célula a célula, sem __file__
    PASTA_SCRIPT = Path.cwd()

# Pastas de entrada e saída (por padrão, a pasta deste arquivo). Dá para trocar pelas variáveis de ambiente.
PASTA_ENTRADA = Path(os.environ.get("BATIMENTO_ENTRADA", PASTA_SCRIPT))
PASTA_SAIDA   = Path(os.environ.get("BATIMENTO_SAIDA", PASTA_SCRIPT))

ARQ_IQ09        = PASTA_ENTRADA / "IQ09.xlsx"
ARQ_GERENCIADOR = PASTA_ENTRADA / "GerenciadorInformacoesEspecificas.csv"
ARQ_ESTOQUE     = PASTA_ENTRADA / "EstoqueLocalPorLote.csv"
ARQ_MAPA        = PASTA_ENTRADA / "Mapa de Locais.xlsx"
ARQ_AVARIAS     = PASTA_ENTRADA / "00 - Avarias.xlsx"
ARQ_DEVOLUCOES  = PASTA_ENTRADA / "01 - Devoluções.xlsx"

faltando = [a.name for a in (ARQ_IQ09, ARQ_GERENCIADOR, ARQ_ESTOQUE, ARQ_MAPA, ARQ_AVARIAS, ARQ_DEVOLUCOES) if not a.exists()]
if faltando:
    raise FileNotFoundError(f"Arquivos não encontrados em {PASTA_ENTRADA}: {', '.join(faltando)}")

CAMINHO_EXCEL = PASTA_SAIDA / f"Batimento_Estoque_{datetime.now():%Y-%m-%d_%H%M}.xlsx"

# %% Importar bases
df_iq09        = pd.read_excel(ARQ_IQ09)
df_gerenciador = pd.read_csv(ARQ_GERENCIADOR, low_memory=False)
df_estoque     = pd.read_csv(ARQ_ESTOQUE, low_memory=False)

# %% Corrige o gerenciador de informação específica
df_gerenciador_corrigido = df_gerenciador.copy()
df_gerenciador_corrigido = df_gerenciador_corrigido.dropna(subset=["Lote"])
df_gerenciador_corrigido["Lote"] = df_gerenciador_corrigido["Lote"].astype("int64")
df_gerenciador_corrigido = df_gerenciador_corrigido[
    df_gerenciador_corrigido["Depositante"] != "MIDEA INDUSTRIA E COMERCIO DO BRASIL LTDA"
]

# %% Corrige o estoque local por lote
df_estoque_corrigido = df_estoque.copy()
df_estoque_corrigido["Data de Fabricação"] = pd.to_datetime(
    df_estoque_corrigido["Data de Fabricação"],
    format="%d/%m/%Y %H:%M:%S",
    errors="coerce",
).astype("datetime64[us]")

df_estoque_corrigido = df_estoque_corrigido.merge(
    df_gerenciador_corrigido[["Lote", "Valor"]],   # chave + coluna que vem do gerenciador
    on="Lote",
    how="left",
)
df_estoque_corrigido = df_estoque_corrigido.dropna(subset="Valor")

df_estoque_corrigido = df_estoque_corrigido[[
    "Local", "Lote", "Código do Produto", "Local Ativo", "Lote Liberado",
    "Data de Fabricação", "Setor", "Valor",
]]
df_estoque_corrigido = df_estoque_corrigido.rename(columns={
    "Código do Produto": "Material",
    "Data de Fabricação": "Data",
    "Setor": "Setor Completo",
    "Valor": "Serial",
})

df_estoque_corrigido["Setor"] = np.where(
    df_estoque_corrigido["Setor Completo"].str.startswith("B", na=False),
    df_estoque_corrigido["Setor Completo"].str[:9],
    df_estoque_corrigido["Setor Completo"],
)
df_estoque_corrigido["Lote"] = df_estoque_corrigido["Lote"].astype("int64")

# %% Corrige a IQ09
df_iq09_corrigido = df_iq09.copy()
df_iq09_corrigido = df_iq09_corrigido[["Material", "UII", "Centro", "Depósito", "Criado em"]]
df_iq09_corrigido = df_iq09_corrigido.rename(columns={"UII": "Serial", "Depósito": "Deposito", "Criado em": "Data"})
df_iq09_corrigido["Setor"] = df_iq09_corrigido["Centro"].astype("str") + "_" + df_iq09_corrigido["Deposito"].astype("str")

# %% Cria a base de análise (todas as séries dos dois estoques, sem repetir)
df_analise = (
    pd.concat(
        [df_estoque_corrigido[["Serial"]], df_iq09_corrigido[["Serial"]]],
        ignore_index=True,
    )
    .drop_duplicates()
    .reset_index(drop=True)
)

# Dados do SAP (IQ09)
df_analise = df_analise.merge(
    df_iq09_corrigido[["Serial", "Material", "Centro", "Deposito", "Data", "Setor"]],
    on="Serial",
    how="left",
)
df_analise = df_analise.rename(columns={
    "Material": "Material SAP", "Centro": "Centro SAP", "Deposito": "Deposito SAP",
    "Data": "Data SAP", "Setor": "Setor SAP",
})

# Dados do WMS (estoque local por lote)
df_analise = df_analise.merge(
    df_estoque_corrigido[["Serial", "Local", "Lote", "Material", "Local Ativo", "Lote Liberado", "Data", "Setor Completo", "Setor"]],
    on="Serial",
    how="left",
)
df_analise = df_analise.rename(columns={
    "Material": "Material WMS", "Data": "Data WMS", "Setor": "Setor WMS", "Setor Completo": "Setor Completo WMS",
})
df_analise["Centro WMS"] = np.where(
    df_analise["Setor WMS"].str.startswith("B", na=False),
    df_analise["Setor WMS"].str[:4],
    df_analise["Setor WMS"],
)
df_analise["Deposito WMS"] = np.where(
    df_analise["Setor WMS"].str.startswith("B", na=False),
    df_analise["Setor WMS"].str[5:9],
    df_analise["Setor WMS"],
)
df_analise["Rua"] = df_analise["Local"].str[:3]

# Tipo de local, pela rua
rua_txt = df_analise["Rua"].astype(str).str.strip().str.upper()
cond_tipo_local = [
    df_analise["Rua"].isna() | (rua_txt == ""),   # sem local (nulo ou vazio)
    rua_txt.str.startswith("R"),
    rua_txt.str.startswith("B"),
    rua_txt.str.startswith("F"),
    rua_txt.str.startswith("AV"),
]
val_tipo_local = ["Sem Local", "Rua", "Blocado", "Falta", "Avaria"]
df_analise["Tipo de Local"] = np.select(cond_tipo_local, val_tipo_local, default="Outros")

df_analise["Lote"] = df_analise["Lote"].astype("Int64")

# %% Mapa de locais
df_mapa = pd.read_excel(ARQ_MAPA)
df_mapa["Tipo de Centro Rua"] = np.where(df_mapa["Centro2"].str.contains("-"), "Mono", "Multi")
df_mapa = df_mapa.rename(columns={
    "Tipo": "Tipo de Material Rua", "Centro1": "Centro 1 da Rua", "Centro2": "Centro 2 da Rua",
})

df_analise = df_analise.merge(
    df_mapa[["Rua", "Tipo de Material Rua", "Centro 1 da Rua", "Centro 2 da Rua", "Tipo de Centro Rua"]],
    on="Rua",
    how="left",
)

# Reordena as colunas
df_analise = df_analise[[
    "Serial", "Material SAP", "Material WMS", "Data SAP", "Data WMS",
    "Centro SAP", "Deposito SAP", "Setor SAP",
    "Centro WMS", "Deposito WMS", "Setor WMS", "Setor Completo WMS",
    "Local", "Local Ativo", "Rua", "Tipo de Local", "Tipo de Material Rua",
    "Centro 1 da Rua", "Centro 2 da Rua", "Tipo de Centro Rua",
    "Lote", "Lote Liberado",
]]

# %% Qualidade: avarias (fica só a data mais recente de cada série)
df_avaria = pd.read_excel(ARQ_AVARIAS)
df_avaria_corrigido = df_avaria.dropna(subset="Nº SÉRIE").copy()

# A coluna DATA vem com dois formatos misturados
data_str = df_avaria_corrigido["DATA"].astype(str).str.strip()
d_br  = pd.to_datetime(data_str, format="%d/%m/%Y %H:%M:%S", errors="coerce")  # 16/10/2026 09:30:55
d_iso = pd.to_datetime(data_str, format="ISO8601", errors="coerce")            # 2026-04-08 09:40:00
df_avaria_corrigido["DATA"] = d_br.fillna(d_iso)

df_avaria_corrigido["Nº SÉRIE"] = df_avaria_corrigido["Nº SÉRIE"].astype(str).str.strip().str.upper()
df_avaria_corrigido = (
    df_avaria_corrigido
    .sort_values("DATA", ascending=False)
    .drop_duplicates(subset="Nº SÉRIE", keep="first")
    .reset_index(drop=True)
)
df_avaria_corrigido = df_avaria_corrigido.rename(columns={"Nº SÉRIE": "Serial", "RESULTADO FINAL ETIQUETA": "Status Av"})

# %% Qualidade: devoluções (fica só a data mais recente de cada série)
df_devolucao = pd.read_excel(ARQ_DEVOLUCOES, header=1)
df_devolucao_corrigido = df_devolucao.dropna(subset="NÚMERO DE SÉRIE").copy()

data_str = df_devolucao_corrigido["DATA"].astype(str).str.strip()
d_br  = pd.to_datetime(data_str, format="%d/%m/%Y %H:%M:%S", errors="coerce")
d_iso = pd.to_datetime(data_str, format="ISO8601", errors="coerce")
df_devolucao_corrigido["DATA"] = d_br.fillna(d_iso)

df_devolucao_corrigido["NÚMERO DE SÉRIE"] = df_devolucao_corrigido["NÚMERO DE SÉRIE"].astype(str).str.strip().str.upper()
df_devolucao_corrigido = (
    df_devolucao_corrigido
    .sort_values("DATA", ascending=False)
    .drop_duplicates(subset="NÚMERO DE SÉRIE", keep="first")
    .reset_index(drop=True)
)
df_devolucao_corrigido = df_devolucao_corrigido.rename(columns={"NÚMERO DE SÉRIE": "Serial", "ESTOQUE FINAL": "Status Dev"})

# %% Qualidade: status final de cada série (Status IQC)
df_qualidade = (
    pd.concat(
        [df_avaria_corrigido[["Serial"]], df_devolucao_corrigido[["Serial"]]],
        ignore_index=True,
    )
    .drop_duplicates()
    .reset_index(drop=True)
)
df_qualidade = df_qualidade.merge(df_avaria_corrigido[["Serial", "Status Av"]], on="Serial", how="left")
df_qualidade = df_qualidade.merge(df_devolucao_corrigido[["Serial", "Status Dev"]], on="Serial", how="left")

# B137 vem com variações (B137-A etc.)
for col in ["Status Av", "Status Dev"]:
    df_qualidade.loc[df_qualidade[col].str.startswith("B137", na=False), col] = "B137"

av  = df_qualidade["Status Av"].astype(object).str.strip().str.upper()
dev = df_qualidade["Status Dev"].astype(object).str.strip().str.upper()

cond_iqc = [
    av.isna() | dev.isna(),                  # algum nulo: vale o preenchido
    (av == dev),                             # iguais: vale o Av
    (av == "B136") | (dev == "B136"),
    (av == "B137") | (dev == "B137"),
    (av == "B149") | (dev == "B149"),
]
val_iqc = [av.combine_first(dev), av, "B136", "B137", "B149"]
df_qualidade["Status IQC"] = np.select(cond_iqc, val_iqc, default="APROVADO")

# Traz o Status IQC para a análise
df_analise = df_analise.merge(df_qualidade[["Serial", "Status IQC"]], on="Serial", how="left")

# Aprovado vira o depósito livre do centro
mapa_livre = {
    "B123": "B110", "B124": "B110",
    "B158": "B114", "B159": "B114",
    "B179": "B413", "B180": "B413",
}
aprovado = df_analise["Status IQC"] == "APROVADO"
livre_do_centro = df_analise["Centro SAP"].map(mapa_livre)
mask = aprovado & livre_do_centro.notna()
df_analise.loc[mask, "Status IQC"] = livre_do_centro[mask]

# Garantia não tem B149 (outlet): vira o livre
mask = df_analise["Centro SAP"].isin(["B179", "B180"]) & (df_analise["Status IQC"] == "B149")
df_analise.loc[mask, "Status IQC"] = "B413"

# %% Regras da análise (ordem de prioridade em "regras")

def limpa(s):
    return s.astype(object).str.strip().str.upper()

def limpa_material(s):
    return (
        s.astype("string")
        .str.strip()
        .str.upper()
        .str.replace(r"\.0$", "", regex=True)
    )

def sem_acento(s):
    return (
        s.astype(object)
        .str.normalize("NFKD")
        .str.encode("ascii", "ignore")
        .str.decode("ascii")
        .str.strip()
        .str.upper()
    )

# Valores de "Tipo de Material Rua" que indicam rua de peças (sem acento, em maiúsculas)
TIPOS_RUA_PECAS = ["PECAS", "PECA"]

# True: vizinha certa com IQC preenchido bloqueia o avanço do local
vizinha_iqc_bloqueia = True

# False: a série com Data WMS vazia é uma ação por série e não bloqueia o local; se ela for a única
#        série errada, as vizinhas certas viram "OK - parcial local". Qualquer outro bloqueio no local continua valendo.
# True:  ela bloqueia o local como os demais bloqueios
fabricacao_bloqueia_local = False

# False: a série com data diferente entre SAP e WMS (corrigir data no WMS) não bloqueia o local
data_bloqueia_local = False

# False: a série em TEMP ou DEVT no SAP (pendente de internalização) não bloqueia o local
internalizacao_bloqueia_local = False

# True: a série com SAP em B135 e outro depósito no WMS bloqueia o local; as vizinhas ficam em análise
b135_bloqueia_local = True

# False: a série com SAP em B111 (falta) e WMS fora de falta (validar físico) não bloqueia o local
validar_fisico_bloqueia_local = False

# False: a série com SAP livre e WMS em avaria, sem IQC (validar no físico se é avaria), não bloqueia o local
validar_avaria_bloqueia_local = False

# False: a série com SAP em B140 e WMS em outro depósito (mover no WMS para B140) não bloqueia o local
b140_bloqueia_local = False

# Rótulo da regra do B140 (depósito de bloqueio da qualidade: o WMS tem que ficar igual ao SAP)
ROTULO_B140 = "Mover no WMS para B140"

# Rótulo da regra do B135
ROTULO_B135 = "Material em IQC, errado no WMS"

# Rótulos das regras de validar no físico
ROTULO_VALIDAR_EXISTE = "Validar no físico se existe"     # SAP em falta (B111) e WMS fora de falta
ROTULO_VALIDAR_AVARIA = "Validar no físico se é avaria"   # SAP livre e WMS em avaria, sem IQC

# Depósitos livres (SAP) e depósitos de avaria (WMS)
DEP_LIVRES = ["B110", "B114", "B413"]
DEP_AVARIA = ["B149", "B137", "B136"]

# Rótulo da regra do setor do local (SAP certo, setor do WMS diferente)
ROTULO_SETOR = "Corrigir setor no WMS"

# Rótulo da regra da rua errada (o centro do SAP não é o da rua onde o pallet está)
ROTULO_RUA = "Mover pallet para a rua do centro correto"

# Rótulo da regra de avaria: série em local de avaria com depósito livre no SAP
ROTULO_AV_B148 = "Mover no SAP para B148"

# Rótulo da regra de avaria: SAP já em depósito de avaria e WMS ainda em B148
ROTULO_AV_WMS = "Mover no WMS para o depósito do SAP"

# Rótulo da regra de avaria: setor do WMS diferente do SAP (o que sobraria em análise na avaria)
ROTULO_SETOR_AV = "Corrigir setor AV"

# Rótulos das regras de desmembrar (ação no físico, no local inteiro). Só valem para locais do tipo Rua ou Blocado.
ROTULO_DESM_FISICO = "Desmembrar físico"        # local com mais de um setor no SAP
ROTULO_DESM_DATA   = "Desmembrar multi data"    # local livre (pelo SAP) com mais de uma data no SAP

# Locais de stage (tipo de local "Outros"): doca, packing e afins. O que sobraria em "Em análise" neles vira "Em stage".
ROTULO_STAGE   = "Em stage"
LOCAIS_STAGE   = ["00000347", "00000348", "OUT001111", "00038456"]
PADRAO_STAGE   = r"^(000\d+|OUT\d+)$"           # outros locais com o mesmo padrão de nome

# Depósitos em que o lote precisa estar liberado (livres e B149)
DEP_LOTE_LIBERADO = ["B110", "B114", "B413", "B149"]

setor_sap  = limpa(df_analise["Setor SAP"])
setor_wms  = limpa(df_analise["Setor WMS"])
centro_sap = limpa(df_analise["Centro SAP"])
centro_wms = limpa(df_analise["Centro WMS"])
dep_wms    = limpa(df_analise["Deposito WMS"])
dep_sap    = limpa(df_analise["Deposito SAP"])
centro1    = limpa(df_analise["Centro 1 da Rua"])
centro2    = limpa(df_analise["Centro 2 da Rua"])
tipo_rua   = limpa(df_analise["Tipo de Centro Rua"])
iqc        = limpa(df_analise["Status IQC"])
mat_sap    = limpa_material(df_analise["Material SAP"])
mat_wms    = limpa_material(df_analise["Material WMS"])
loc_ativo  = limpa(df_analise["Local Ativo"])
lote_lib   = limpa(df_analise["Lote Liberado"])
tipo_mat_rua = sem_acento(df_analise["Tipo de Material Rua"])
tipo_local   = sem_acento(df_analise["Tipo de Local"])

# Locais do tipo Rua ou Blocado: são os que estão no mapa de ruas, e os únicos que podem ser desmembrados
tipo_fisico = tipo_local.isin(["RUA", "BLOCADO"])

# Local de stage: tipo "Outros" e nome na lista ou no padrão
local_txt = limpa(df_analise["Local"])
em_stage  = (tipo_local == "OUTROS") & (local_txt.isin(LOCAIS_STAGE) | local_txt.str.match(PADRAO_STAGE, na=False))

# Local de avaria: não trava o local inteiro, só as séries dele. Por isso, nos locais de avaria cada série
# forma o seu próprio grupo; nos demais locais o grupo continua sendo o local.
em_avaria = tipo_local == "AVARIA"
g = df_analise["Local"].where(~em_avaria, "AV#" + df_analise["Serial"].astype(str))

# Serial duplicado no WMS (as duas linhas; o SAP não repete serial)
serial_dup_wms = df_analise["Serial"].duplicated(keep=False)

# Existência do serial em cada estoque
so_wms   = df_analise["Local"].notna() & df_analise["Setor SAP"].isna()    # sem Setor SAP
so_sap   = df_analise["Local"].isna()  & df_analise["Setor SAP"].notna()   # sem Local
nos_dois = df_analise["Local"].notna() & df_analise["Setor SAP"].notna()

# Material divergente: nos dois estoques, material preenchido nos dois lados e diferente
material_div = (
    nos_dois
    & mat_sap.notna()
    & mat_wms.notna()
    & (mat_sap != mat_wms).fillna(False).astype(bool)
)

# Falta: setor FALTA ou SOBRA/FALTA no WMS
falta_wms       = setor_wms.isin(["FALTA", "SOBRA/FALTA"])
falta_b111      = falta_wms & (dep_sap == "B111")
falta_fora_b111 = falta_wms & dep_sap.notna() & (dep_sap != "B111")

# Validar no físico se existe: o SAP está em falta (B111) e o WMS não. A falta prevalece sobre o IQC,
# e a série precisa ter o físico validado antes de qualquer outra correção.
validar_fisico = nos_dois & (dep_sap == "B111") & ~falta_wms

# B140 (bloqueio da qualidade) é absoluto: se o SAP está em B140, o WMS tem que estar em B140.
# Série em falta no WMS mantém as regras de falta.
b140_errado = nos_dois & (dep_sap == "B140") & dep_wms.notna() & (dep_wms != "B140") & ~falta_wms

# B140 só no WMS (SAP em outro depósito) também é erro, porque o B140 sempre nasce no SAP.
# Ainda sem ação definida: a série fica em análise e trava o local.
b140_so_wms = nos_dois & (dep_wms == "B140") & dep_sap.notna() & (dep_sap != "B140")

# Fabricação: série nos dois estoques, com Data SAP, sem Data WMS e depósito SAP diferente de TEMP e B145
data_sap_raw = pd.to_datetime(df_analise["Data SAP"])
data_wms_raw = pd.to_datetime(df_analise["Data WMS"])
preencher_fab = (
    nos_dois
    & data_sap_raw.notna()
    & data_wms_raw.isna()
    & dep_sap.notna()
    & ~dep_sap.isin(["TEMP", "B145"])
)

# Pendente de internalização no SAP: depósito TEMP (inbound) ou DEVT (devolução)
pend_inbound = dep_sap == "TEMP"
pend_devol   = dep_sap == "DEVT"
pend_sap     = pend_inbound | pend_devol

# Produto em rua de peças: a rua do local é de peças, então o pallet inteiro está no endereço errado
produto_em_pecas = df_analise["Local"].notna() & tipo_mat_rua.isin(TIPOS_RUA_PECAS)

# Condições por série
setor_ok  = setor_sap.notna() & (setor_sap == setor_wms)

# Local de avaria não depende do mapa de ruas: para ele, basta ter centro no WMS
centro_ok = centro_wms.notna() & (
    (centro_wms == centro1) |
    ((tipo_rua == "MULTI") & (centro_wms == centro2)) |
    em_avaria
)

data_sap = pd.to_datetime(df_analise["Data SAP"]).dt.normalize()
data_wms = pd.to_datetime(df_analise["Data WMS"]).dt.normalize()
data_ok  = data_sap.notna() & (data_sap == data_wms)

material_ok = mat_sap.notna() & (mat_sap == mat_wms).fillna(False).astype(bool)

# Data diferente: nos dois estoques, as duas datas preenchidas e diferentes (só o dia)
corrigir_data = nos_dois & data_sap.notna() & data_wms.notna() & ~data_ok

# B135 (análise de qualidade): SAP em B135 e WMS em outro depósito, e essa é a única divergência
# (mesmo centro, centro da rua, data e material iguais). A ação é trocar o setor do local no WMS.
# B135 só no WMS (SAP em outro depósito) está errado e fica em análise.
b135_sap = dep_sap == "B135"
b135_wms = dep_wms == "B135"
b135_errado = (
    nos_dois
    & ~serial_dup_wms
    & ~falta_wms
    & ~pend_sap
    & b135_sap
    & dep_wms.notna()
    & ~b135_wms
    & (centro_sap == centro_wms)
    & centro_ok
    & data_ok
    & material_ok
)
b135_so_wms = nos_dois & b135_wms & dep_sap.notna() & ~b135_sap

# Garantia (B179/B180): não existe B149 (sem outlet), o esperado é o livre B413
gar_sap     = centro_sap.isin(["B179", "B180"])
b149_em_gar = gar_sap & ((dep_sap == "B149") | (dep_wms == "B149"))

iqc_ef = iqc.where(~(b149_em_gar & iqc.isna()), "B413")          # sem IQC -> B413
iqc_ef = iqc_ef.where(~(gar_sap & (iqc_ef == "B149")), "B413")   # IQC B149 -> B413
iqc_ef = iqc_ef.where(dep_sap != "B140")                          # B140 é absoluto: o IQC é ignorado

# IQC que veio da qualidade de fato (sem o B413 inferido para garantia)
iqc_real = iqc_ef.where(iqc.notna())

# Se houver IQC (efetivo), ele deve ser igual ao depósito do WMS
iqc_ok = iqc_ef.isna() | (iqc_ef == dep_wms)

# Validar no físico se é avaria: o SAP está livre e o WMS em avaria, sem IQC que confirme a avaria.
# Com IQC preenchido valem as regras do IQC (ex.: IQC igual ao depósito do WMS é "Mover no SAP para o depósito do IQC").
validar_avaria = nos_dois & dep_sap.isin(DEP_LIVRES) & dep_wms.isin(DEP_AVARIA) & iqc_ef.isna()

# Mover no SAP para B148: a série está em um local de avaria, o depósito do SAP é livre e não há IQC.
# Em local de avaria esta regra vale no lugar de "Validar no físico se é avaria", porque o próprio local já é de avaria.
mover_b148 = nos_dois & em_avaria & dep_sap.isin(DEP_LIVRES) & iqc.isna()

# O IQC prevalece: série em local de avaria com SAP livre e IQC de avaria (B149, B137 ou B136)
# vai no SAP para o depósito do IQC, e não para o B148.
mover_iqc_av = nos_dois & em_avaria & dep_sap.isin(DEP_LIVRES) & iqc_real.isin(DEP_AVARIA) & (iqc_real != dep_sap)

# Validar no físico (local de avaria): SAP livre e IQC livre. A qualidade aprovou o material,
# mas ele está em um local de avaria, então o físico precisa ser conferido.
validar_av_local = nos_dois & em_avaria & dep_sap.isin(DEP_LIVRES) & iqc_real.isin(DEP_LIVRES)

# Mover no WMS para o depósito do SAP: série em local de avaria com o SAP já em depósito de avaria (B149, B137 ou B136)
# e o WMS ainda em B148. Vale sem IQC ou com IQC igual ao depósito do SAP, e com o mesmo centro nos dois lados.
# Em garantia o B149 não existe, então fica de fora.
mover_wms_av = (
    nos_dois & em_avaria
    & dep_sap.isin(DEP_AVARIA) & (dep_wms == "B148")
    & (centro_sap == centro_wms)
    & (iqc_real.isna() | (iqc_real == dep_sap))
    & ~(gar_sap & (dep_sap == "B149"))
)

# Corrigir setor AV: série em local de avaria, nos dois estoques, com o setor do WMS diferente do SAP e
# data e material iguais. O IQC, se existir, concorda com o depósito do SAP (senão o IQC prevalece e a série fica em análise).
# Fica no fim da lista de regras, então só vale para o que sobraria em "Em análise".
corrigir_setor_av = (
    nos_dois & em_avaria
    & ~setor_ok
    & data_ok & material_ok
    & ~b135_sap & ~b135_wms & ~b140_so_wms & ~falta_wms
    & (iqc_ef.isna() | (iqc_ef == dep_sap))
)

# Desmembrar multi data: local livre com mais de uma data no SAP. O FIFO exige uma única data por local livre.
# "Livre" vale pelo ponto de vista do SAP: só entram na conta as séries com depósito do SAP livre (B110, B114, B413).
# O depósito do WMS não é critério. Só vale para locais do tipo Rua ou Blocado. Todas as séries do local recebem o rótulo.
# Local de falta (setor FALTA ou SOBRA/FALTA no WMS) não é um endereço físico, então não entra.
livre_sap        = nos_dois & dep_sap.isin(DEP_LIVRES)
datas_livres_sap = data_sap.where(livre_sap).groupby(g, dropna=False).transform("nunique")
multi_data       = nos_dois & tipo_fisico & ~falta_wms & (datas_livres_sap > 1)

# Desmembrar físico: local do tipo Rua ou Blocado com mais de um setor no SAP.
# Série sem SAP (só no WMS) não conta como setor.
setores_no_local = setor_sap.where(nos_dois).groupby(g, dropna=False).transform("nunique")
desm_fisico      = nos_dois & tipo_fisico & (setores_no_local > 1)

# O IQC prevalece sobre o B135: o rótulo só vale sem IQC. Com IQC igual ao depósito do WMS, o caso é do
# "Mover no SAP para o depósito do IQC"; com IQC diferente do depósito do WMS, a série fica em análise.
b135_errado = b135_errado & iqc_ef.isna()

# Posição na sequência de qualidade (B148 antecede qualquer etapa)
ordem = {"B148": 0, "B110": 1, "B114": 1, "B413": 1, "B149": 2, "B137": 3, "B136": 4}
pos_dep = dep_wms.map(ordem)
pos_dep = pos_dep.where(~(gar_sap & (dep_wms == "B149")), 0)   # em garantia, B149 antecede o livre
pos_iqc = iqc_ef.map(ordem)

# Base comum: nos dois estoques, sem duplicado, fora de TEMP/DEVT, B135 errado, B140 errado ou só no WMS
# e validar no físico, centro da rua, data e material iguais
base_serie       = nos_dois & ~serial_dup_wms & ~pend_sap & ~b135_errado & ~b140_errado & ~b140_so_wms & ~validar_fisico & ~validar_avaria & centro_ok & data_ok & material_ok
mesmo_centro_dep = (centro_sap == centro_wms) & (dep_sap == dep_wms)

# A) série certa
serie_A = base_serie & setor_ok & iqc_ok

# B) mover só no SAP: WMS já está igual ao IQC
serie_B = (
    base_serie
    & (centro_sap == centro_wms)
    & iqc_ef.notna()
    & (iqc_ef == dep_wms)
    & (dep_sap != dep_wms)
)

# D) avançar no WMS e no SAP: SAP = WMS, e o depósito está antes do IQC na sequência
serie_D = (
    base_serie & mesmo_centro_dep
    & pos_dep.notna() & pos_iqc.notna() & (pos_dep < pos_iqc)
)

# E) exceção: SAP = WMS, e o depósito está depois do IQC (mais grave)
serie_E = (
    base_serie & mesmo_centro_dep
    & pos_dep.notna() & pos_iqc.notna() & (pos_dep > pos_iqc)
)

# F) garantia: só o WMS está em B149 e o SAP já está no livre B413
serie_F = (
    base_serie
    & (centro_sap == centro_wms)
    & gar_sap
    & (dep_wms == "B149")
    & (iqc_ef == "B413")
    & (dep_sap == "B413")
)

# Séries cuja ação move o local inteiro no WMS
serie_M = serie_D | serie_F

# S) corrigir o setor do local no WMS: o pallet inteiro tem o mesmo setor no SAP, o centro do SAP é o da rua,
#    data e material batem, e o setor do WMS é diferente. A correção é trocar o setor do local no WMS.
#    Não vale para local de avaria (para a avaria existe o "Corrigir setor AV").
centro_sap_ok = centro_sap.notna() & (
    (centro_sap == centro1) |
    ((tipo_rua == "MULTI") & (centro_sap == centro2)) |
    em_avaria
)

serie_S = (
    nos_dois
    & ~serial_dup_wms
    & ~pend_sap
    & ~b135_sap
    & ~b135_wms
    & ~falta_wms
    & ~validar_fisico
    & ~validar_avaria
    & ~b140_errado
    & ~b140_so_wms
    & ~produto_em_pecas
    & ~em_avaria
    & centro_sap_ok
    & data_ok
    & material_ok
    & ~setor_ok
    & (iqc_ef.isna() | (iqc_ef == dep_sap))      # o IQC, se existir, concorda com o depósito do SAP
)

# O resultado vale para o local inteiro: todas as séries do local devem ser "S",
# com o mesmo setor no SAP e um único setor no WMS
g_setor = g
local_todo_S          = serie_S.groupby(g_setor, dropna=False).transform("all")
local_setor_sap_unico = setor_sap.fillna("(SEM SAP)").groupby(g_setor, dropna=False).transform("nunique").eq(1)
local_setor_wms_unico = setor_wms.fillna("(SEM WMS)").groupby(g_setor, dropna=False).transform("nunique").eq(1)
corrigir_setor = serie_S & local_todo_S & local_setor_sap_unico & local_setor_wms_unico

# T) mover o pallet para a rua certa: o centro do SAP NÃO é o da rua onde o pallet está. O setor do SAP ser
#    igual ou diferente do setor do WMS é irrelevante. A rua precisa estar no mapa (sem mapa não há centro de referência).
serie_T = (
    nos_dois
    & ~serial_dup_wms
    & ~pend_sap
    & ~b135_sap
    & ~b135_wms
    & ~falta_wms
    & ~validar_fisico
    & ~validar_avaria
    & ~b140_errado
    & ~b140_so_wms
    & ~produto_em_pecas
    & centro1.notna()
    & ~centro_sap_ok
    & data_ok
    & material_ok
    & (iqc_ef.isna() | (iqc_ef == dep_sap))
)
local_todo_T = serie_T.groupby(g_setor, dropna=False).transform("all")
mover_rua    = serie_T & local_todo_T & local_setor_sap_unico & local_setor_wms_unico

# Séries cujo rótulo é uma ação própria por série (fabricação, data, internalização, B135, validar no físico, B140)
acao_serie = preencher_fab | corrigir_data | pend_sap | b135_errado | validar_fisico | validar_avaria | b140_errado

# Dessas, as que não bloqueiam o local. Duplicado e material divergente têm rótulo pior e continuam bloqueando.
acao_sem_bloqueio = (
    (
        (preencher_fab & (not fabricacao_bloqueia_local))
        | (corrigir_data & (not data_bloqueia_local))
        | (pend_sap & (not internalizacao_bloqueia_local))
        | (b135_errado & (not b135_bloqueia_local))
        | (validar_fisico & (not validar_fisico_bloqueia_local))
        | (validar_avaria & (not validar_avaria_bloqueia_local))
        | (b140_errado & (not b140_bloqueia_local))
    )
    & ~serial_dup_wms
    & ~material_div
)

# C) bloqueio: qualquer série que não seja A, B, D, E, F, S ou T, nem uma ação por série sem bloqueio.
# O B135 (errado no WMS ou só no WMS) e o B140 só no WMS bloqueiam o local mesmo que a série tenha outro rótulo de ação.
serie_C = (
    (~(serie_A | serie_B | serie_D | serie_E | serie_F | corrigir_setor | mover_rua) & ~acao_sem_bloqueio)
    | ((b135_errado | b135_so_wms) & b135_bloqueia_local)
    | (b140_errado & b140_bloqueia_local)
    | b140_so_wms
)

# O resultado vale para o local inteiro (em local de avaria, vale só para a própria série)
def por_local(m):
    return m.groupby(g, dropna=False).transform("any")

# Lote: o depósito avaliado é o de destino, quando o local vai ser movido no WMS
destino_local = iqc_ef.where(serie_M).groupby(g, dropna=False).transform("first")
dep_lote      = destino_local.fillna(dep_wms)
lote_bloq     = nos_dois & dep_lote.isin(DEP_LOTE_LIBERADO) & (lote_lib != "S")

# L) desbloquear lote: série certa, a mover só no SAP ou em exceção, com lote bloqueado
serie_L = lote_bloq & (serie_A | serie_B | serie_E)

local_tem_C = por_local(serie_C)
local_tem_B = por_local(serie_B)
local_tem_M = por_local(serie_M)
local_tem_L = por_local(serie_L | (acao_sem_bloqueio & nos_dois))

# Movimento no WMS: todas as séries do local devem ir para o mesmo IQC
alvo_unico = iqc_ef.where(serie_M).groupby(g, dropna=False).transform("nunique") <= 1

# Vizinha certa com IQC preenchido seria arrastada para um status pior que o da qualidade
vizinha_com_iqc  = por_local(serie_A & iqc_ef.notna())
conflito_vizinha = vizinha_com_iqc if vizinha_iqc_bloqueia else pd.Series(False, index=df_analise.index)

local_bloqueado = (
    local_tem_C
    | (local_tem_M & local_tem_B)       # mover no WMS e mover só no SAP no mesmo local
    | (local_tem_M & ~alvo_unico)       # IQC diferentes entre as séries a mover
    | (local_tem_M & conflito_vizinha)  # vizinha com IQC próprio
)

# Local Ativo: o local inteiro precisa estar ativo (o valor é único por local)
local_inativo = (df_analise["Local"].notna() & (loc_ativo != "S")).groupby(g, dropna=False).transform("any")

# Séries que recebem a ação de ativar o local (duplicado, material, falta e ações por série mantêm o rótulo)
elegivel_local = nos_dois & ~serial_dup_wms & ~material_div & ~falta_wms & ~acao_serie
lote_visivel   = lote_bloq & ~serie_M      # se a série vai ser movida, o mover sobrepõe o lote

ativar_e_lote = elegivel_local & local_inativo & lote_visivel
ativar_local  = elegivel_local & local_inativo

ok          = nos_dois & serie_A & ~local_bloqueado & ~local_inativo & ~local_tem_B & ~local_tem_M & ~local_tem_L
ok_parcial  = nos_dois & serie_A & ~local_bloqueado & ~local_inativo & (local_tem_B | local_tem_M | local_tem_L)
mover_iqc   = nos_dois & serie_B & ~local_bloqueado
avancar     = nos_dois & serie_D & ~local_bloqueado
mover_wms   = nos_dois & serie_F & ~local_bloqueado
desbloquear = nos_dois & serie_L & ~local_bloqueado & ~acao_serie
excecao     = nos_dois & serie_E & ~local_bloqueado

# A ordem das regras é a ordem de gravidade (a primeira verdadeira vence). Ordem definida:
#   0) integridade
#   1) ações só no sistêmico (WMS antes do SAP), com as verificações no físico
#   2) desmembrar (físico e multi data)
#   3) mudar o pallet de local
#   4) demais (OK e o que sobraria em análise)
regras = [
    # 0) integridade
    (serial_dup_wms,                  "Serial duplicado no WMS"),
    (material_div,                    "Material divergente entre SAP e WMS"),
    (so_wms,                          "Serial existe apenas no WMS"),
    # 1a) sistêmico no WMS (as verificações no físico mantêm a posição que já tinham: logo depois de fabricação e data)
    (preencher_fab,                   "Preencher fabricação no WMS"),
    (corrigir_data,                   "Corrigir data no WMS"),
    (validar_av_local,                ROTULO_VALIDAR_AVARIA),
    (validar_fisico,                  ROTULO_VALIDAR_EXISTE),
    (validar_avaria & ~mover_b148,    ROTULO_VALIDAR_AVARIA),     # em local de avaria o "Mover no SAP para B148" substitui esta validação
    (ativar_e_lote,                   "Ativar local e desbloquear lote no WMS"),
    (ativar_local,                    "Ativar local no WMS"),
    (corrigir_setor,                  ROTULO_SETOR),
    (b135_errado,                     ROTULO_B135),
    (desbloquear,                     "Desbloquear lote no WMS"),
    # 1b) sistêmico no SAP
    (pend_inbound,                    "Pendente de internalização inbound"),
    (pend_devol,                      "Pendente de internalização de devolução"),
    (mover_b148,                      ROTULO_AV_B148),
    (mover_iqc_av,                    "Mover no SAP para o depósito do IQC"),
    (mover_iqc,                       "Mover no SAP para o depósito do IQC"),
    (falta_fora_b111,                 "Mover no SAP para B111"),
    (falta_b111,                      "Falta - desconsiderar"),
    (so_sap & (dep_sap == "B145"),    "Extravio no SAP - desconsiderar"),
    (so_sap,                          "Serial existe apenas no SAP"),
    # 2) desmembrar (o multi data vem antes do desmembrar físico)
    (multi_data,                      ROTULO_DESM_DATA),
    (desm_fisico,                     ROTULO_DESM_FISICO),
    # 3) mudar o pallet de local
    (b140_errado,                     ROTULO_B140),
    (produto_em_pecas,                "Produto em rua de peças"),
    (mover_wms_av,                    ROTULO_AV_WMS),
    (mover_rua,                       ROTULO_RUA),
    (avancar,                         "Mover no WMS e no SAP para o depósito do IQC"),
    (mover_wms,                       "Mover no WMS para o depósito do IQC"),
    # 4) demais
    (ok,                              "OK"),
    (ok_parcial,                      "OK - parcial local"),
    (excecao,                         "OK - exceção status mais grave que o IQC"),
    (corrigir_setor_av,               ROTULO_SETOR_AV),
]
condicoes = [c for c, _ in regras]
valores   = [v for _, v in regras]

df_analise["Analise"] = np.select(condicoes, valores, default="Em análise")

# Stage: o que sobraria em "Em análise" nos locais de stage vira "Em stage"
df_analise.loc[(df_analise["Analise"] == "Em análise") & em_stage, "Analise"] = ROTULO_STAGE

# ---------- Colunas derivadas da Analise ----------
mapa_acao = {
    "OK":                                           ("Nenhum",    "Nenhum"),
    "OK - parcial local":                           ("Nenhum",    "Nenhum"),
    "OK - exceção status mais grave que o IQC":     ("Nenhum",    "Nenhum"),
    "Falta - desconsiderar":                        ("Nenhum",    "Nenhum"),
    "Extravio no SAP - desconsiderar":              ("SAP",       "Série"),
    "Mover no SAP para B111":                       ("SAP",       "Série"),
    "Mover no SAP para o depósito do IQC":          ("SAP",       "Série"),
    "Pendente de internalização inbound":           ("SAP",       "Série"),
    "Pendente de internalização de devolução":      ("SAP",       "Série"),
    "Mover no WMS para o depósito do IQC":          ("WMS",       "Local"),
    "Mover no WMS e no SAP para o depósito do IQC": ("WMS e SAP", "Local"),
    "Ativar local no WMS":                          ("WMS",       "Local"),
    "Produto em rua de peças":                      ("WMS",       "Local"),
    "Desbloquear lote no WMS":                      ("WMS",       "Série"),
    "Ativar local e desbloquear lote no WMS":       ("WMS",       "Local e série"),
    "Serial duplicado no WMS":                      ("WMS",       "Série"),
    "Preencher fabricação no WMS":                  ("WMS",       "Série"),
    "Corrigir data no WMS":                         ("WMS",       "Série"),
    ROTULO_B140:                                    ("WMS",       "Série"),
    ROTULO_VALIDAR_EXISTE:                          ("Físico",    "Série"),
    ROTULO_VALIDAR_AVARIA:                          ("Físico",    "Série"),
    ROTULO_B135:                                    ("WMS",       "Local"),
    ROTULO_SETOR:                                   ("WMS",       "Local"),
    ROTULO_RUA:                                     ("WMS",       "Local"),
    ROTULO_AV_B148:                                 ("SAP",       "Série"),
    ROTULO_AV_WMS:                                  ("WMS",       "Série"),
    ROTULO_SETOR_AV:                                ("WMS",       "Série"),
    ROTULO_DESM_FISICO:                             ("Físico",    "Local"),
    ROTULO_DESM_DATA:                               ("Físico",    "Local"),
    ROTULO_STAGE:                                   ("A definir", "A definir"),
    "Material divergente entre SAP e WMS":          ("A definir", "A definir"),
    "Serial existe apenas no WMS":                  ("A definir", "A definir"),
    "Serial existe apenas no SAP":                  ("A definir", "Série"),
    "Em análise":                                   ("A definir", "A definir"),
}
tabela_acao = pd.DataFrame.from_dict(mapa_acao, orient="index", columns=["Onde agir", "Nível"])
df_analise["Onde agir"] = df_analise["Analise"].map(tabela_acao["Onde agir"])
df_analise["Nível"]     = df_analise["Analise"].map(tabela_acao["Nível"])

# Em local de avaria, os movimentos no WMS agem só na série (e não no local inteiro)
mov_wms_local = df_analise["Analise"].isin([
    "Mover no WMS e no SAP para o depósito do IQC",
    "Mover no WMS para o depósito do IQC",
])
df_analise.loc[mov_wms_local & em_avaria, "Nível"] = "Série"

# Motivo da análise: lista as regras que falharam, só para quem ficou em "Em análise"
em_analise = df_analise["Analise"] == "Em análise"
nao_c = ~serie_C

flags = pd.DataFrame({
    "Data WMS vazia (SAP em B145)":                 nos_dois & data_sap_raw.notna() & data_wms_raw.isna() & (dep_sap == "B145"),
    "Data SAP ≠ Data WMS":                          nos_dois & data_sap_raw.notna() & data_wms_raw.notna() & ~data_ok,
    "B135 só no WMS (SAP em outro depósito)":       b135_so_wms,
    "B140 só no WMS (SAP em outro depósito)":       b140_so_wms,
    "Rua fora do mapa":                             nos_dois & centro1.isna() & ~em_avaria,
    "Centro da rua ≠ Centro WMS":                   nos_dois & ~centro_ok,
    "Material vazio no SAP ou no WMS":              nos_dois & (mat_sap.isna() | mat_wms.isna()),
    "Centro SAP ≠ Centro da rua":                   nos_dois & ~centro_sap_ok,
    "Local com setor SAP misto":                    nos_dois & ~local_setor_sap_unico,
    "Centro SAP ≠ Centro WMS":                      nos_dois & (centro_sap != centro_wms),
    "Depósito SAP ≠ Depósito WMS":                  nos_dois & (dep_sap != dep_wms),
    "IQC ≠ depósito WMS":                           nos_dois & ~iqc_ok,
    "Local com série em bloqueio":                  nos_dois & nao_c & local_tem_C,
    "Local mistura mover no WMS e só no SAP":       nos_dois & nao_c & local_tem_M & local_tem_B,
    "Local com IQC diferentes a mover":             nos_dois & nao_c & local_tem_M & ~alvo_unico,
    "Local com vizinha com IQC próprio":            nos_dois & nao_c & local_tem_M & conflito_vizinha,
})

df_analise["Motivo da análise"] = ""
if em_analise.any():
    texto = flags[em_analise].apply(lambda r: " + ".join(r.index[r.values]), axis=1)
    df_analise.loc[em_analise, "Motivo da análise"] = texto.replace("", "Outro motivo")


# %% Exporta para Excel: aba "Analise" (df_analise inteiro) e aba "Resumo" (contagem por status)
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FONTE = "Arial"
COR_CABECALHO = "1F3864"   # azul escuro


def montar_resumo(df):
    """Uma linha por status, com a ação e o nível correspondentes."""
    def juntar(s):
        return " / ".join(pd.unique(s.dropna()))

    resumo = (
        df.groupby("Analise")
        .agg(Qtd=("Serial", "size"), Onde_agir=("Onde agir", juntar), Nivel=("Nível", juntar))
        .reset_index()
        .sort_values(["Qtd", "Analise"], ascending=[False, True])
        .reset_index(drop=True)
    )
    return resumo


def exportar_excel(df, caminho):
    resumo = montar_resumo(df)
    n_linhas = len(df)
    col_analise = get_column_letter(df.columns.get_loc("Analise") + 1)

    with pd.ExcelWriter(caminho, engine="openpyxl", datetime_format="DD/MM/YYYY", date_format="DD/MM/YYYY") as writer:
        # ----- aba Analise -----
        df.to_excel(writer, sheet_name="Analise", index=False)
        wb = writer.book
        wb._named_styles["Normal"].font = Font(name=FONTE, size=10)   # fonte padrão do arquivo

        ws = wb["Analise"]
        for c in ws[1]:
            c.font = Font(name=FONTE, size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", start_color=COR_CABECALHO)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = "B2"                       # cabeçalho e Serial sempre visíveis
        ws.auto_filter.ref = ws.dimensions           # filtro em todas as colunas

        amostra = df.head(2000)
        for i, nome in enumerate(df.columns, start=1):
            maior = max([len(str(nome))] + [len(str(v)) for v in amostra[nome].dropna().unique()[:300]])
            ws.column_dimensions[get_column_letter(i)].width = min(max(maior + 3, 12), 48)

        # ----- aba Resumo -----
        rs = wb.create_sheet("Resumo")
        rs["A1"] = "Resumo da análise - batimento de estoque SAP x WMS"
        rs["A1"].font = Font(name=FONTE, size=12, bold=True)
        rs["A2"] = f"Gerado em {datetime.now():%d/%m/%Y %H:%M}. A coluna Qtd conta as linhas da aba Analise."
        rs["A2"].font = Font(name=FONTE, size=9, italic=True, color="595959")

        cab_row = 4
        for j, titulo in enumerate(["Analise", "Qtd", "% do total", "Onde agir", "Nível"], start=1):
            c = rs.cell(cab_row, j, titulo)
            c.font = Font(name=FONTE, size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", start_color=COR_CABECALHO)
            c.alignment = Alignment(horizontal="center", vertical="center")

        primeira = cab_row + 1
        ultima = primeira + len(resumo) - 1
        total = ultima + 1
        faixa = f"Analise!${col_analise}$2:${col_analise}${n_linhas + 1}"

        for i, linha in resumo.iterrows():
            r = primeira + i
            rs.cell(r, 1, linha["Analise"])
            rs.cell(r, 2, f"=COUNTIF({faixa},A{r})").number_format = "#,##0"
            rs.cell(r, 3, f"=B{r}/$B${total}").number_format = "0.0%"
            rs.cell(r, 4, linha["Onde_agir"])
            rs.cell(r, 5, linha["Nivel"])

        rs.cell(total, 1, "Total")
        rs.cell(total, 2, f"=SUM(B{primeira}:B{ultima})").number_format = "#,##0"
        rs.cell(total, 3, f"=SUM(C{primeira}:C{ultima})").number_format = "0.0%"
        for j in range(1, 6):
            rs.cell(total, j).font = Font(name=FONTE, size=10, bold=True)
            rs.cell(total, j).fill = PatternFill("solid", start_color="D9E1F2")

        for letra, largura in zip("ABCDE", [48, 12, 12, 16, 16]):
            rs.column_dimensions[letra].width = largura
        rs.freeze_panes = f"A{primeira}"

        wb.active = 0                                # abre na aba Analise


exportar_excel(df_analise, CAMINHO_EXCEL)

# %% Resumo no console
print(df_analise["Analise"].value_counts().to_string())
print("Total de linhas:", len(df_analise))
print("Arquivo gerado:", CAMINHO_EXCEL)
