# -*- coding: utf-8 -*-
"""
Resposta 03 - Projeção de indicadores EMBRAPII para 2027
Autor: Kennedy Anderson Nascimento Santos

Objetivo do script
------------------
Este script lê a base histórica disponibilizada na Questão 03 e estima,
para o ano de 2027, os três indicadores solicitados no enunciado:

1. Número de projetos contratados no ano;
2. Valor total dos projetos contratados no ano;
3. Número de projetos concluídos no ano.

Além de calcular as projeções, o script também gera todos os arquivos de apoio
utilizados no relatório técnico: tabelas em CSV e imagens dos gráficos/tabelas.
Ao final da execução, os gráficos são exibidos na tela para conferência visual.

Como executar
-------------
1. Coloque este script na mesma pasta da planilha:
   Embrapii_seleção_analista_2026_questao03_Estimativa.xlsx

2. Instale as dependências, caso necessário:
   pip install pandas numpy matplotlib openpyxl

3. Execute no terminal:
   python kennedy_anderson_resposta03_projecao_2027.py

Observação:
- Por padrão, o script cria uma pasta chamada "outputs_questao03".
- Para rodar sem abrir os gráficos na tela, use:
  python kennedy_anderson_resposta03_projecao_2027.py --no-show
"""



from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# =============================================================================
# 1. PARÂMETROS GERAIS DA ANÁLISE
# =============================================================================

ANO_PREVISAO_PADRAO = 2027
N_BOOTSTRAP_PADRAO = 5000
SEMENTE_ALEATORIA_PADRAO = 123

# Nome esperado da planilha entregue no enunciado da questão.
# O script também possui uma busca alternativa por padrão de nome, caso o arquivo
# seja salvo com pequena diferença de acentuação no sistema operacional.
NOME_ARQUIVO_PADRAO = "Embrapii_selecao_analista_2026_questao03_Estimativa.xlsx"

# Colunas obrigatórias da aba "dados".
COLUNAS_OBRIGATORIAS = [
    "ano",
    "mes",
    "novos_projetos_contratados",
    "valor_projetos_contratados",
    "projetos_concluidos",
]

# Configuração dos indicadores para evitar repetição de código.
# Cada indicador recebe um rótulo amigável e uma unidade de apresentação.
INDICADORES = {
    "novos_projetos_contratados": {
        "rotulo": "Projetos contratados no ano",
        "rotulo_curto": "Projetos contratados",
        "unidade": "projetos",
    },
    "valor_projetos_contratados": {
        "rotulo": "Valor contratado no ano",
        "rotulo_curto": "Valor contratado",
        "unidade": "R$",
    },
    "projetos_concluidos": {
        "rotulo": "Projetos concluídos no ano",
        "rotulo_curto": "Projetos concluídos",
        "unidade": "projetos",
    },
}


# =============================================================================
# 2. ESTRUTURAS DE RESULTADO
# =============================================================================

@dataclass
class ResultadoIndicador:
    """Armazena o resultado anual e mensal de um indicador projetado."""

    indicador: str
    estimativa_anual: float
    limite_inferior_95: float
    limite_superior_95: float
    previsao_mensal: pd.DataFrame
    valores_ajustados: pd.DataFrame
    erro_medio_absoluto: float


# =============================================================================
# 3. FUNÇÕES DE ENTRADA, VALIDAÇÃO E PREPARAÇÃO DOS DADOS
# =============================================================================


def localizar_arquivo_entrada(caminho_informado: str | None = None) -> Path:
    """
    Localiza a planilha de entrada.

    A função prioriza o caminho informado pelo usuário. Se nenhum caminho for
    informado, procura a planilha na mesma pasta do script. Caso o nome exato não
    seja encontrado, tenta localizar um arquivo Excel com termos compatíveis.
    """

    pasta_script = Path(__file__).resolve().parent

    if caminho_informado:
        caminho = Path(caminho_informado).expanduser().resolve()
        if not caminho.exists():
            raise FileNotFoundError(f"Arquivo informado não encontrado: {caminho}")
        return caminho

    caminho_padrao = pasta_script / NOME_ARQUIVO_PADRAO
    if caminho_padrao.exists():
        return caminho_padrao

    # Busca alternativa para reduzir risco de erro por acentuação/digitação.
    candidatos = list(pasta_script.glob("*Embrapii*questao03*Estimativa*.xlsx"))
    candidatos += list(pasta_script.glob("*Embrapii*questão03*Estimativa*.xlsx"))
    if candidatos:
        return candidatos[0]

    raise FileNotFoundError(
        "Planilha de entrada não encontrada. Coloque o arquivo Excel na mesma "
        "pasta do script ou informe o caminho com o argumento --input."
    )


def carregar_dados(caminho_excel: Path) -> pd.DataFrame:
    """
    Lê a aba 'dados' da planilha e valida se as colunas necessárias existem.

    O enunciado da questão pede reprodutibilidade. Por isso, a validação é feita
    logo no início: se a estrutura do arquivo mudar, o erro fica claro para quem
    for executar o script.
    """

    df = pd.read_excel(caminho_excel, sheet_name="dados")

    colunas_faltantes = [col for col in COLUNAS_OBRIGATORIAS if col not in df.columns]
    if colunas_faltantes:
        raise ValueError(
            "A planilha não possui todas as colunas obrigatórias. "
            f"Colunas faltantes: {colunas_faltantes}"
        )

    # Mantém somente as colunas necessárias e garante ordenação cronológica.
    df = df[COLUNAS_OBRIGATORIAS].copy()
    df["ano"] = df["ano"].astype(int)
    df["mes"] = df["mes"].astype(int)
    df["data"] = pd.to_datetime(dict(year=df["ano"], month=df["mes"], day=1))
    df = df.sort_values("data").reset_index(drop=True)

    # Índice temporal mensal. Ex.: janeiro/2015 = 0, fevereiro/2015 = 1, etc.
    # Ele permite ao modelo capturar tendência de crescimento ou queda no tempo.
    df["t"] = np.arange(len(df), dtype=float)

    return df


def criar_pastas_saida(output_dir: Path) -> Dict[str, Path]:
    """Cria a estrutura de pastas de saída para tabelas e visuais."""

    caminhos = {
        "raiz": output_dir,
        "tabelas": output_dir / "tabelas",
        "visuais": output_dir / "visuais",
    }

    for caminho in caminhos.values():
        caminho.mkdir(parents=True, exist_ok=True)

    return caminhos


# =============================================================================
# 4. MODELO ESTATÍSTICO
# =============================================================================


def montar_matriz_modelo(df: pd.DataFrame, colunas_dummies: Iterable[str] | None = None) -> Tuple[np.ndarray, List[str]]:
    """
    Monta a matriz de regressão usada pelo modelo.

    Técnica escolhida:
    - Intercepto: nível médio da série;
    - Tendência temporal: variação ao longo dos meses;
    - Dummies mensais: efeito típico de cada mês, capturando sazonalidade.

    A dummy de janeiro é omitida por ser a categoria de referência.
    """

    dummies_mes = pd.get_dummies(df["mes"].astype(int), prefix="m", drop_first=True, dtype=float)

    if colunas_dummies is not None:
        # Garante que a base futura tenha exatamente as mesmas colunas da base
        # histórica, mesmo que algum mês não apareça por algum motivo.
        for coluna in colunas_dummies:
            if coluna not in dummies_mes.columns:
                dummies_mes[coluna] = 0.0
        dummies_mes = dummies_mes[list(colunas_dummies)]

    nomes_colunas = ["intercepto", "tendencia"] + list(dummies_mes.columns)
    matriz_x = np.column_stack(
        [
            np.ones(len(df), dtype=float),
            df["t"].to_numpy(dtype=float),
            dummies_mes.to_numpy(dtype=float),
        ]
    )

    return matriz_x, nomes_colunas


def projetar_indicador(
    df: pd.DataFrame,
    indicador: str,
    ano_previsao: int,
    n_bootstrap: int,
    seed: int,
) -> ResultadoIndicador:
    """
    Ajusta o modelo e projeta os 12 meses do ano desejado.

    Para estimar a incerteza, é aplicado bootstrap dos resíduos:
    1. calcula-se o erro histórico do modelo;
    2. sorteiam-se erros observados no passado com reposição;
    3. esses erros são somados à previsão mensal;
    4. a soma anual é repetida várias vezes;
    5. os percentis 2,5% e 97,5% formam a faixa de 95%.
    """

    if indicador not in INDICADORES:
        raise ValueError(f"Indicador não mapeado: {indicador}")

    x_hist, nomes_x = montar_matriz_modelo(df)
    y_hist = df[indicador].to_numpy(dtype=float)

    # Ajuste por mínimos quadrados ordinários.
    # A função lstsq estima os coeficientes que minimizam o erro quadrático.
    beta, *_ = np.linalg.lstsq(x_hist, y_hist, rcond=None)

    y_ajustado = x_hist @ beta
    residuos = y_hist - y_ajustado
    erro_medio_absoluto = float(np.mean(np.abs(residuos)))

    # Base futura: 12 meses do ano de previsão.
    meses_futuros = pd.date_range(f"{ano_previsao}-01-01", f"{ano_previsao}-12-01", freq="MS")
    futuro = pd.DataFrame({"data": meses_futuros})
    futuro["ano"] = futuro["data"].dt.year
    futuro["mes"] = futuro["data"].dt.month

    data_inicial = df["data"].min()
    futuro["t"] = (
        (futuro["data"].dt.year - data_inicial.year) * 12
        + (futuro["data"].dt.month - data_inicial.month)
    ).astype(float)

    colunas_dummies = nomes_x[2:]
    x_futuro, _ = montar_matriz_modelo(futuro, colunas_dummies=colunas_dummies)

    previsao_mensal = x_futuro @ beta

    # Como os três indicadores não podem ser negativos, aplica-se piso zero.
    previsao_mensal = np.maximum(previsao_mensal, 0)

    # Bootstrap dos resíduos para estimar faixa anual de incerteza.
    rng = np.random.default_rng(seed)
    amostras_anuais = np.empty(n_bootstrap, dtype=float)

    for i in range(n_bootstrap):
        residuos_sorteados = rng.choice(residuos, size=len(futuro), replace=True)
        previsao_com_ruido = np.maximum(previsao_mensal + residuos_sorteados, 0)
        amostras_anuais[i] = previsao_com_ruido.sum()

    resultado_mensal = futuro[["ano", "mes", "data"]].copy()
    resultado_mensal["indicador"] = indicador
    resultado_mensal["valor_previsto"] = previsao_mensal

    valores_ajustados = df[["ano", "mes", "data"]].copy()
    valores_ajustados["indicador"] = indicador
    valores_ajustados["valor_observado"] = y_hist
    valores_ajustados["valor_ajustado"] = y_ajustado
    valores_ajustados["residuo"] = residuos

    return ResultadoIndicador(
        indicador=indicador,
        estimativa_anual=float(previsao_mensal.sum()),
        limite_inferior_95=float(np.quantile(amostras_anuais, 0.025)),
        limite_superior_95=float(np.quantile(amostras_anuais, 0.975)),
        previsao_mensal=resultado_mensal,
        valores_ajustados=valores_ajustados,
        erro_medio_absoluto=erro_medio_absoluto,
    )


# =============================================================================
# 5. FORMATAÇÃO DOS RESULTADOS
# =============================================================================


def formatar_numero_ptbr(valor: float, casas_decimais: int = 0) -> str:
    """Formata números no padrão brasileiro."""

    texto = f"{valor:,.{casas_decimais}f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_moeda_ptbr(valor: float) -> str:
    """Formata valor monetário em reais, sem casas decimais para leitura executiva."""

    return f"R$ {formatar_numero_ptbr(valor, 0)}"


def formatar_valor_indicador(indicador: str, valor: float) -> str:
    """Aplica a formatação correta conforme o tipo de indicador."""

    if indicador == "valor_projetos_contratados":
        return formatar_moeda_ptbr(valor)
    return formatar_numero_ptbr(valor, 0)


def criar_tabela_resultados(resultados: List[ResultadoIndicador]) -> pd.DataFrame:
    """Consolida as projeções anuais em uma tabela de resultado."""

    linhas = []
    for resultado in resultados:
        linhas.append(
            {
                "indicador": resultado.indicador,
                "Indicador": INDICADORES[resultado.indicador]["rotulo"],
                "Estimativa 2027": formatar_valor_indicador(resultado.indicador, resultado.estimativa_anual),
                "Limite inferior 95%": formatar_valor_indicador(resultado.indicador, resultado.limite_inferior_95),
                "Limite superior 95%": formatar_valor_indicador(resultado.indicador, resultado.limite_superior_95),
                "estimativa_num": resultado.estimativa_anual,
                "limite_inferior_num": resultado.limite_inferior_95,
                "limite_superior_num": resultado.limite_superior_95,
                "erro_medio_absoluto": resultado.erro_medio_absoluto,
            }
        )

    return pd.DataFrame(linhas)


def criar_tabela_previsao_mensal(resultados: List[ResultadoIndicador]) -> pd.DataFrame:
    """Consolida as previsões mensais dos três indicadores."""

    mensal = pd.concat([resultado.previsao_mensal for resultado in resultados], ignore_index=True)

    tabela = mensal.pivot_table(
        index=["ano", "mes", "data"],
        columns="indicador",
        values="valor_previsto",
        aggfunc="sum",
    ).reset_index()

    # Reordena colunas no mesmo padrão do enunciado.
    tabela = tabela[["ano", "mes", "data"] + list(INDICADORES.keys())]
    return tabela


def criar_historico_anual(df: pd.DataFrame) -> pd.DataFrame:
    """Cria o histórico anual e marca se o ano está completo ou parcial."""

    historico = (
        df.groupby("ano", as_index=False)
        .agg(
            meses_observados=("mes", "nunique"),
            novos_projetos_contratados=("novos_projetos_contratados", "sum"),
            valor_projetos_contratados=("valor_projetos_contratados", "sum"),
            projetos_concluidos=("projetos_concluidos", "sum"),
        )
        .sort_values("ano")
    )
    historico["tipo_ano"] = np.where(historico["meses_observados"] == 12, "Ano completo", "Ano parcial")
    return historico


# =============================================================================
# 6. GERAÇÃO DAS TABELAS E GRÁFICOS USADOS NA DOCUMENTAÇÃO
# =============================================================================


def salvar_tabelas(
    df_original: pd.DataFrame,
    tabela_resultados: pd.DataFrame,
    tabela_mensal: pd.DataFrame,
    historico_anual: pd.DataFrame,
    caminhos: Dict[str, Path],
) -> None:
    """
    Salva as tabelas da análise em CSV com separador ';'.

    O separador ';' facilita a abertura em ambientes configurados em português,
    especialmente no Excel/Power BI em padrão brasileiro.
    """

    tabela_resultados.to_csv(
        caminhos["tabelas"] / "resultado_projecao_2027.csv",
        index=False,
        sep=";",
        decimal=",",
        encoding="utf-8-sig",
    )

    tabela_mensal.to_csv(
        caminhos["tabelas"] / "previsao_mensal_2027.csv",
        index=False,
        sep=";",
        decimal=",",
        encoding="utf-8-sig",
    )

    historico_anual.to_csv(
        caminhos["tabelas"] / "historico_anual_observado.csv",
        index=False,
        sep=";",
        decimal=",",
        encoding="utf-8-sig",
    )

    resumo_base = pd.DataFrame(
        [
            {"Item": "Período inicial da base", "Valor": df_original["data"].min().strftime("%Y/%m")},
            {"Item": "Período final da base", "Valor": df_original["data"].max().strftime("%Y/%m")},
            {"Item": "Registros mensais", "Valor": str(len(df_original))},
            {"Item": "Anos completos observados", "Valor": str((historico_anual["meses_observados"] == 12).sum())},
            {"Item": "Técnica", "Valor": "Regressão linear com tendência + sazonalidade mensal"},
            {"Item": "Incerteza", "Valor": "Bootstrap dos resíduos, faixa de 95%"},
        ]
    )

    resumo_base.to_csv(
        caminhos["tabelas"] / "resumo_base_metodo.csv",
        index=False,
        sep=";",
        encoding="utf-8-sig",
    )


def salvar_tabela_como_imagem(
    df_tabela: pd.DataFrame,
    caminho_saida: Path,
    titulo: str,
    colunas: List[str],
    largura: float = 12.0,
    altura: float = 2.8,
) -> plt.Figure:
    """Transforma uma tabela em imagem para uso direto no relatório."""

    fig, ax = plt.subplots(figsize=(largura, altura))
    ax.axis("off")

    # O título é desenhado na figura e a tabela recebe uma área fixa.
    # Isso evita muito espaço em branco quando a imagem é inserida no Word.
    fig.text(0.5, 0.92, titulo, ha="center", va="center", fontsize=14, fontweight="bold")

    tabela_plot = ax.table(
        cellText=df_tabela[colunas].values,
        colLabels=colunas,
        cellLoc="left",
        colLoc="left",
        bbox=[0.01, 0.05, 0.98, 0.72],
    )

    tabela_plot.auto_set_font_size(False)
    tabela_plot.set_fontsize(9)

    # Ajuste visual simples: cabeçalho destacado e bordas suaves.
    for (linha, coluna), celula in tabela_plot.get_celld().items():
        celula.set_edgecolor("#D0D7DE")
        if linha == 0:
            celula.set_facecolor("#1F4E79")
            celula.set_text_props(color="white", weight="bold")
        else:
            celula.set_facecolor("#F8FAFC" if linha % 2 == 0 else "white")

    fig.savefig(caminho_saida, dpi=200, bbox_inches="tight")
    return fig


def grafico_historico_anual_projecao(
    historico_anual: pd.DataFrame,
    tabela_resultados: pd.DataFrame,
    ano_previsao: int,
    caminho_saida: Path,
) -> plt.Figure:
    """Gera gráfico anual com histórico observado e projeção de 2027."""

    fig, ax1 = plt.subplots(figsize=(12, 6))
    ax2 = ax1.twinx()

    completo = historico_anual[historico_anual["tipo_ano"] == "Ano completo"]
    parcial = historico_anual[historico_anual["tipo_ano"] == "Ano parcial"]

    # Linhas de quantidade de projetos.
    ax1.plot(
        completo["ano"],
        completo["novos_projetos_contratados"],
        marker="o",
        linewidth=2,
        label="Projetos contratados - observado",
    )
    ax1.plot(
        completo["ano"],
        completo["projetos_concluidos"],
        marker="o",
        linewidth=2,
        label="Projetos concluídos - observado",
    )

    # Valor contratado em bilhões para facilitar leitura.
    ax2.plot(
        completo["ano"],
        completo["valor_projetos_contratados"] / 1e9,
        marker="o",
        linestyle="--",
        linewidth=2,
        label="Valor contratado - observado (R$ bi)",
    )

    # Ano parcial, quando existir, aparece como ponto separado e transparente.
    if not parcial.empty:
        ax1.scatter(
            parcial["ano"],
            parcial["novos_projetos_contratados"],
            marker="x",
            s=80,
            label="Projetos contratados - ano parcial",
        )
        ax1.scatter(
            parcial["ano"],
            parcial["projetos_concluidos"],
            marker="x",
            s=80,
            label="Projetos concluídos - ano parcial",
        )
        ax2.scatter(
            parcial["ano"],
            parcial["valor_projetos_contratados"] / 1e9,
            marker="x",
            s=80,
            label="Valor contratado - ano parcial (R$ bi)",
        )

    def linha_resultado(indicador: str) -> pd.Series:
        return tabela_resultados.loc[tabela_resultados["indicador"] == indicador].iloc[0]

    r_contratados = linha_resultado("novos_projetos_contratados")
    r_valor = linha_resultado("valor_projetos_contratados")
    r_concluidos = linha_resultado("projetos_concluidos")

    # Projeções com intervalo de 95%.
    ax1.errorbar(
        [ano_previsao],
        [r_contratados["estimativa_num"]],
        yerr=[[
            r_contratados["estimativa_num"] - r_contratados["limite_inferior_num"]
        ], [
            r_contratados["limite_superior_num"] - r_contratados["estimativa_num"]
        ]],
        fmt="o",
        capsize=5,
        markersize=8,
        label="Projetos contratados - projeção 2027",
    )
    ax1.errorbar(
        [ano_previsao],
        [r_concluidos["estimativa_num"]],
        yerr=[[
            r_concluidos["estimativa_num"] - r_concluidos["limite_inferior_num"]
        ], [
            r_concluidos["limite_superior_num"] - r_concluidos["estimativa_num"]
        ]],
        fmt="s",
        capsize=5,
        markersize=8,
        label="Projetos concluídos - projeção 2027",
    )
    ax2.errorbar(
        [ano_previsao],
        [r_valor["estimativa_num"] / 1e9],
        yerr=[[
            (r_valor["estimativa_num"] - r_valor["limite_inferior_num"]) / 1e9
        ], [
            (r_valor["limite_superior_num"] - r_valor["estimativa_num"]) / 1e9
        ]],
        fmt="D",
        capsize=5,
        markersize=7,
        label="Valor contratado - projeção 2027 (R$ bi)",
    )

    ax1.set_title("Histórico anual observado e projeção para 2027", fontsize=15, fontweight="bold")
    ax1.set_xlabel("Ano")
    ax1.set_ylabel("Quantidade de projetos")
    ax2.set_ylabel("Valor contratado (R$ bilhões)")
    ax1.grid(True, linestyle="--", alpha=0.35)

    # Une legendas dos dois eixos.
    linhas1, labels1 = ax1.get_legend_handles_labels()
    linhas2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(linhas1 + linhas2, labels1 + labels2, loc="upper left", fontsize=8)

    fig.tight_layout()
    fig.savefig(caminho_saida, dpi=200, bbox_inches="tight")
    return fig


def grafico_intervalos_confianca(
    tabela_resultados: pd.DataFrame,
    caminho_saida: Path,
) -> plt.Figure:
    """Gera gráfico simples com estimativa e faixa de 95% de cada indicador."""

    fig, axes = plt.subplots(nrows=3, ncols=1, figsize=(10, 7))

    for ax, indicador in zip(axes, INDICADORES.keys()):
        linha = tabela_resultados.loc[tabela_resultados["indicador"] == indicador].iloc[0]
        estimativa = linha["estimativa_num"]
        limite_inf = linha["limite_inferior_num"]
        limite_sup = linha["limite_superior_num"]

        ax.errorbar(
            estimativa,
            0,
            xerr=[[estimativa - limite_inf], [limite_sup - estimativa]],
            fmt="o",
            capsize=5,
            markersize=8,
        )
        ax.axvline(estimativa, linestyle="--", alpha=0.45)
        ax.set_yticks([])
        ax.set_title(INDICADORES[indicador]["rotulo"], fontsize=11, loc="left")
        ax.grid(True, axis="x", linestyle="--", alpha=0.25)

        if indicador == "valor_projetos_contratados":
            ax.set_xlabel("R$ bilhões")
            ax.set_xlim(max(0, limite_inf * 0.85) / 1e9, limite_sup * 1.1 / 1e9)
            ax.clear()
            ax.errorbar(
                estimativa / 1e9,
                0,
                xerr=[[(estimativa - limite_inf) / 1e9], [(limite_sup - estimativa) / 1e9]],
                fmt="o",
                capsize=5,
                markersize=8,
            )
            ax.axvline(estimativa / 1e9, linestyle="--", alpha=0.45)
            ax.set_yticks([])
            ax.set_title(INDICADORES[indicador]["rotulo"], fontsize=11, loc="left")
            ax.set_xlabel("R$ bilhões")
            ax.grid(True, axis="x", linestyle="--", alpha=0.25)
        else:
            ax.set_xlabel("Quantidade de projetos")

    fig.suptitle("Estimativa 2027 e faixa de 95%", fontsize=15, fontweight="bold", y=0.98)
    fig.tight_layout(rect=[0, 0, 1, 0.96])
    fig.savefig(caminho_saida, dpi=200, bbox_inches="tight")
    return fig


def grafico_distribuicao_mensal(
    tabela_mensal: pd.DataFrame,
    caminho_saida: Path,
) -> plt.Figure:
    """Gera a distribuição mensal prevista para 2027."""

    meses = tabela_mensal["data"].dt.strftime("%b/%y")
    fig, axes = plt.subplots(nrows=3, ncols=1, figsize=(12, 8), sharex=True)

    for ax, indicador in zip(axes, INDICADORES.keys()):
        valores = tabela_mensal[indicador].copy()
        ylabel = "Projetos" if indicador != "valor_projetos_contratados" else "R$ milhões"
        if indicador == "valor_projetos_contratados":
            valores = valores / 1e6

        ax.bar(meses, valores)
        ax.set_title(INDICADORES[indicador]["rotulo_curto"], fontsize=11, loc="left")
        ax.set_ylabel(ylabel)
        ax.grid(True, axis="y", linestyle="--", alpha=0.25)

    axes[-1].tick_params(axis="x", rotation=45)
    fig.suptitle("Distribuição mensal prevista para 2027", fontsize=15, fontweight="bold", y=0.98)
    fig.tight_layout(rect=[0, 0, 1, 0.96])
    fig.savefig(caminho_saida, dpi=200, bbox_inches="tight")
    return fig


# =============================================================================
# 7. ORQUESTRAÇÃO DA EXECUÇÃO
# =============================================================================


def executar_analise(
    caminho_excel: Path,
    output_dir: Path,
    ano_previsao: int,
    n_bootstrap: int,
    seed: int,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, List[plt.Figure]]:
    """Executa todas as etapas da análise e retorna as tabelas principais."""

    caminhos = criar_pastas_saida(output_dir)
    df = carregar_dados(caminho_excel)

    resultados = [
        projetar_indicador(
            df=df,
            indicador=indicador,
            ano_previsao=ano_previsao,
            n_bootstrap=n_bootstrap,
            seed=seed,
        )
        for indicador in INDICADORES.keys()
    ]

    tabela_resultados = criar_tabela_resultados(resultados)
    tabela_mensal = criar_tabela_previsao_mensal(resultados)
    historico_anual = criar_historico_anual(df)

    salvar_tabelas(df, tabela_resultados, tabela_mensal, historico_anual, caminhos)

    figuras = []

    figuras.append(
        salvar_tabela_como_imagem(
            df_tabela=tabela_resultados,
            caminho_saida=caminhos["visuais"] / "tabela_resultados_2027.png",
            titulo="Resultados estimados para 2027",
            colunas=["Indicador", "Estimativa 2027", "Limite inferior 95%", "Limite superior 95%"],
            largura=12,
            altura=3.0,
        )
    )

    resumo_base = pd.read_csv(caminhos["tabelas"] / "resumo_base_metodo.csv", sep=";")
    figuras.append(
        salvar_tabela_como_imagem(
            df_tabela=resumo_base,
            caminho_saida=caminhos["visuais"] / "tabela_resumo_base_metodo.png",
            titulo="Resumo técnico da base e do método",
            colunas=["Item", "Valor"],
            largura=10,
            altura=3.5,
        )
    )

    figuras.append(
        grafico_historico_anual_projecao(
            historico_anual=historico_anual,
            tabela_resultados=tabela_resultados,
            ano_previsao=ano_previsao,
            caminho_saida=caminhos["visuais"] / "grafico_historico_anual_projecao_2027.png",
        )
    )

    figuras.append(
        grafico_intervalos_confianca(
            tabela_resultados=tabela_resultados,
            caminho_saida=caminhos["visuais"] / "grafico_intervalos_confianca_2027.png",
        )
    )

    figuras.append(
        grafico_distribuicao_mensal(
            tabela_mensal=tabela_mensal,
            caminho_saida=caminhos["visuais"] / "grafico_distribuicao_mensal_2027.png",
        )
    )

    return tabela_resultados, tabela_mensal, historico_anual, figuras


def imprimir_resumo_execucao(
    tabela_resultados: pd.DataFrame,
    tabela_mensal: pd.DataFrame,
    historico_anual: pd.DataFrame,
    output_dir: Path,
) -> None:
    """Imprime no terminal o resumo final da execução."""

    print("\n" + "=" * 80)
    print("PROJEÇÃO DE INDICADORES EMBRAPII PARA 2027")
    print("=" * 80)
    print("\nResultados anuais estimados:\n")
    print(
        tabela_resultados[
            ["Indicador", "Estimativa 2027", "Limite inferior 95%", "Limite superior 95%"]
        ].to_string(index=False)
    )

    print("\nResumo da base histórica:\n")
    print(
        f"- Período: {historico_anual['ano'].min()} a {historico_anual['ano'].max()} "
        f"({int(historico_anual['meses_observados'].sum())} registros mensais)."
    )
    print(
        f"- Ano parcial identificado: "
        f"{', '.join(map(str, historico_anual.loc[historico_anual['tipo_ano'] == 'Ano parcial', 'ano'].tolist())) or 'não há'}"
    )

    print("\nArquivos gerados:\n")
    print(f"- Tabelas: {output_dir / 'tabelas'}")
    print(f"- Visuais: {output_dir / 'visuais'}")
    print("\nTabelas principais geradas:")
    print("- resultado_projecao_2027.csv")
    print("- previsao_mensal_2027.csv")
    print("- historico_anual_observado.csv")
    print("- resumo_base_metodo.csv")
    print("\nGráficos/tabelas visuais gerados para a documentação:")
    print("- tabela_resultados_2027.png")
    print("- tabela_resumo_base_metodo.png")
    print("- grafico_historico_anual_projecao_2027.png")
    print("- grafico_intervalos_confianca_2027.png")
    print("- grafico_distribuicao_mensal_2027.png")
    print("\nAmostra da previsão mensal de 2027:\n")
    print(tabela_mensal.head(12).to_string(index=False))
    print("=" * 80 + "\n")


def montar_parser_argumentos() -> argparse.ArgumentParser:
    """Cria os argumentos opcionais de linha de comando."""

    parser = argparse.ArgumentParser(
        description="Projeção dos indicadores da Questão 03 da EMBRAPII para 2027."
    )
    parser.add_argument(
        "--input",
        type=str,
        default=None,
        help="Caminho da planilha de entrada. Se omitido, o script procura na mesma pasta.",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default="outputs_questao03",
        help="Pasta onde tabelas e gráficos serão salvos.",
    )
    parser.add_argument(
        "--ano-previsao",
        type=int,
        default=ANO_PREVISAO_PADRAO,
        help="Ano que será projetado. Padrão: 2027.",
    )
    parser.add_argument(
        "--n-bootstrap",
        type=int,
        default=N_BOOTSTRAP_PADRAO,
        help="Número de reamostragens para estimar a faixa de 95%. Padrão: 5000.",
    )
    parser.add_argument(
        "--seed",
        type=int,
        default=SEMENTE_ALEATORIA_PADRAO,
        help="Semente aleatória para reprodutibilidade. Padrão: 123.",
    )
    parser.add_argument(
        "--no-show",
        action="store_true",
        help="Salva os visuais, mas não abre os gráficos na tela ao final.",
    )
    return parser


def main() -> None:
    """Função principal do script."""

    parser = montar_parser_argumentos()
    args = parser.parse_args()

    caminho_excel = localizar_arquivo_entrada(args.input)
    output_dir = Path(args.output_dir).expanduser().resolve()

    tabela_resultados, tabela_mensal, historico_anual, figuras = executar_analise(
        caminho_excel=caminho_excel,
        output_dir=output_dir,
        ano_previsao=args.ano_previsao,
        n_bootstrap=args.n_bootstrap,
        seed=args.seed,
    )

    imprimir_resumo_execucao(tabela_resultados, tabela_mensal, historico_anual, output_dir)

    # Exibe todos os gráficos/tabelas visuais que foram salvos e utilizados na documentação.
    # Em ambientes sem interface gráfica, utilize o argumento --no-show.
    if not args.no_show:
        plt.show()
    else:
        for figura in figuras:
            plt.close(figura)


if __name__ == "__main__":
    main()
