"""
GERADOR DE DATA BOOK DIGITAL - EXCEL + GITHUB PAGES

Este script lê dados_equipamentos.xlsx e cria:
- 1 página HTML por equipamento
- 1 página inicial com busca
- 1 QR Code por equipamento
- Registros repetidos por TAG são consolidados; linhas sem TAG usam a OS como identificador.

INSTALAÇÃO:
pip install pandas openpyxl qrcode pillow

USO:
1. Preencha equipcaldeira1-INPASA.xlsx
2. Altere BASE_URL abaixo
3. Rode: python gerar_site.py
4. Envie o conteúdo da pasta para o GitHub
"""

from pathlib import Path
import builtins as _builtins
import pandas as pd
import qrcode
import re
import html
import unicodedata

# Alguns diagnósticos do editor/linters podem perder a resolução dos builtins
# em projetos com configuração local incompleta. Mantemos aliases explícitos.
str = _builtins.str
FileNotFoundError = _builtins.FileNotFoundError
ValueError = _builtins.ValueError
print = _builtins.print

# ALTERAR AQUI depois que o GitHub Pages estiver criado.
# Exemplo: "https://goldredutoresinpasa-web.github.io/DATABOOK-CALDEIRA-1"
BASE_URL = "https://goldredutoresinpasa-web.github.io/DATABOOK-CALDEIRA-1"

BASE_DIR = Path(__file__).resolve().parent
ARQUIVO_EXCEL = BASE_DIR / "equipcaldeira1-INPASA.xlsx"
PASTA_EQUIPAMENTOS = BASE_DIR / "equipamentos"
PASTA_QRCODES = BASE_DIR / "assets" / "qrcodes"


def limpar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def primeiro_valor_nao_vazio(*valores):
    for valor in valores:
        texto = limpar_texto(valor)
        if texto:
            return texto
    return ""


def normalizar_nome_coluna(nome):
    nome = unicodedata.normalize("NFKD", str(nome))
    nome = nome.encode("ascii", "ignore").decode("ascii").casefold()
    return re.sub(r"[^a-z0-9]", "", nome)


def obter_observacoes(dados):
    for coluna, valor in dados.items():
        if normalizar_nome_coluna(coluna) == "observacoes":
            return limpar_texto(valor)
    return ""


def obter_numero_os(dados):
    observacoes = obter_observacoes(dados)
    correspondencia = re.search(r"\bOS\s*(\d+)\b", observacoes, re.IGNORECASE)
    return correspondencia.group(1) if correspondencia else ""


def obter_valor_coluna(dados, *nomes_colunas):
    nomes_alvos = {normalizar_nome_coluna(nome) for nome in nomes_colunas}

    for coluna, valor in dados.items():
        if normalizar_nome_coluna(coluna) in nomes_alvos:
            return limpar_texto(valor)

    for nome in nomes_colunas:
        if nome in dados:
            return limpar_texto(dados[nome])

    return ""


def preparar_equipamentos(df):
    registros = [row.to_dict() for _, row in df.iterrows()]
    grupos_por_os = {}

    for dados in registros:
        tag = limpar_texto(dados.get("TAG", ""))
        if tag:
            numero_os = obter_numero_os(dados)
            if numero_os:
                grupos_por_os.setdefault(numero_os, set()).add(tag_para_pasta(tag))

    grupos = {}
    for indice, dados in enumerate(registros, start=2):
        tag = limpar_texto(dados.get("TAG", ""))
        numero_os = obter_numero_os(dados)

        grupos_correspondentes = grupos_por_os.get(numero_os, set())
        if not tag and len(grupos_correspondentes) == 1:
            tag = next(iter(grupos_correspondentes))
        elif not tag and numero_os:
            tag = f"SEM-TAG-OS{numero_os}"
        elif not tag:
            tag = f"SEM-TAG-LINHA-{indice}"

        dados["TAG"] = tag
        chave = tag_para_pasta(tag)
        if chave not in grupos:
            grupos[chave] = dados
            continue

        registro_existente = grupos[chave]
        for coluna, valor in dados.items():
            texto = limpar_texto(valor)
            if not texto:
                continue

            if normalizar_nome_coluna(coluna) == "observacoes":
                observacoes = limpar_texto(registro_existente.get(coluna, ""))
                linhas_existentes = set(observacoes.splitlines())
                novas_linhas = [linha for linha in texto.splitlines() if linha not in linhas_existentes]
                if novas_linhas:
                    registro_existente[coluna] = "\n".join(filter(None, (observacoes, "\n".join(novas_linhas))))
            else:
                registro_existente[coluna] = valor

    return pd.DataFrame(grupos.values())


def tag_para_pasta(tag):
    tag = limpar_texto(tag).upper()
    tag = tag.replace(" ", "-")
    tag = re.sub(r"[^A-Z0-9_-]", "-", tag)
    tag = re.sub(r"-+", "-", tag)
    return tag.strip("-")


def link_relativo(caminho):
    caminho = limpar_texto(caminho)
    if caminho == "":
        return ""
    if caminho.startswith("http://") or caminho.startswith("https://"):
        return caminho
    if caminho.startswith("../../"):
        return caminho
    return "../../" + caminho.lstrip("/")


def criar_botao(titulo, caminho, classe="botao"):
    caminho = limpar_texto(caminho)
    if caminho == "":
        return f'<span class="botao inativo">{html.escape(titulo)} indisponível</span>'
    return f'<a class="{classe}" href="{html.escape(link_relativo(caminho))}" target="_blank">{html.escape(titulo)}</a>'


def gerar_html_equipamento(dados):
    tag = limpar_texto(dados.get("TAG", ""))
    tag_pasta = tag_para_pasta(tag)

    foto = link_relativo(obter_valor_coluna(dados, "Foto", "FOTO"))
    if foto:
        bloco_foto = f'<img src="{html.escape(foto)}" alt="Foto do equipamento {html.escape(tag)}" class="foto-equipamento">'
    else:
        bloco_foto = '<div class="obs">Foto não cadastrada na planilha.</div>'

    campos = [
        ("TAG", "TAG"),
        ("Fabricante", "Fabricante"),
        ("Tipo do equipamento", "Tipo"),
        ("Modelo", "Modelo"),
        ("Potência", "Potencia"),
        ("Relação de redução", "Relacao"),
        ("RPM entrada/saída", "RPM"),
        # ("Tipo de montagem", "Montagem"),
        ("Lubrificante", "Lubrificante"),
        ("Quantidade de óleo", "Quantidade_Oleo"),
        ("Local / Área", "Local"),
        ("Observações", "Observacoes"),
    ]

    linhas = ""
    for titulo, coluna in campos:
        valor = obter_observacoes(dados) if coluna == "Observacoes" else limpar_texto(obter_valor_coluna(dados, coluna, coluna.replace("_", " ")))
        linhas += f"""
        <tr>
          <th>{html.escape(titulo)}</th>
          <td>{html.escape(valor)}</td>
        </tr>"""

    botao_lista = criar_botao("Lista de Peças", obter_valor_coluna(dados, "Lista_Pecas", "Lista de Peças", "LISTA DE PEÇAS", "LISTA PECAS"))
    botao_manual = criar_botao("Manual", obter_valor_coluna(dados, "Manual", "MANUAL"))

    caminho_relatorio = obter_valor_coluna(
        dados,
        "RELATÓRIO FINAL",
        "RELATORIO FINAL",
        "Relatorio Final",
        "Relatório Final",
        "Certificado",
        "CERTIFICADA",
        "Certificada",
    )
    botao_relatorio = criar_botao("RELATORIO FINAL", caminho_relatorio)
    botao_fotos = criar_botao("Fotos adicionais", obter_valor_coluna(dados, "Fotos_Adicionais", "FOTOS ADICIONAIS"))

    link_publicado = BASE_URL.rstrip("/") + "/equipamentos/" + tag_pasta + "/"

    logo_gold = "../../assets/img/transparent/logo_gold_transparent.png"
    logo_inpasa = "../../assets/img/logo_inpasa.png"

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Mini Data Book - {html.escape(tag)}</title>
  <link rel="stylesheet" href="../../assets/css/style.css">
</head>
<body>

  <header class="topo">
    <div class="header-row">
      <div class="logos">
         <!-- Ajuste o tamanho do logo INPASA aqui via .logo-inpasa no CSS. -->
        <img class="logo-inpasa" src="{html.escape(logo_inpasa)}" alt="INPASA">
       
      </div>
      <div class="header-text">
        <div class="tipo">{html.escape(limpar_texto(dados.get("Tipo", "Equipamento")))}</div>
        <div class="tag">{html.escape(tag)}</div>
      </div>
      <div class="logos">
      <img class="logo-gold" src="{html.escape(logo_gold)}" alt="Gold Redutores">
      </div>
    </div>
  </header>

  <main class="container">
    <section class="photo-card">
      {bloco_foto}
    </section>

    <section class="card">
      <h2>Dados principais</h2>
      <table class="tabela-dados">
        {linhas}
      </table>
    </section>

    <section class="card">
      <h2>Acessos rápidos</h2>
      <div class="botoes">
        {botao_relatorio}
        {botao_lista}
        {botao_manual}
        {botao_fotos}
        <a href="../../index.html" class="botao secundario">Voltar para lista geral</a>
      </div>
    </section>

    <section class="card">
      <h2>Link desta página</h2>
      <div class="obs">{html.escape(link_publicado)}</div>
    </section>
  </main>

  <footer>
    Gold Redutores - Mini Data Book Digital Caldeira 1 - INPASA| Página individual por equipamento
  </footer>

</body>
</html>
"""


def gerar_index_principal(df):
    cards = ""

    for _, row in df.iterrows():
        tag = limpar_texto(row.get("TAG", ""))
        if not tag:
            continue

        tag_pasta = tag_para_pasta(tag)
        tipo = limpar_texto(row.get("Tipo", "Equipamento"))
        modelo = limpar_texto(row.get("Modelo", ""))
        local = limpar_texto(row.get("Local", ""))
        observacoes = obter_observacoes(row)
        busca = (tag + " " + tipo + " " + modelo + " " + local + " " + observacoes).lower()

        cards += f"""
      <a class="equipamento-card" href="equipamentos/{html.escape(tag_pasta)}/index.html" data-busca="{html.escape(busca)}">
        <strong>{html.escape(tag)}</strong>
        <span>{html.escape(tipo)}</span>
        <span>{html.escape(modelo)}</span>
        <small class="observacoes-card">Observações: {html.escape(observacoes)}</small>
      </a>
"""

    logo_gold = "assets/img/transparent/logo_gold_transparent.png"
    logo_inpasa = "assets/img/logo_inpasa.png"

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Data Book Digital - Equipamentos Caldeira 1</title>
  <link rel="stylesheet" href="assets/css/style.css">
  <style>
    /* Estilos Integrados do Pure CSS3 Slider - Copiado do Relatório */
    .header-slider-wrapper {{ position: sticky; top: 0; z-index: 1000; }}
    #slider {{ width: 90%; max-width: 1080px; height: 70px; position: relative; overflow: hidden; margin: 8px auto 15px auto; border: 1px solid #333; background: #fff; }}
    .slides {{ width: 200%; height: 100%; position: relative; -webkit-animation: slide 12s infinite; -moz-animation: slide 12s infinite; animation: slide 12s infinite; }}
    .slider {{ width: 16.6666%; height: 100%; float: left; position: relative; z-index: 1; overflow: hidden; box-sizing: border-box; border-right: 2px solid #333; }}
    .image {{ width: 100%; height: 100%; position: relative; }}
    .image img {{ width: 100%; height: 100%; object-fit: contain; display: block; }}
    .switch {{ width: 120px; height: 10px; position: absolute; bottom: 18px; z-index: 99; left: 20px; }}
    .switch>ul {{ list-style: none; overflow: hidden; padding: 0; margin: 0; }}
    .switch>ul>li {{ width: 6px; height: 6px; border-radius: 50%; background: #333; float: left; margin-right: 9px; cursor: pointer; }}
    .on {{ width: 100%; height: 100%; border-radius: 50%; background: rgba(23, 117, 177, 1); position: relative; -webkit-animation: on 12s infinite; -moz-animation: on 12s infinite; animation: on 12s infinite; }}
    @-webkit-keyframes slide {{ 0%, 20% {{ margin-left: 0%; }} 25%, 45% {{ margin-left: -33.333%; }} 50%, 70% {{ margin-left: -66.666%; }} 75%, 95% {{ margin-left: -100%; }} 100% {{ margin-left: 0%; }} }}
    @keyframes slide {{ 0%, 20% {{ margin-left: 0%; }} 25%, 45% {{ margin-left: -33.333%; }} 50%, 70% {{ margin-left: -66.666%; }} 75%, 95% {{ margin-left: -100%; }} 100% {{ margin-left: 0%; }} }}
    @-webkit-keyframes on {{ 0%, 20% {{ margin-left: 0%; }} 25%, 45% {{ margin-left: 15px; }} 50%, 70% {{ margin-left: 30px; }} 75%, 95% {{ margin-left: 45px; }} 100% {{ margin-left: 0%; }} }}
    @keyframes on {{ 0%, 20% {{ margin-left: 0%; }} 25%, 45% {{ margin-left: 15px; }} 50%, 70% {{ margin-left: 30px; }} 75%, 95% {{ margin-left: 45px; }} 100% {{ margin-left: 0%; }} }}
  </style>
</head>
<body>

  <div class="header-slider-wrapper">
    <header class="topo">
      <div class="header-row">
        <div class="logos">
        <!-- Ajuste o tamanho do logo INPASA aqui via .logo-inpasa no CSS. -->
          <img class="logo-inpasa" src="{html.escape(logo_inpasa)}" alt="INPASA">
        </div>
          <div class="header-text">
          <div class="tipo">Data Book Digital</div>
          <div class="tag">Equipamentos - Caldeira 1 </div>
        </div>
        <div class="logos">
        <img class="logo-gold" src="{html.escape(logo_gold)}" alt="Gold Redutores">
        </div>
      </div>
    </header>

    <!-- Carrossel de Imagens -->
    <div id="slider">
        <div class="slides">
            <div class="slider">
                <div class="image">
                    <img src="https://i.ibb.co/V07bLNT1/GOLD-MAQUINAS.png" alt="Logo Gold Redutores">
                </div>
            </div>
            <div class="slider">
                <div class="image">
                    <img src="https://i.ibb.co/JjQNQPwJ/logo-inpasa.png" alt="Logo Inpasa" style="object-fit: contain; padding: 20px; box-sizing: border-box;">
                </div>
            </div>
            <div class="slider">
                <div class="image">
                    <img src="https://i.ibb.co/r9DRsZq/GOLD-REDUTORES.png" alt="GOLD REDUTORES" border="0">
                </div>
            </div>
            <div class="slider">
                <div class="image">
                    <img src="https://i.ibb.co/JjQNQPwJ/logo-inpasa.png" alt="Logo Inpasa" style="object-fit: contain; padding: 20px; box-sizing: border-box;">
                </div>
            </div>
            <div class="slider">
                <div class="image">
                    <img src="https://i.ibb.co/3mVp95wL/GOLD-AUTOMA-O.png" alt="GOLD AUTOMAÇÃO" border="0">
                </div>
            </div>
            <div class="slider">
                <div class="image">
                    <img src="https://i.ibb.co/JjQNQPwJ/logo-inpasa.png" alt="Logo Inpasa" style="object-fit: contain; padding: 20px; box-sizing: border-box;">
                </div>
            </div>
        </div>
        <div class="switch">
            <ul>
                <li><div class="on"></div></li>
                <li></li>
                <li></li>
                <li></li>
            </ul>
        </div>
    </div>
  </div>

  <main class="container">
    <section class="card">
      <h2>Equipamentos cadastrados</h2>
      <p>Digite a TAG, modelo, tipo ou local para localizar o equipamento.</p>
      <input class="busca" id="busca" type="text" placeholder="Exemplo: ME-631011, motoredutor, WEG...">
    </section>

    <section class="grid-equipamentos" id="listaEquipamentos">
      {cards}
    </section>
  </main>

  <footer>Gold Redutores - Data Book Digital - Caldeira 1</footer>

  <script>
    const campoBusca = document.getElementById("busca");
    const cards = document.querySelectorAll(".equipamento-card");

    campoBusca.addEventListener("input", function() {{
      const termo = campoBusca.value.toLowerCase().trim();
      cards.forEach(card => {{
        const texto = card.getAttribute("data-busca");
        card.style.display = texto.includes(termo) ? "block" : "none";
      }});
    }});
  </script>

</body>
</html>
"""


def main():
    arquivo_excel = ARQUIVO_EXCEL
    if not arquivo_excel.exists():
        fallback = BASE_DIR / "dados_equipamentos-geral.xlsx"
        if fallback.exists():
            arquivo_excel = fallback
        else:
            raise FileNotFoundError(f"Arquivo não encontrado: {arquivo_excel}")

    df = pd.read_excel(arquivo_excel)

    if "TAG" not in df.columns:
        raise ValueError("A planilha precisa ter uma coluna chamada TAG.")

    df = preparar_equipamentos(df)

    PASTA_EQUIPAMENTOS.mkdir(exist_ok=True)
    PASTA_QRCODES.mkdir(parents=True, exist_ok=True)

    total = 0
    tags_esperadas = set()

    for _, row in df.iterrows():
        tag = limpar_texto(row.get("TAG", ""))

        tag_pasta = tag_para_pasta(tag)
        tags_esperadas.add(tag_pasta)
        pasta = PASTA_EQUIPAMENTOS / tag_pasta
        pasta.mkdir(parents=True, exist_ok=True)

        html_equip = gerar_html_equipamento(row)
        (pasta / "index.html").write_text(html_equip, encoding="utf-8")

        link_equipamento = BASE_URL.rstrip("/") + f"/equipamentos/{tag_pasta}/"

        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=10,
            border=4,
        )
        qr.add_data(link_equipamento)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        img.save(PASTA_QRCODES / f"QR-{tag_pasta}.png")

        total += 1

    (BASE_DIR / "index.html").write_text(gerar_index_principal(df), encoding="utf-8")

    for pagina in PASTA_EQUIPAMENTOS.glob("*/index.html"):
        if pagina.parent.name not in tags_esperadas:
            pagina.unlink()

    for qr in PASTA_QRCODES.glob("QR-*.png"):
        if qr.stem[3:] not in tags_esperadas:
            qr.unlink()

    print(f"Concluído: {total} páginas criadas.")
    print("QR Codes salvos em assets/qrcodes/")
    print("Página inicial criada: index.html")


if __name__ == "__main__":
    main()
