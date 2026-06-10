
"""
GERADOR DE DATA BOOK DIGITAL - EXCEL + GITHUB PAGES

Este script lê dados_equipamentos.xlsx e cria:
- 1 página HTML por equipamento
- 1 página inicial com busca
- 1 QR Code por equipamento

INSTALAÇÃO:
pip install pandas openpyxl qrcode pillow

USO:
1. Preencha dados_equipamentos.xlsx
2. Altere BASE_URL abaixo
3. Rode: python gerar_site.py
4. Envie o conteúdo da pasta para o GitHub
"""

from pathlib import Path
import pandas as pd
import qrcode
import re
import html

# ALTERAR AQUI depois que o GitHub Pages estiver criado.
# Exemplo: "https://gold-redutores.github.io/DATABOOK-INPASA"
BASE_URL = " https://goldredutoresinpasa-web.github.io/DATABOOK-INPASA/"

ARQUIVO_EXCEL = "dados_equipamentos.xlsx"
PASTA_EQUIPAMENTOS = Path("equipamentos")
PASTA_QRCODES = Path("assets/qrcodes")


def limpar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


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

    foto = link_relativo(dados.get("Foto", ""))
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
        valor = limpar_texto(dados.get(coluna, ""))
        linhas += f"""
        <tr>
          <th>{html.escape(titulo)}</th>
          <td>{html.escape(valor)}</td>
        </tr>"""

    botao_lista = criar_botao("Lista de Peças", dados.get("Lista_Pecas", ""))
    botao_manual = criar_botao("Manual", dados.get("Manual", ""))
    # botao_certificado = criar_botao("Certificado", dados.get("Certificado", ""))
    botao_fotos = criar_botao("Fotos adicionais", dados.get("Fotos_Adicionais", ""))

    link_publicado = BASE_URL.rstrip("/") + "/equipamentos/" + tag_pasta + "/"

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
    <div>
      <h1>Mini Data Book - {html.escape(tag)}</h1>
      <p>{html.escape(limpar_texto(dados.get("Tipo", "Equipamento")))} | Consulta rápida via QR Code</p>
    </div>
    <div class="logo">GOLD<br>REDUTORES</div>
  </header>

  <main class="container">
    <section class="card">
      <h2>Foto do equipamento</h2>
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
    Gold Redutores - Mini Data Book Digital | Página individual por equipamento
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
        busca = (tag + " " + tipo + " " + modelo + " " + local).lower()

        cards += f"""
      <a class="equipamento-card" href="equipamentos/{html.escape(tag_pasta)}/index.html" data-busca="{html.escape(busca)}">
        <strong>{html.escape(tag)}</strong>
        <span>{html.escape(tipo)}</span>
        <span>{html.escape(modelo)}</span>
        <small>{html.escape(local)}</small>
      </a>
"""

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Data Book Digital - Equipamentos</title>
  <link rel="stylesheet" href="assets/css/style.css">
</head>
<body>

  <header class="topo">
    <div>
      <h1>Data Book Digital</h1>
      <p>Consulta de equipamentos por TAG</p>
    </div>
    <div class="logo">GOLD<br>REDUTORES</div>
  </header>

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

  <footer>Gold Redutores - Data Book Digital</footer>

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
    if not Path(ARQUIVO_EXCEL).exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {ARQUIVO_EXCEL}")

    df = pd.read_excel(ARQUIVO_EXCEL)

    if "TAG" not in df.columns:
        raise ValueError("A planilha precisa ter uma coluna chamada TAG.")

    PASTA_EQUIPAMENTOS.mkdir(exist_ok=True)
    PASTA_QRCODES.mkdir(parents=True, exist_ok=True)

    total = 0

    for _, row in df.iterrows():
        tag = limpar_texto(row.get("TAG", ""))
        if not tag:
            continue

        tag_pasta = tag_para_pasta(tag)
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

    Path("index.html").write_text(gerar_index_principal(df), encoding="utf-8")

    print(f"Concluído: {total} páginas criadas.")
    print("QR Codes salvos em assets/qrcodes/")
    print("Página inicial criada: index.html")


if __name__ == "__main__":
    main()
