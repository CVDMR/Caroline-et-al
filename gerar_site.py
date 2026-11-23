"""Gera o site estático (docs/index.html) com menu lateral por variável.

Lê as figuras (PNG) de `figuras/` e as tabelas (.docx) de `tabelas/` — as mesmas
que o app Streamlit salva — e monta uma página única, no estilo de um relatório
Quarto: menu à esquerda com as variáveis; ao clicar, aparecem as figuras e tabelas
relacionadas. Só usa resultados agregados (a planilha original não é lida).

Uso:  python gerar_site.py        (depois: git add docs && git commit && git push)
"""
from __future__ import annotations

import html
from datetime import date
import shutil
from pathlib import Path

from docx import Document

RAIZ = Path(__file__).parent
FIGURAS = RAIZ / "figuras"
TABELAS = RAIZ / "tabelas"
SAIDA = RAIZ / "docs"

# Menu: grupo -> lista de (id, nome da variável, [figuras], [tabelas])
# Os nomes são os dos arquivos salvos pelo app (sem extensão).
MENU = {
    "Caracterização": [
        ("subgrupo", "Subgrupo", ["produtos_por_subgrupo"], ["produtos_por_subgrupo"]),
        ("marca", "Marca/empresa", ["produtos_por_marca_empresa"], ["produtos_por_marca_empresa"]),
        ("preco", "Preço por 100 g",
         ["preco_por_100_g", "histograma_preco_100g", "boxplot_preco_100g"],
         ["preco_100g_por_subgrupo", "explorador_preco_100g"]),
        ("peso", "Peso adquirido (g)", ["peso_adquirido_g"], ["peso_gramas_por_subgrupo"]),
    ],
    "Ingredientes e aditivos": [
        ("ingredientes", "Nº de ingredientes", ["no_de_ingredientes"], ["numero_ingredientes_por_subgrupo"]),
        ("aditivos", "Nº de aditivos", ["no_de_aditivos"], ["n_aditivos_por_subgrupo"]),
        ("glutamato", "Glutamato, nitrato e nitrito",
         ["presenca_de_glutamato_nitrato_e_nitrito"], ["presenca_aditivos"]),
    ],
    "Rotulagem": [
        ("conformidade", "Itens obrigatórios",
         ["conformidade_itens_obrigatorios", "conformidade_por_subgrupo"],
         ["conformidade_itens_obrigatorios", "conformidade_por_subgrupo"]),
        ("rnf", "Rotulagem nutricional frontal (RNF)", [], ["rnf_resumo"]),
        ("rnf_alertas", "Nº de alertas na RNF", ["no_de_alertas_na_rnf"], ["rnf_n_alertas"]),
        ("rnf_nutrientes", "Nutriente(s) indicado(s) na RNF", ["nutriente_s_indicado_s_na_rnf"], ["rnf_nutrientes"]),
    ],
}


def tabela_html(nome: str) -> str:
    """Converte um .docx de tabela (título, tabela, nota) em HTML no estilo de artigo."""
    doc = Document(TABELAS / f"{nome}.docx")
    textos = [p.text for p in doc.paragraphs if p.text.strip()]
    titulo = textos[0] if textos else nome
    nota = next((t for t in textos if t.startswith("Nota.")), "")
    tab = doc.tables[0]
    linhas = [[c.text for c in r.cells] for r in tab.rows]

    esc = html.escape
    num, _, resto = titulo.partition(". ")
    out = [f'<figure class="tabela"><figcaption><b>{esc(num)}.</b> {esc(resto)}</figcaption>',
           '<div class="rolagem"><table><thead><tr>']
    out += [f"<th>{esc(c)}</th>" for c in linhas[0]]
    out.append("</tr></thead><tbody>")
    for r in linhas[1:]:
        cls = ' class="total"' if r[0] == "Total" else ""
        out.append(f"<tr{cls}>" + "".join(f"<td>{esc(c)}</td>" for c in r) + "</tr>")
    out.append("</tbody></table></div>")
    if nota:
        out.append(f'<p class="nota"><i>Nota.</i> {esc(nota.removeprefix("Nota. "))}</p>')
    out.append("</figure>")
    return "".join(out)


def figura_html(nome: str) -> str:
    return (f'<figure class="figura"><a href="figuras/{nome}.png" target="_blank">'
            f'<img src="figuras/{nome}.png" alt="{html.escape(nome.replace("_", " "))}" loading="lazy"></a></figure>')


def pagina(vid: str, titulo: str, figs: list[str], tabs: list[str]) -> str:
    partes = [f'<section class="pagina" id="p-{vid}"><h1>{html.escape(titulo)}</h1>']
    if figs:
        partes.append("<h2>Figuras</h2>")
        partes += [figura_html(f) for f in figs]
    if tabs:
        partes.append("<h2>Tabelas</h2>")
        partes += [tabela_html(t) for t in tabs]
    partes.append("</section>")
    return "".join(partes)


CSS = """
:root{--bg:#fff;--fg:#1f2328;--mut:#656d76;--lin:#d0d7de;--side:#f6f8fa;--ac:#0b5cad;--hov:#e8eef6}
@media (prefers-color-scheme:dark){:root{--bg:#0f1318;--fg:#e6edf3;--mut:#8b949e;--lin:#30363d;--side:#161b22;--ac:#58a6ff;--hov:#1f2937}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:16px/1.55 system-ui,"Segoe UI",sans-serif;display:flex;min-height:100vh}
nav{width:280px;flex:none;background:var(--side);border-right:1px solid var(--lin);padding:16px 12px;position:sticky;top:0;height:100vh;overflow-y:auto}
nav .titulo{font-weight:700;font-size:1.05rem;margin:4px 8px 12px}
nav h3{font-size:.72rem;letter-spacing:.06em;text-transform:uppercase;color:var(--mut);margin:18px 8px 4px}
nav a{display:block;padding:6px 10px;border-radius:6px;color:var(--fg);text-decoration:none;font-size:.93rem}
nav a:hover{background:var(--hov)}
nav a.ativo{background:var(--hov);color:var(--ac);font-weight:600;box-shadow:inset 3px 0 var(--ac)}
main{flex:1;min-width:0;max-width:980px;padding:28px 40px 60px;margin:0 auto}
.pagina{display:none}.pagina.ativa{display:block}
h1{font-size:1.7rem;margin:0 0 6px;border-bottom:1px solid var(--lin);padding-bottom:8px}
h2{font-size:1.1rem;color:var(--mut);margin:26px 0 8px}
figure{margin:12px 0 22px}
.figura img{max-width:100%;height:auto;border:1px solid var(--lin);border-radius:6px;background:#fff}
.tabela figcaption{font-size:.92rem;margin-bottom:4px}
.rolagem{overflow-x:auto}
table{border-collapse:collapse;font-family:"Times New Roman",Times,serif;font-size:.98rem;min-width:50%}
thead th{border-top:2px solid var(--fg);border-bottom:1px solid var(--fg);text-align:center;padding:4px 12px}
td{padding:3px 12px;text-align:center}
td:first-child,th:first-child{text-align:left}
tbody tr:last-child td{border-bottom:2px solid var(--fg)}
tr.total td{border-top:1px solid var(--fg)}
.nota{font-size:.82rem;color:var(--mut);margin:4px 0 0}
.inicio p{max-width:68ch}
.inicio .grade{display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:10px;margin-top:14px}
.inicio .cartao{border:1px solid var(--lin);border-radius:8px;padding:10px 12px;text-decoration:none;color:var(--fg)}
.inicio .cartao:hover{background:var(--hov)}
.inicio .cartao small{display:block;color:var(--mut)}
footer{margin-top:40px;font-size:.8rem;color:var(--mut)}
#abre{display:none}
@media (max-width:760px){
 body{display:block}
 nav{position:static;width:auto;height:auto;border-right:0;border-bottom:1px solid var(--lin);max-height:45vh}
 main{padding:18px 16px 40px}
}
"""

JS = """
const paginas=[...document.querySelectorAll('.pagina')], links=[...document.querySelectorAll('nav a[data-id]')];
function mostra(){
  const id=(location.hash||'#inicio').slice(1);
  let achou=false;
  paginas.forEach(p=>{const ok=p.id==='p-'+id;p.classList.toggle('ativa',ok);achou=achou||ok;});
  if(!achou){document.getElementById('p-inicio').classList.add('ativa');}
  links.forEach(a=>a.classList.toggle('ativo',a.dataset.id===id));
  window.scrollTo(0,0);
}
addEventListener('hashchange',mostra);mostra();
"""


def monta() -> str:
    nav = ['<nav><div class="titulo">Rotulagem de produtos cárneos</div>',
           '<a href="#inicio" data-id="inicio">Início</a>']
    paginas = []
    cartoes = []
    for grupo, itens in MENU.items():
        nav.append(f"<h3>{html.escape(grupo)}</h3>")
        for vid, nome, figs, tabs in itens:
            figs = [f for f in figs if (FIGURAS / f"{f}.png").exists()]
            tabs = [t for t in tabs if (TABELAS / f"{t}.docx").exists()]
            nav.append(f'<a href="#{vid}" data-id="{vid}">{html.escape(nome)}</a>')
            paginas.append(pagina(vid, nome, figs, tabs))
            cartoes.append(f'<a class="cartao" href="#{vid}">{html.escape(nome)}'
                           f'<small>{len(figs)} figura(s) · {len(tabs)} tabela(s)</small></a>')
    nav.append("</nav>")

    inicio = ('<section class="pagina inicio" id="p-inicio"><h1>Análise descritiva</h1>'
              "<p>Resultados da análise descritiva da rotulagem e da composição de produtos cárneos "
              "comercializados na Bahia. Escolha uma variável no menu à esquerda para ver as figuras "
              "e as tabelas correspondentes.</p>"
              f'<div class="grade">{"".join(cartoes)}</div></section>')

    return ('<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1">'
            "<title>Análise descritiva</title>"
            f"<style>{CSS}</style></head><body>{''.join(nav)}"
            f"<main>{inicio}{''.join(paginas)}"
            f"<footer>Resultados agregados; gerado por <code>gerar_site.py</code> em {date.today():%d/%m/%Y}.</footer></main>"
            f"<script>{JS}</script></body></html>")


def main() -> None:
    SAIDA.mkdir(exist_ok=True)
    destino = SAIDA / "figuras"
    destino.mkdir(exist_ok=True)
    for png in FIGURAS.glob("*.png"):
        shutil.copy2(png, destino / png.name)
    (SAIDA / ".nojekyll").write_text("")
    (SAIDA / "index.html").write_text(monta(), encoding="utf-8")
    print(f"Site gerado em {SAIDA / 'index.html'}")


if __name__ == "__main__":
    main()
