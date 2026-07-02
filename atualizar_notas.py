import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
import os

TOKEN = os.environ["CANVAS_TOKEN"]

BASE_URL = "https://pucpr.instructure.com/api/v1"

HEADERS = {"Authorization": f"Bearer {TOKEN}"}

COURSES = {
    "Aspectos Legais": 61646,
    "Experiencia Criativa": 61655,
    "Filosofia": 61640,
    "Raciocinio Algoritmico": 61650,
    "Fundamentos Sistemas Ciberfisicos": 61660,
}


def get_all_pages(url):
    dados = []

    while url:
        r = requests.get(url, headers=HEADERS)

        if r.status_code != 200:
            raise Exception(f"Erro {r.status_code}: {r.text}")

        dados.extend(r.json())

        if "next" in r.links:
            url = r.links["next"]["url"]
        else:
            url = None

    return dados


arquivo_excel = "Painel_Academico_PUCPR.xlsx"

writer = pd.ExcelWriter(arquivo_excel, engine="openpyxl")

resumo = []
pendentes = []
nao_contabilizadas = []
agenda = []

for curso_nome, course_id in COURSES.items():
    print(f"Processando {curso_nome}")

    try:
        assignments = get_all_pages(
            f"{BASE_URL}/courses/{course_id}/assignments?per_page=100"
        )
        print(f"{curso_nome}: {len(assignments)} atividades encontradas")

        linhas = []
        total_obtido = 0.0
        total_maximo = 0.0
        total_futuro = 0.0

        for a in assignments:
            if a.get("omit_from_final_grade", False):
                nao_contabilizadas.append(
                    {"Disciplina": curso_nome, "Atividade": a["name"]}
                )

                continue

            pontos = a.get("points_possible")
            due_at = a.get("due_at")
            html_url = a.get("html_url")

            if pontos is None or pontos <= 0:
                continue

            try:
                sub = requests.get(
                    f"{BASE_URL}/courses/{course_id}/assignments/{a['id']}/submissions/self",
                    headers=HEADERS,
                ).json()

                score = sub.get("score")

            except Exception:
                score = None

            if score is None:
                total_futuro += float(pontos)

                pendentes.append(
                    {
                        "Disciplina": curso_nome,
                        "Atividade": a["name"],
                        "Entrega": due_at, 
                        "Valor Maximo": pontos,
                        "Link": html_url,                       
                    }
                )
              
                agenda.append(
                  {
                     "Data": due_at,
                     "Disciplina": curso_nome,
                     "Atividade": a["name"],
                     "Status": "Pendente",
                     "Link": html_url,
                  }
                )

                linhas.append(
                {
                   "Atividade": a["name"],
                   "Entrega": due_at,
                   "Obtido": "",
                  "Maximo": pontos,
                  "%": "",
                  "Status": "Pendente",
                 "Link": html_url,
                }
                )

                continue

            percentual = round((float(score) / float(pontos)) * 100, 2)

            total_obtido += float(score)
            total_maximo += float(pontos)

            if percentual >= 70:
                status = "Bom"
            else:
                status = "Atencao"

            linhas.append(
                {
                    "Atividade": a["name"],
                    "Obtido": score,
                    "Maximo": pontos,
                    "%": percentual,
                    "Status": status,
                }
            )

        media = 0

        if total_maximo > 0:
            media = round((total_obtido / total_maximo) * 100, 2)
            projecao = media

        if total_futuro > 0 and total_maximo > 0:
            desempenho_atual = total_obtido / total_maximo

            nota_futura_esperada = total_futuro * desempenho_atual

            projecao = round(
                (total_obtido + nota_futura_esperada)
                / (total_maximo + total_futuro)
                * 100,
                2,
            )

            resumo.append(
                {
                    "Disciplina": curso_nome,
                    "Obtido": round(total_obtido, 2),
                    "Maximo Atual": round(total_maximo, 2),
                    "Pontos Pendentes": round(total_futuro, 2),
                    "Media Atual (%)": media,
                    "Projecao Final (%)": projecao,
                }
            )

        print(f"{curso_nome}: {len(linhas)} linhas geradas")

        df = pd.DataFrame(linhas)

        if len(df) > 0:
            df.to_excel(writer, sheet_name=curso_nome[:31], index=False)

    except Exception as e:
        print(f"ERRO EM {curso_nome}: {e}")

        resumo.append(
            {
                "Disciplina": curso_nome,
                "Obtido": "ERRO",
                "Maximo Atual": "",
                "Pontos Pendentes": "",
                "Media Atual (%)": "",
                "Projecao Final (%)": str(e),
            }
        )

pd.DataFrame(nao_contabilizadas).to_excel(
    writer, sheet_name="Nao_Contabilizadas", index=False
)

pd.DataFrame(resumo).to_excel(writer, sheet_name="Resumo", index=False)

pd.DataFrame(pendentes).to_excel(writer, sheet_name="Pendentes", index=False)

writer.close()

wb = load_workbook(arquivo_excel)

for ws in wb.worksheets:
    for cell in ws[1]:
        cell.font = Font(bold=True)

        cell.fill = PatternFill(fill_type="solid", fgColor="D9EAD3")

    for col in ws.columns:
        tamanho = 0

        for cell in col:
            try:
                tamanho = max(tamanho, len(str(cell.value)))
            except:  # noqa: E722
                pass

        ws.column_dimensions[col[0].column_letter].width = tamanho + 3

wb.save(arquivo_excel)

print()
print("=" * 50)
print("PLANILHA GERADA COM SUCESSO")
print(arquivo_excel)
print("=" * 50)
