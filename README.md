# Painel Acadêmico PUCPR – Canvas Grade Tracker

Automatically fetch your grades from Canvas LMS (PUCPR) and build a clear, formatted Excel dashboard – updated every 12 hours via GitHub Actions.

## 📊 Overview

This project connects to the **Canvas API** of PUCPR (Pontifícia Universidade Católica do Paraná), retrieves all assignments for pre-configured courses, and generates a spreadsheet (`Painel_Academico_PUCPR.xlsx`) containing:

- A per-course breakdown of each assignment (score obtained, maximum points, percentage, status).
- A summary with current average and projected final grade, considering future (ungraded) assignments.
- Lists of pending activities and assignments excluded from the final grade.

The included GitHub Actions workflow runs the script automatically and pushes the updated spreadsheet back to the repository.

## ✨ Features

- 📥 Fetches all assignments (paginated) from multiple Canvas courses.
- 📈 Calculates **current average** and a **projected final grade** assuming future performance mirrors past results.
- ⚠️ Highlights assignments with score < 70% as "Atenção" (warning).
- 📋 Separate sheets: **Resumo** (summary), **Pendentes** (pending), **Não Contabilizadas** (not counted).
- 🎨 Excel formatting: bold headers, auto-column width, green header fill.
- ⏱ Scheduled execution via GitHub Actions (every 12 hours) – no manual run needed.
- 🔒 Securely stores your Canvas token as a GitHub secret.

## 🧰 Prerequisites

- Python 3.11 or higher (for local runs)
- A **Canvas API access token** (see [How to get a Canvas token](#-how-to-get-a-canvas-token))
- Git (for cloning the repository)
- A GitHub account (if you want to use the automated workflow)

## 🚀 Setup (Local Execution)

1. **Clone the repository**
   ```bash
   git clone https://github.com/your-username/your-repo.git
   cd your-repo
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set your Canvas token**
   - **Linux/macOS**: `export CANVAS_TOKEN="seu_token_aqui"`
   - **Windows (PowerShell)**: `$env:CANVAS_TOKEN="seu_token_aqui"`
   - Or create a `.env` file (not included, but you could extend the script).

4. **Configure your courses**  
   Open `atualizar_notas.py` and edit the `COURSES` dictionary:
   ```python
   COURSES = {
       "Nome da Disciplina 1": course_id_1,
       "Nome da Disciplina 2": course_id_2,
       # Add more entries as needed
   }
   ```
   - The **course ID** is the number in the Canvas URL when you enter a course: `https://pucpr.instructure.com/courses/XXXXX`
   - You can remove, add, or rename courses – just keep the key as the display name and the value as the numeric ID.

5. **Run the script**
   ```bash
   python atualizar_notas.py
   ```
   The file `Painel_Academico_PUCPR.xlsx` will be created/updated in the same folder.

## 🤖 Automation with GitHub Actions

The included workflow (`.github/workflows/atualizar.yml`) automatically:
- Runs every 12 hours (cron: `0 */12 * * *`)
- Can be triggered manually via the “Actions” tab (`workflow_dispatch`)

To enable this:

1. **Fork or create a new repository** with all the project files.
2. **Add your Canvas token as a secret** in the repository:
   - Go to **Settings → Secrets and variables → Actions → New repository secret**
   - Name: `CANVAS_TOKEN`
   - Value: your Canvas API token
3. The workflow will start running according to the schedule. Each successful run commits the updated `Painel_Academico_PUCPR.xlsx` back to the repo.

> ⚠️ Note: The spreadsheet file is binary; frequent commits may increase repository size over time. To keep it clean, you can occasionally squash history or purge the file from Git history.

## 📄 Output Structure

The Excel file contains the following sheets:

| Sheet                    | Description |
|--------------------------|-------------|
| **Resumo**               | Discipline, points obtained, max points so far, pending points, current average (%), projected final grade (%) |
| **Pendentes**            | Assignments not yet graded (score is null), with their maximum points |
| **Nao_Contabilizadas**   | Assignments marked as “omit from final grade” |
| **{Disciplina}**         | One sheet per course with columns: Atividade, Obtido, Máximo, %, Status |

- **Status** values:  
  - `Bom` (≥70%)  
  - `Atenção` (<70%)  
  - `Pendente` (not graded)

## 🔑 How to Get a Canvas Token

1. Log in to [Canvas PUCPR](https://pucpr.instructure.com).
2. Click on your avatar (top left) → **Settings**.
3. Scroll down to **Approved Integrations** and click **+ New Access Token**.
4. Set a purpose (e.g., “Grade Dashboard”) and an expiration date.
5. **Copy the token immediately** – it will not be shown again.

Keep this token private. Never commit it to public repositories.

## 🧪 Troubleshooting

- **“Erro 401” / Unauthorized**: Your token is invalid or expired. Generate a new one.
- **Course not found**: Double‑check the `COURSES` IDs. The IDs must match exactly the course numbers in the Canvas URL.
- **Empty sheets**: The course might have no published assignments, or all assignments may be set as “omit from final grade”.
- **GitHub Actions fails to push**: Ensure the workflow has `permissions: write` for contents and that branch protection rules (if any) allow the push.

## 📝 Notes

- The script and output are in **Portuguese**, matching PUCPR’s language.  
  - `Resumo` = Summary, `Pendentes` = Pending, `Nao_Contabilizadas` = Not Counted, `Atividade` = Assignment, `Obtido` = Obtained, `Máximo` = Max.
- The projection formula assumes you will maintain the same performance on future assignments as you have on already graded ones. This is an estimate, not an official grade.
- This project is not affiliated with PUCPR or Instructure.

## 📬 Contributing

Feel free to open issues or pull requests if you find a bug or want to add a feature (e.g., email notifications, a web dashboard, support for multiple users).

---

**Built with ❤️ for PUCPR students who want to stay on top of their grades.**
