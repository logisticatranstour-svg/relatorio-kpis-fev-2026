# relatorio-kpis-fev-2026

Dashboard e relatório HTML de KPIs de satisfação de entregas (Transtour x Hapvida) com publicação via GitHub Pages e atualização automática diária.

## ✅ Link do Dashboard (GitHub Pages)
Acesse: https://logisticatranstour-svg.github.io/relatorio-kpis-fev-2026/

---

## 🔄 Como funciona (fluxo automático)
1) Paciente responde no **Tally**
2) Respostas são salvas no **Google Sheets**
3) Um workflow do **GitHub Actions** roda **todo dia às 23:59 (Brasil)**:
   - baixa o CSV publicado do Sheets
   - aplica a regra “Recusou responder = Sim + Nota 5”
   - gera o `index.html` com seletor de mês
   - faz commit automaticamente
4) O **GitHub Pages** publica a versão mais recente no mesmo link

---

## ▶️ Como atualizar manualmente (sem esperar o horário)
1) Vá em **Actions**
2) Abra **Atualizar dashboard diariamente**
3) Clique em **Run workflow**
4) Aguarde ficar verde ✅
5) Abra o link do Pages e dê **Ctrl + F5** (atualização forçada)

---

## 🖼️ Logos (como trocar)
Os arquivos de logo ficam na raiz do repositório:
- `logo-transtour.png`
- `logo-cliente.png`

Para trocar:
1) **Add file → Upload files**
2) Envie as novas imagens com os mesmos nomes
3) Commit changes
4) Rode o workflow (opcional) ou aguarde o horário

---

## 📌 Fonte de dados (Sheets publicado)
O script usa o CSV publicado do Google Sheets:
`SHEET_CSV_URL` dentro do arquivo `build_dashboard.py`.

---

## ⚙️ Arquivos importantes
- `build_dashboard.py` → gera o `index.html` automaticamente
- `.github/workflows/atualizar-dashboard.yml` → agenda a atualização diária
- `index.html` → gerado automaticamente (não editar manualmente)

---

## 🧠 Regra acordada com o cliente
Se houver “Recusou responder”:
- critérios operacionais são tratados como **Sim**
- satisfação é tratada como **Nota 5**
(conforme alinhamento operacional com o cliente)
