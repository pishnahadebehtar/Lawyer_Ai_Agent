
# ⚖️ Vakil Jibi (وکیل جیبی) — Open-Source AI Legal Assistant

[![Website](https://img.shields.io/badge/Website-vakiljibi.ir-yellow?style=for-the-badge)](https://vakiljibi.ir)
[![Cafe Bazaar](https://img.shields.io/badge/Download-Cafe_Bazaar-green?style=for-the-badge)](https://cafebazaar.ir/app/ir.vakiljibi)
[![GitHub License](https://img.shields.io/github/license/pishnahadebehtar/Lawyer_Ai_Agent?style=for-the-badge)](LICENSE)
[![GitHub Stars](https://img.shields.io/github/stars/pishnahadebehtar/Lawyer_Ai_Agent?style=for-the-badge)](https://github.com/pishnahadebehtar/Lawyer_Ai_Agent/stargazers)

**Vakil Jibi (Pocket Lawyer)** is an open-source, production-ready AI Legal Agent specifically designed for the Iranian legal system. Using Retrieval-Augmented Generation (RAG) and advanced LLM architectures (Gemini 2.5 Flash), Vakil Jibi analyzes Persian legal inquiries, searches the official database of Iranian laws, and generates custom legal drafts.

The system is fully cross-platform, offering a Next.js PWA frontend, an Android mobile application, and a Telegram Bot interface.

---

## 🚀 Live Services & Portals (SEO Directory)

You can explore the live, free-to-use legal portals powered by the Vakil Jibi engine below:

*   **[مشاوره حقوقی هوش مصنوعی (Free Legal AI Consulting)](https://vakiljibi.ir/)** — Our main landing page and web application portal.
*   **[وکیل هوش مصنوعی چک برگشتی (Dishonored Checks)](https://vakiljibi.ir/cheque)** — AI drafting and consulting for bounced or commercial checks.
*   **[سفته و تعهدات مالی (Promissory Notes)](https://vakiljibi.ir/promissory-note)** — Dispute resolution and legal guides for promissory notes.
*   **[تنظیم آنلاین انواع قرارداد (Online Contract Drafting)](https://vakiljibi.ir/contracts)** — Automatic custom contract generation and term analysis.
*   **[قوانین و مطالبه مهریه (Dowry/Marriage Laws)](https://vakiljibi.ir/mehrieh)** — Financial marriage calculations and asset seizure pathways.
*   **[مراحل طلاق و حقوق خانواده (Divorce & Custody)](https://vakiljibi.ir/talagh)** — Mutual and unilateral family dispute guides.
*   **[طلاق توافقی با هوش مصنوعی (Mutual Divorce Helper)](https://vakiljibi.ir/divorce)** — Seamless drafting of family agreements.
*   **[انحصار وراثت آنلاین (Inheritance & Probate)](https://vakiljibi.ir/enhesar-verasat)** — Legal estate divisions and heir share calculators.
*   **[شکایت از کارفرما اداره کار (Labor & Employment)](https://vakiljibi.ir/labor-complaint)** — Dispute filings, wage calculations, and arbitration help.
*   **[تخلیه ملک و اختلافات اجاره (Rental & Eviction)](https://vakiljibi.ir/property)** — Immediate eviction notifications and rent collection guides.
*   **[تصرف عدوانی ملکی (Forcible Property Entry)](https://vakiljibi.ir/tasarof-odvani)** — Civil and criminal legal pathways to reclaim property.
*   **[شکایت کلاهبرداری اینترنتی (Cybersecurity & Fraud)](https://vakiljibi.ir/cyber-fraud)** — Emergency bank blocking procedures and cyber complaint templates.
*   **[خیانت در امانت (Breach of Legal Trust)](https://vakiljibi.ir/breach-of-trust)** — Proof of deposit, demand notices, and official complaints.
*   **[توقیف اموال و تامین خواسته (Asset Seizure)](https://vakiljibi.ir/toqif-amval)** — Asset freezing, bank freezes, and legal notices.
*   **[اعسار و تقسیط بدهی (Insolvency & Installments)](https://vakiljibi.ir/easar)** — Court petition files for insolvencies.

---

## 🌟 Key Features

*   **RAG Engine connected to Iranian Law:** Directly fetches relevant clauses from the Civil Code (قانون مدنی), Penal Code (قانون مجازات اسلامی), Commercial Code (قانون تجارت), and Supreme Court Rulings (آرای وحدت رویه).
*   **Automated Document Drafting:** Generates customized legal petitions, defense bills, complaints, and official contracts in `.docx` (Microsoft Word) formats based on conversational Persian inputs.
*   **Source Transparency:** Generates a companion `.xlsx` (Microsoft Excel) spreadsheet with each legal consultation, mapping every claim to its exact legal citation, definitions, and court cases.
*   **Voice & File Analysis:** Accepts Persian voice commands (speech-to-text) and supports PDF/Word/Excel document analysis for contract reviews.
*   **No Sign-up / 100% Free:** No login steps required, ensuring private access to legal information.

---

## 🛠️ System Architecture

```text
               +----------------------------------------+
               |          CLIENT INTERFACES             |
               | (Next.js PWA / Android / Telegram Bot) |
               +-------------------+--------------------+
                                   |
                                   v  (JSON / REST API)
                       +-----------+-----------+
                       |   Nginx Reverse Proxy |
                       +-----------+-----------+
                                   |
                                   v
                       +-----------+-----------+
                       |      Appwrite /       |
                       |   Supabase Backend    |
                       +-----------+-----------+
                                   |
                     +-------------+-------------+
                     |                           |
                     v                           v
         +-----------+-----------+   +-----------+-----------+
         |     Gemini AI Brain   |   |   PostgreSQL / Vector |
         |   (LLM Reasoning)     |   |    (pgvector Store)   |
         +-----------------------+   +-----------------------+

Tech Stack

  - Frontend: Next.js (App Router), React, Material UI (MUI).
  - Backend & DB: Supabase & Appwrite for database operations, authentication,
    and vector embeddings.
  - AI Integration: Gemini API (using Gemini 1.5 & 2.5 Flash engines).
  - RAG Source Parsing: Python FastAPI microservice (handling document
    splitting, vectorization, and legal databases).

💻 Getting Started Locally

Prerequisites

  - Node.js (v18.x or higher)
  - Python 3.10+ (for RAG and parsing services)
  - Supabase or Appwrite project with PGVector enabled
  - Gemini API Key

Installation

1.  Clone the repository:

    git clone https://github.com/pishnahadebehtar/Lawyer_Ai_Agent.git
    cd Lawyer_Ai_Agent

2.  Configure environment variables: Create a .env.local in your root folder:

    NEXT_PUBLIC_SUPABASE_URL=your_supabase_url
    NEXT_PUBLIC_SUPABASE_ANON_KEY=your_supabase_anon_key
    GEMINI_API_KEY=your_gemini_api_key

3.  Install dependencies and run the Next.js frontend:

    npm install
    npm run dev

    Open http://localhost:3000 to inspect the local deployment.

🤝 Contributing

We welcome contributions from legal scholars, developers, and UI/UX designers
[2].

1.  Fork this repository.
2.  Create your feature branch (git checkout -b feature/AmazingFeature).
3.  Commit your changes (git commit -m 'Add some AmazingFeature').
4.  Push to the branch (git push origin feature/AmazingFeature).
5.  Open a Pull Request.

📄 License

This project is licensed under the MIT License - see the LICENSE file for
details.


---

### Step 2: How to Deploy This `README.md` to GitHub

You have two easy ways to update your GitHub repository with this new file:

#### Method A: Using the GitHub Website Interface (Easiest & Quickest)
1. Go to your repository page: `https://github.com/pishnahadebehtar/Lawyer_Ai_Agent`.
2. Click on the **`README.md`** file in your repository's file list.
3. Click on the **Pencil Icon** (Edit this file) in the top-right corner of the file viewer.
4. Delete all existing content, and paste the entire Markdown block from Step 1 above.
5. Click **Commit changes...** at the top right, write a short commit message (e.g., `docs: update roadmap, add SEO links and responsive design details`), and save.

#### Method B: Using your Local Git & PowerShell (Standard Dev Workflow)
If you manage the repository via Git locally, open your PowerShell inside `E:\Vakiljibi\vakiljibi-main`:

1. Open your local project's `README.md` file in VS Code or any text editor and paste the content [2].
2. In PowerShell, commit and push [2]:
   ```powershell
   git add README.md
   git commit -m "docs: update sitemap, dynamic pages, and add SEO backlink directory"
   git push origin main

