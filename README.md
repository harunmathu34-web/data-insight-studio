# Insight Studio — setup guide

This is a Vite + React site. Same stack you already used for Veleminour and
Carlos Graphics, so the workflow will feel familiar — the only new piece is
the AI assistant, which needs its own API key (explained in Step 5).

## 1. Get the project onto your machine

1. Unzip this folder wherever you keep your projects, e.g.
   `C:\Users\Wambugu\Desktop\HARUN MATHU\projects\data-insight-studio`
2. Open VS Code → File → Open Folder → select that folder.
3. Open a terminal inside VS Code: Terminal menu → New Terminal.

## 2. Install and run it locally

In the VS Code terminal, type:

```
npm install
npm run dev
```

Wait for it to print a `Local:` link (something like `http://localhost:5173`),
then Ctrl+click that link — it opens in your browser. Try uploading the
`duka_sales_practice_data.xlsx` file and watch the charts build.

**Note:** the AI Assistant tab and "Generate AI insights" button will NOT
work yet at this step — they need Step 5 first, and even then they only work
once deployed (or once you run `vercel dev` instead of `npm run dev`, see
Step 6). Everything else — upload, auto-charts, data quality — works right
now.

## 3. Put it on GitHub

Same as veleminour-studios:

1. Go to github.com → New repository → name it `data-insight-studio` →
   Create repository (leave it empty, no README).
2. Back in the VS Code terminal:

```
git init
git add .
git commit -m "Initial commit — Insight Studio"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/data-insight-studio.git
git push -u origin main
```

Replace `YOUR-USERNAME` with your actual GitHub username.

## 4. Deploy to Vercel

1. Go to vercel.com → sign in with GitHub.
2. Click "Add New" → "Project".
3. Pick `data-insight-studio` from your repo list → Import.
4. Vercel auto-detects it's a Vite project — leave the build settings as
   default.
5. **Before clicking Deploy**, open "Environment Variables" on that same
   screen and add:
   - Name: `ANTHROPIC_API_KEY`
   - Value: (your key from Step 5 below)
6. Click Deploy. In about a minute you'll get a live URL like
   `data-insight-studio.vercel.app`.

## 5. Get an Anthropic API key (needed for the AI Assistant to work)

This is different from your normal Claude.ai chat — it's a separate,
pay-as-you-go API account, billed by usage (typically fractions of a cent
per question asked, for a tool like this).

1. Go to console.anthropic.com and sign up / log in.
2. Add a payment method under Billing (it won't charge anything until the
   AI assistant is actually used).
3. Go to "API Keys" → "Create Key" → copy the key (starts with `sk-ant-`).
4. Paste that into Vercel's environment variable in Step 4.5 above.
5. If you already deployed without it: go to your Vercel project → Settings
   → Environment Variables → add it there → then Deployments → "..." on the
   latest one → Redeploy.

## 6. (Optional) Test the AI assistant locally before deploying

```
npm install -g vercel
vercel login
vercel link
vercel env pull .env.local
vercel dev
```

This runs the site AND the `/api/chat` function together locally, so you
can test the AI Assistant tab before pushing live.

## What's in this folder

- `src/DataInsightStudio.jsx` — the whole app: upload, auto-charts, AI chat
- `api/chat.js` — the one server file: holds your API key safely and talks
  to Anthropic on the browser's behalf (the browser never sees your key)
- everything else is standard Vite/React/Tailwind scaffolding

## Next features to add (not built yet)

- Kenya map view (Performance by County/Site)
- PDF/report export
- Multi-sheet Excel support
- Saved dashboards / multiple datasets at once

Tell Claude which one you want next and we'll build it into this same repo.
