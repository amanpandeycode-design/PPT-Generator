# AI Deck Generator — Vercel Deploy

Same tool jo pehle tha, ab Vercel pe deploy karne ke liye restructured:
- `api/index.py` — Flask app, Vercel serverless Python function ban jaata hai
- `public/index.html` — static frontend, Vercel automatically serve karta hai
- `vercel.json` — dono ko connect karta hai (`/api/*` → Python function, baaki sab → static file)

## Deploy karne ke steps

### Option A: Vercel dashboard se (easiest)
1. Is poore folder ko GitHub repo me push karo
2. [vercel.com](https://vercel.com) pe jaao → "Add New Project" → apna repo import karo
3. Framework preset: **Other** choose karo (auto-detect na ho to)
4. **Environment Variables** me jaake add karo:
   - `GEMINI_API_KEY` = apni free Gemini key ([aistudio.google.com/apikey](https://aistudio.google.com/apikey))
5. Deploy dabao

### Option B: Vercel CLI se
```bash
npm install -g vercel
cd ppt-generator-vercel
vercel login
vercel env add GEMINI_API_KEY    # apni key paste karo jab poochhe
vercel --prod
```

## ⚠️ Important cheezein jo dhyan me rakhni hain

1. **Function timeout**: Vercel ke free (Hobby) plan pe default serverless function timeout **10 seconds** hota hai. Outline generation (Gemini call) + image generation (Pollinations) mila ke kabhi-kabhi isse zyada time le sakta hai, especially cold start pe. Agar timeout error aaye:
   - Vercel dashboard → Project → Settings → Functions me jaake max duration badhao (Hobby plan pe limited hai, Pro plan pe 60s+ mil sakta hai)
   - Ya `num_slides` kam rakho taaki generation fast ho

2. **Cold starts**: Pehli request thodi slow ho sakti hai (function "wake up" hone me time leta hai) — ye normal hai serverless ka.

3. **Ye maine test nahi kiya hai end-to-end** — mere paas Vercel tak network access nahi hai is environment me, isliye syntax aur structure verify kiya hai (Python compiles clean, JSX valid hai, config JSON valid hai), lekin actual live deployment khud verify nahi kar saka. Agar deploy karte waqt koi error aaye — poora error message paste kar dena, turant fix karenge.

4. **Local testing** ab bhi kaam karta hai bina Vercel ke:
   ```bash
   cd api
   pip install -r requirements.txt
   export GEMINI_API_KEY=your-key
   python index.py
   ```
   Phir `public/index.html` seedha browser me kholo — ye automatically `localhost:5000` use karega jab file:// se khola jaaye.
