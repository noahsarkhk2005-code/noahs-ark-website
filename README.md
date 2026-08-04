# Noah's Ark Official Website

[![Netlify Status](https://api.netlify.com/api/v1/badges/e708e6a7-5592-4403-a036-487e4f8a85d9/deploy-status)](https://app.netlify.com/projects/noahsarkhk2005/deploys)

香港現場音樂 / 樂隊文化推廣平台官方網站（靜態版）

**Live:** https://noahsarkhk2005.netlify.app  
**Repo:** https://github.com/noahsarkhk2005-code/noahs-ark-website

## 技術
- 純 HTML + 建置後 Tailwind CSS（約 15KB）
- WebP + PNG fallback 圖片
- Tally form embed
- 中英雙語切換
- IntersectionObserver 淡入動畫

## 頁面
- `index.html` — 首頁
- `tickets-noah-2-5.html` — Noah's Ark 2.5 購票
- `merch.html` — 商品頁

## 本地預覽
直接用瀏覽器打開 `index.html`，或：

```bash
npx serve .
```

## 重新建置 CSS（如有修改 HTML class）
```bash
npx tailwindcss -i ./input.css -o ./assets/style.css --minify
```

## 部署
- **Production:** push to `main` → Netlify auto-deploys  
- **Manual:** `netlify deploy --prod --dir .`

Site: `noahsarkhk2005` · Branch: `main` · Publish dir: `.`

## 官方 Instagram
https://www.instagram.com/noahs_ark_hk/
