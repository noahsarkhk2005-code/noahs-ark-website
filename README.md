# Noah's Ark Official Website

香港極端金屬活動平台官方網站（靜態版）

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
直接用瀏覽器打開 `index.html`

## 重新建置 CSS（如有修改 HTML class）
```bash
/tmp/twcss -i ./input.css -o ./assets/style.css --minify
# 或使用 npx tailwindcss
```

## 部署到 Netlify
1. 打開 https://app.netlify.com/drop
2. 把 zip 拖進去（確保 index.html 在根目錄）
3. 獲得免費網址

## 官方 Instagram
https://www.instagram.com/noahs_ark_hk/
