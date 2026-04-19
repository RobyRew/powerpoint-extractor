# App Icons

Place the following icon files here for PWA support:

- `icon-32.png` — 32×32 favicon
- `icon-192.png` — 192×192 PWA icon
- `icon-512.png` — 512×512 PWA icon
- `icon-maskable-512.png` — 512×512 maskable icon (with safe zone padding)

Generate from `../favicon.svg` using any SVG-to-PNG tool or:

```bash
# Using sharp-cli (npm install -g sharp-cli)
npx sharp -i ../favicon.svg -o icon-32.png resize 32
npx sharp -i ../favicon.svg -o icon-192.png resize 192
npx sharp -i ../favicon.svg -o icon-512.png resize 512
npx sharp -i ../favicon.svg -o icon-maskable-512.png resize 512
```
