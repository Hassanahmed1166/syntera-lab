# Deploying syntera.au on Cloudflare Pages

The site is static. Build output is committed in `docs/`, so no build step is needed.

## 1. Create the Pages project
Cloudflare dashboard → Workers & Pages → Create → Pages → Connect to Git → pick `hassanahmed1166/syntera-lab`.
- Production branch: `main`
- Framework preset: None
- Build command: (leave empty)
- Build output directory: `docs`

## 2. Connect the domain
Project → Custom domains → Set up a domain → `syntera.au`, then add `www.syntera.au`.
Easiest route is to move DNS to Cloudflare:
1. Cloudflare → Add a site → `syntera.au` (Free plan). It shows two nameservers.
2. VentraIP → My Services → Domains → syntera.au → Nameservers → replace with the two Cloudflare ones.
3. Wait for Cloudflare to report the domain Active (minutes to a few hours).
4. Add the domain in the Pages project; Cloudflare creates the DNS records automatically.

## 3. Redirect www → apex
Cloudflare → syntera.au → Rules → Redirect Rules → hostname equals `www.syntera.au` → dynamic redirect to
`concat("https://syntera.au", http.request.uri.path)`, status 301.
Also enable SSL/TLS → Edge Certificates → Always Use HTTPS.

## 4. Regenerate after content changes
`python build_site.py`, commit `docs/`, push to `main`. Pages redeploys automatically.

## 5. After going live
- Add the site in Google Search Console and submit `https://syntera.au/sitemap.xml`.
- Update the Google Scholar / LinkedIn / UniSC profile links to the new URL.
- Optionally set the old GitHub Pages site to redirect to syntera.au.
