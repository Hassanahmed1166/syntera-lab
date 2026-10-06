# SYNTERA SEO playbook

What is already built into the site (rebuild with `python build_site.py` from `D:\SYNTERA`) and what only you can do.

## Done in code (on-page + technical)
- New page **/collaborate**: academic collaboration, joint grants, commissioned/paid projects, industry pilots, student projects, FAQ (FAQPage + Service schema).
- Keyword-aware titles and meta descriptions on every page; collaboration call-to-action on home, research, each research area, About, Join.
- Structured data: Organization + WebSite (home), BreadcrumbList (all pages), Person (every profile), FAQPage + Service (collaborate).
- Publications are now server-rendered (all 155 papers readable by crawlers, with DOI links), not JavaScript-only.
- Clean internal links (`/about` not `/about.html`), canonical URLs, `en-AU`, robots meta with large-snippet permission, sitemap with priorities, `robots.txt`, `llms.txt` for AI assistants, HTTPS + security headers.

## Search Console and Bing (you, 15 minutes, free) — do this first
1. Google Search Console: add property `syntera.au` (Domain), verify with the DNS TXT record in Cloudflare DNS, submit `https://syntera.au/sitemap.xml`, then "Request indexing" for `/`, `/collaborate`, `/research`, `/publications`.
2. Bing Webmaster Tools: import from Search Console, submit the sitemap.
3. Check "Page indexing" weekly; fix anything listed under "Why pages aren't indexed".

## Off-page: backlinks and entity signals (this is what moves rankings)
Search engines rank sites other trusted sites point to. Ask for a link to `https://syntera.au` from each of:
- Director/co-director/co-founder staff pages at the University of the Sunshine Coast and each member's own university profile ("Research group: SYNTERA, syntera.au").
- ORCID, Google Scholar, ResearchGate, LinkedIn, GitHub, X/Bluesky bios of all members: add the website field.
- A **LinkedIn company page** and a Google Scholar / ResearchGate lab page for SYNTERA.
- Partner institutions' pages that mention the collaboration (the partners listed on the site).
- UniSC news / media office: a short news item about the group launch or the LAUNCH Partnership Grant, linking to the site.
- Conference and journal pages: put "SYNTERA Research Group, syntera.au" in affiliation/acknowledgement lines of new papers.
- Directories: Queensland and Australian research/industry-engagement directories, university "find an expert" pages, AI Australia / ACS / IEEE chapters, Wikidata item for the group.
Quality over quantity: 10 relevant academic or institutional links beat 500 directory links. Never buy links.

## Content to keep adding (rankings follow content)
- One page or post per funded project, partner or case study ("AI knee MRI segmentation with ...").
- News/updates page with each new paper, grant, talk. Publish at least monthly.
- Project and area pages in plain language: what problem, which partners, what outcome.
- Add a physical-world signal where truthful: UniSC address and a Google Business Profile only if the group has a public location.

## Target searches (realistic)
Single words such as "funding" or "projects" are dominated by governments and portals and cannot be won. Target phrases people with budgets actually search:
`academic collaboration AI Australia`, `industry partnership university AI Queensland`, `commission AI research project university`, `AI research partner for grant application`, `university research consultancy machine learning`, `IoT research collaboration`, `AI for health research partner Australia`, `smart home ageing AI research`, `precision agriculture AI research partner`, `learning analytics research collaboration`.
Realistic timeline: indexed in days, long-tail rankings in 4-12 weeks, competitive phrases 6-12 months and depends on backlinks.

## Facts to keep accurate
Claims on /collaborate (paid projects, grants, IP terms in writing) describe how the group offers to work. Confirm with UniSC research office that contract research and fees go through the university's process before advertising prices.
