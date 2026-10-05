# SYNTERA Lab website

Static website for SYNTERA Lab (Applied AI & Connected Systems), University of the Sunshine Coast.

- `site/` – the built website (served by GitHub Pages)
- `build_site.py` – generator (pages, people, publications data)
- `import_pubs.py`, `fetch_pub_meta.py` – publication import and DOI/keyword/abstract lookup
- `pub_meta.json`, `pubs_extra.json` – publication data

Preview locally: `python -m http.server --directory site`
