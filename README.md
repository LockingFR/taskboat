# Taskboat

Exportez gratuitement vos listes Microsoft To Do avant leur retrait (30 juin 2027) : JSON, CSV, Todoist,
Markdown (Obsidian), ICS. Tout se passe dans le navigateur, aucune tâche ne passe par un serveur.

*Free, in-browser Microsoft To Do exporter (JSON, CSV, Todoist, Markdown, ICS). Your tasks never leave your browser.*

## Lancer en local
```bash
python -m http.server 8000   # puis http://localhost:8000
```

## Tester
```bash
pip install playwright && playwright install chromium
python tests/run_test.py
```

## Déployer
Site statique sur Cloudflare Pages, sans build. Voir `CLAUDE.md` pour l'application Azure (connexion
Microsoft) et les marqueurs à remplacer.

Taskboat n'est pas affilié à Microsoft. Licence MIT.
