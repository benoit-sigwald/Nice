# Nice, capitale de l'IA — instructions

Plan métropolitain d'efficience budgétaire, de souveraineté et de sécurité urbaine par l'IA, pour la **Ville de Nice et la Métropole Nice Côte d'Azur (MNCA)**. Auteur : Benoît Sigwald, Senior AI Architect & AMO IA métropolitain. Août 2026. Dépôt `Nice.git`.

> Le dossier s'appelait `Dev/Einstein` — nom hérité et trompeur, **renommé en `Dev/Nice` le 14/08/2026**. Rien ici ne concerne le projet Einstein, qui vit dans `G:\My Drive\Einstein\` (livrables) et `G:\My Drive\Einstein-mcp\` (serveur MCP).

## Référence
`README.md` puis `docs/`.

## Nature du projet
Livrable institutionnel destiné à une collectivité territoriale. Chaîne de production : `generate_charts.py` → `generate_docx.py` / `generate_pdf.py` → `Le_Pacte_Nice_IA.docx` / `.pdf`. `auto_translate_watcher.py` gère la version traduite.

## Règles
- **Document destiné à une collectivité publique** : aucun chiffre, aucune économie budgétaire, aucun gain annoncé sans source citée. Une estimation non sourcée engage l'auteur auprès d'une institution.
- Les `.docx`, `.pdf` et `.png` de la racine sont des **artefacts régénérés** par les scripts. Modifier la source, jamais l'artefact.
- 33 Mo, surtout des assets et des rendus : ne pas lancer de `Glob`/`Grep` récursif à la racine.

## Style de réponse
3-6 lignes max. Donner le résultat, s'arrêter.
