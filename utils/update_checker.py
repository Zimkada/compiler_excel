"""
Vérification de mise à jour via l'API GitHub Releases.

L'application interroge la dernière Release publiée du dépôt et compare son tag
à la version courante. La vérification est non-bloquante et silencieuse en cas
d'échec (hors-ligne, dépôt privé, API indisponible) : elle ne doit JAMAIS
empêcher d'utiliser l'application.

Ce module contient la logique PURE (comparaison de versions, parsing de la
réponse) testable sans réseau. L'appel réseau proprement dit est effectué par
un worker Qt séparé (ui/workers/update_worker.py).
"""

import json
import re
from dataclasses import dataclass
from typing import Optional, Tuple


# Dépôt hébergeant les Releases. L'API renvoie la dernière release publiée.
GITHUB_OWNER = "Zimkada"
GITHUB_REPO = "compiler_excel"
LATEST_RELEASE_URL = (
    f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
)
RELEASES_PAGE_URL = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"


@dataclass
class UpdateInfo:
    """Résultat d'une vérification de mise à jour."""
    update_available: bool
    latest_version: str = ""
    current_version: str = ""
    download_url: str = ""      # page de la release (où télécharger)


def parse_version(text: str) -> Optional[Tuple[int, ...]]:
    """Extrait un tuple de version comparable depuis une chaîne.

    Tolère les préfixes courants (« v3.2 », « 3.2.0 », « version 3.2 ») et
    ignore tout suffixe non numérique (« 3.2.1-beta » -> (3, 2, 1)). Renvoie
    None si aucune version numérique n'est trouvée.
    """
    if not text:
        return None
    # Premier groupe de type X.Y ou X.Y.Z… dans la chaîne.
    match = re.search(r'(\d+(?:\.\d+)*)', str(text))
    if not match:
        return None
    try:
        return tuple(int(p) for p in match.group(1).split('.'))
    except ValueError:
        return None


def is_newer(latest: str, current: str) -> bool:
    """Vrai si ``latest`` est une version strictement plus récente que
    ``current``. Comparaison numérique par composant, robuste aux longueurs
    différentes (« 3.2 » vs « 3.2.1 »). En cas de version illisible, renvoie
    False (on ne propose jamais une mise à jour douteuse).
    """
    lv = parse_version(latest)
    cv = parse_version(current)
    if lv is None or cv is None:
        return False
    # Aligner les longueurs en complétant par des zéros (3.2 -> 3.2.0).
    length = max(len(lv), len(cv))
    lv += (0,) * (length - len(lv))
    cv += (0,) * (length - len(cv))
    return lv > cv


def evaluate_release(response_body: str, current_version: str) -> UpdateInfo:
    """Analyse le corps JSON d'une réponse GitHub `releases/latest` et décide
    s'il existe une mise à jour. Silencieux (update_available=False) sur toute
    donnée absente ou malformée — jamais d'exception propagée.

    Args:
        response_body: Corps brut de la réponse HTTP (JSON GitHub).
        current_version: Version actuelle de l'application (ex. « 3.2 »).
    """
    try:
        data = json.loads(response_body)
    except (ValueError, TypeError):
        return UpdateInfo(update_available=False, current_version=current_version)

    # GitHub renvoie tag_name (ex. « v3.3 ») ; on tolère aussi « name ».
    tag = data.get("tag_name") or data.get("name") or ""
    html_url = data.get("html_url") or RELEASES_PAGE_URL

    if not tag or not is_newer(tag, current_version):
        return UpdateInfo(update_available=False,
                          latest_version=str(tag),
                          current_version=current_version,
                          download_url=html_url)

    return UpdateInfo(update_available=True,
                      latest_version=str(tag),
                      current_version=current_version,
                      download_url=html_url)
