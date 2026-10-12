#!/usr/bin/env python3
"""
Contrôle et delta du fichier TVA généré par le spécifique Sage X3.

Deux volets :

  1) controle : contrôle d'un seul fichier (par défaut file.csv à côté du script)
       python tva_check.py controle
       python tva_check.py controle -o rapport_controle.xlsx

  2) delta : comparaison ancien / nouveau fichier (détection de régressions)
       python tva_check.py delta ancien.csv nouveau.csv -o rapport_delta.xlsx

Dépendances : pandas, openpyxl   (pip install pandas openpyxl)
"""
import argparse
import sys
from pathlib import Path

import pandas as pd

# --- Colonnes ---------------------------------------------------------------
COL_PIECE = "N° pièce règlement"
COL_LIGNE = "Ligne règlement"
COL_REGLE = "Montant réglé"
COL_TVA = "TVA à déclaré"
COL_SIGNE = "Montant signé"
COL_TYPE_TVA = "Type de TVA"
COL_TYPE_REGL = "Type règlement"
COL_REMARQUE = "Remarque"

NUM_COLS = [COL_REGLE, COL_TVA, COL_SIGNE]
REQUIRED = [COL_PIECE, COL_LIGNE, COL_REGLE, COL_TVA, COL_SIGNE,
            COL_TYPE_TVA, COL_TYPE_REGL, COL_REMARQUE]

# --- Valeurs métier ---------------------------------------------------------
T_VENTE = "Collectée ventes"
T_ACHAT = "Déductible achats"
T_IMMO = "Déductible immos"

R_LETTRE = "RAPPROCHE LETTRE"
R_IMPAY = "IMPAY"
R_ETAPE = "ETAPE INTERMEDIAIRE NON LETTRE"
R_RECLASS = "Reclassement multiple"   # lignes sans TVA, à ajouter au détail
R_FRAIS = "FRAIS BANCAIRE"        # côté achat, mais à récupérer pour contrôler les ventes
REM_VENTE = [R_LETTRE, R_IMPAY, R_ETAPE]

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_FILE = BASE_DIR / "file.csv"   # fichier contrôlé par défaut : file.csv à la racine du projet

TOL = 0.01  # tolérance sur les comparaisons de montants


# --- Utilitaires ------------------------------------------------------------
def to_num(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip().str.replace(" ", "", regex=False)
    s = s.str.replace(" ", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(s.replace("", "0"), errors="raise").fillna(0.0)


def load(path: Path) -> pd.DataFrame:
    for enc in ("utf-8-sig", "cp1252"):
        try:
            df = pd.read_csv(path, sep=";", dtype=str, encoding=enc, keep_default_na=False)
            break
        except UnicodeDecodeError:
            continue
    else:
        sys.exit(f"Impossible de décoder {path}")

    df.columns = [c.strip() for c in df.columns]
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        sys.exit(f"{path.name} : colonnes manquantes {missing}")

    for c in df.columns:
        df[c] = df[c].str.strip()
    for c in NUM_COLS:
        df[c] = to_num(df[c])
    return df


def fmt(x) -> str:
    """Montant au format français : 1 234 567,89"""
    if pd.isna(x):
        return ""
    return f"{x:,.2f}".replace(",", " ").replace(".", ",")


def fmt_int(x) -> str:
    return "" if pd.isna(x) else f"{int(x):,}".replace(",", " ")


def table(df: pd.DataFrame, money=(), ints=(), max_text: int = 60) -> str:
    """Tableau avec bordures (caractères box-drawing). Nombres à droite, texte à gauche."""
    cols = list(df.columns)
    money, ints = set(money), set(ints)

    def kind(c):
        if c in money:
            return "money"
        if c in ints:
            return "int"
        s = df[c]
        if pd.api.types.is_bool_dtype(s) or not pd.api.types.is_numeric_dtype(s):
            return "text"
        nn = s.dropna()
        return "int" if (nn == nn.round()).all() and nn.abs().max() < 1e6 and c.startswith(("nb_", "delta_nb", "Lignes")) else "money"

    kinds = [kind(c) for c in cols]
    rows = []
    for _, r in df.iterrows():
        out = []
        for c, k in zip(cols, kinds):
            v = r[c]
            if k == "money":
                out.append(fmt(v))
            elif k == "int":
                out.append(fmt_int(v))
            else:
                t = "" if pd.isna(v) else str(v)
                out.append(t if len(t) <= max_text else t[: max_text - 1] + "…")
        rows.append(out)

    widths = [max([len(str(c))] + [len(r[i]) for r in rows]) for i, c in enumerate(cols)]

    def sep(l, m, r):
        return l + m.join("─" * (w + 2) for w in widths) + r

    def line(vals, header=False):
        cells = []
        for v, w, k in zip(vals, widths, kinds):
            if header:
                cells.append(" " + str(v).center(w) + " ")
            elif k == "text":
                cells.append(" " + v.ljust(w) + " ")
            else:
                cells.append(" " + v.rjust(w) + " ")
        return "│" + "│".join(cells) + "│"

    lines = [sep("┌", "┬", "┐"), line(cols, header=True), sep("├", "┼", "┤")]
    lines += [line(r) for r in rows]
    lines.append(sep("└", "┴", "┘"))
    return "\n".join(lines)


# =============================================================================
# VOLET 1 : contrôle d'un seul fichier
# =============================================================================
def resume_tva(df: pd.DataFrame):
    """TVA à déclarer : ventes, achats (déductible achats + immos fusionnés), net."""
    t = df[COL_TYPE_TVA]
    ventes = df[t == T_VENTE]
    achats = df[t.isin([T_ACHAT, T_IMMO])]
    v = round(ventes[COL_TVA].sum(), 2)
    a = round(achats[COL_TVA].sum(), 2)
    out = pd.DataFrame(
        [("Ventes (collectée)", len(ventes), v),
         ("Achats (déductible achats + immos)", len(achats), a),
         ("TVA nette (ventes - achats)", None, round(v - a, 2))],
        columns=["Type de TVA", "Lignes", "TVA à déclarer"],
    )
    # lignes sans Type de TVA mais avec de la TVA : elles sortiraient du récapitulatif
    sans_type = df[(t == "") & (df[COL_TVA].abs() > TOL)]
    return out, sans_type


def detail_tva(df: pd.DataFrame) -> pd.DataFrame:
    """TVA à déclarer par Type de TVA, sans fusion (y compris le type vide)."""
    g = df.groupby(COL_TYPE_TVA, sort=False)
    out = pd.DataFrame({"Lignes": g.size(), "TVA à déclarer": g[COL_TVA].sum().round(2)})
    ordre = [T_VENTE, T_ACHAT, T_IMMO, ""]
    out = out.reindex([t for t in ordre if t in out.index]
                      + [t for t in out.index if t not in ordre])
    out.index = out.index.map(lambda t: t if t else "(vide)")
    return out.rename_axis("Type de TVA").reset_index()


def tableau_tva(df: pd.DataFrame) -> pd.DataFrame:
    """Un seul tableau : détail par type, total achats (achats + immos) et TVA nette."""
    d = detail_tva(df)
    tva = d.set_index("Type de TVA")["TVA à déclarer"]
    lignes = d.set_index("Type de TVA")["Lignes"]
    v = tva.get(T_VENTE, 0.0)
    achats_types = [t for t in (T_ACHAT, T_IMMO) if t in tva.index]
    a = round(sum(tva[t] for t in achats_types), 2)
    extra = pd.DataFrame([
        ("Total achats (déductible achats + immos)",
         int(sum(lignes[t] for t in achats_types)), a),
        ("TVA nette (ventes - achats)", None, round(v - a, 2)),
    ], columns=d.columns)
    return pd.concat([d, extra], ignore_index=True)


def controle_ventes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Pour chaque règlement de la partie vente, compare Montant signé au détail :

        retenu = Σ Montant réglé des lignes vente RAPPROCHE LETTRE
               + Σ lignes vente IMPAY
               + Σ lignes vente ETAPE INTERMEDIAIRE NON LETTRE
               + Σ lignes Reclassement multiple (sans TVA) qui ne portent qu'une partie
                 du règlement (celles égales au montant complet sont des doublons)
               - Σ lignes FRAIS BANCAIRE (côté achat, prises sur tout le règlement)

        écart = |retenu| - |Montant signé|     (valeur absolue : les décaissements
                                                 sont signés négatifs)

    Seules ces remarques sont prises en compte ; les autres lignes du règlement
    sont ignorées.
    """
    is_vente = df[COL_TYPE_TVA] == T_VENTE
    scope = set(df.loc[is_vente & df[COL_REMARQUE].isin(REM_VENTE), COL_PIECE])
    d = df[df[COL_PIECE].isin(scope)]
    v = d[COL_TYPE_TVA] == T_VENTE
    rem = d[COL_REMARQUE]

    g = d.groupby(COL_PIECE, sort=True)
    res = pd.DataFrame(index=g.size().index)
    res["type_reglement"] = g[COL_TYPE_REGL].first()
    res["montant_signe"] = g[COL_SIGNE].first().round(2)
    res["signe_constant"] = g[COL_SIGNE].nunique() == 1

    def somme(mask: pd.Series) -> pd.Series:
        return (d[mask].groupby(COL_PIECE)[COL_REGLE].sum()
                .reindex(res.index, fill_value=0.0).round(2))

    res["rapproche_lettre"] = somme(v & (rem == R_LETTRE))
    res["impay"] = somme(v & (rem == R_IMPAY))
    res["etape_interm"] = somme(v & (rem == R_ETAPE))
    # Reclassement multiple : à ajouter seulement s'il porte une PARTIE du règlement.
    # Une ligne qui porte le montant complet (= |Montant signé|) double les lignes lettrées.
    partiel = (d[COL_REGLE].abs() - d[COL_SIGNE].abs()).abs() > TOL
    res["reclassement"] = somme((rem == R_RECLASS) & partiel)
    res["frais_bancaire"] = somme(rem == R_FRAIS)
    res["retenu"] = (res["rapproche_lettre"] + res["impay"] + res["etape_interm"]
                     + res["reclassement"] - res["frais_bancaire"]).round(2)
    res["ecart"] = (res["retenu"].abs() - res["montant_signe"].abs()).round(2)

    return res.reset_index()


def ecarts_ventes(res: pd.DataFrame) -> pd.DataFrame:
    ko = res[(res["ecart"].abs() > TOL) | (~res["signe_constant"])].copy()
    return ko.reindex(ko["ecart"].abs().sort_values(ascending=False).index)


def cmd_controle(args) -> None:
    if not args.fichier.exists():
        sys.exit(f"Fichier introuvable : {args.fichier}\n"
                 f"Place le CSV dans le dossier du projet sous le nom file.csv.")
    df = load(args.fichier)
    print(f"\nFichier : {args.fichier.name}  ({len(df)} lignes)")

    # 1. Total TVA par type
    resume, sans_type = resume_tva(df)
    tableau = tableau_tva(df)
    print("\n=== TOTAL TVA PAR TYPE ===")
    print(table(tableau, money=["TVA à déclarer"], ints=["Lignes"]))
    if len(sans_type):
        print(f"\n/!\\ {len(sans_type)} ligne(s) sans Type de TVA avec une TVA non nulle "
              f"(total {fmt(sans_type[COL_TVA].sum())}) : hors récapitulatif ci-dessus.")

    # 2. Contrôle partie vente
    res = controle_ventes(df)
    ko = ecarts_ventes(res)
    print(f"\n=== CONTRÔLE VENTES : Montant signé vs détail du règlement ===")
    print(f"{len(res)} règlements contrôlés, {len(ko)} en écart.")
    if len(ko):
        cols = ["N° pièce règlement", "type_reglement", "montant_signe",
                "rapproche_lettre", "impay", "etape_interm", "reclassement",
                "frais_bancaire", "retenu", "ecart"]
        money = ["montant_signe", "rapproche_lettre", "impay", "etape_interm",
                 "reclassement", "frais_bancaire", "retenu", "ecart"]
        print(table(ko[cols], money=money))
        if (~ko["signe_constant"]).any():
            print("\n/!\\ Montant signé non constant sur toutes les lignes pour : "
                  + ", ".join(ko.loc[~ko["signe_constant"], COL_PIECE]))
        print(f"\nTotal des écarts : {fmt(ko['ecart'].sum())}  "
              f"(valeur absolue cumulée : {fmt(ko['ecart'].abs().sum())})")

    if args.output:
        with pd.ExcelWriter(args.output, engine="openpyxl") as xw:
            tableau.to_excel(xw, sheet_name="TVA_par_type", index=False)
            ko.to_excel(xw, sheet_name="Ecarts_ventes", index=False)
            res.to_excel(xw, sheet_name="Detail_ventes", index=False)
        print(f"\nRapport écrit : {args.output}")


# =============================================================================
# VOLET 2 : delta entre deux fichiers
# =============================================================================
def aggregate_reglements(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(COL_PIECE, sort=True)
    return pd.DataFrame({
        "nb_lignes": g.size(),
        "montant_regle": g[COL_REGLE].sum().round(2),
        "tva_a_declarer": g[COL_TVA].sum().round(2),
        "montant_signe": g[COL_SIGNE].first().round(2),
    })


def delta_reglements(old: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    j = old.add_suffix("_ancien").join(new.add_suffix("_nouveau"), how="outer")
    in_old = j["nb_lignes_ancien"].notna()
    in_new = j["nb_lignes_nouveau"].notna()

    j["statut"] = "IDENTIQUE"
    j.loc[~in_old & in_new, "statut"] = "NOUVEAU"
    j.loc[in_old & ~in_new, "statut"] = "SUPPRIME"

    j["delta_nb_lignes"] = j["nb_lignes_nouveau"] - j["nb_lignes_ancien"]
    j["delta_montant_regle"] = (j["montant_regle_nouveau"] - j["montant_regle_ancien"]).round(2)
    j["delta_tva"] = (j["tva_a_declarer_nouveau"] - j["tva_a_declarer_ancien"]).round(2)
    j["delta_signe"] = (j["montant_signe_nouveau"] - j["montant_signe_ancien"]).round(2)

    changed = in_old & in_new & (
        (j["delta_nb_lignes"] != 0)
        | (j["delta_montant_regle"].abs() > TOL)
        | (j["delta_tva"].abs() > TOL)
        | (j["delta_signe"].abs() > TOL)
    )
    j.loc[changed, "statut"] = "MODIFIE"

    cols = ["statut",
            "nb_lignes_ancien", "nb_lignes_nouveau", "delta_nb_lignes",
            "montant_regle_ancien", "montant_regle_nouveau", "delta_montant_regle",
            "tva_a_declarer_ancien", "tva_a_declarer_nouveau", "delta_tva",
            "montant_signe_ancien", "montant_signe_nouveau", "delta_signe"]
    return j[cols].reset_index().rename(columns={"index": COL_PIECE})


def delta_resume(old: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    a, _ = resume_tva(old)
    b, _ = resume_tva(new)
    j = a.merge(b, on="Type de TVA", suffixes=("_ancien", "_nouveau"))
    j["delta_tva"] = (j["TVA à déclarer_nouveau"] - j["TVA à déclarer_ancien"]).round(2)
    return j


def cmd_delta(args) -> None:
    old_df, new_df = load(args.ancien), load(args.nouveau)
    d = delta_reglements(aggregate_reglements(old_df), aggregate_reglements(new_df))
    diff = d[d["statut"] != "IDENTIQUE"]
    dr = delta_resume(old_df, new_df)

    print(f"\nAncien  : {args.ancien.name}  ({len(old_df)} lignes)")
    print(f"Nouveau : {args.nouveau.name}  ({len(new_df)} lignes)")
    print("\n=== DELTA TVA PAR TYPE ===")
    print(table(dr, money=["TVA à déclarer_ancien", "TVA à déclarer_nouveau", "delta_tva"],
                ints=["Lignes_ancien", "Lignes_nouveau"]))
    print("\n=== RÈGLEMENTS ===")
    print(d["statut"].value_counts().to_string())
    if len(diff):
        print("\n=== RÈGLEMENTS EN ÉCART ===")
        print(table(diff))

    # Régressions de contrôle : écarts ventes apparus dans le nouveau fichier
    ko_old = set(ecarts_ventes(controle_ventes(old_df))[COL_PIECE])
    ko_new = ecarts_ventes(controle_ventes(new_df))
    regress = ko_new[~ko_new[COL_PIECE].isin(ko_old)]
    print(f"\n=== ÉCARTS DE CONTRÔLE VENTES APPARUS DANS LE NOUVEAU : {len(regress)} ===")
    if len(regress):
        print(table(regress))

    with pd.ExcelWriter(args.output, engine="openpyxl") as xw:
        dr.to_excel(xw, sheet_name="Delta_TVA_type", index=False)
        diff.to_excel(xw, sheet_name="Ecarts_reglements", index=False)
        regress.to_excel(xw, sheet_name="Nouveaux_ecarts_ventes", index=False)
    print(f"\nRapport écrit : {args.output}")


# =============================================================================
def main() -> None:
    ap = argparse.ArgumentParser(description="Contrôle / delta du fichier TVA X3")
    sub = ap.add_subparsers(dest="cmd", required=True)

    c = sub.add_parser("controle", help="contrôler un seul fichier")
    c.add_argument("fichier", type=Path, nargs="?", default=DEFAULT_FILE,
                   help="fichier CSV (défaut : file.csv à côté du script)")
    c.add_argument("-o", "--output", type=Path, help="rapport Excel (optionnel)")
    c.set_defaults(func=cmd_controle)

    d = sub.add_parser("delta", help="comparer ancien et nouveau fichier")
    d.add_argument("ancien", type=Path)
    d.add_argument("nouveau", type=Path)
    d.add_argument("-o", "--output", type=Path, default=Path("rapport_delta_tva.xlsx"))
    d.set_defaults(func=cmd_delta)

    args = ap.parse_args()
    try:  # caractères des bordures et accents, même si la console est en cp1252
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    args.func(args)


if __name__ == "__main__":
    main()
