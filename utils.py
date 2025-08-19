import os
import pandas as pd

def load_sav_preserving_codes(path):
    """
    (Opcional – não usado no fluxo em memória.)
    Lê .sav preservando CÓDIGOS (sem aplicar rótulos) e retorna (df, meta).
    """
    import pyreadstat
    df, meta = pyreadstat.read_sav(path, apply_value_formats=False)
    df.columns = [str(c) for c in df.columns]
    return df, meta

def spds_value_labels_map(meta):
    """
    Constrói um dict {var_name: {codigo: label}} a partir do meta do SPSS.
    Retorna {} se não houver labels.
    """
    labels = {}
    if not meta:
        return labels
    for var, labelset in (meta.variable_to_label or {}).items():
        if labelset and meta.value_labels and labelset in meta.value_labels:
            labels[var] = meta.value_labels[labelset]
    return labels

def detect_ext(filename: str) -> str:
    return os.path.splitext(filename)[1].lower()
