#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genere des PDF de facture de test (avec couche texte).

Ces fichiers servent a tester l'import et l'extraction de l'outil. Ils
contiennent une vraie couche texte : le niveau 1 d'extraction (lecture via
Word) doit donc les traiter sans OCR.
"""


def make_pdf(path, lignes):
    """Ecrit un PDF minimal mono-page contenant les lignes de texte donnees."""
    # --- flux de contenu ---
    flux = ["BT", "/F1 11 Tf", "50 800 Td"]
    for i, ln in enumerate(lignes):
        esc = ln.replace("\\", r"\\").replace("(", r"\(").replace(")", r"\)")
        if i > 0:
            flux.append("0 -16 Td")
        flux.append(f"({esc}) Tj")
    flux.append("ET")
    contenu = "\n".join(flux).encode("latin-1", "replace")

    objets = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] "
        b"/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
        b"<< /Length " + str(len(contenu)).encode() + b" >>\nstream\n"
        + contenu + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]

    out = bytearray(b"%PDF-1.4\n")
    offsets = []
    for i, obj in enumerate(objets, start=1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n".encode() + obj + b"\nendobj\n"

    xref_pos = len(out)
    out += f"xref\n0 {len(objets) + 1}\n".encode()
    out += b"0000000000 65535 f \n"
    for off in offsets:
        out += f"{off:010d} 00000 n \n".encode()
    out += (f"trailer\n<< /Size {len(objets) + 1} /Root 1 0 R >>\n"
            f"startxref\n{xref_pos}\n%%EOF").encode()

    with open(path, "wb") as f:
        f.write(out)
    print("genere :", path)


FACTURE_1 = [
    "ACME Travaux SARL",
    "12 rue des Chantiers - 75011 Paris",
    "SIRET 12345678900012",
    "",
    "FACTURE N0 2026-0412",
    "Date facture : 14/05/2026",
    "Client : SCI Immeuble Haussmann",
    "",
    "Refection toiture - lot 3",
    "",
    "Total HT : 18 500,00 EUR",
    "Montant TVA : 3 700,00 EUR",
    "Total TTC : 22 200,00 EUR",
    "",
    "Reglement a 30 jours.",
]

FACTURE_2 = [
    "Energie Plus SA",
    "5 avenue de la Republique - 69003 Lyon",
    "SIRET 98765432100021",
    "",
    "FACTURE N0 EP-2026-00891",
    "Date facture : 02/05/2026",
    "Client : Societe de Gestion ABC",
    "",
    "Fourniture electricite - avril 2026",
    "",
    "Total HT : 1 240,50 EUR",
    "Montant TVA : 248,10 EUR",
    "Total TTC : 1 488,60 EUR",
]

if __name__ == "__main__":
    make_pdf("facture_ACME_2026-0412.pdf", FACTURE_1)
    make_pdf("facture_EnergiePlus_EP-2026-00891.pdf", FACTURE_2)
