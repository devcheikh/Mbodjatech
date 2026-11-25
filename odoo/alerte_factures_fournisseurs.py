# -*- coding: utf-8 -*-
# =============================================================================
# Alerte J-2 : factures fournisseurs arrivant à échéance
# =============================================================================
#
# Code à coller dans une ACTION PLANIFIÉE Odoo
# (Paramètres > Technique > Automatisation > Actions planifiées),
# modèle « Écriture comptable » (account.move), action « Exécuter du code ».
#
# Ce code s'exécute dans le bac à sable d'Odoo (safe_eval) :
#   - pas d'« import » possible (ni csv, ni base64, ni io...) ;
#   - on utilise uniquement les objets fournis par Odoo :
#     env, datetime, b64encode, log.
#
# Fonctionnement :
#   1. Cherche les factures fournisseurs validées et non (ou partiellement)
#      payées dont l'échéance tombe dans JOURS_AVANT_ECHEANCE jours.
#   2. Ne garde que les fournisseurs dont la condition de paiement est
#      « 7 jours » ou « hebdomadaire » (voir MOTS_CLES_CONDITIONS).
#   3. Envoie UN seul e-mail récapitulatif à DESTINATAIRE avec un CSV en
#      pièce jointe (séparateur « ; », encodage UTF-8 avec BOM pour Excel).
#   4. Si aucune facture n'est trouvée, envoie un e-mail de contrôle
#      (désactivable avec ENVOYER_SI_AUCUNE_FACTURE = False).
# =============================================================================

# ----------------------------------------------------------------------------
# Paramètres
# ----------------------------------------------------------------------------
DESTINATAIRE = "compta@clubtiossane.sn"
JOURS_AVANT_ECHEANCE = 2
# Mots recherchés (sans tenir compte des majuscules) dans le nom de la
# condition de paiement fournisseur du partenaire.
MOTS_CLES_CONDITIONS = ["7 jours", "7 days", "hebdomadaire", "weekly"]
ENVOYER_SI_AUCUNE_FACTURE = True
SEPARATEUR = ";"


# ----------------------------------------------------------------------------
# Fonctions utilitaires
# ----------------------------------------------------------------------------
def cellule_csv(valeur):
    """Formate une valeur pour une cellule CSV (guillemets échappés)."""
    if valeur is None or valeur is False:
        texte = ""
    else:
        texte = str(valeur)
    besoin_guillemets = (
        SEPARATEUR in texte or '"' in texte or "\n" in texte or "\r" in texte
    )
    texte = texte.replace('"', '""')
    if besoin_guillemets:
        texte = '"' + texte + '"'
    return texte


def ligne_csv(valeurs):
    return SEPARATEUR.join([cellule_csv(v) for v in valeurs])


def echapper_html(valeur):
    texte = str(valeur or "")
    return (
        texte.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def montant(valeur):
    return "%.2f" % (valeur or 0.0)


def date_fr(valeur):
    return valeur.strftime("%d/%m/%Y") if valeur else ""


def condition_fournisseur(facture):
    partenaire = facture.partner_id.commercial_partner_id
    return partenaire.property_supplier_payment_term_id


def condition_ciblee(condition):
    if not condition:
        return False
    nom = (condition.name or "").lower()
    for mot in MOTS_CLES_CONDITIONS:
        if mot.lower() in nom:
            return True
    return False


def envoyer_mail(sujet, corps_html, piece_jointe=None):
    valeurs = {
        "subject": sujet,
        "body_html": corps_html,
        "email_to": DESTINATAIRE,
        "email_from": env.company.email_formatted or env.user.email_formatted,
        "auto_delete": True,
    }
    if piece_jointe:
        valeurs["attachment_ids"] = [(4, piece_jointe.id)]
    mail = env["mail.mail"].sudo().create(valeurs)
    mail.send()
    return mail


# ----------------------------------------------------------------------------
# Recherche des factures
# ----------------------------------------------------------------------------
aujourd_hui = datetime.date.today()
date_echeance = aujourd_hui + datetime.timedelta(days=JOURS_AVANT_ECHEANCE)

factures = env["account.move"].search(
    [
        ("move_type", "=", "in_invoice"),
        ("state", "=", "posted"),
        ("payment_state", "in", ["not_paid", "partial"]),
        ("invoice_date_due", "=", date_echeance),
    ],
    order="partner_id, invoice_date_due, name",
)
factures = factures.filtered(lambda f: condition_ciblee(condition_fournisseur(f)))

log(
    "Alerte factures fournisseurs : %s facture(s) à échéance le %s"
    % (len(factures), date_fr(date_echeance)),
    level="info",
)

# ----------------------------------------------------------------------------
# Aucune facture : e-mail de contrôle
# ----------------------------------------------------------------------------
if not factures:
    if ENVOYER_SI_AUCUNE_FACTURE:
        envoyer_mail(
            "[Odoo] Aucune facture fournisseur à échéance le %s"
            % date_fr(date_echeance),
            "<p>Bonjour,</p>"
            "<p>Aucune facture fournisseur (condition 7 jours ou hebdomadaire) "
            "non payée n'arrive à échéance le <b>%s</b>.</p>"
            "<p>Ce message confirme que l'alerte automatique fonctionne.</p>"
            % date_fr(date_echeance),
        )

# ----------------------------------------------------------------------------
# Factures trouvées : CSV + e-mail récapitulatif
# ----------------------------------------------------------------------------
else:
    entetes = [
        "Fournisseur",
        "N° facture",
        "Référence fournisseur",
        "Date facture",
        "Échéance",
        "Condition de paiement",
        "Montant total",
        "Reste à payer",
        "Devise",
        "Statut paiement",
    ]
    statuts = {"not_paid": "Non payée", "partial": "Partiellement payée"}

    lignes = [ligne_csv(entetes)]
    lignes_html = []
    totaux = {}

    for f in factures:
        condition = condition_fournisseur(f)
        devise = f.currency_id.name or ""
        valeurs = [
            f.partner_id.commercial_partner_id.name,
            f.name,
            f.ref,
            date_fr(f.invoice_date),
            date_fr(f.invoice_date_due),
            condition.name,
            montant(f.amount_total),
            montant(f.amount_residual),
            devise,
            statuts.get(f.payment_state, f.payment_state),
        ]
        lignes.append(ligne_csv(valeurs))
        lignes_html.append(
            "<tr>"
            + "".join(
                ['<td style="padding:4px 8px;border:1px solid #ccc;">%s</td>'
                 % echapper_html(v) for v in valeurs]
            )
            + "</tr>"
        )
        totaux[devise] = totaux.get(devise, 0.0) + (f.amount_residual or 0.0)

    # BOM UTF-8 pour qu'Excel affiche correctement les accents.
    contenu_csv = "﻿" + "\r\n".join(lignes) + "\r\n"
    nom_fichier = "factures_fournisseurs_echeance_%s.csv" % date_echeance.strftime(
        "%Y-%m-%d"
    )
    piece_jointe = env["ir.attachment"].sudo().create(
        {
            "name": nom_fichier,
            "type": "binary",
            "datas": b64encode(contenu_csv.encode("utf-8")),
            "mimetype": "text/csv",
        }
    )

    texte_totaux = "<br/>".join(
        ["%s %s" % (montant(totaux[d]), echapper_html(d)) for d in sorted(totaux)]
    )
    entetes_html = "".join(
        ['<th style="padding:4px 8px;border:1px solid #ccc;background:#f2f2f2;">%s</th>'
         % echapper_html(e) for e in entetes]
    )
    corps = (
        "<p>Bonjour,</p>"
        "<p><b>%s</b> facture(s) fournisseur arrivent à échéance le <b>%s</b> "
        "(J-%s).</p>"
        "<p>Reste à payer :<br/>%s</p>"
        '<table style="border-collapse:collapse;font-size:12px;">'
        "<thead><tr>%s</tr></thead><tbody>%s</tbody></table>"
        "<p>Le détail est joint au format CSV (%s).</p>"
        % (
            len(factures),
            date_fr(date_echeance),
            JOURS_AVANT_ECHEANCE,
            texte_totaux,
            entetes_html,
            "".join(lignes_html),
            echapper_html(nom_fichier),
        )
    )
    envoyer_mail(
        "[Odoo] %s facture(s) fournisseur à échéance le %s"
        % (len(factures), date_fr(date_echeance)),
        corps,
        piece_jointe,
    )
