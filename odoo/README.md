# Alerte J-2 des factures fournisseurs (Odoo)

Le script [`alerte_factures_fournisseurs.py`](alerte_factures_fournisseurs.py) est fait pour tourner dans une **action planifiée** Odoo. Chaque jour, il :

1. cherche les factures fournisseurs (`account.move`, `move_type = in_invoice`) **validées** (`state = posted`) et **non payées ou partiellement payées** (`payment_state` = `not_paid` ou `partial`) ;
2. garde celles dont l'**échéance tombe dans 2 jours** (J-2) ;
3. ne retient que les fournisseurs dont la **condition de paiement fournisseur** contient « 7 jours » ou « hebdomadaire » ;
4. envoie **un seul e-mail récapitulatif** à `compta@clubtiossane.sn`, avec un tableau dans le corps du message et un **fichier CSV** en pièce jointe ;
5. si aucune facture n'est trouvée, envoie un **e-mail de contrôle** pour confirmer que l'alerte fonctionne.

Le CSV est construit à la main puis encodé avec `b64encode`, car le bac à sable d'Odoo (`safe_eval`) interdit les `import` (donc pas de module `csv` ni `base64`). Il utilise le séparateur `;` et un encodage UTF-8 avec BOM, pour s'ouvrir correctement dans Excel en français.

## Prérequis

- Odoo 16, 17 ou 18 avec le module **Comptabilité / Facturation**.
- Un **serveur d'e-mails sortant** configuré (Paramètres > Technique > Serveurs de messagerie sortants).
- Une adresse e-mail renseignée sur la société (elle sert d'expéditeur).
- Les fournisseurs concernés ont une **condition de paiement fournisseur** (fiche contact > onglet Ventes et achats) dont le nom contient « 7 jours » ou « Hebdomadaire ».

## Créer l'action planifiée

1. Activer le **mode développeur** (Paramètres > Activer le mode développeur).
2. Aller dans **Paramètres > Technique > Automatisation > Actions planifiées**, puis cliquer sur **Nouveau**.
3. Remplir :
   | Champ | Valeur |
   |---|---|
   | Nom | Alerte J-2 factures fournisseurs |
   | Modèle | Écriture comptable (`account.move`) |
   | Exécuter tous les | 1 Jour(s) |
   | Prochaine exécution | demain, par exemple à 07:00 |
   | Utilisateur | un utilisateur ayant les droits Comptabilité |
   | Action à faire | Exécuter du code |
4. Coller tout le contenu de `alerte_factures_fournisseurs.py` dans l'onglet **Code**.
5. **Enregistrer**, puis cliquer sur **Exécuter manuellement** pour tester. Un e-mail doit arriver à `compta@clubtiossane.sn` (récapitulatif ou e-mail de contrôle).

> En Odoo 16, le menu s'appelle **Paramètres > Technique > Automatisation > Actions planifiées** et le champ « Action à faire » porte la valeur **Exécuter le code Python**.

## Paramètres modifiables

En haut du script :

| Variable | Rôle | Valeur par défaut |
|---|---|---|
| `DESTINATAIRE` | Adresse qui reçoit l'alerte | `compta@clubtiossane.sn` |
| `JOURS_AVANT_ECHEANCE` | Nombre de jours avant l'échéance | `2` |
| `MOTS_CLES_CONDITIONS` | Mots cherchés dans le nom de la condition de paiement | `7 jours`, `7 days`, `hebdomadaire`, `weekly` |
| `ENVOYER_SI_AUCUNE_FACTURE` | Envoyer un e-mail de contrôle quand il n'y a rien | `True` |
| `SEPARATEUR` | Séparateur du CSV | `;` |

Une fois l'alerte validée en production, on peut passer `ENVOYER_SI_AUCUNE_FACTURE` à `False` pour ne recevoir un e-mail que lorsqu'il y a des factures.

## Contenu du CSV

Colonnes : Fournisseur, N° facture, Référence fournisseur, Date facture, Échéance, Condition de paiement, Montant total, Reste à payer, Devise, Statut paiement.

## Vérifier et dépanner

- **Journal** : chaque exécution écrit une ligne « Alerte factures fournisseurs : N facture(s)… » dans Paramètres > Technique > Journalisation (`ir.logging`).
- **E-mails envoyés ou en échec** : Paramètres > Technique > E-mails. Un e-mail en état « Exception » indique en général un problème de serveur sortant.
- **Aucune facture alors qu'il devrait y en avoir** : vérifier la date d'échéance de la facture, son statut de paiement et le nom exact de la condition de paiement du fournisseur.
- **Erreur « forbidden opcode » ou « import »** : le code a été modifié avec une instruction interdite par le bac à sable. Repartir de la version du dépôt.
