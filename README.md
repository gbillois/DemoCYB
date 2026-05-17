# DemoCYB

Une mini **arcade cybersécurité** en HTML, CSS et JavaScript vanilla.

## Jouer

Ouvre `index.html` dans un navigateur moderne. L'écran de démarrage permet de choisir entre deux modes :

- **Casse-brique** : le jeu cyberpunk classique où tu détruis des pare-feux néon avec une balle plasma.
- **Snake Cyber** : un Snake compatible mobile où le serpent mange des mots de cybersécurité, puis répond à un quiz.

## Contrôles

### Casse-brique

- `←` / `→` ou `A` / `D` : déplacer la raquette
- Souris ou tactile : déplacer la raquette
- `Espace` : pause / reprise

### Snake Cyber

- `←` / `↑` / `↓` / `→` ou `W` / `A` / `S` / `D` : diriger le serpent
- Gestes tactiles sur le canvas ou pavé directionnel mobile : diriger le serpent
- `Espace` : pause / reprise

## Quiz Snake Cyber

Le serpent mange des mots autour de quatre thèmes :

- **ransomware**
- **ISO 27001**
- **cloud security**
- **AI security**

Après chaque mot mangé, un popup de quiz propose trois réponses :

- une bonne réponse,
- une réponse presque bonne,
- une réponse vraiment mauvaise.

Une bonne réponse fait grandir le serpent et augmente le score. Une mauvaise réponse réduit le serpent et enlève des points.

## Bonus casse-brique

Certaines briques libèrent des bonus néon à attraper avec la raquette :

- **Raquette XL** : élargit temporairement la raquette
- **Ralenti plasma** : ralentit temporairement la balle
- **Bouclier** : grossit temporairement la balle pour la rendre plus facile à jouer
