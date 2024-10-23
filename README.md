
Full de càlcul de Google amb funcionalitats per fer el seguiment dels alumnes i comunicar-los les qualificacions.


Us de clasp:
===========

Els projectes d'Apps Script es poden gestionar localment a través de clasp. En aquest web es fa un resum de les funcionalitats més importants:

 https://developers.google.com/apps-script/guides/clasp

Clasp permet sincronitzar un projecte en local amb un sol projecte Apps Script associat a un fitxer del Drive (el contenidor, en llenguatge Google).

Pot passar que un mateix projecte Apps Script es vulgui fer servir en diferents fitxers del Drive. Per exemple, potser el nostre projecte es fa servir a diversos fulls de càlcul. En aquest cas, cada full de càlcul requerirà el seu propi projecte (independent), és a dir que el fitxer Drive contindrà una còpia del codi (es poden veure a script.google.com o des del menú del mateix fitxer via Extensions->Apps Script). Cadascun d'aquests projectes té la seva pròpia scriptId. Amb aquest valor es pot canviar el projecte amb el que clasp realitza la sincronització.

Per canviar la ID del projecte a clasp cal executar la següent comanda:
 clasp setting scriptId AQUÍ_LA_ID_NOVA


Us de git:
=========

Com que la gestió del projecte es fa localment i no a través de l'editor de scripts de Google, es pot fer servir git per gestionar les versions del codi. En concret el codi està versionat a través de github:

 https://github.com/cumalon/qualificacions
