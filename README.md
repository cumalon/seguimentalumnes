# Introducció

Aquest projecte consisteix en una extensió de fulls de càlcul de Google.

L'extensió té dues funcionalitats principals:
- Generar informes a partir de les dades d'un full de càlcul. 
- Visualització dels informes en una webapp.

Els informes són documents de text de Google.

La webapp genera una pàgina html on s'enllaça els informes generats. Hi ha dues visualitzacions, la vista admin o d'accés complet i la vista user o d'accés limitat.


## Com es fa?

Aquesta eina realitza un procés merge entre les dades d'un full de càlcul i una plantilla predefinida.

El procés merge consisteix en combinar la informació que conté el full de càlcul amb la plantilla, amb la finalitat de generar documents o informes personalitzats. Aquest procés permet visualitzar dades estructurades (com noms, adreces, qualificacions o altres valors) en diferents documents o formats mantenint un estil uniforme.

**Full de càlcul:** conté les dades en forma de taula, on cada fila representa una entrada individual (com una persona o un projecte) i cada columna un camp d'informació (com el nom, l'adreça, etc.).

**Plantilla:** és un document pre-formatat amb un disseny estàtic, on s'han definit espais reservats o placeholders (com {{Nom}}, {{Data}}, etc.) que seran substituïts per les dades del full de càlcul.

**Procés de merge:** s’utilitza un programa (com Microsoft Word amb el seu correu massiu o aplicacions de Google Workspace) per fer el merge, agafant cada fila del full de càlcul i substituint els placeholders de la plantilla amb els valors corresponents de cada fila.

El resultat és una sèrie de documents generats automàticament on cada un conté les dades personalitzades de cada fila del full de càlcul.

### Fulls i camps importants

Les dades d'un merge han d'estar totes elles en una sola pestanya del full de càlcul. El full pot tenir qualsevol nom. Els noms de les capçaleres són lliures per realitzar un procés de merge, però no és així quan s'executa la webapp. La webapp buscarà els camps **Email** i **Cognoms, Nom**.

**Email:** quan posem aquest camp en una de les columnes d'un merge, el document generat de cada fila es comparteix en mode lectura amb el correu proporcionat. La webapp s'ha d'executar des d'un usuari Google, de manera que, per defecte, només es visualitza els informes generats i compartits amb l'usuari en concret. Això permet que, en una graella de qualificacions d'un grup d'alumnes per exemple, un estudiant visualitzi únicament les seves qualificacions.

**Cognoms, Nom**: quan un usuari admin executa la webapp aquest pot accedir als informes generats per a múltiples files del full de dades. En aquest cas, apareix un selector on es pot escollit el Cognom, Nom del qual es vol visualitzar els informes. Aquesta seria la vista que podria veure el professorat de l'exemple acadèmic mencionat abans.

Els informes accessibles des de la webapp es llisten a una taula de dades del full amb noma **webapp** (cal veure exemples). Aquest mateix full també conté la taula d'usuaris Google amb accés admin i la taula amb els Cognoms, Nom que es podran seleccionar des de la visualització admin.


# Us de clasp

Els projectes d'Apps Script es poden gestionar localment a través de clasp. En aquest web es fa un resum de les funcionalitats més importants:

 https://developers.google.com/apps-script/guides/clasp

Clasp permet sincronitzar un projecte en local amb un sol projecte Apps Script associat a un fitxer del Drive (el contenidor, en llenguatge Google).

Pot passar que un mateix projecte Apps Script es vulgui fer servir en diferents fitxers del Drive. Per exemple, potser el nostre projecte es fa servir a diversos fulls de càlcul. En aquest cas, cada full de càlcul requerirà el seu propi projecte (independent), és a dir que el fitxer Drive contindrà una còpia del codi (es poden veure a script.google.com o des del menú del mateix fitxer via Extensions->Apps Script). Cadascun d'aquests projectes té la seva pròpia scriptId. Amb aquest valor es pot canviar el projecte amb el que clasp realitza la sincronització.

Per canviar la ID del projecte a clasp cal executar la següent comanda:
 clasp setting scriptId AQUÍ_LA_ID_NOVA


# Us de git:

Com que la gestió del projecte es fa localment i no a través de l'editor de scripts de Google, es pot fer servir git per gestionar les versions del codi. En concret el codi està versionat a través de github:

 https://github.com/cumalon/qualificacions
