# Bloc militaire du code de DyMEMDS



Prérequis dans les fichiers Excel :
- StocksMilitaryMods.xlsx : feuille 'Stocks' avec colonnes A=code, B=family, C=unit_name, D..= années (1900, 1970..2014, 2100).
- RawMaterialsV2.xlsx : onglets d'intensité dynamique par famille et onglet 'Military_MassUnit' (D=mass_t).
- l'ordre des matériels dans le subscript MilitaryTypeUnit doit être STRICTEMENT identique à l'ordre des lignes dans les feuilles Excel. 


## Subscripts

On crée tous nos objets militaires pour lesquels on va définir des équations, et on les regroupe en familles

```vensim
MilitaryArmouredUnit:
	MilitaryArmoured Aml,MilitaryArmoured Amx10,MilitaryArmoured AmxVtt,\
		MilitaryArmoured Amx13,MilitaryArmoured Amx30,MilitaryArmoured Erc90,\
		MilitaryArmoured Leclerc,MilitaryArmoured Vab,MilitaryArmoured Vbci,\
		MilitaryArmoured VblM11
	~	
	~	Sous-plage : 10 matériels de la famille MilitaryArmoured (ordre de la feuille "Stocks").
	|

MilitaryAircraftsUnit:
	MilitaryAircrafts A300Variants,MilitaryAircrafts A400MAtlas,MilitaryAircrafts Alize,\
		MilitaryAircrafts Alphajet,MilitaryAircrafts Atlantic,MilitaryAircrafts Breguet941,\
		MilitaryAircrafts C130H,MilitaryAircrafts C212,MilitaryAircrafts C47,\
		MilitaryAircrafts Cn235,MilitaryAircrafts Caravelle,MilitaryAircrafts CessnaCaravanIi,\
		MilitaryAircrafts Crusader,MilitaryAircrafts Dc6,MilitaryAircrafts Dc8,\
		MilitaryAircrafts Dhc6,MilitaryAircrafts E2CHawkeye,MilitaryAircrafts E3F,\
		MilitaryAircrafts Emb121,MilitaryAircrafts Emb312,MilitaryAircrafts EtendardIv,\
		MilitaryAircrafts F100D,MilitaryAircrafts Flamant,MilitaryAircrafts FougaMs,\
		MilitaryAircrafts Jaguar,MilitaryAircrafts Kc135,MilitaryAircrafts Mirage2000,\
		MilitaryAircrafts MirageF1,MilitaryAircrafts MirageIii,MilitaryAircrafts MirageIv,\
		MilitaryAircrafts MirageV,MilitaryAircrafts MystereFalcon,MilitaryAircrafts Navajo,\
		MilitaryAircrafts Noratlas,MilitaryAircrafts Nord264,MilitaryAircrafts Rafale,\
		MilitaryAircrafts Sp2HNeptune,MilitaryAircrafts SuperEtendard,\
		MilitaryAircrafts SuperMystere,MilitaryAircrafts T33A,MilitaryAircrafts Tbm700,\
		MilitaryAircrafts TransallC160,MilitaryAircrafts VautourIi,MilitaryAircrafts LightTraining
	~	
	~	Sous-plage : 44 matériels de la famille MilitaryAircrafts (ordre de la feuille "Stocks").
	|

MilitaryHelicoptersUnit:
	MilitaryHelicopters Alouette,MilitaryHelicopters Caracal,MilitaryHelicopters Colibiri,\
		MilitaryHelicopters Cougar,MilitaryHelicopters Dauphin,MilitaryHelicopters Ecureuil,\
		MilitaryHelicopters Fennec,MilitaryHelicopters Gazelle,MilitaryHelicopters Nh90Tth,\
		MilitaryHelicopters Panther,MilitaryHelicopters Puma,MilitaryHelicopters SuperFrelon,\
		MilitaryHelicopters SuperPuma,MilitaryHelicopters Tigre
	~	
	~	Sous-plage : 14 matériels de la famille MilitaryHelicopters (ordre de la feuille "Stocks").
	|

MilitaryPatrolUnit:
	MilitaryPatrol Albatros,MilitaryPatrol Athos,MilitaryPatrol ClasseBay,\
		MilitaryPatrol DEstienneDOrves,MilitaryPatrol Grebe,MilitaryPatrol LeopardInstruction,\
		MilitaryPatrol Opv54,MilitaryPatrol P400,MilitaryPatrol PatrouilleursSp,\
		MilitaryPatrol Sterne,MilitaryPatrol Trident,MilitaryPatrol TypeMso,MilitaryPatrol TypeD
	~	
	~	Sous-plage : 13 matériels de la famille MilitaryPatrol (ordre de la feuille "Stocks").
	|

MilitarySurfaceUnit:
	MilitarySurface C65,MilitarySurface CharlesDeGaulle,MilitarySurface Clemenceau,\
		MilitarySurface Colbert,MilitarySurface CommandantRiviere,MilitarySurface DEstienneDOrves,\
		MilitarySurface Escorteurs,MilitarySurface F67,MilitarySurface Floreal,\
		MilitarySurface JeanneDArc,MilitarySurface LaFayette,MilitarySurface Suffren,\
		MilitarySurface TypeE52,MilitarySurface TypeF70
	~	
	~	Sous-plage : 14 matériels de la famille MilitarySurface (ordre de la feuille "Stocks").
	|

MilitarySubmarinesUnit:
	MilitarySubmarines Agosta,MilitarySubmarines Arethuse,MilitarySubmarines Daphne,\
		MilitarySubmarines LeRedoutable,MilitarySubmarines LeTriomphant,MilitarySubmarines Narval,\
		MilitarySubmarines Rubis
	|

MilitaryTypeUnit:
	MilitaryArmoured Aml,MilitaryArmoured Amx10,MilitaryArmoured AmxVtt,\
		MilitaryArmoured Amx13,MilitaryArmoured Amx30,MilitaryArmoured Erc90,\
		MilitaryArmoured Leclerc,MilitaryArmoured Vab,MilitaryArmoured Vbci,\
		MilitaryArmoured VblM11,MilitaryAircrafts A300Variants,MilitaryAircrafts A400MAtlas,\
		MilitaryAircrafts Alize,MilitaryAircrafts Alphajet,MilitaryAircrafts Atlantic,\
		MilitaryAircrafts Breguet941,MilitaryAircrafts C130H,MilitaryAircrafts C212,\
		MilitaryAircrafts C47,MilitaryAircrafts Cn235,MilitaryAircrafts Caravelle,\
		MilitaryAircrafts CessnaCaravanIi,MilitaryAircrafts Crusader,MilitaryAircrafts Dc6,\
		MilitaryAircrafts Dc8,MilitaryAircrafts Dhc6,MilitaryAircrafts E2CHawkeye,\
		MilitaryAircrafts E3F,MilitaryAircrafts Emb121,MilitaryAircrafts Emb312,\
		MilitaryAircrafts EtendardIv,MilitaryAircrafts F100D,MilitaryAircrafts Flamant,\
		MilitaryAircrafts FougaMs,MilitaryAircrafts Jaguar,MilitaryAircrafts Kc135,\
		MilitaryAircrafts Mirage2000,MilitaryAircrafts MirageF1,MilitaryAircrafts MirageIii,\
		MilitaryAircrafts MirageIv,MilitaryAircrafts MirageV,MilitaryAircrafts MystereFalcon,\
		MilitaryAircrafts Navajo,MilitaryAircrafts Noratlas,MilitaryAircrafts Nord264,\
		MilitaryAircrafts Rafale,MilitaryAircrafts Sp2HNeptune,MilitaryAircrafts SuperEtendard,\
		MilitaryAircrafts SuperMystere,MilitaryAircrafts T33A,MilitaryAircrafts Tbm700,\
		MilitaryAircrafts TransallC160,MilitaryAircrafts VautourIi,\
		MilitaryAircrafts LightTraining,MilitaryHelicopters Alouette,MilitaryHelicopters Caracal,\
		MilitaryHelicopters Colibiri,MilitaryHelicopters Cougar,MilitaryHelicopters Dauphin,\
		MilitaryHelicopters Ecureuil,MilitaryHelicopters Fennec,MilitaryHelicopters Gazelle,\
		MilitaryHelicopters Nh90Tth,MilitaryHelicopters Panther,MilitaryHelicopters Puma,\
		MilitaryHelicopters SuperFrelon,MilitaryHelicopters SuperPuma,MilitaryHelicopters Tigre,\
		MilitaryPatrol Albatros,MilitaryPatrol Athos,MilitaryPatrol ClasseBay,\
		MilitaryPatrol DEstienneDOrves,MilitaryPatrol Grebe,MilitaryPatrol LeopardInstruction,\
		MilitaryPatrol Opv54,MilitaryPatrol P400,MilitaryPatrol PatrouilleursSp,\
		MilitaryPatrol Sterne,MilitaryPatrol Trident,MilitaryPatrol TypeMso,MilitaryPatrol TypeD,\
		MilitarySurface C65,MilitarySurface CharlesDeGaulle,MilitarySurface Clemenceau,\
		MilitarySurface Colbert,MilitarySurface CommandantRiviere,MilitarySurface DEstienneDOrves,\
		MilitarySurface Escorteurs,MilitarySurface F67,MilitarySurface Floreal,\
		MilitarySurface JeanneDArc,MilitarySurface LaFayette,MilitarySurface Suffren,\
		MilitarySurface TypeE52,MilitarySurface TypeF70,MilitarySubmarines Agosta,\
		MilitarySubmarines Arethuse,MilitarySubmarines Daphne,MilitarySubmarines LeRedoutable,\
		MilitarySubmarines LeTriomphant,MilitarySubmarines Narval,MilitarySubmarines Rubis
	~	
	~	102 matériels militaires détaillés (union des 6 sous-plages).
		L'ORDRE est celui des lignes de la feuille "Stocks" du fichier StocksMilitaryMods.xlsx.
	|

CountriesNotFrance:
	World
	~	
	~	Pays autres que France (ici : agregat World), pour mettre le militaire a \
		zero
	|
```

## STOCKS PAR MATÉRIEL (lus depuis Excel)

on fait lire à vensim toutes les lignes du fichier de données, où chaque ligne est dans le même ordre d'appel que le subscript MilitaryTypeUnit

```vensim
StockTUMilitaryUnit[France,MilitaryTypeUnit]=
	GET DIRECT LOOKUPS('StocksMilitaryMods.xlsx', 'Stocks', '1', 'D2') ~~|
StockTUMilitaryUnit[CountriesNotFrance,MilitaryTypeUnit](
	[(1900,0)-(2100,1)],(1900,0),(2100,0))
	~	TU
	~	Stock du parc militaire français par matériel, série annuelle 1900..2100.
		D1 = ligne des années ; D2 = première cellule de la première série.
		Vensim itère sur MilitaryTypeUnit en descendant les lignes.
	|

PreviousStockTUMilitaryUnit[Country,MilitaryTypeUnit]= INTEG (
	StockTUMilitaryUnit[Country,MilitaryTypeUnit](Time)
	- PreviousStockTUMilitaryUnit[Country,MilitaryTypeUnit],
		0)
	~	TU
	~	Retard d'un pas de temps sur le stock (analogue à PreviousStockTUinUse).
	|
```

les variables New et EoL sont calculées à partir des séries de données directement (pas d'hypothèse de durée de vie nécessaire)

```vensim
"FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryTypeUnit]=
	MAX(0, StockTUMilitaryUnit[Country,MilitaryTypeUnit](Time)
	- PreviousStockTUMilitaryUnit[Country,MilitaryTypeUnit])
	~	TU/Year
	~	Livraisons observées = hausses du stock d'une année sur l'autre.
	|

"FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryTypeUnit]=
	MAX(0, PreviousStockTUMilitaryUnit[Country,MilitaryTypeUnit]
	- StockTUMilitaryUnit[Country,MilitaryTypeUnit](Time))
	~	TU/Year
	~	Décommissionnements observés = baisses du stock d'une année sur l'autre.
	|


"FlowEoLTUMilitaryFam(/yr)"[Country,MilitaryAircrafts]=
	SUM("FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryAircraftsUnit!]) ~~|
"FlowEoLTUMilitaryFam(/yr)"[Country,MilitaryPrincipalSurface]=
	SUM("FlowEoLTUMilitaryUnit(/yr)"[Country,MilitarySurfaceUnit!]) ~~|
"FlowEoLTUMilitaryFam(/yr)"[Country,MilitaryPatrol]=
	SUM("FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryPatrolUnit!]) ~~|
"FlowEoLTUMilitaryFam(/yr)"[Country,MilitarySubmarines]=
	SUM("FlowEoLTUMilitaryUnit(/yr)"[Country,MilitarySubmarinesUnit!]) ~~|
"FlowEoLTUMilitaryFam(/yr)"[Country,MilitaryArmoured]=
	SUM("FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryArmouredUnit!]) ~~|
"FlowEoLTUMilitaryFam(/yr)"[Country,MilitaryHelicopters]=
	SUM("FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryHelicoptersUnit!])
	~	1/Year
	~	agrégateur famille ; SUM(MAX(0,-dS_u)) et non MAX(0,-d(SUM S_u)).
	|
```

## MASSE UNITAIRE (tonnes par matériel, lue depuis Excel)

```vensim
"MassUnitMilitaryUnit(t/TU)"[MilitaryTypeUnit]=
	GET DIRECT CONSTANTS('RawMaterialsV2.xlsx', 'Military_MassUnit', 'D2*')
	~	t/TU
	~	Masse unitaire de chaque matériel militaire (colonne D de l'onglet
		Military_MassUnit, à partir de la ligne 2).
	|
```

## CHAÎNE MATIÈRE

on importe les intensités en pourcentages, qu'on a déclarées par famille

```vensim
"RMIntensityFamilyPct(pct)"[RMAll,MilitaryArmoured]=
	GET DIRECT LOOKUPS('RawMaterialsV2.xlsx', 'Military_Armoured', '1', 'B2') ~~|
"RMIntensityFamilyPct(pct)"[RMAll,MilitaryAircrafts]=
	GET DIRECT LOOKUPS('RawMaterialsV2.xlsx', 'Military_Aircrafts', '1', 'B2') ~~|
"RMIntensityFamilyPct(pct)"[RMAll,MilitaryHelicopters]=
	GET DIRECT LOOKUPS('RawMaterialsV2.xlsx', 'Military_Helicopters', '1', 'B2') ~~|
"RMIntensityFamilyPct(pct)"[RMAll,MilitaryPatrol]=
	GET DIRECT LOOKUPS('RawMaterialsV2.xlsx', 'Military_Patrol', '1', 'B2') ~~|
"RMIntensityFamilyPct(pct)"[RMAll,MilitaryPrincipalSurface]=
	GET DIRECT LOOKUPS('RawMaterialsV2.xlsx', 'Military_PrincipSurface', '1', 'B2') ~~|
"RMIntensityFamilyPct(pct)"[RMAll,MilitarySubmarines]=
	GET DIRECT LOOKUPS('RawMaterialsV2.xlsx', 'Military_Submarines', '1', 'B2')
	~	percent
	~	Intensité matière (% de la masse unitaire) par famille et par matériau,
		en série temporelle. Chaque équation ne fait varier qu'une seule plage
		(RMAll), la famille étant un élément fixe : 18 séries par feuillet.
		Vensim interpole linéairement entre 1980, 1990 et 2000, et tient la
		valeur aux bornes en dehors (colonnes 1900 et 2050).
	|
```

on assigne la valeur de la famille à chaque unité

```vensim
"RMIntensityMilitaryUnitPct(pct)"[RMAll,MilitaryArmouredUnit]=
	"RMIntensityFamilyPct(pct)"[RMAll,MilitaryArmoured](Time) ~~|
"RMIntensityMilitaryUnitPct(pct)"[RMAll,MilitaryAircraftsUnit]=
	"RMIntensityFamilyPct(pct)"[RMAll,MilitaryAircrafts](Time) ~~|
"RMIntensityMilitaryUnitPct(pct)"[RMAll,MilitaryHelicoptersUnit]=
	"RMIntensityFamilyPct(pct)"[RMAll,MilitaryHelicopters](Time) ~~|
"RMIntensityMilitaryUnitPct(pct)"[RMAll,MilitaryPatrolUnit]=
	"RMIntensityFamilyPct(pct)"[RMAll,MilitaryPatrol](Time) ~~|
"RMIntensityMilitaryUnitPct(pct)"[RMAll,MilitarySurfaceUnit]=
	"RMIntensityFamilyPct(pct)"[RMAll,MilitaryPrincipalSurface](Time) ~~|
"RMIntensityMilitaryUnitPct(pct)"[RMAll,MilitarySubmarinesUnit]=
	"RMIntensityFamilyPct(pct)"[RMAll,MilitarySubmarines](Time)
	~	percent
	~	Intensité par matériel. La sous-plage ...Unit n'apparaît qu'à gauche :
		Vensim diffuse la valeur famille sur tous ses membres. C'est légal parce
		que le membre droit ne comporte plus qu'une plage variable (RMAll).
		Noter l'asymétrie de nommage : sous-plage MilitarySurfaceUnit (ligne
		276) mais élément de famille MilitaryPrincipalSurface (ligne 204).
	|
```

conversion pourcentages > t/TU

```vensim
"RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryTypeUnit]=
	"RMIntensityMilitaryUnitPct(pct)"[RMAll,MilitaryTypeUnit]
	* "MassUnitMilitaryUnit(t/TU)"[MilitaryTypeUnit] / 100
	~	t/TU
	~	IMPORTANT : "RMIntensityMilitaryUnitPct(pct)" n'est plus un Lookup mais
		une auxiliaire DÉJÀ évaluée au temps courant en 4b. L'appel (Time) qui
		figurait ici doit disparaître, sinon Vensim tente d'appeler une
		auxiliaire comme une fonction tabulée.
	|
```

Limite connue et assumée de la version ci-dessus : les flux de fin de vie sont évalués à la composition CONTEMPORAINE DU DÉMANTÈLEMENT, et non à celle de la construction.
Effet : surestimation de la récupération des matériaux en croissance (titane, cuivre, néodyme, argent, or), sous-estimation de l'aluminium. Le régime est dominant sur 1990-2010, puisque les décommissionnements de cette période portent sur du matériel des années 1960-1970. Le flux ENTRANT n'est pas concerné : une livraison a lieu au temps courant, donc Time est bien le millésime correct côté primaire et recyclé.
Si repris plus tard, 2 implémentations
- option1
	- colonne E de Military_MassUnit = année moyenne de livraison des 102 matériels (pondérée par les hausses annuelles de stock lues dans StocksMilitaryMods.xlsx) ;
	- "RMIntensityEoLPct(pct)" = profil famille appelé avec cette année au lieu de Time, en six équations calquées sur la section 4b ;
	- substitution de "RMIntensityEoL(t/TU)" à "RMIntensityMilitaryUnit" dans "FlowRMTUEoLMilitaryUnit" et "FlowRMTULostMilitaryUnit" UNIQUEMENT — pas dans FlowRMTUPrimary ni FlowRMTURecy.
- option2 Variante endogène : millésime calculé par INTEG de FlowNew*Time calcul de toutes les variantes de FlowRM natives du modèle

On continue avec toutes les variantes de flux de RM : primaire, recyc, EoL etc.
```vensim
CollectingRateMilitaryTypeUnit[MilitaryArmouredUnit]=
	0.7 ~~|
CollectingRateMilitaryTypeUnit[MilitaryAircraftsUnit]=
	0.7 ~~|
CollectingRateMilitaryTypeUnit[MilitaryHelicoptersUnit]=
	0.7 ~~|
CollectingRateMilitaryTypeUnit[MilitaryPatrolUnit]=
	0.7 ~~|
CollectingRateMilitaryTypeUnit[MilitarySurfaceUnit]=
	0.7 ~~|
CollectingRateMilitaryTypeUnit[MilitarySubmarinesUnit]=
	0.7
	~	Dmnl
	~	Taux de collecte au démantèlement. Uniforme par famille pour l'instant,
		pourra être affiné par matériel.
	|


"FlowRMTUPrimaryMilitaryUnit(t/yr)"[Country,RMAll,MilitaryTypeUnit]=
	"FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryTypeUnit]
	* "RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryTypeUnit]
	* (1 - RecyclingRateRM[RMAll](Time) * CorrecRR)
	~	t/Year
	~	Matière primaire pour les nouvelles livraisons de chaque matériel,
		net du recyclage disponible.
	|

"FlowRMTURecyMilitaryUnit(t/yr)"[Country,RMAll,MilitaryTypeUnit]=
	"FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryTypeUnit]
	* "RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryTypeUnit]
	* RecyclingRateRM[RMAll](Time) * CorrecRR
	~	t/Year
	~	Fraction recyclée injectée dans les nouvelles livraisons.
	|

"FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitaryTypeUnit]=
	"FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryTypeUnit]
	* "RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryTypeUnit]
	* RecyclingRateRM[RMAll](Time) * CorrecRR
	* CollectingRateMilitaryTypeUnit[MilitaryTypeUnit]
	~	t/Year
	~	Matière récupérable en fin de vie (baisses de stock observées × intensité
		× taux de recyclage × taux de collecte).
	|

"FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitaryTypeUnit]=
	"FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryTypeUnit]
	* "RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryTypeUnit]
	* (1 - RecyclingRateRM[RMAll](Time) * CorrecRR
	* CollectingRateMilitaryTypeUnit[MilitaryTypeUnit])
	~	t/Year
	~	Matière perdue en fin de vie (non recyclée ou non collectée).
	|
```

agrégation à l'échelle du secteur militaire total

```vensim
FlowRMMilitaryPR[Country,RMAll]=
	SUM("FlowRMTUPrimaryMilitaryUnit(t/yr)"[Country,RMAll,MilitaryTypeUnit!]) + SUM("FlowRMTURecyMilitaryUnit(t/yr)"\
		[Country,RMAll,MilitaryTypeUnit!])
	~	t/Year
	~	Somme des flux primaire + recyclé au niveau matériel.
		Remplace l'ancienne agrégation SUM(FlowRMTUPR[..., MilitaryType!])
		qui produisait des valeurs biaisées (pipeline générique avec LifeTime
		militaire neutralisé).
	|

FlowRMMilitaryPR smooth[Country,RMAll]=
	SMOOTH(FlowRMMilitaryPR[Country,RMAll], 3)
	~	t/Year
	~	moyenne glissante 3 ans pour les graphiques
	|
```

agrégations intermédiaires

```vensim
"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitaryAircrafts]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitaryAircraftsUnit!]) ~~|
"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitaryPrincipalSurface]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitarySurfaceUnit!]) ~~|
"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitaryPatrol]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitaryPatrolUnit!]) ~~|
"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitarySubmarines]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitarySubmarinesUnit!]) ~~|
"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitaryArmoured]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitaryArmouredUnit!]) ~~|
"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitaryHelicopters]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitaryHelicoptersUnit!])
	~	t/Year
	~	Tonnage recycle en fin de vie, par famille.
	|

"FlowRMLostMilitaryFam(t/yr)"[Country,RMAll,MilitaryAircrafts]=
	SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitaryAircraftsUnit!]) ~~|
"FlowRMLostMilitaryFam(t/yr)"[Country,RMAll,MilitaryPrincipalSurface]=
	SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitarySurfaceUnit!]) ~~|
"FlowRMLostMilitaryFam(t/yr)"[Country,RMAll,MilitaryPatrol]=
	SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitaryPatrolUnit!]) ~~|
"FlowRMLostMilitaryFam(t/yr)"[Country,RMAll,MilitarySubmarines]=
	SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitarySubmarinesUnit!]) ~~|
"FlowRMLostMilitaryFam(t/yr)"[Country,RMAll,MilitaryArmoured]=
	SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitaryArmouredUnit!]) ~~|
"FlowRMLostMilitaryFam(t/yr)"[Country,RMAll,MilitaryHelicopters]=
	SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitaryHelicoptersUnit!])
	~	t/Year
	~	Tonnage perdu en fin de vie, par famille.
	|
```

ancienne variable ??

```vensim
FlowRMTULost[Country,RMAll, MilitaryType]=
	"FlowRMLostMilitaryFam(t/yr)"[Country,RMAll,MilitaryType] ~~|
FlowRMTULost[Country,RMAll,InfrastOthers]=
	StockRMTUInfrastructure[Country,RMAll, InfrastOthers]* ( 1- RecyclingRateRM[RMAll](Time\
		)*CorrecRR) / LifeTimeOthers[InfrastOthers
	]
	~	t/Year
	~		|
	
```

## CHAÎNE ÉNERGIE

### ERM

```vensim
"ERMMilitaryUnitPrimary(PJ/yr)"[Country,RMAll,MilitaryTypeUnit]=
	"FlowRMTUPrimaryMilitaryUnit(t/yr)"[Country,RMAll,MilitaryTypeUnit]
	* "RMPrimary-Energy(GJ/t)"[RMAll] / 1e+06
	~	PJ/Year
	~	Énergie de production de la matière PRIMAIRE entrant dans les livraisons
		de chaque matériel militaire.
	|

"ERMMilitaryUnitRecy(PJ/yr)"[Country,RMAll,MilitaryTypeUnit]=
	"FlowRMTURecyMilitaryUnit(t/yr)"[Country,RMAll,MilitaryTypeUnit]
	* "RMRecycl-Energy(GJ/t)"[RMAll] / 1e+06
	~	PJ/Year
	~	Énergie de production de la matière RECYCLÉE entrant dans les livraisons
		de chaque matériel militaire.
	|

"ERMMilitaryUnit(PJ/yr)"[Country,RMAll,MilitaryTypeUnit]=
	"ERMMilitaryUnitPrimary(PJ/yr)"[Country,RMAll,MilitaryTypeUnit]
	+ "ERMMilitaryUnitRecy(PJ/yr)"[Country,RMAll,MilitaryTypeUnit]
	~	PJ/Year
	~	Énergie matière totale par matériel et par matériau. Analogue militaire
		de ERMTU[Country,RMAll,TechoUnitextend].
	|



"ERMMilitary(PJ/yr)"[Country,RMAll]=
	SUM("ERMMilitaryUnit(PJ/yr)"[Country,RMAll,MilitaryTypeUnit!])
	~	PJ/Year
	~	Énergie matière du secteur militaire par matériau. C'est la variable
		principale, homologue de ERMConstruction, ERMTransport, ERMAppliance.
	|

"ERMMilitaryPrimary(PJ/yr)"[Country,RMAll]=
	SUM("ERMMilitaryUnitPrimary(PJ/yr)"[Country,RMAll,MilitaryTypeUnit!]) ~~|
"ERMMilitaryRecy(PJ/yr)"[Country,RMAll]=
	SUM("ERMMilitaryUnitRecy(PJ/yr)"[Country,RMAll,MilitaryTypeUnit!])
	~	PJ/Year
	~	Décomposition primaire / recyclé, utile parce que RecyclingRateRM est un
		paramètre économie-wide mal adapté aux aciers de blindage et de coque et
		aux alliages aéronautiques certifiés : ces deux courbes montrent
		directement ce que coûte cette hypothèse.
	|

"ERMMilitaryFamily(PJ/yr)"[Country,MilitaryArmoured]=
	SUM("ERMMilitaryUnit(PJ/yr)"[Country,RMAll!,MilitaryArmouredUnit!]) ~~|
"ERMMilitaryFamily(PJ/yr)"[Country,MilitaryAircrafts]=
	SUM("ERMMilitaryUnit(PJ/yr)"[Country,RMAll!,MilitaryAircraftsUnit!]) ~~|
"ERMMilitaryFamily(PJ/yr)"[Country,MilitaryHelicopters]=
	SUM("ERMMilitaryUnit(PJ/yr)"[Country,RMAll!,MilitaryHelicoptersUnit!]) ~~|
"ERMMilitaryFamily(PJ/yr)"[Country,MilitaryPatrol]=
	SUM("ERMMilitaryUnit(PJ/yr)"[Country,RMAll!,MilitaryPatrolUnit!]) ~~|
"ERMMilitaryFamily(PJ/yr)"[Country,MilitaryPrincipalSurface]=
	SUM("ERMMilitaryUnit(PJ/yr)"[Country,RMAll!,MilitarySurfaceUnit!]) ~~|
"ERMMilitaryFamily(PJ/yr)"[Country,MilitarySubmarines]=
	SUM("ERMMilitaryUnit(PJ/yr)"[Country,RMAll!,MilitarySubmarinesUnit!])
	~	PJ/Year
	~	Énergie matière par famille d'équipement. Rappel de nommage : l'élément
		de MilitaryType est MilitaryPrincipalSurface, la sous-plage de matériels
		MilitarySurfaceUnit.
	|

"ERMMilitaryByUnit(PJ/yr)"[Country,MilitaryTypeUnit]=
	SUM("ERMMilitaryUnit(PJ/yr)"[Country,RMAll!,MilitaryTypeUnit])
	~	PJ/Year
	~	Énergie matière par matériel, tous matériaux. Permet d'identifier les
		plateformes qui font les pics (Charles de Gaulle, Le Redoutable...),
		comme "FlowRMTotByUnit(t/yr)" le fait côté tonnage.
	|


"ERMMilitaryTot(PJ/yr)"[Country]=
	SUM("ERMMilitary(PJ/yr)"[Country,RMAll!])
	~	PJ/Year
	~	Énergie matière totale du secteur militaire, toutes familles et tous
		matériaux confondus.
	|

"ERMMilitaryTot smooth(PJ/yr)"[Country]=
	SMOOTH("ERMMilitaryTot(PJ/yr)"[Country], 3)
	~	PJ/Year
	~	Moyenne glissante 3 ans. Le flux militaire brut est très discontinu
		(livraisons par tranches), la version lissée est celle à afficher.
	|
```

**Attention ici**

1. NE PAS ajouter "ERMMilitaryTot" à "ERMTOTAL(PJ/yr)". FlowRMMilitaryPR (ligne 181) est déjà branché en amont, donc la matière militaire est déjà comptée dans ENonFerrousMetals / ESteel&Fe / ENonMetallicMinerals. Ce bloc est une DÉCOMPOSITION de reporting, pas un terme supplémentaire. L'ajouter doublerait le militaire.

2. NE PAS confondre avec ERMTU[Country,RMAll,MilitaryType]. Le pipeline générique produit toujours des valeurs pour les six familles MilitaryX, mais à partir de flux biaisés (LiFeTimeTU = 1e9). Ces valeurs sont mortes et ne doivent jamais être sommées avec celles de ce bloc.

3. NE PAS additionner avec EmbEnergyMilitary. Ce sont deux périmètres différents : ERM est l'énergie de PRODUCTION DES MATIÈRES, EmbEnergy est l'énergie incorporée estimée à partir des prix (capital industriel, assemblage, R&D). Elles se recouvrent partiellement.

4. Périmètre. Comme pour la matière, il s'agit des 102 plateformes majeures et non du secteur militaire : ni munitions, ni armement léger, ni logistique non blindée, ni bases, ni spatial, ni MCO.

### EmbTU

définition des prix des équipements, qui est l'entrée de Dymemds pour l'énergie embarquée. Hypothèses importantes et discutables

```vensim
"EmbTU(MJ/TU)"[Country,xx]=
	"PriceRef(1998$/TU)"[xx,Country]*"aFPIndLOCAL(MJ/$)"[Country]
	~	MJ/TU
	~		|

"EmbTU(MJ/yr)"[Country, xx]=
	"EmbTU(MJ/TU)"[Country, xx]*"FlowNewTU(/yr)"[Country, xx]
	~	MJ/Year
	~		|

"EmbTU(PJ/yr)"[Country, xx]=
	"EmbTU(MJ/yr)"[Country, xx]/1e+09 ~~|

EmbEnergyMilitaryTotal[Country]=
	SUM("EmbTU(PJ/yr)"[Country,MilitaryType!])
	~	PJ/Year
	~	total energie embarquee militaire
	|

EmbEnergyMilitaryTotal smooth[Country]=
	SMOOTH(EmbEnergyMilitaryTotal[Country], 3)
	~	PJ/Year
	~	lissage causal 3 ans
	|

"EmbTU(PJ/yr) smooth"[Country,MilitaryType]=
	SMOOTH("EmbTU(PJ/yr)"[Country,MilitaryType], 3)
	~	PJ/Year
	~	energie embarquee par type, lissee
	|
```


### Masses et prix des proxys civils

```
"MassRefCivilAircraft(t/TU)"=
	"weightAircraft(t)"
	~	t/TU
	~	Avion civil de reference DyMEMDS, declare a 120 t.
	|

"MassRefCivilHV(t/TU)"=
	6
	~	t/TU
	~	Poids lourd civil de reference. DyMEMDS pose HV = 4 x LV, et LV
	|

"MassRefCivilVessel(t/TU)"=
	10000
	~	t/TU
	~	HYPOTHESE A DISCUTER (pas de source dans le modele). PriceRef[vessel] = 2e8 $1998 ; 
	10 000 t lege + 200 M$1998 => 20 000 $1998/t, ordre de
		grandeur d'un navire complexe (methanier, paquebot)
	|
```


```
"PriceRefCivilAircraft(1998$/t)"[Country]=
	"PriceRef(1998$/TU)"[Aircarft,Country]/"MassRefCivilAircraft(t/TU)"
	~	$/t
	~		|

"PriceRefCivilHV(1998$/t)"[Country]=
	"PriceRef(1998$/TU)"["HV-ICV",Country]/"MassRefCivilHV(t/TU)"
	~	$/t
	~		|

"PriceRefCivilVessel(1998$/t)"[Country]=
	"PriceRef(1998$/TU)"[vessel,Country]/"MassRefCivilVessel(t/TU)"
	~	$/t
	~		|
```

### Masse pondérée annuelle par famille
 Rend le prix de famille coherent avec le mix annuel. Prix_famille x Livraisons_famille = tonnage livre x prix/tonne, donc la valeur totale est conservee sans passer par un prix par materiel.                                                        

```
"MassRefMilitaryFam(t/TU)"[Country,MilitaryArmoured]=
	XIDZ(SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryArmouredUnit!]
	*"MassUnitMilitaryUnit(t/TU)"[MilitaryArmouredUnit!]),
	SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryArmouredUnit!]),
	"MassMeanMilitaryFam(t/TU)"[MilitaryArmoured]) ~~|
"MassRefMilitaryFam(t/TU)"[Country,MilitaryAircrafts]=
	XIDZ(SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryAircraftsUnit!]
	*"MassUnitMilitaryUnit(t/TU)"[MilitaryAircraftsUnit!]),
	SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryAircraftsUnit!]),
	"MassMeanMilitaryFam(t/TU)"[MilitaryAircrafts]) ~~|
"MassRefMilitaryFam(t/TU)"[Country,MilitaryHelicopters]=
	XIDZ(SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryHelicoptersUnit!]
	*"MassUnitMilitaryUnit(t/TU)"[MilitaryHelicoptersUnit!]),
	SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryHelicoptersUnit!]),
	"MassMeanMilitaryFam(t/TU)"[MilitaryHelicopters]) ~~|
"MassRefMilitaryFam(t/TU)"[Country,MilitaryPatrol]=
	XIDZ(SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryPatrolUnit!]
	*"MassUnitMilitaryUnit(t/TU)"[MilitaryPatrolUnit!]),
	SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryPatrolUnit!]),
	"MassMeanMilitaryFam(t/TU)"[MilitaryPatrol]) ~~|
"MassRefMilitaryFam(t/TU)"[Country,MilitaryPrincipalSurface]=
	XIDZ(SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitarySurfaceUnit!]
	*"MassUnitMilitaryUnit(t/TU)"[MilitarySurfaceUnit!]),
	SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitarySurfaceUnit!]),
	"MassMeanMilitaryFam(t/TU)"[MilitaryPrincipalSurface]) ~~|
"MassRefMilitaryFam(t/TU)"[Country,MilitarySubmarines]=
	XIDZ(SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitarySubmarinesUnit!]
	*"MassUnitMilitaryUnit(t/TU)"[MilitarySubmarinesUnit!]),
	SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitarySubmarinesUnit!]),
	"MassMeanMilitaryFam(t/TU)"[MilitarySubmarines])
	~	t/TU
	~	Masse unitaire moyenne des materiels effectivement livres dans l'annee,
		par famille. Rappel de nommage : element de famille
		MilitaryPrincipalSurface, sous-plage MilitarySurfaceUnit.
	|
```

(Masse moyenne non pondérée, variable secours si pas de livraison)

Moyenne arithmetique des masses unitaires lues dans Military_MassUnit.        
 Les diviseurs sont les cardinaux des sous-plages : 10 blindes, 44 avions,     
14 helicos, 13 patrouilleurs, 14 batiments de surface, 7 sous-marins.         
```
"MassMeanMilitaryFam(t/TU)"[MilitaryArmoured]=
	SUM("MassUnitMilitaryUnit(t/TU)"[MilitaryArmouredUnit!])/10 ~~|
"MassMeanMilitaryFam(t/TU)"[MilitaryAircrafts]=
	SUM("MassUnitMilitaryUnit(t/TU)"[MilitaryAircraftsUnit!])/44 ~~|
"MassMeanMilitaryFam(t/TU)"[MilitaryHelicopters]=
	SUM("MassUnitMilitaryUnit(t/TU)"[MilitaryHelicoptersUnit!])/14 ~~|
"MassMeanMilitaryFam(t/TU)"[MilitaryPatrol]=
	SUM("MassUnitMilitaryUnit(t/TU)"[MilitaryPatrolUnit!])/13 ~~|
"MassMeanMilitaryFam(t/TU)"[MilitaryPrincipalSurface]=
	SUM("MassUnitMilitaryUnit(t/TU)"[MilitarySurfaceUnit!])/14 ~~|
"MassMeanMilitaryFam(t/TU)"[MilitarySubmarines]=
	SUM("MassUnitMilitaryUnit(t/TU)"[MilitarySubmarinesUnit!])/7
	~	t/TU
	~			Sert de valeur de repli les annees sans livraison.
	|
```


###  Coef Civil-> Militaire  
```
PremiumMilitary[MilitaryAircrafts]=
	1.5 ~~|
PremiumMilitary[MilitaryHelicopters]=
	1.2 ~~|
PremiumMilitary[MilitaryArmoured]=
	3.5 ~~|
PremiumMilitary[MilitaryPatrol]=
	1.5 ~~|
PremiumMilitary[MilitaryPrincipalSurface]=
	2 ~~|
PremiumMilitary[MilitarySubmarines]=
	12
	~	
	~	 helicoptere proche du civil, la
		cellule est souvent la meme, blinde piece massive mais peu de matiere, SNA/SNLE
		coque epaisse + chaufferie nucleaire.
	|
```

### Calcul final prix militaires (regle de trois)
```
"PriceRef(1998$/TU)"[MilitaryAircrafts,Country]=
	"PriceRefCivilAircraft(1998$/t)"[Country]
	*"MassRefMilitaryFam(t/TU)"[Country,MilitaryAircrafts]
	*PremiumMilitary[MilitaryAircrafts] ~~|
"PriceRef(1998$/TU)"[MilitaryHelicopters,Country]=
	"PriceRefCivilAircraft(1998$/t)"[Country]
	*"MassRefMilitaryFam(t/TU)"[Country,MilitaryHelicopters]
	*PremiumMilitary[MilitaryHelicopters] ~~|
"PriceRef(1998$/TU)"[MilitaryArmoured,Country]=
	"PriceRefCivilHV(1998$/t)"[Country]
	*"MassRefMilitaryFam(t/TU)"[Country,MilitaryArmoured]
	*PremiumMilitary[MilitaryArmoured] ~~|
"PriceRef(1998$/TU)"[MilitaryPatrol,Country]=
	"PriceRefCivilVessel(1998$/t)"[Country]
	*"MassRefMilitaryFam(t/TU)"[Country,MilitaryPatrol]
	*PremiumMilitary[MilitaryPatrol] ~~|
"PriceRef(1998$/TU)"[MilitaryPrincipalSurface,Country]=
	"PriceRefCivilVessel(1998$/t)"[Country]
	*"MassRefMilitaryFam(t/TU)"[Country,MilitaryPrincipalSurface]
	*PremiumMilitary[MilitaryPrincipalSurface] ~~|
"PriceRef(1998$/TU)"[MilitarySubmarines,Country]=
	"PriceRefCivilVessel(1998$/t)"[Country]
	*"MassRefMilitaryFam(t/TU)"[Country,MilitarySubmarines]
	*PremiumMilitary[MilitarySubmarines] ~~|
```




## Branchements sur les boucles dymemds

```vensim
StockTU[MilitaryArmoured,Country]=
	SUM(StockTUMilitaryUnit[Country,MilitaryArmouredUnit!](Time)) ~~|
StockTU[MilitaryAircrafts,Country]=
	SUM(StockTUMilitaryUnit[Country,MilitaryAircraftsUnit!](Time)) ~~|
StockTU[MilitaryHelicopters,Country]=
	SUM(StockTUMilitaryUnit[Country,MilitaryHelicoptersUnit!](Time)) ~~|
StockTU[MilitaryPatrol,Country]=
	SUM(StockTUMilitaryUnit[Country,MilitaryPatrolUnit!](Time)) ~~|
StockTU[MilitaryPrincipalSurface,Country]=
	SUM(StockTUMilitaryUnit[Country,MilitarySurfaceUnit!](Time)) ~~|
StockTU[MilitarySubmarines,Country]=
	SUM(StockTUMilitaryUnit[Country,MilitarySubmarinesUnit!](Time))
	~	
	~	
	|

FlowTUEOL[Country, TechoUnitextendNoMil]=
	StockTUInfrastructure[Country,TechoUnitextendNoMil]/LiFeTimeTU[TechoUnitextendNoMil]\
		 ~~|
FlowTUEOL[Country, MilitaryType]=
	"FlowEoLTUMilitaryFam(/yr)"[Country,MilitaryType] ~~|
```

*Bloc conservé en commentaire dans le fichier source — pour info seulement, présent plus bas dans le modèle. Code désactivé (la version active de ces variables se trouve plus bas dans le modèle complet) :*

```vensim
FlowRMTUEoL[Country,RMAll, MilitaryType]=
	"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitaryType] ~~|

FlowRMTUEoL[Country,RMAll, TechoUnitextendNoMil]=
	(StockRMTUInfrastructure[Country,RMAll, TechoUnitextendNoMil]/LiFeTimeTU[TechoUnitextendNoMil\
		])*(RecyclingRateRM[RMAll](Time)*CorrecRR) ~~|
FlowRMTUEoL[Country,RMAll, MilitaryType]=
	"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitaryType] ~~|

StockTUinUse[Country,MilitaryType]=
	StockTU[MilitaryType,Country]
	~	
	~		|

SumFlowRMPRSectors[RMAll,Country]=
	FlowRMConstructionPR[RMAll, Country]+FlowRMEnergyPR[RMAll, Country]+FlowRMTransportPR
	[RMAll, Country]+FlowRMApplianceElectronicPR[RMAll, Country]+FlowRMOthersPR[RMAll,Others\
		,Country]+FlowRMMilitaryPR[Country,RMAll] ~~|

XMilitary1=
	0
	~	Dmnl
	~	Interrupteur du militaire top-down (element Military1 de InfrastOthers).
		0 = militaire compte uniquement en bottom-up par les 102 plateformes
		    (FlowRMMilitaryPR), branche sur SumFlowRMPRSectors. CHOIX ACTUEL.
		1 = ancienne part top-down complete (perimetre large : munitions,
		    armement leger, bases, logistique, spatial...). ATTENTION : ne pas
		    mettre a 1 sans neutraliser le branchement bottom-up, sinon double
		    comptage partiel sur la France.
		L'element Military1 est volontairement CONSERVE dans le subscript
		InfrastOthers : le supprimer decalerait la lecture de LifeTimeOthers
		('C10', 6 valeurs en ligne) et des autres colonnes de l'onglet.
	|

EmbTU:
	"LV-ICV","LV-Hyb","LV-PHEV","LV-BEV","HV-ICV","HV-Hyb","HV-PHEV","HV-BEV",HighSpeed,\
		IntercityElec,IntercityFossil,Aircarft,vessel,Residential,CivilEngineering,WashingMachine\
		,TV,Fridge,DishWasher,Mobile,NoteBook,Tablet,
		MilitaryAircrafts,MilitaryPrincipalSurface,MilitaryPatrol,MilitarySubmarines,MilitaryArmoured\
		,MilitaryHelicopters
	~	
	~		|

TechoUnitextend:
	LV,HV,Locomotive,vessel,Aircarft,Residential,CivilEngineering,service,WashingMachine\
		,TV,Fridge,DishWasher,Mobile,NoteBook,Tablet,"LV-ICV","LV-Hyb","LV-PHEV","LV-BEV","HV-ICV"\
		,"HV-Hyb","HV-PHEV","HV-BEV",HighSpeed,IntercityFossil,IntercityElec,RailHighspeed,\
		RailIntercityElec,RailIntercityFossil,BatteryHyb, BatteryEV,BatteryPHEV,BatteryMobile\
		,BatteryNotebook,BatteryTablet,Roadkm,
	MilitaryAircrafts,MilitaryPrincipalSurface,MilitaryPatrol,MilitarySubmarines,MilitaryArmoured\
		,MilitaryHelicopters
	~	
	~		|

TechoUnitextendNoMil:
	LV,HV,Locomotive,vessel,Aircarft,Residential,CivilEngineering,service,WashingMachine\
		,TV,Fridge,DishWasher,Mobile,NoteBook,Tablet,"LV-ICV","LV-Hyb","LV-PHEV","LV-BEV","HV-ICV"\
		,"HV-Hyb","HV-PHEV","HV-BEV",HighSpeed,IntercityFossil,IntercityElec,RailHighspeed,\
		RailIntercityElec,RailIntercityFossil,BatteryHyb, BatteryEV,BatteryPHEV,BatteryMobile\
		,BatteryNotebook,BatteryTablet,Roadkm
	~	
	~	TechoUnitextend privé des 6 familles MilitaryType (36 elements).
	|

TechnoUnits:
	vessel,Aircarft, Residential,CivilEngineering,service,WashingMachine,TV,Fridge,DishWasher\
		,Mobile,NoteBook,Tablet,"LV-ICV","LV-Hyb","LV-PHEV","LV-BEV","HV-ICV","HV-Hyb","HV-PHEV"\
		,"HV-BEV",HighSpeed,IntercityFossil,IntercityElec,RailHighspeed,RailIntercityElec,RailIntercityFossil\
		,BatteryHyb, BatteryEV,BatteryPHEV,BatteryMobile,BatteryNotebook,BatteryTablet,Roadkm\
		,
	MilitaryAircrafts,MilitaryPrincipalSurface,MilitaryPatrol,MilitarySubmarines,
	        MilitaryArmoured,MilitaryHelicopters
	~	
	~		|

InfrastructureEcons:
	RoadV,LV,HV,Locomotive,Railway,Aircarft,AircarftD,AircarftI,vessel, VesselD,vesselI,\
		Transport,TransportD, TransportI,Building,Residential,CivilEngineering,Industry,HeavyInd\
		,Manufacture,clothing,furniture,agriculture,Food,service,"LV-ICV","LV-Hyb","LV-PHEV"\
		,"LV-BEV","HV-ICV","HV-Hyb","HV-PHEV","HV-BEV",HighSpeed,intercity,IntercityFossil,\
		IntercityElec,RailHighspeed,RailIntercityElec,RailIntercityFossil,ProdFossil,ProdElec\
		,BatteryHyb, BatteryEV,BatteryPHEV,ApplianceElectronic,WashingMachine,TV,Fridge,DishWasher\
		,Mobile,NoteBook,Tablet,BatteryMobile,BatteryNotebook,BatteryTablet,Others,TransportOthers1\
		,ChemistryProduction1,EngineeringMechanic1,Packaging1,Military1,Others1,Roadkm,
	MilitaryAircrafts,MilitaryPrincipalSurface,MilitaryPatrol,MilitarySubmarines,MilitaryArmoured\
		,MilitaryHelicopters
	~	
	~		|
```

# Annexes

## VARIABLES DE VÉRIFICATION

```vensim
StockTUMilitaryUnitAg[Country]=
	SUM(StockTUMilitaryUnit[Country,MilitaryTypeUnit!](Time))
	~	TU
	~	Total du parc militaire (toutes familles, tous matériels) — check.
	|

"FlowRMMilitaryUnitPR(t/yr)"[Country,RMAll]=
	SUM("FlowRMTUPrimaryMilitaryUnit(t/yr)"[Country,RMAll,MilitaryTypeUnit!])
	+ SUM("FlowRMTURecyMilitaryUnit(t/yr)"[Country,RMAll,MilitaryTypeUnit!])
	~	t/Year
	~	Flux matière primaire+recyclée militaire consolidé (analogue à
		l'ancien FlowRMMilitaryPR mais recalculé au niveau matériel).
		Doit coïncider avec SUM(FlowRMTUPR[Country,RMAll,MilitaryType!]).
	|

"FlowRMTotByUnit(t/yr)"[Country,MilitaryTypeUnit]=
	SUM("FlowRMTUPrimaryMilitaryUnit(t/yr)"[Country,RMAll!,MilitaryTypeUnit])
	+ SUM("FlowRMTURecyMilitaryUnit(t/yr)"[Country,RMAll!,MilitaryTypeUnit])
	~	t/Year
	~	Flux matiere entrant total (primaire + recyclee, tous materiaux) par
		materiel militaire. Permet d'identifier les unites qui contribuent
		aux grosses variations du flux militaire global.
	|

"FlowRMEoLByUnit(t/yr)"[Country,MilitaryTypeUnit]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll!,MilitaryTypeUnit])
	+ SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll!,MilitaryTypeUnit])
	~	t/Year
	~	Flux matiere sortant (recyclee + perdue en fin de vie, tous materiaux)
		par materiel militaire. Utile pour reperer les decommissionnements massifs.
	|

"FlowRMTotByUnit smooth(t/yr)"[Country,MilitaryTypeUnit]=
	SMOOTH("FlowRMTotByUnit(t/yr)"[Country,MilitaryTypeUnit], 3)
	~	t/Year
	~	Moyenne glissante 3 ans du flux matiere entrant par materiel.
	|

"FlowRMEoLByUnit smooth(t/yr)"[Country,MilitaryTypeUnit]=
	SMOOTH("FlowRMEoLByUnit(t/yr)"[Country,MilitaryTypeUnit], 3)
	~	t/Year
	~	Moyenne glissante 3 ans du flux matiere sortant par materiel.
	|

"FlowRMTUPrimaryMilitaryAg(t/yr)"[Country,RMAll]=
	SUM("FlowRMTUPrimaryMilitaryUnit(t/yr)"[Country,RMAll,MilitaryTypeUnit!])
	~	t/Year
	~	Flux matiere primaire militaire agrege (par materiau), somme sur tous
		les materiels. Utile pour graphes Primary vs Recy vs EoL.
	|

"FlowRMTURecyMilitaryAg(t/yr)"[Country,RMAll]=
	SUM("FlowRMTURecyMilitaryUnit(t/yr)"[Country,RMAll,MilitaryTypeUnit!])
	~	t/Year
	~	Flux matiere recyclee militaire agrege (par materiau).
	|

"FlowRMTUEoLMilitaryAg(t/yr)"[Country,RMAll]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitaryTypeUnit!])
	~	t/Year
	~	Flux matiere EoL recuperee militaire agrege (par materiau).
	|

"FlowRMTULostMilitaryAg(t/yr)"[Country,RMAll]=
	SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitaryTypeUnit!])
	~	t/Year
	~	Flux matiere perdue militaire agrege (par materiau).
	|


"RMIntensityTU(t/TU)"[Country,RMAll,MilitaryArmoured]=
	ZIDZ( SUM("RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryArmouredUnit!]
	          * StockTUMilitaryUnit[Country,MilitaryArmouredUnit!](Time)),
	      SUM(StockTUMilitaryUnit[Country,MilitaryArmouredUnit!](Time)) ) ~~|
"RMIntensityTU(t/TU)"[Country,RMAll,MilitaryAircrafts]=
	ZIDZ( SUM("RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryAircraftsUnit!]
	          * StockTUMilitaryUnit[Country,MilitaryAircraftsUnit!](Time)),
	      SUM(StockTUMilitaryUnit[Country,MilitaryAircraftsUnit!](Time)) ) ~~|
"RMIntensityTU(t/TU)"[Country,RMAll,MilitaryHelicopters]=
	ZIDZ( SUM("RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryHelicoptersUnit!]
	          * StockTUMilitaryUnit[Country,MilitaryHelicoptersUnit!](Time)),
	      SUM(StockTUMilitaryUnit[Country,MilitaryHelicoptersUnit!](Time)) ) ~~|
"RMIntensityTU(t/TU)"[Country,RMAll,MilitaryPatrol]=
	ZIDZ( SUM("RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryPatrolUnit!]
	          * StockTUMilitaryUnit[Country,MilitaryPatrolUnit!](Time)),
	      SUM(StockTUMilitaryUnit[Country,MilitaryPatrolUnit!](Time)) ) ~~|
"RMIntensityTU(t/TU)"[Country,RMAll,MilitaryPrincipalSurface]=
	ZIDZ( SUM("RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitarySurfaceUnit!]
	          * StockTUMilitaryUnit[Country,MilitarySurfaceUnit!](Time)),
	      SUM(StockTUMilitaryUnit[Country,MilitarySurfaceUnit!](Time)) ) ~~|
"RMIntensityTU(t/TU)"[Country,RMAll,MilitarySubmarines]=
	ZIDZ( SUM("RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitarySubmarinesUnit!]
	          * StockTUMilitaryUnit[Country,MilitarySubmarinesUnit!](Time)),
	      SUM(StockTUMilitaryUnit[Country,MilitarySubmarinesUnit!](Time)) )
	~	t/TU
	~	Intensité matière moyenne du parc, pondérée par le stock des matériels
		de chaque famille. ZIDZ renvoie 0 quand le stock est nul (évite les NaN).
		Utilisée seulement pour alimenter StockRMTUInfrastp (traçage du stock
		matière militaire) — les flux FlowRMTUPrimary/Recy/EoL/Lost du militaire
		sont recalculés directement au niveau matériel dans le bloc
		militaire détaillé.
	|
```

## Anciens blocs

Deux Pipelines de calcul militaire actuellement, un à part du modèle avec variables nommées xxMilitary, et un via les variables natives de DyMEMDS et l'intégration des technologies militaires dans le subscript TechoUnitextend.

Le pipeline générique TechoUnitextend continue de fonctionner sur les 6 familles MilitaryType, qui sont maintenant recalculées comme la somme des matériels détaillés du subscript MilitaryTypeUnit. les variables génériques FlowRMTUPrimary/Recy/EoL/Lost produisent des valeurs pour les familles de MilitaryX, mais ces valeurs sont biaisées :
- LiFeTimeTU[MilitaryType] = 1e9 -> FlowRMTUEoL générique ~ 0
- collectingRateTU[MilitaryType] = 0 -> FlowRMTURecy générique ~ 0
- FlowRMTUPrimary générique = dStockRM + FlowRMTULost + XSRMTU capture la matière primaire pour la croissance nette du parc uniquement. Ces valeurs génériques ne sont plus consommées par aucun agrégateur militaire (voir FlowRMMilitaryPR rerouté ci-dessous) — donc leur biais est inoffensif.
Les VRAIS flux militaires sont calculés au niveau matériel dans les variables "FlowRMTUPrimaryMilitaryUnit", "FlowRMTURecyMilitaryUnit", "FlowRMTUEoLMilitaryUnit" et "FlowRMTULostMilitaryUnit" (voir §7), et agrégés vers FlowRMMilitaryPR via une redirection explicite.

L'ancienne construction appelait des variantes sans xxxUnit, dont les stocks étaient lus depuis un fichier excel sommé sur la famille

```vensim
MilitaryType:
	MilitaryAircrafts,MilitaryPrincipalSurface,MilitaryPatrol,MilitarySubmarines,MilitaryArmoured\
		,MilitaryHelicopters
	~	
	~		|

LiFeTimeTU[MilitaryType]=
	1e+09
	~	
	~		|

collectingRateTU[MilitaryType]=
	0.7
	~	
	~		|
```