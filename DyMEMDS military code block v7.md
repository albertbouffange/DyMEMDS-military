# Bloc militaire détaillé

Prérequis dans les fichiers Excel :
- StocksMilitaryMods.xlsx : feuille 'Stocks_smooth' avec colonnes A=code, B=family, C=unit_name, D..= années (1900, 1970..2014, 2100).
- RawMaterialsV2.xlsx : onglets d'intensité dynamique par famille et onglet 'Military_MassUnit' (D=mass_t).
Attention Vensim : l'ordre des matériels dans le subscript MilitaryUnits doit être STRICTEMENT identique à l'ordre des lignes dans les feuilles Excel.

## Subscripts

On crée tous nos objets militaires pour lesquels on va définir des équations, et on les regroupe en familles

```vensim
MilitaryArmouredUnit:
	MilitaryArmoured Aml,MilitaryArmoured Amx10,MilitaryArmoured AmxVtt,MilitaryArmoured Amx13\
		,MilitaryArmoured Amx30,MilitaryArmoured Erc90,MilitaryArmoured Leclerc,MilitaryArmoured Vab\
		,MilitaryArmoured Vbci,MilitaryArmoured VblM11
	~	
	~	Sous-plage : 10 matériels de la famille MilitaryArmoured (ordre de la \
		feuille "Stocks").
	|

MilitaryAircraftsUnit:
	MilitaryAircrafts A300Variants,MilitaryAircrafts A400MAtlas,MilitaryAircrafts Alize,\
		MilitaryAircrafts Alphajet,MilitaryAircrafts Atlantic,MilitaryAircrafts Breguet941,\
		MilitaryAircrafts C130H,MilitaryAircrafts C212,MilitaryAircrafts C47,MilitaryAircrafts Cn235\
		,MilitaryAircrafts Caravelle,MilitaryAircrafts CessnaCaravanIi,MilitaryAircrafts Crusader\
		,MilitaryAircrafts Dc6,MilitaryAircrafts Dc8,MilitaryAircrafts Dhc6,MilitaryAircrafts E2CHawkeye\
		,MilitaryAircrafts E3F,MilitaryAircrafts Emb121,MilitaryAircrafts Emb312,MilitaryAircrafts EtendardIv\
		,MilitaryAircrafts F100D,MilitaryAircrafts Flamant,MilitaryAircrafts FougaMs,MilitaryAircrafts Jaguar\
		,MilitaryAircrafts Kc135,MilitaryAircrafts Mirage2000,MilitaryAircrafts MirageF1,MilitaryAircrafts MirageIii\
		,MilitaryAircrafts MirageIv,MilitaryAircrafts MirageV,MilitaryAircrafts MystereFalcon\
		,MilitaryAircrafts Navajo,MilitaryAircrafts Noratlas,MilitaryAircrafts Nord264,MilitaryAircrafts Rafale\
		,MilitaryAircrafts Sp2HNeptune,MilitaryAircrafts SuperEtendard,MilitaryAircrafts SuperMystere\
		,MilitaryAircrafts T33A,MilitaryAircrafts Tbm700,MilitaryAircrafts TransallC160,MilitaryAircrafts VautourIi\
		,MilitaryAircrafts LightTraining
	~	
	~	Sous-plage : 44 matériels de la famille MilitaryAircrafts (ordre de la \
		feuille "Stocks").
	|

MilitaryHelicoptersUnit:
	MilitaryHelicopters Alouette,MilitaryHelicopters Caracal,MilitaryHelicopters Colibiri\
		,MilitaryHelicopters Cougar,MilitaryHelicopters Dauphin,MilitaryHelicopters Ecureuil\
		,MilitaryHelicopters Fennec,MilitaryHelicopters Gazelle,MilitaryHelicopters Nh90Tth\
		,MilitaryHelicopters Panther,MilitaryHelicopters Puma,MilitaryHelicopters SuperFrelon\
		,MilitaryHelicopters SuperPuma,MilitaryHelicopters Tigre
	~	
	~	Sous-plage : 14 matériels de la famille MilitaryHelicopters (ordre de la \
		feuille "Stocks").
	|

MilitaryPatrolUnit:
	MilitaryPatrol Albatros,MilitaryPatrol Athos,MilitaryPatrol ClasseBay,MilitaryPatrol DEstienneDOrves\
		,MilitaryPatrol Grebe,MilitaryPatrol LeopardInstruction,MilitaryPatrol Opv54,MilitaryPatrol P400\
		,MilitaryPatrol PatrouilleursSp,MilitaryPatrol Sterne,MilitaryPatrol Trident,MilitaryPatrol TypeMso\
		,MilitaryPatrol TypeD
	~	
	~	Sous-plage : 13 matériels de la famille MilitaryPatrol (ordre de la \
		feuille "Stocks").
	|

MilitarySurfaceUnit:
	MilitarySurface C65,MilitarySurface CharlesDeGaulle,MilitarySurface Clemenceau,MilitarySurface Colbert\
		,MilitarySurface CommandantRiviere,MilitarySurface DEstienneDOrves,MilitarySurface Escorteurs\
		,MilitarySurface F67,MilitarySurface Floreal,MilitarySurface JeanneDArc,MilitarySurface LaFayette\
		,MilitarySurface Suffren,MilitarySurface TypeE52,MilitarySurface TypeF70
	~	
	~	Sous-plage : 14 matériels de la famille MilitarySurface (ordre de la \
		feuille "Stocks").
	|

MilitarySubmarinesUnit:
	MilitarySubmarines Agosta,MilitarySubmarines Arethuse,MilitarySubmarines Daphne,MilitarySubmarines LeRedoutable\
		,MilitarySubmarines LeTriomphant,MilitarySubmarines Narval,MilitarySubmarines Rubis
	~	
	~		|

MilitaryUnits:
	MilitaryArmoured Aml,MilitaryArmoured Amx10,MilitaryArmoured AmxVtt,MilitaryArmoured Amx13\
		,MilitaryArmoured Amx30,MilitaryArmoured Erc90,MilitaryArmoured Leclerc,MilitaryArmoured Vab\
		,MilitaryArmoured Vbci,MilitaryArmoured VblM11,MilitaryAircrafts A300Variants,MilitaryAircrafts A400MAtlas\
		,MilitaryAircrafts Alize,MilitaryAircrafts Alphajet,MilitaryAircrafts Atlantic,MilitaryAircrafts Breguet941\
		,MilitaryAircrafts C130H,MilitaryAircrafts C212,MilitaryAircrafts C47,MilitaryAircrafts Cn235\
		,MilitaryAircrafts Caravelle,MilitaryAircrafts CessnaCaravanIi,MilitaryAircrafts Crusader\
		,MilitaryAircrafts Dc6,MilitaryAircrafts Dc8,MilitaryAircrafts Dhc6,MilitaryAircrafts E2CHawkeye\
		,MilitaryAircrafts E3F,MilitaryAircrafts Emb121,MilitaryAircrafts Emb312,MilitaryAircrafts EtendardIv\
		,MilitaryAircrafts F100D,MilitaryAircrafts Flamant,MilitaryAircrafts FougaMs,MilitaryAircrafts Jaguar\
		,MilitaryAircrafts Kc135,MilitaryAircrafts Mirage2000,MilitaryAircrafts MirageF1,MilitaryAircrafts MirageIii\
		,MilitaryAircrafts MirageIv,MilitaryAircrafts MirageV,MilitaryAircrafts MystereFalcon\
		,MilitaryAircrafts Navajo,MilitaryAircrafts Noratlas,MilitaryAircrafts Nord264,MilitaryAircrafts Rafale\
		,MilitaryAircrafts Sp2HNeptune,MilitaryAircrafts SuperEtendard,MilitaryAircrafts SuperMystere\
		,MilitaryAircrafts T33A,MilitaryAircrafts Tbm700,MilitaryAircrafts TransallC160,MilitaryAircrafts VautourIi\
		,MilitaryAircrafts LightTraining,MilitaryHelicopters Alouette,MilitaryHelicopters Caracal\
		,MilitaryHelicopters Colibiri,MilitaryHelicopters Cougar,MilitaryHelicopters Dauphin\
		,MilitaryHelicopters Ecureuil,MilitaryHelicopters Fennec,MilitaryHelicopters Gazelle\
		,MilitaryHelicopters Nh90Tth,MilitaryHelicopters Panther,MilitaryHelicopters Puma,MilitaryHelicopters SuperFrelon\
		,MilitaryHelicopters SuperPuma,MilitaryHelicopters Tigre,MilitaryPatrol Albatros,MilitaryPatrol Athos\
		,MilitaryPatrol ClasseBay,MilitaryPatrol DEstienneDOrves,MilitaryPatrol Grebe,MilitaryPatrol LeopardInstruction\
		,MilitaryPatrol Opv54,MilitaryPatrol P400,MilitaryPatrol PatrouilleursSp,MilitaryPatrol Sterne\
		,MilitaryPatrol Trident,MilitaryPatrol TypeMso,MilitaryPatrol TypeD,MilitarySurface C65\
		,MilitarySurface CharlesDeGaulle,MilitarySurface Clemenceau,MilitarySurface Colbert\
		,MilitarySurface CommandantRiviere,MilitarySurface DEstienneDOrves,MilitarySurface Escorteurs\
		,MilitarySurface F67,MilitarySurface Floreal,MilitarySurface JeanneDArc,MilitarySurface LaFayette\
		,MilitarySurface Suffren,MilitarySurface TypeE52,MilitarySurface TypeF70,MilitarySubmarines Agosta\
		,MilitarySubmarines Arethuse,MilitarySubmarines Daphne,MilitarySubmarines LeRedoutable\
		,MilitarySubmarines LeTriomphant,MilitarySubmarines Narval,MilitarySubmarines Rubis
	~	
	~	102 matériels militaires détaillés (union des 6 sous-plages).
		L'ORDRE est celui des lignes de la feuille "Stocks" du fichier \
		StocksMilitaryMods.xlsx.
	|


MilitaryFam:
	MilitaryArmouredFam,MilitaryHelicoptersFam,MilitaryAircraftsFam,MilitaryPatrolFam,MilitarySubmarinesFam,MilitarySurfaceFam ~~	|

CountriesNotFrance:
	World
	~	
	~	Pays autres que France (ici : agregat World), pour mettre le militaire a \
		zero
	|
```

## Stocks par matériel (lus depuis Excel)

on fait lire à vensim toutes les lignes du fichier de données, où chaque ligne est dans le même ordre d'appel que le subscript MilitaryUnits

```vensim
StockTUMilitaryUnit[France,MilitaryUnits]=
	GET DIRECT LOOKUPS('StocksMilitaryMods.xlsx', 'Stocks_smooth', '1', 'D2') ~~|
StockTUMilitaryUnit[CountriesNotFrance,MilitaryUnits](
	[(1900,0)-(2100,1)],(1900,0),(2100,0))
	~	TU
	~	Stock du parc militaire français par matériel, série annuelle 1900..2100.
		D1 = ligne des années ; D2 = première cellule de la première série.
		Vensim itère sur MilitaryUnits en descendant les lignes.
	|

PreviousStockTUMilitaryUnit[Country,MilitaryUnits]= INTEG (
	StockTUMilitaryUnit[Country,MilitaryUnits](Time)
	- PreviousStockTUMilitaryUnit[Country,MilitaryUnits],
		0)
	~	TU
	~	Retard d'un pas de temps sur le stock (analogue à PreviousStockTUinUse).
	|
```

les variables New et EoL sont calculées à partir des séries de données directement (pas d'hypothèse de durée de vie nécessaire)

```vensim
"FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryUnits]=
	MAX(0, StockTUMilitaryUnit[Country,MilitaryUnits](Time)
	- PreviousStockTUMilitaryUnit[Country,MilitaryUnits])
	~	TU/Year
	~	Livraisons observées = hausses du stock d'une année sur l'autre.
	|

"FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryUnits]=
	MAX(0, PreviousStockTUMilitaryUnit[Country,MilitaryUnits]
	- StockTUMilitaryUnit[Country,MilitaryUnits](Time))
	~	TU/Year
	~	Décommissionnements observés = baisses du stock d'une année sur l'autre.
	|

"FlowEoLTUMilitaryFam(/yr)"[Country,MilitaryAircraftsFam]=
	SUM("FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryAircraftsUnit!]) ~~|
"FlowEoLTUMilitaryFam(/yr)"[Country,MilitarySurfaceFam]=
	SUM("FlowEoLTUMilitaryUnit(/yr)"[Country,MilitarySurfaceUnit!]) ~~|
"FlowEoLTUMilitaryFam(/yr)"[Country,MilitaryPatrolFam]=
	SUM("FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryPatrolUnit!]) ~~|
"FlowEoLTUMilitaryFam(/yr)"[Country,MilitarySubmarinesFam]=
	SUM("FlowEoLTUMilitaryUnit(/yr)"[Country,MilitarySubmarinesUnit!]) ~~|
"FlowEoLTUMilitaryFam(/yr)"[Country,MilitaryArmouredFam]=
	SUM("FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryArmouredUnit!]) ~~|
"FlowEoLTUMilitaryFam(/yr)"[Country,MilitaryHelicoptersFam]=
	SUM("FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryHelicoptersUnit!])
	~	1/Year
	~	agrégateur famille ; SUM(MAX(0,-dS_u)) et non MAX(0,-d(SUM S_u)).
	|
```

## Masse unitaire (tonnes par matériel, lue depuis Excel)

```vensim
"MassUnitMilitaryUnit(t/TU)"[MilitaryUnits]=
	GET DIRECT CONSTANTS('RawMaterialsV2.xlsx', 'Military_MassUnit', 'D2*')
	~	t/TU
	~	Masse unitaire de chaque matériel militaire (colonne D de l'onglet
		Military_MassUnit, à partir de la ligne 2).
	|
```

## Chaîne matière

on importe les intensités en pourcentages, qu'on a déclarées par famille

```vensim
"RMIntensityFamilyPct(pct)"[RMAll,MilitaryArmouredFam]=
	GET DIRECT LOOKUPS('RawMaterialsV2.xlsx', 'Military_Armoured', '1', 'B2') ~~|
"RMIntensityFamilyPct(pct)"[RMAll,MilitaryAircraftsFam]=
	GET DIRECT LOOKUPS('RawMaterialsV2.xlsx', 'Military_Aircrafts', '1', 'B2') ~~|
"RMIntensityFamilyPct(pct)"[RMAll,MilitaryHelicoptersFam]=
	GET DIRECT LOOKUPS('RawMaterialsV2.xlsx', 'Military_Helicopters', '1', 'B2') ~~|
"RMIntensityFamilyPct(pct)"[RMAll,MilitaryPatrolFam]=
	GET DIRECT LOOKUPS('RawMaterialsV2.xlsx', 'Military_Patrol', '1', 'B2') ~~|
"RMIntensityFamilyPct(pct)"[RMAll,MilitarySurfaceFam]=
	GET DIRECT LOOKUPS('RawMaterialsV2.xlsx', 'Military_PrincipSurface', '1', 'B2') ~~|
"RMIntensityFamilyPct(pct)"[RMAll,MilitarySubmarinesFam]=
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
	"RMIntensityFamilyPct(pct)"[RMAll,MilitaryArmouredFam](Time) ~~|
"RMIntensityMilitaryUnitPct(pct)"[RMAll,MilitaryAircraftsUnit]=
	"RMIntensityFamilyPct(pct)"[RMAll,MilitaryAircraftsFam](Time) ~~|
"RMIntensityMilitaryUnitPct(pct)"[RMAll,MilitaryHelicoptersUnit]=
	"RMIntensityFamilyPct(pct)"[RMAll,MilitaryHelicoptersFam](Time) ~~|
"RMIntensityMilitaryUnitPct(pct)"[RMAll,MilitaryPatrolUnit]=
	"RMIntensityFamilyPct(pct)"[RMAll,MilitaryPatrolFam](Time) ~~|
"RMIntensityMilitaryUnitPct(pct)"[RMAll,MilitarySurfaceUnit]=
	"RMIntensityFamilyPct(pct)"[RMAll,MilitarySurfaceFam](Time) ~~|
"RMIntensityMilitaryUnitPct(pct)"[RMAll,MilitarySubmarinesUnit]=
	"RMIntensityFamilyPct(pct)"[RMAll,MilitarySubmarinesFam](Time)
	~	percent
	~	Intensité par matériel. La sous-plage ...Unit n'apparaît qu'à gauche :
		Vensim diffuse la valeur famille sur tous ses membres. C'est légal parce
		que le membre droit ne comporte plus qu'une plage variable (RMAll).
		Noter l'asymétrie de nommage : sous-plage MilitarySurfaceUnit (ligne
		276) mais élément de famille MilitarySurfaceFam (ligne 204).
	|
```

conversion pourcentages > t/TU

```vensim
"RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryUnits]=
	"RMIntensityMilitaryUnitPct(pct)"[RMAll,MilitaryUnits]
	* "MassUnitMilitaryUnit(t/TU)"[MilitaryUnits] / 100
	~	t/TU
	~	IMPORTANT : "RMIntensityMilitaryUnitPct(pct)" n'est plus un Lookup mais
		une auxiliaire DÉJÀ évaluée au temps courant en 4b. L'appel (Time) qui
		figurait ici doit disparaître, sinon Vensim tente d'appeler une
		auxiliaire comme une fonction tabulée.
	|
```

Limite connue et assumée de la version ci-dessus : les flux de fin de vie sont évalués à la composition CONTEMPORAINE DU DÉMANTÈLEMENT, et non à celle de la construction.
calcul de toutes les variantes de FlowRM natives du modèle

```vensim
CollectingRateMilitaryUnits[MilitaryArmouredUnit]=
	0.7 ~~|
CollectingRateMilitaryUnits[MilitaryAircraftsUnit]=
	0.7 ~~|
CollectingRateMilitaryUnits[MilitaryHelicoptersUnit]=
	0.7 ~~|
CollectingRateMilitaryUnits[MilitaryPatrolUnit]=
	0.7 ~~|
CollectingRateMilitaryUnits[MilitarySurfaceUnit]=
	0.7 ~~|
CollectingRateMilitaryUnits[MilitarySubmarinesUnit]=
	0.7
	~	Dmnl
	~	Taux de collecte au démantèlement. Uniforme par famille pour l'instant,
		pourra être affiné par matériel.
	|

"FlowRMTUPrimaryMilitaryUnit(t/yr)"[Country,RMAll,MilitaryUnits]=
	"FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryUnits]
	* "RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryUnits]
	* (1 - RecyclingRateRM[RMAll](Time) * CorrecRR)
	~	t/Year
	~	Matière primaire pour les nouvelles livraisons de chaque matériel,
		net du recyclage disponible.
	|

"FlowRMTURecyMilitaryUnit(t/yr)"[Country,RMAll,MilitaryUnits]=
	"FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryUnits]
	* "RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryUnits]
	* RecyclingRateRM[RMAll](Time) * CorrecRR
	~	t/Year
	~	Fraction recyclée injectée dans les nouvelles livraisons.
	|

"FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitaryUnits]=
	"FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryUnits]
	* "RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryUnits]
	* RecyclingRateRM[RMAll](Time) * CorrecRR
	* CollectingRateMilitaryUnits[MilitaryUnits]
	~	t/Year
	~	Matière récupérable en fin de vie (baisses de stock observées × intensité
		× taux de recyclage × taux de collecte).
	|

"FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitaryUnits]=
	"FlowEoLTUMilitaryUnit(/yr)"[Country,MilitaryUnits]
	* "RMIntensityMilitaryUnit(t/TU)"[RMAll,MilitaryUnits]
	* (1 - RecyclingRateRM[RMAll](Time) * CorrecRR
	* CollectingRateMilitaryUnits[MilitaryUnits])
	~	t/Year
	~	Matière perdue en fin de vie (non recyclée ou non collectée).
	|
```

agrégation à l'échelle du secteur militaire total

```vensim
FlowRMMilitaryPR[Country,RMAll]=
	SUM("FlowRMTUPrimaryMilitaryUnit(t/yr)"[Country,RMAll,MilitaryUnits!]) + SUM("FlowRMTURecyMilitaryUnit(t/yr)"\
		[Country,RMAll,MilitaryUnits!])
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
"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitaryAircraftsFam]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitaryAircraftsUnit!]) ~~|
"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitarySurfaceFam]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitarySurfaceUnit!]) ~~|
"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitaryPatrolFam]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitaryPatrolUnit!]) ~~|
"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitarySubmarinesFam]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitarySubmarinesUnit!]) ~~|
"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitaryArmouredFam]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitaryArmouredUnit!]) ~~|
"FlowRMEoLMilitaryFam(t/yr)"[Country,RMAll,MilitaryHelicoptersFam]=
	SUM("FlowRMTUEoLMilitaryUnit(t/yr)"[Country,RMAll,MilitaryHelicoptersUnit!])
	~	t/Year
	~	Tonnage recycle en fin de vie, par famille.
	|

"FlowRMLostMilitaryFam(t/yr)"[Country,RMAll,MilitaryAircraftsFam]=
	SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitaryAircraftsUnit!]) ~~|
"FlowRMLostMilitaryFam(t/yr)"[Country,RMAll,MilitarySurfaceFam]=
	SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitarySurfaceUnit!]) ~~|
"FlowRMLostMilitaryFam(t/yr)"[Country,RMAll,MilitaryPatrolFam]=
	SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitaryPatrolUnit!]) ~~|
"FlowRMLostMilitaryFam(t/yr)"[Country,RMAll,MilitarySubmarinesFam]=
	SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitarySubmarinesUnit!]) ~~|
"FlowRMLostMilitaryFam(t/yr)"[Country,RMAll,MilitaryArmouredFam]=
	SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitaryArmouredUnit!]) ~~|
"FlowRMLostMilitaryFam(t/yr)"[Country,RMAll,MilitaryHelicoptersFam]=
	SUM("FlowRMTULostMilitaryUnit(t/yr)"[Country,RMAll,MilitaryHelicoptersUnit!])
	~	t/Year
	~	Tonnage perdu en fin de vie, par famille.
	|
```

## Chaîne énergie

### ERM

```vensim
"ERMMilitaryUnitPrimary(PJ/yr)"[Country,RMAll,MilitaryUnits]=
	"FlowRMTUPrimaryMilitaryUnit(t/yr)"[Country,RMAll,MilitaryUnits]
	* "RMPrimary-Energy(GJ/t)"[RMAll] / 1e+06
	~	PJ/Year
	~	Énergie de production de la matière PRIMAIRE entrant dans les livraisons
		de chaque matériel militaire.
	|

"ERMMilitaryUnitRecy(PJ/yr)"[Country,RMAll,MilitaryUnits]=
	"FlowRMTURecyMilitaryUnit(t/yr)"[Country,RMAll,MilitaryUnits]
	* "RMRecycl-Energy(GJ/t)"[RMAll] / 1e+06
	~	PJ/Year
	~	Énergie de production de la matière RECYCLÉE entrant dans les livraisons
		de chaque matériel militaire.
	|

"ERMMilitaryUnit(PJ/yr)"[Country,RMAll,MilitaryUnits]=
	"ERMMilitaryUnitPrimary(PJ/yr)"[Country,RMAll,MilitaryUnits]
	+ "ERMMilitaryUnitRecy(PJ/yr)"[Country,RMAll,MilitaryUnits]
	~	PJ/Year
	~	Énergie matière totale par matériel et par matériau. Analogue militaire
		de ERMTU[Country,RMAll,TechoUnitextend].
	|

"ERMMilitary(PJ/yr)"[Country,RMAll]=
	SUM("ERMMilitaryUnit(PJ/yr)"[Country,RMAll,MilitaryUnits!])
	~	PJ/Year
	~	Énergie matière du secteur militaire par matériau. C'est la variable
		principale, homologue de ERMConstruction, ERMTransport, ERMAppliance.
	|

"ERMMilitaryPrimary(PJ/yr)"[Country,RMAll]=
	SUM("ERMMilitaryUnitPrimary(PJ/yr)"[Country,RMAll,MilitaryUnits!])
	~	
	~		|

"ERMMilitaryRecy(PJ/yr)"[Country,RMAll]=
	SUM("ERMMilitaryUnitRecy(PJ/yr)"[Country,RMAll,MilitaryUnits!])
	~	PJ/Year
	~	Décomposition primaire / recyclé, utile parce que RecyclingRateRM est un
		paramètre économie-wide mal adapté aux aciers de blindage et de coque et
		aux alliages aéronautiques certifiés : ces deux courbes montrent
		directement ce que coûte cette hypothèse.
	|


"ERMMilitaryByUnit(PJ/yr)"[Country,MilitaryUnits]=
	SUM("ERMMilitaryUnit(PJ/yr)"[Country,RMAll!,MilitaryUnits])
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
```

**Attention ici**

1. NE PAS ajouter "ERMMilitaryTot" à "ERMTOTAL(PJ/yr)". FlowRMMilitaryPR (ligne 181) est déjà branché en amont, donc la matière militaire est déjà comptée dans ENonFerrousMetals / ESteel&Fe / ENonMetallicMinerals. Ce bloc est une DÉCOMPOSITION de reporting, pas un terme supplémentaire. L'ajouter doublerait le militaire.

2. NE PAS additionner avec EmbEnergyMilitary. Ce sont deux périmètres différents : ERM est l'énergie de PRODUCTION DES MATIÈRES, EmbEnergy est l'énergie incorporée estimée à partir des prix (capital industriel, assemblage, R&D). Elles se recouvrent partiellement.

3. Périmètre. Comme pour la matière, il s'agit des 102 plateformes majeures et non du secteur militaire : ni munitions, ni armement léger, ni logistique non blindée, ni bases, ni spatial, ni MCO.

### EmbTU

définition des prix des équipements, qui est l'entrée de Dymemds pour l'énergie embarquée. Hypothèses importantes et discutables

```vensim
"EmbTUMilitary(MJ/TU)"[Country,MilitaryUnits]=
	"PriceRefMilitary(1998$/TU)"[MilitaryUnits,Country]*"aFPIndLOCAL(MJ/$)"[Country] ~~|


"EmbTUMilitary(MJ/yr)"[Country, MilitaryUnits]=
	"EmbTUMilitary(MJ/TU)"[Country, MilitaryUnits]*"FlowNewTUMilitaryUnit(/yr)"[Country, MilitaryUnits] ~~|


"EmbTUMilitary(PJ/yr)"[Country, MilitaryUnits]=
	"EmbTUMilitary(MJ/yr)"[Country, MilitaryUnits]/1e+09 ~~|
```

## Masses de référence des proxys civils

```vensim
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

## Prix par tonne des proxys civils

```vensim
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

## Masse moyenne par famille, non pondérée (variable secours)

Moyenne arithmetique des masses unitaires lues dans Military_MassUnit.
Les diviseurs sont les cardinaux des sous-plages : 10 blindes, 44 avions, 14 helicos, 13 patrouilleurs, 14 batiments de surface, 7 sous-marins.
A METTRE A JOUR si on ajoute un materiel a une sous-plage.

```vensim
"MassMeanMilitaryFam(t/TU)"[MilitaryArmouredFam]=
	SUM("MassUnitMilitaryUnit(t/TU)"[MilitaryArmouredUnit!])/10 ~~|
"MassMeanMilitaryFam(t/TU)"[MilitaryAircraftsFam]=
	SUM("MassUnitMilitaryUnit(t/TU)"[MilitaryAircraftsUnit!])/44 ~~|
"MassMeanMilitaryFam(t/TU)"[MilitaryHelicoptersFam]=
	SUM("MassUnitMilitaryUnit(t/TU)"[MilitaryHelicoptersUnit!])/14 ~~|
"MassMeanMilitaryFam(t/TU)"[MilitaryPatrolFam]=
	SUM("MassUnitMilitaryUnit(t/TU)"[MilitaryPatrolUnit!])/13 ~~|
"MassMeanMilitaryFam(t/TU)"[MilitarySurfaceFam]=
	SUM("MassUnitMilitaryUnit(t/TU)"[MilitarySurfaceUnit!])/14 ~~|
"MassMeanMilitaryFam(t/TU)"[MilitarySubmarinesFam]=
	SUM("MassUnitMilitaryUnit(t/TU)"[MilitarySubmarinesUnit!])/7
	~	t/TU
	~			Sert de valeur de repli les annees sans livraison.
	|
```

## Masse de référence par famille, pondérée par les livraisons

Rend le prix de famille coherent avec le mix annuel.
Prix_famille x Livraisons_famille = tonnage livre x prix/tonne donc la valeur totale est conservee sans passer par un prix par materiel.

```vensim
"MassRefMilitaryFam(t/TU)"[Country,MilitaryArmouredFam]=
	XIDZ(SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryArmouredUnit!]
	*"MassUnitMilitaryUnit(t/TU)"[MilitaryArmouredUnit!]),
	SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryArmouredUnit!]),
	"MassMeanMilitaryFam(t/TU)"[MilitaryArmouredFam]) ~~|
"MassRefMilitaryFam(t/TU)"[Country,MilitaryAircraftsFam]=
	XIDZ(SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryAircraftsUnit!]
	*"MassUnitMilitaryUnit(t/TU)"[MilitaryAircraftsUnit!]),
	SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryAircraftsUnit!]),
	"MassMeanMilitaryFam(t/TU)"[MilitaryAircraftsFam]) ~~|
"MassRefMilitaryFam(t/TU)"[Country,MilitaryHelicoptersFam]=
	XIDZ(SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryHelicoptersUnit!]
	*"MassUnitMilitaryUnit(t/TU)"[MilitaryHelicoptersUnit!]),
	SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryHelicoptersUnit!]),
	"MassMeanMilitaryFam(t/TU)"[MilitaryHelicoptersFam]) ~~|
"MassRefMilitaryFam(t/TU)"[Country,MilitaryPatrolFam]=
	XIDZ(SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryPatrolUnit!]
	*"MassUnitMilitaryUnit(t/TU)"[MilitaryPatrolUnit!]),
	SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitaryPatrolUnit!]),
	"MassMeanMilitaryFam(t/TU)"[MilitaryPatrolFam]) ~~|
"MassRefMilitaryFam(t/TU)"[Country,MilitarySurfaceFam]=
	XIDZ(SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitarySurfaceUnit!]
	*"MassUnitMilitaryUnit(t/TU)"[MilitarySurfaceUnit!]),
	SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitarySurfaceUnit!]),
	"MassMeanMilitaryFam(t/TU)"[MilitarySurfaceFam]) ~~|
"MassRefMilitaryFam(t/TU)"[Country,MilitarySubmarinesFam]=
	XIDZ(SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitarySubmarinesUnit!]
	*"MassUnitMilitaryUnit(t/TU)"[MilitarySubmarinesUnit!]),
	SUM("FlowNewTUMilitaryUnit(/yr)"[Country,MilitarySubmarinesUnit!]),
	"MassMeanMilitaryFam(t/TU)"[MilitarySubmarinesFam])
	~	t/TU
	~	Masse unitaire moyenne des materiels effectivement livres dans l'annee,
		par famille. Rappel de nommage : element de famille
		MilitarySurfaceFam, sous-plage MilitarySurfaceUnit.
	|
```

## Coef civil → militaire

```vensim
PremiumMilitary[MilitaryAircraftsFam]=
	1.5 ~~|
PremiumMilitary[MilitaryHelicoptersFam]=
	1.2 ~~|
PremiumMilitary[MilitaryArmouredFam]=
	3.5 ~~|
PremiumMilitary[MilitaryPatrolFam]=
	1.5 ~~|
PremiumMilitary[MilitarySurfaceFam]=
	2 ~~|
PremiumMilitary[MilitarySubmarinesFam]=
	12
	~	
	~	 helicoptere proche du civil, la
		cellule est souvent la meme, blinde piece massive mais peu de matiere, SNA/SNLE
		coque epaisse + chaufferie nucleaire.
	|
```

## Prix de référence militaire (règle de trois)

```vensim
"PriceRefMilitary(1998$/TU)"[MilitaryAircraftsUnit,Country]=
	"PriceRefCivilAircraft(1998$/t)"[Country]
	*"MassRefMilitaryFam(t/TU)"[Country,MilitaryAircraftsFam]
	*PremiumMilitary[MilitaryAircraftsFam] ~~|
"PriceRefMilitary(1998$/TU)"[MilitaryHelicoptersUnit,Country]=
	"PriceRefCivilAircraft(1998$/t)"[Country]
	*"MassRefMilitaryFam(t/TU)"[Country,MilitaryHelicoptersFam]
	*PremiumMilitary[MilitaryHelicoptersFam] ~~|
"PriceRefMilitary(1998$/TU)"[MilitaryArmouredUnit,Country]=
	"PriceRefCivilHV(1998$/t)"[Country]
	*"MassRefMilitaryFam(t/TU)"[Country,MilitaryArmouredFam]
	*PremiumMilitary[MilitaryArmouredFam] ~~|
"PriceRefMilitary(1998$/TU)"[MilitaryPatrolUnit,Country]=
	"PriceRefCivilVessel(1998$/t)"[Country]
	*"MassRefMilitaryFam(t/TU)"[Country,MilitaryPatrolFam]
	*PremiumMilitary[MilitaryPatrolFam] ~~|
"PriceRefMilitary(1998$/TU)"[MilitarySurfaceUnit,Country]=
	"PriceRefCivilVessel(1998$/t)"[Country]
	*"MassRefMilitaryFam(t/TU)"[Country,MilitarySurfaceFam]
	*PremiumMilitary[MilitarySurfaceFam] ~~|
"PriceRefMilitary(1998$/TU)"[MilitarySubmarinesUnit,Country]=
	"PriceRefCivilVessel(1998$/t)"[Country]
	*"MassRefMilitaryFam(t/TU)"[Country,MilitarySubmarinesFam]
	*PremiumMilitary[MilitarySubmarinesFam] ~
	~ Prix unitaire de reference par famille, obtenu par regle de trois sur le
		prix a la tonne du proxy civil. Difference de fond avec la version
		precedente : le coefficient n'est plus fige sur le mix de livraisons
		1985-2005, il suit le mix reel annee par annee via
		"MassRefMilitaryFam(t/TU)". Seule la prime au kilo reste exogene.
	|
	
```
